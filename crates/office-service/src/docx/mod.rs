//! DOCX sessions over `docx_edit::EngineSession`, as office-checkpoint.ts
//! drives the docx-edit WASM session, with the save projection
//! (yrsToDocument.ts) ported in `project`.

mod comments;
mod format;
mod item;
mod paragraph;
mod project;
pub(crate) mod read;
mod sdt;
mod styles;

use std::collections::HashMap;

use docx_edit::{EditCtx, EngineSession, StoryRange, UndoSession};

use crate::common::{asset_from_data_url, changed_span, check_replacement};
use crate::env;
use crate::error::{EditCode, Error, Result};
use crate::jsv::{V, canonical, text};
use crate::obj;
use crate::session::{Applied, Item};
use crate::types::{
    Asset, Command, Determinism, EffectKind, Entry, Format, ObjectKind, ObjectRef, Target,
};

pub(crate) use comments::comment_ooxml_ids;
pub(crate) use item::clear_unit_breaks;
use project::{Hooks, SaveContext, project_document};

pub(crate) struct DocxSession {
    engine: EngineSession,
    /// The TS facade tracks Undo before an edit, which keeps deleted items
    /// from garbage collection; the state must keep them too.
    undo: UndoSession,
    base: Vec<u8>,
    /// The materialized source package (`materializeDocx`), read on first use.
    document: Option<V>,
    projected: Option<Projected>,
}

struct Projected {
    document: V,
    stories: Vec<String>,
    /// `onEmbed` content by `${story}:${offset}`.
    embeds: HashMap<String, V>,
}

/// One paragraph's replaceable text (`docxParagraphs`).
#[derive(Clone)]
struct Paragraph {
    para_id: String,
    /// Story offset of the replaceable text, after the breaks it opens with.
    start: usize,
    length: usize,
    text: String,
    plain: bool,
}

fn open_engine(
    base: &[u8],
    seed: bool,
) -> Result<(EngineSession, Option<docx_parse::S9WireEnvelope>)> {
    let engine = EngineSession::new(env::next_client());
    let envelope = docx_edit::parse_docx_for_edit(base).map_err(Error::Engine)?;
    engine.set_media(docx_edit::package_media(&envelope));
    if seed {
        docx_edit::seed_parsed_docx(engine.doc(), envelope).map_err(Error::Engine)?;
        return Ok((engine, None));
    }
    Ok((engine, Some(envelope)))
}

/// `seedOffice` for DOCX.
pub(crate) fn seed(base: &[u8]) -> Result<Vec<u8>> {
    let (engine, _) = open_engine(base, true)?;
    Ok(engine.doc().encode_state_as_update_v1())
}

/// The materialized source as decodeS9Envelope gives the TS its `Document`.
fn decode_document(envelope: &docx_parse::S9WireEnvelope) -> Result<V> {
    let json = serde_json::to_string(envelope).map_err(Error::engine)?;
    let wire = V::parse(&json).map_err(Error::Engine)?;
    let document = wire.get("document");
    let package = document.get("package");
    let body = package.get("document");
    let content = body.get("content");
    let decoded_body = obj! { "content": content.clone() };
    let body_object = decoded_body.obj().cloned().unwrap_or_default();
    if let Some(sections) = body.get("sections").arr() {
        let blocks = content.items();
        body_object.set(
            "sections",
            V::array(
                sections
                    .items()
                    .iter()
                    .map(|section| {
                        let decoded = crate::jsv::Obj::new();
                        if !section.get("id").is_undef() {
                            decoded.set("id", section.get("id"));
                        }
                        decoded.set("properties", section.get("properties"));
                        let start = section.get("contentStart").to_number() as usize;
                        let end = section.get("contentEnd").to_number() as usize;
                        decoded.set(
                            "content",
                            V::array(
                                blocks[start.min(blocks.len())..end.min(blocks.len())].to_vec(),
                            ),
                        );
                        V::Obj(decoded)
                    })
                    .collect(),
            ),
        );
    }
    for key in ["finalSectionProperties", "customRootBindings", "comments"] {
        if !body.get(key).is_undef() {
            body_object.set(key, body.get(key));
        }
    }
    let decoded = crate::jsv::Obj::new();
    decoded.set("document", decoded_body);
    for (key, from) in [
        ("theme", "theme"),
        ("numbering", "numbering"),
        ("settings", "settings"),
        ("fontTable", "fontTable"),
        ("relationships", "relationshipEntries"),
        ("media", "mediaEntries"),
        ("charts", "chartEntries"),
    ] {
        decoded.set(key, package.get(from));
    }
    for (key, from) in [
        ("styles", "styles"),
        ("headers", "headerEntries"),
        ("footers", "footerEntries"),
        ("footnotes", "footnotes"),
        ("endnotes", "endnotes"),
        ("footnoteSeparators", "footnoteSeparators"),
        ("endnoteSeparators", "endnoteSeparators"),
    ] {
        if !package.get(from).is_undef() {
            decoded.set(key, package.get(from));
        }
    }
    let result = obj! { "package": V::Obj(decoded) };
    if !document.get("warnings").is_undef()
        && let Some(object) = result.obj()
    {
        object.set("warnings", document.get("warnings"));
    }
    Ok(result)
}

/// `authoredText`: the text a projected node shows.
fn authored_text(value: &V) -> String {
    if let Some(items) = value.arr() {
        return items.items().iter().map(authored_text).collect();
    }
    if value.obj().is_none() {
        return String::new();
    }
    let string = |key: &str| {
        value
            .get(key)
            .as_str()
            .map(|text| text.to_string())
            .unwrap_or_default()
    };
    match &*item::ty(value) {
        "text" => string("text"),
        "tab" => "\t".into(),
        "break" => "\n".into(),
        "softHyphen" => "\u{ad}".into(),
        "noBreakHyphen" => "\u{2011}".into(),
        "symbol" => string("char"),
        "paragraph" => authored_text(&value.get("content")) + "\n",
        "run" | "inlineSdt" | "blockSdt" | "simpleField" | "tableCell" => {
            authored_text(&value.get("content"))
        }
        "hyperlink" => authored_text(
            &value
                .get("structuredChildren")
                .or_else(|| value.get("children")),
        ),
        "complexField" => authored_text(&value.get("fieldResult")),
        "mathEquation" => string("plainText"),
        "shape" => {
            if value.get("shape").truthy() {
                authored_text(&value.get("shape"))
            } else {
                authored_text(
                    &value
                        .get("textBody")
                        .get("paragraphs")
                        .or_else(|| value.get("children")),
                )
            }
        }
        "table" => authored_text(&value.get("rows")),
        "tableRow" => {
            let cells = value.get("cells");
            let joined = match cells.arr() {
                Some(cells) => cells
                    .items()
                    .iter()
                    .map(authored_text)
                    .collect::<Vec<_>>()
                    .join("\t"),
                None => "undefined".into(),
            };
            joined + "\n"
        }
        _ => String::new(),
    }
}

impl DocxSession {
    pub fn open(base: &[u8], state: Option<&[u8]>) -> Result<Self> {
        let (engine, envelope) = open_engine(base, state.is_none())?;
        if let Some(state) = state {
            engine.doc().load_state_v1(state).map_err(Error::engine)?;
        }
        let envelope = match envelope {
            Some(envelope) => envelope,
            None => docx_edit::parse_docx_for_edit(base).map_err(Error::Engine)?,
        };
        Ok(Self {
            document: Some(decode_document(&envelope)?),
            engine,
            undo: UndoSession::new(),
            base: base.to_vec(),
            projected: None,
        })
    }

    pub fn state(&self) -> Vec<u8> {
        self.engine.doc().encode_state_as_update_v1()
    }

    fn base_document(&self) -> V {
        self.document.clone().unwrap_or_default()
    }

    fn project(&mut self) -> Result<&Projected> {
        if self.projected.is_none() {
            let base = self.base_document();
            let hooks = Hooks {
                on_story: true,
                embeds: Some(Vec::new()),
                ..Hooks::default()
            };
            let mut context = SaveContext::new(self.engine.doc(), &base, hooks)?;
            let document = project_document(&mut context, &base);
            clear_unit_breaks();
            let document = document?;
            let hooks = context.into_hooks();
            let mut embeds = HashMap::new();
            for (story, offset, content) in hooks.embeds.unwrap_or_default() {
                embeds.insert(
                    format!("{story}:{}", crate::js::number_to_string(offset)),
                    content,
                );
            }
            self.projected = Some(Projected {
                document,
                stories: hooks.stories,
                embeds,
            });
        }
        Ok(self.projected.as_ref().expect("projected above"))
    }

    pub fn entries(&mut self) -> Result<Vec<Item>> {
        let media = media_data_urls(&self.base_document());
        self.project()?;
        let projected = self.projected.as_ref().expect("projected above");
        let doc = self.engine.doc();
        let mut entries = Vec::new();
        for story in &projected.stories {
            let objects = read::story_object_ids(doc, story)?;
            let mut object_index = 0;
            let mut paragraph_index = 0;
            let mut paragraph_text = String::new();
            let mut formats: Vec<V> = Vec::new();
            let mut offset = 0.0;
            for segment in read::story_segments(doc, story)? {
                let attributes = segment.get("attributes");
                if item::is_kind(&segment, "text") {
                    let source = segment.get("text").to_js_string();
                    if !attributes.get("del").truthy() {
                        if attributes
                            .obj()
                            .is_some_and(|attributes| !attributes.is_empty())
                        {
                            formats.push(obj! {
                                "offset": text::len(&paragraph_text) as f64,
                                "length": text::len(&source) as f64,
                                "attributes": attributes.clone(),
                            });
                        }
                        paragraph_text.push_str(&source);
                    }
                    offset += text::len(&source) as f64;
                    continue;
                }
                let stable_id = objects
                    .get(object_index)
                    .filter(|id| !id.is_empty())
                    .cloned()
                    .ok_or_else(|| Error::engine("Missing DOCX object identity"))?;
                object_index += 1;
                let at = format!("{story}:{}", crate::js::number_to_string(offset));
                if item::is_kind(&segment, "pilcrow") {
                    let id = format!("{story}:paragraph:{}", segment.get("paraId").to_js_string());
                    entries.push(Item::text(
                        &id,
                        format!("{story}, paragraph {}", paragraph_index + 1),
                        std::mem::take(&mut paragraph_text),
                        format!("{story}:{paragraph_index}"),
                    ));
                    entries.push(Item::visual_v(
                        format!("{id}:format"),
                        format!("{story}, paragraph {} formatting", paragraph_index + 1),
                        &obj! {
                            "properties": segment.get("properties"),
                            "runs": V::array(std::mem::take(&mut formats)),
                        },
                        "",
                    ));
                    paragraph_index += 1;
                } else {
                    let embed_kind = segment.get("embedKind").to_js_string();
                    let payload = segment.get("payload");
                    if embed_kind == "image" && !attributes.get("del").truthy() {
                        let id = format!("{story}:image:{stable_id}");
                        let asset = docx_image_asset(&payload.get("src"), &media)?;
                        entries.push(Item {
                            id: id.clone(),
                            kind: EffectKind::Image,
                            label: format!("{story}, image"),
                            value: asset
                                .as_ref()
                                .map(|asset| asset.sha256.clone())
                                .unwrap_or_else(|| canonical(&payload)),
                            position: at.clone(),
                            asset_ref: Some(ObjectRef {
                                format: Format::Docx,
                                kind: ObjectKind::Image,
                                id: stable_id.clone(),
                                story_id: Some(story.clone()),
                                sheet_id: None,
                                slide_id: None,
                            }),
                            asset,
                        });
                        let geometry = payload
                            .obj()
                            .map(crate::jsv::Obj::spread)
                            .unwrap_or_default();
                        geometry.delete("src");
                        entries.push(Item::visual_v(
                            format!("{id}:format"),
                            format!("{story}, image formatting"),
                            &V::Obj(geometry),
                            "",
                        ));
                    } else if payload.get("modelKind").as_str().as_deref()
                        == Some("commentReference")
                    {
                        // A comment's reference mark: its comment entry carries the change.
                    } else {
                        let id = format!("{story}:object:{stable_id}");
                        let readable =
                            authored_text(&projected.embeds.get(&at).cloned().unwrap_or_default());
                        let textual = embed_kind != "table"
                            && (!readable.is_empty()
                                || ["sdt", "blockSdt", "field", "math", "shape"]
                                    .contains(&embed_kind.as_str()));
                        if textual && !attributes.get("del").truthy() {
                            entries.push(Item::text(
                                &id,
                                format!("{story}, {embed_kind}"),
                                readable,
                                at.clone(),
                            ));
                        }
                        entries.push(Item::visual_v(
                            if textual {
                                format!("{id}:format")
                            } else {
                                id.clone()
                            },
                            format!("{story}, {embed_kind} formatting"),
                            &payload,
                            at.clone(),
                        ));
                    }
                }
                offset += 1.0;
            }
        }
        for comment in read::list_comments(doc)? {
            let id = comment.get("id").to_js_string();
            let parent = comment.get("parentId");
            entries.push(Item::text(
                format!("comment:{id}"),
                if parent.truthy() {
                    "Comment reply"
                } else {
                    "Comment"
                },
                canonical(&obj! { "author": comment.get("author"), "body": comment.get("body") }),
                parent.clone().or_else(|| V::str("")).to_js_string(),
            ));
            entries.push(Item::visual_v(
                format!("comment:{id}:status"),
                "Comment status",
                &obj! { "done": comment.get("done"), "parentId": parent },
                "",
            ));
        }
        Ok(entries)
    }

    fn paragraphs(&self, story: &str) -> Result<Vec<Paragraph>> {
        let segments = read::story_segments(self.engine.doc(), story)
            .map_err(|_| Error::edit(EditCode::UnavailableTarget, "unknown DOCX story"))?;
        let mut paragraphs = Vec::new();
        let (mut start, mut cursor) = (0, 0);
        let mut paragraph_text = String::new();
        let mut plain = true;
        for segment in segments {
            if item::is_kind(&segment, "text") {
                let attributes = segment.get("attributes");
                if attributes.get("del").truthy() || attributes.get("ins").truthy() {
                    plain = false;
                }
                let source = segment.get("text").to_js_string();
                cursor += text::len(&source);
                paragraph_text.push_str(&source);
                continue;
            }
            if item::is_kind(&segment, "pilcrow") {
                paragraphs.push(Paragraph {
                    para_id: segment.get("paraId").to_js_string(),
                    start,
                    length: cursor - start,
                    text: std::mem::take(&mut paragraph_text),
                    plain,
                });
                cursor += 1;
                start = cursor;
                plain = true;
                continue;
            }
            // A page or column break the paragraph opens with stays in place.
            let kind = segment.get("embedKind").as_str();
            if cursor == start && matches!(kind.as_deref(), Some("pageBreak" | "columnBreak")) {
                start += 1;
            } else {
                plain = false;
            }
            cursor += 1;
        }
        Ok(paragraphs)
    }

    fn paragraph(&self, id: &str) -> Result<(String, Paragraph)> {
        let marker = id
            .rfind(":paragraph:")
            .filter(|marker| *marker > 0)
            .ok_or_else(|| {
                Error::edit(
                    EditCode::UnavailableTarget,
                    "target_id is not a DOCX paragraph",
                )
            })?;
        let story = &id[..marker];
        let para_id = &id[marker + ":paragraph:".len()..];
        let paragraph = self
            .paragraphs(story)?
            .into_iter()
            .find(|paragraph| paragraph.para_id == para_id)
            .ok_or_else(|| {
                Error::edit(
                    EditCode::UnavailableTarget,
                    "the paragraph is no longer in the document",
                )
            })?;
        Ok((story.to_owned(), paragraph))
    }

    pub fn editable(&mut self) -> Result<Vec<Entry>> {
        let stories = self.project()?.stories.clone();
        let mut entries = Vec::new();
        for story in stories {
            for (index, paragraph) in self.paragraphs(&story)?.into_iter().enumerate() {
                if paragraph.plain {
                    entries.push(Entry {
                        id: format!("{story}:paragraph:{}", paragraph.para_id),
                        label: format!("{story}, paragraph {}", index + 1),
                        value: paragraph.text,
                        position: format!("{story}:{index}"),
                    });
                }
            }
        }
        Ok(entries)
    }

    pub fn apply(&mut self, command: &Command) -> Result<Applied> {
        let Command::ReplaceText { target_id, .. } = command else {
            return Err(Error::edit(
                EditCode::UnsupportedOperation,
                "DOCX sources support replace_text only",
            ));
        };
        let (story, paragraph) = self.paragraph(target_id)?;
        if !paragraph.plain {
            return Err(Error::edit(
                EditCode::UnavailableTarget,
                "the paragraph holds objects or tracked changes",
            ));
        }
        let text = check_replacement(&paragraph.text, command, "DOCX")?.to_owned();
        let span = changed_span(&paragraph.text, &text);
        self.undo.track(self.engine.doc());
        self.undo.select_story(&story);
        if span.start < span.end || !span.text.is_empty() {
            let range = StoryRange::new(
                story.as_str(),
                (paragraph.start + span.start) as u32,
                (paragraph.start + span.end) as u32,
            );
            self.engine
                .doc()
                .replace_range(
                    &EditCtx::local(String::new(), String::new()),
                    range,
                    &span.text,
                )
                .map_err(Error::engine)?;
        }
        self.projected = None;
        Ok(Applied {
            id: target_id.clone(),
            inverse: Command::ReplaceText {
                target_id: target_id.clone(),
                expected_text: text,
                text: paragraph.text,
            },
        })
    }

    pub fn locate(&mut self, id: &str) -> Result<Target> {
        let (story, paragraph) = self.paragraph(id)?;
        Ok(Target {
            id: id.to_owned(),
            path: vec!["stories".into(), story],
            range: Some([
                paragraph.start as u32,
                (paragraph.start + paragraph.length) as u32,
            ]),
        })
    }

    pub fn export(&mut self, determinism: Determinism) -> Result<Vec<u8>> {
        let document = self.project()?.document.clone();
        save(&document, &self.base, determinism)
    }

    /// The projected document with `onParagraph` (office-checkpoint.ts `docxRawInlines`).
    pub(crate) fn raw_inlines(&mut self) -> Result<HashMap<String, String>> {
        let base = self.base_document();
        let hooks = Hooks {
            paragraphs: Some(Vec::new()),
            ..Hooks::default()
        };
        let mut context = SaveContext::new(self.engine.doc(), &base, hooks)?;
        let projected = project_document(&mut context, &base);
        clear_unit_breaks();
        projected?;
        let mut paragraphs = HashMap::new();
        for (story, _, paragraph, id) in context.into_hooks().paragraphs.unwrap_or_default() {
            let key = format!("{story}\u{0}{id}");
            // A repeated id cannot name one source paragraph.
            let value = if paragraphs.contains_key(&key) {
                "repeated".to_owned()
            } else {
                crate::jsv::stringify(&V::array(
                    paragraph
                        .get("content")
                        .items()
                        .into_iter()
                        .filter(|child| item::is(child, "rawXml"))
                        .collect(),
                ))
                .unwrap_or_default()
            };
            paragraphs.insert(key, value);
        }
        Ok(paragraphs)
    }

    pub(crate) fn engine(&self) -> &EngineSession {
        &self.engine
    }
}

/// The source package's media by alias: its display data URL.
fn media_data_urls(document: &V) -> HashMap<String, String> {
    styles::map_entries(&document.get("package").get("media"))
        .into_iter()
        .filter_map(|(alias, file)| Some((alias, file.get("dataUrl").as_str()?.to_string())))
        .collect()
}

/// A seeded source image names its package part (`media:<part>`); an inserted one carries its bytes.
fn docx_image_asset(src: &V, media: &HashMap<String, String>) -> Result<Option<Asset>> {
    let Some(src) = src.as_str() else {
        return Ok(None);
    };
    let Some(part) = src.strip_prefix("media:") else {
        return asset_from_data_url(&src).map(Some);
    };
    let data_url = media
        .get(part)
        .filter(|url| !url.is_empty())
        .ok_or_else(|| {
            Error::engine(format!(
                "DOCX image part {part} is absent from the source package"
            ))
        })?;
    asset_from_data_url(data_url).map(Some)
}

/// rustSaveFacade.ts `writeDocumentWithRust` with `updateModifiedDate: false`.
fn save(document: &V, original: &[u8], determinism: Determinism) -> Result<Vec<u8>> {
    let package = document.get("package");
    let entries = |key: &str| package.get(key).or_else(|| V::array(Vec::new()));
    let request = obj! {
        "determinism": obj! { "seed": determinism.seed, "now": determinism.now },
        "document": package.get("document"),
        "headerEntries": entries("headers"),
        "footerEntries": entries("footers"),
        "footnotes": entries("footnotes"),
        "endnotes": entries("endnotes"),
        "footnoteSeparators": entries("footnoteSeparators"),
        "endnoteSeparators": entries("endnoteSeparators"),
        "relationshipEntries": entries("relationships"),
    };
    if let Some(object) = request.obj() {
        if !package.get("numbering").is_undef() {
            object.set("numbering", package.get("numbering"));
        }
        object.set("options", obj! { "updateModifiedDate": false });
    }
    let json = serde_json::to_string(&request).map_err(Error::engine)?;
    let request: docx_parse::serializer::S13SaveRequest =
        serde_json::from_str(&json).map_err(Error::engine)?;
    docx_parse::serializer::write_docx_s13(request, original).map_err(Error::engine)
}
