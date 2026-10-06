//! PPTX sessions over `pptx_edit::DeckSession`, as office-checkpoint.ts
//! drives `PptxDocument`.

use std::collections::HashMap;

use pptx_edit::{DeckSession, EditCtx, TextStyle};

use crate::common::{changed_span, check_replacement, image_mime_type, len16, sha256_hex};
use crate::env;
use crate::error::{EditCode, Error, Result};
use crate::js::{J, Obj, utf16_len};
use crate::session::{Applied, Item};
use crate::types::{Asset, Command, EffectKind, Entry, Format, ObjectKind, ObjectRef, Target};

pub(crate) struct PptxSession {
    deck: DeckSession,
    /// Paragraph id to story id; a session serves one call, and replace_text
    /// never moves a paragraph to another story.
    stories: Option<HashMap<String, String>>,
}

struct Paragraph {
    start: u32,
    length: u32,
    text: String,
    runs: Vec<J>,
}

impl PptxSession {
    pub fn open(base: &[u8], state: Option<&[u8]>) -> Result<Self> {
        let deck = match state {
            Some(state) => {
                DeckSession::open_from_update_with_source(state, base, env::next_client())
            }
            None => DeckSession::open(base, env::next_client()),
        }
        .map_err(Error::engine)?;
        Ok(Self {
            deck,
            stories: None,
        })
    }

    pub fn state(&self) -> Vec<u8> {
        self.deck.encode_state_as_update_v1()
    }

    /// `JSON.parse(doc.snapshotJson())`.
    pub fn snapshot(&self) -> Result<J> {
        deck_json(&self.deck)
    }

    pub fn entries(&self) -> Result<Vec<Item>> {
        let deck = self.snapshot()?;
        let mut size = Obj::new();
        for (key, from) in [("width", "widthEmu"), ("height", "heightEmu")] {
            if let Some(value) = deck.get(from) {
                size.set(key, value.clone());
            }
        }
        let mut entries = vec![Item::visual(
            "deck:size",
            "Presentation size",
            &J::Obj(size),
            "",
        )];
        for (slide_index, slide) in array(deck.get("slides")).iter().enumerate() {
            let slide_id = text(slide.get("id"));
            let mut identity = Obj::new();
            for (key, from) in [("name", "name"), ("layout", "layoutPartPath")] {
                if let Some(value) = slide.get(from) {
                    identity.set(key, value.clone());
                }
            }
            entries.push(Item::visual(
                &slide_id,
                format!("Slide {}", slide_index + 1),
                &J::Obj(identity),
                slide_index.to_string(),
            ));
            self.visit(
                &mut entries,
                array(slide.get("shapes")),
                &slide_id,
                slide_index,
                &slide_id,
            )?;
        }
        Ok(entries)
    }

    fn visit(
        &self,
        entries: &mut Vec<Item>,
        shapes: &[J],
        parent: &str,
        slide_index: usize,
        slide_id: &str,
    ) -> Result<()> {
        for (shape_index, shape) in shapes.iter().enumerate() {
            let shape_id = text(shape.get("id"));
            let label = format!("Slide {}, {}", slide_index + 1, text(shape.get("name")));
            let mut geometry = shape.as_obj().cloned().unwrap_or_default();
            geometry.remove("children");
            geometry.remove("textStories");
            entries.push(Item::visual(
                &shape_id,
                &label,
                &J::Obj(geometry),
                format!("{parent}:{shape_index}"),
            ));
            for story in array(shape.get("textStories")) {
                let story_id = text(story.get("id"));
                for (index, paragraph) in array(story.get("paragraphs")).iter().enumerate() {
                    let paragraph_id = text(paragraph.get("id"));
                    let runs = array(paragraph.get("runs"));
                    entries.push(Item::text(
                        &paragraph_id,
                        &label,
                        runs_text(runs),
                        format!("{story_id}:{index}"),
                    ));
                    let mut format = paragraph.as_obj().cloned().unwrap_or_default();
                    format.set(
                        "runs",
                        J::Arr(
                            runs.iter()
                                .map(|run| run.get("style").cloned().unwrap_or(J::Null))
                                .collect(),
                        ),
                    );
                    entries.push(Item::visual(
                        format!("{paragraph_id}:format"),
                        format!("{label} text formatting"),
                        &J::Obj(format),
                        "",
                    ));
                }
            }
            if let Some(path) = shape.get("mediaPartPath").filter(|path| path.truthy()) {
                let path = crate::xlsx::js_string(path);
                let bytes = self.deck.media_bytes(&path).map_err(Error::engine)?;
                let mime_type = image_mime_type(&path)?;
                let asset = Asset {
                    sha256: sha256_hex(&bytes),
                    mime_type: mime_type.to_owned(),
                    bytes,
                };
                entries.push(Item {
                    id: format!("{shape_id}:image"),
                    kind: EffectKind::Image,
                    label: label.clone(),
                    value: asset.sha256.clone(),
                    position: format!("{slide_id}:{shape_index}"),
                    asset_ref: Some(ObjectRef {
                        format: Format::Pptx,
                        kind: ObjectKind::Image,
                        id: shape_id.clone(),
                        story_id: None,
                        sheet_id: None,
                        slide_id: Some(slide_id.to_owned()),
                    }),
                    asset: Some(asset),
                });
            }
            self.visit(
                entries,
                array(shape.get("children")),
                &shape_id,
                slide_index,
                slide_id,
            )?;
        }
        Ok(())
    }

    /// office-checkpoint.ts `pptxTextEntries`, with each entry's story.
    fn text_entries(&self) -> Result<Vec<(Entry, String)>> {
        let deck = self.snapshot()?;
        let mut entries = Vec::new();
        fn visit(entries: &mut Vec<(Entry, String)>, shapes: &[J], slide_index: usize) {
            for shape in shapes {
                for story in array(shape.get("textStories")) {
                    let story_id = text(story.get("id"));
                    for (index, paragraph) in array(story.get("paragraphs")).iter().enumerate() {
                        entries.push((
                            Entry {
                                id: text(paragraph.get("id")),
                                label: format!(
                                    "Slide {}, {}",
                                    slide_index + 1,
                                    text(shape.get("name"))
                                ),
                                value: runs_text(array(paragraph.get("runs"))),
                                position: format!("{story_id}:{index}"),
                            },
                            story_id.clone(),
                        ));
                    }
                }
                visit(entries, array(shape.get("children")), slide_index);
            }
        }
        for (slide_index, slide) in array(deck.get("slides")).iter().enumerate() {
            visit(&mut entries, array(slide.get("shapes")), slide_index);
        }
        Ok(entries)
    }

    pub fn editable(&self) -> Result<Vec<Entry>> {
        Ok(self
            .text_entries()?
            .into_iter()
            .map(|(entry, _)| entry)
            .collect())
    }

    fn story_paragraphs(&self, story_id: &str) -> Result<Vec<(String, Paragraph)>> {
        let story = self.deck.story(story_id).map_err(Error::engine)?;
        let story = J::parse(&serde_json::to_string(&story).map_err(Error::engine)?)
            .map_err(Error::Engine)?;
        let mut cursor = 0;
        Ok(array(story.get("paragraphs"))
            .iter()
            .map(|paragraph| {
                let runs = array(paragraph.get("runs")).to_vec();
                let text = runs_text(&runs);
                let length = len16(&text);
                let item = Paragraph {
                    start: cursor,
                    length,
                    text,
                    runs,
                };
                cursor += length + 1;
                (
                    crate::xlsx::js_string(paragraph.get("id").unwrap_or(&J::Null)),
                    item,
                )
            })
            .collect())
    }

    fn paragraph(&mut self, id: &str) -> Result<(String, Paragraph)> {
        if self.stories.is_none() {
            self.stories = Some(
                self.text_entries()?
                    .into_iter()
                    .map(|(entry, story)| (entry.id, story))
                    .collect(),
            );
        }
        let missing = || {
            Error::edit(
                EditCode::UnavailableTarget,
                "the paragraph is no longer in the deck",
            )
        };
        let story_id = self
            .stories
            .as_ref()
            .and_then(|stories| stories.get(id))
            .cloned()
            .ok_or_else(missing)?;
        let paragraph = self
            .story_paragraphs(&story_id)?
            .into_iter()
            .find(|(paragraph_id, _)| paragraph_id == id)
            .map(|(_, paragraph)| paragraph)
            .ok_or_else(missing)?;
        Ok((story_id, paragraph))
    }

    pub fn apply(&mut self, command: &Command) -> Result<Applied> {
        let Command::ReplaceText { target_id, .. } = command else {
            return Err(Error::edit(
                EditCode::UnsupportedOperation,
                "PPTX sources support replace_text only",
            ));
        };
        let (story_id, paragraph) = self.paragraph(target_id)?;
        let text = check_replacement(&paragraph.text, command, "PPTX")?.to_owned();
        let span = changed_span(&paragraph.text, &text);
        // A pure insertion replaces nothing; it takes the style of the character before it.
        let style = style_at(
            &paragraph.runs,
            if span.start < span.end {
                span.start
            } else {
                span.start.saturating_sub(1)
            },
        );
        let start = paragraph.start + span.start as u32;
        let ctx = EditCtx::local("wasm");
        if span.start < span.end {
            self.deck
                .delete_text(&ctx, &story_id, start, paragraph.start + span.end as u32)
                .map_err(Error::engine)?;
        }
        if !span.text.is_empty() {
            let style: TextStyle = match style {
                None => TextStyle::default(),
                Some(style) => serde_json::to_string(&style)
                    .and_then(|json| serde_json::from_str(&json))
                    .map_err(Error::engine)?,
            };
            self.deck
                .insert_text(&ctx, &story_id, start, &span.text, &style)
                .map_err(Error::engine)?;
        }
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
        let (story_id, paragraph) = self.paragraph(id)?;
        Ok(Target {
            id: id.to_owned(),
            path: vec!["pptx:stories".into(), story_id],
            range: Some([paragraph.start, paragraph.start + paragraph.length]),
        })
    }

    pub fn export(&self) -> Result<Vec<u8>> {
        self.deck.save().map_err(Error::engine)
    }
}

/// `JSON.parse(doc.snapshotJson())` of any deck session.
pub(crate) fn deck_json(deck: &DeckSession) -> Result<J> {
    let snapshot = deck.snapshot().map_err(Error::engine)?;
    J::parse(&serde_json::to_string(&snapshot).map_err(Error::engine)?).map_err(Error::Engine)
}

pub(crate) fn seed(base: &[u8]) -> Result<Vec<u8>> {
    Ok(PptxSession::open(base, None)?.state())
}

/// Style of the run holding the character at `offset`, else of the last run;
/// none stands for `undefined`.
fn style_at(runs: &[J], offset: usize) -> Option<J> {
    let mut end = 0;
    for run in runs {
        end += utf16_len(run.get("text").and_then(J::as_str).unwrap_or_default());
        if offset < end {
            return run.get("style").cloned();
        }
    }
    Some(
        runs.last()
            .and_then(|run| run.get("style").filter(|style| !style.is_null()).cloned())
            .unwrap_or_else(|| J::Obj(Obj::new())),
    )
}

/// `runs.map((run) => run.text).join("")`.
fn runs_text(runs: &[J]) -> String {
    runs.iter()
        .map(|run| match run.get("text") {
            None | Some(J::Null) => String::new(),
            Some(text) => crate::xlsx::js_string(text),
        })
        .collect()
}

fn array(value: Option<&J>) -> &[J] {
    value.and_then(J::as_arr).map(Vec::as_slice).unwrap_or(&[])
}

fn text(value: Option<&J>) -> String {
    value
        .map(crate::xlsx::js_string)
        .unwrap_or_else(|| "undefined".into())
}
