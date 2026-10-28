//! PPTX sessions over `pptx_edit::DeckSession`, as office-checkpoint.ts
//! drives `PptxDocument`.

mod agent;

use std::collections::HashMap;

use pptx_edit::{DeckSession, EditCtx, TextStyle};

use crate::common::{
    check_expected_paragraph, check_paragraph_text, check_replacement, image_mime_type,
    insert_anchor, len16, sha256_hex, storable_text,
};
use crate::env;
use crate::error::{EditCode, Error, Result};
use crate::js::{J, Obj, utf16_len};
use crate::session::{Applied, Item};
use crate::touched::{self, Kept, Touched};
use crate::types::{Asset, Command, EffectKind, Entry, Format, ObjectKind, ObjectRef, Target};

/// The root map of PPTX stories (`crate::document_roots`).
const STORIES: &str = "pptx:stories";

pub(crate) struct PptxSession {
    deck: DeckSession,
    /// Paragraph id to story id; a session serves one call, replace_text
    /// never moves a paragraph to another story, and a paragraph insert or
    /// delete clears it.
    stories: Option<HashMap<String, String>>,
    /// For each paragraph an insert went after, the id of the last paragraph
    /// this session put there, so the next lands after it.
    after: HashMap<String, String>,
    /// What the last command wrote that its Redo re-creates, read after it
    /// applied, for [`Self::created`].
    created: Option<Touched>,
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
            after: HashMap::new(),
            created: None,
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
                format!("deck:{slide_index}"),
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

    /// A paragraph id never changes story, so the cache only misses the
    /// paragraphs a command added (or finds one gone): a miss reads the
    /// deck again once (review L5).
    fn paragraph(&mut self, id: &str) -> Result<(String, Paragraph)> {
        let mut fresh = false;
        loop {
            if self.stories.is_none() {
                self.stories = Some(
                    self.text_entries()?
                        .into_iter()
                        .map(|(entry, story)| (entry.id, story))
                        .collect(),
                );
                fresh = true;
            }
            let found = self
                .stories
                .as_ref()
                .and_then(|stories| stories.get(id))
                .cloned()
                .and_then(|story_id| {
                    let paragraph = self
                        .story_paragraphs(&story_id)
                        .ok()?
                        .into_iter()
                        .find(|(paragraph_id, _)| paragraph_id == id)?
                        .1;
                    Some((story_id, paragraph))
                });
            match found {
                Some(found) => return Ok(found),
                None if fresh => return Err(self.paragraph_refusal(id)),
                None => self.stories = None,
            }
        }
    }

    /// [`crate::CommandResult::touched`]: what the command's Undo
    /// re-creates (the Editor's rule, 2026-10-22): the text a `replace_text`
    /// takes out, a deleted paragraph's text and mark, the text and format
    /// markers a `format_text` span covers, the alignment a `set_alignment`
    /// writes over; for the slide commands [`Self::touched_agent`]; nothing
    /// for the rest, nor where the target is not found.
    pub fn touched(&mut self, command: &Command) -> Touched {
        let (Command::ReplaceText { target_id, .. }
        | Command::DeleteParagraph { target_id, .. }
        | Command::FormatText { target_id, .. }
        | Command::SetAlignment { target_id, .. }) = command
        else {
            return self.touched_agent(command);
        };
        let Ok((story, paragraph)) = self.paragraph(target_id) else {
            return Touched::default();
        };
        let doc = self.deck.yrs_doc();
        let (start, end, kept) = match command {
            Command::ReplaceText { .. } => {
                match check_replacement(&paragraph.text, command, "PPTX") {
                    Ok((span, _)) => (span.start as u32, span.end as u32, Kept::Content),
                    Err(_) => return Touched::default(),
                }
            }
            Command::FormatText { text, .. } => {
                let Some(at) = paragraph.text.find(text.as_str()) else {
                    return Touched::default();
                };
                let at = len16(&paragraph.text[..at]);
                (at, at + len16(text), Kept::Formatting)
            }
            Command::SetAlignment { .. } => {
                return touched::mark_entries(
                    doc,
                    STORIES,
                    &story,
                    paragraph.start + paragraph.length,
                    &["alignment"],
                );
            }
            _ => (0, paragraph.length + 1, Kept::Content),
        };
        touched::story_range(
            doc,
            STORIES,
            &story,
            paragraph.start + start,
            paragraph.start + end,
            kept,
        )
    }

    pub fn apply(&mut self, command: &Command) -> Result<Applied> {
        let target_id = match command {
            Command::ReplaceText { target_id, .. } => target_id,
            Command::InsertParagraph {
                after_id,
                before_id,
                text,
                style,
                list,
            } => {
                if list.is_some() {
                    return Err(Error::edit(
                        EditCode::InvalidInput,
                        "list sets a DOCX paragraph's list; a new PPTX paragraph takes its neighbour's",
                    ));
                }
                let (anchor, after) = insert_anchor(after_id.as_deref(), before_id.as_deref())?;
                return self.insert_paragraph(anchor, after, text, style.as_deref());
            }
            Command::DeleteParagraph {
                target_id,
                expected_text,
            } => return self.delete_paragraph(target_id, expected_text),
            _ => return self.apply_agent(command),
        };
        let (story_id, paragraph) = self.paragraph(target_id)?;
        let (span, text) = check_replacement(&paragraph.text, command, "PPTX")?;
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
                .map_err(refused)?;
        }
        if !span.text.is_empty() {
            self.deck
                .insert_text(&ctx, &story_id, start, &span.text, &text_style(style)?)
                .map_err(refused)?;
        }
        Ok(Applied {
            ids: vec![target_id.clone()],
            inverse: Some(Command::ReplaceText {
                target_id: target_id.clone(),
                expected_text: text,
                text: storable_text(&paragraph.text),
            }),
        })
    }

    /// A paragraph right after or before `anchor` in its text, taking that
    /// paragraph's properties (level, bullet, alignment) as a split does and,
    /// for its text, the style of its character next to the new one. Several
    /// inserts after one paragraph land in call order. The inverse deletes
    /// it, unless it is its story's last paragraph, which delete_paragraph
    /// refuses.
    fn insert_paragraph(
        &mut self,
        anchor: &str,
        after: bool,
        text: &str,
        style: Option<&str>,
    ) -> Result<Applied> {
        if style.is_some() {
            return Err(Error::edit(
                EditCode::InvalidInput,
                "style names a DOCX paragraph style; a new PPTX paragraph takes its neighbour's formatting",
            ));
        }
        check_paragraph_text(text)?;
        let (story_id, _) = self.paragraph(anchor)?;
        let paragraphs = self.story_paragraphs(&story_id)?;
        let position = |id: &str| paragraphs.iter().position(|(other, _)| other == id);
        let ctx = EditCtx::local("wasm");
        // After a paragraph a new one is split off its end, after the
        // paragraph this call last put there; before one it copies that
        // paragraph's mark, which keeps its id.
        let (new_id, start, style, undoable) = if after {
            let index = match self.after.get(anchor).and_then(|id| position(id)) {
                Some(index) => index,
                None => position(anchor).ok_or_else(missing_paragraph)?,
            };
            let neighbour = &paragraphs[index].1;
            let end = neighbour.start + neighbour.length;
            self.deck
                .insert_paragraph_break(&ctx, &story_id, end)
                .map_err(refused)?;
            let (id, _) = self
                .story_paragraphs(&story_id)?
                .into_iter()
                .nth(index + 1)
                .ok_or_else(|| Error::engine("the split paragraph is missing"))?;
            let style = style_at(&neighbour.runs, neighbour.length as usize);
            (id, end + 1, style, index + 1 < paragraphs.len())
        } else {
            let paragraph = &paragraphs[position(anchor).ok_or_else(missing_paragraph)?].1;
            let id = self
                .deck
                .insert_paragraph_before(&ctx, &story_id, anchor)
                .map_err(refused)?;
            (id, paragraph.start, style_at(&paragraph.runs, 0), true)
        };
        if !text.is_empty() {
            self.deck
                .insert_text(&ctx, &story_id, start, text, &text_style(style)?)
                .map_err(refused)?;
        }
        self.stories = None;
        if after {
            self.after.insert(anchor.to_owned(), new_id.clone());
        }
        Ok(Applied {
            ids: vec![new_id.clone()],
            inverse: undoable.then(|| Command::DeleteParagraph {
                target_id: new_id,
                expected_text: text.to_owned(),
            }),
        })
    }

    /// Refused unless the paragraph reads `expected_text` and is not its
    /// story's last. No inverse: putting it back exactly is the live Undo's.
    fn delete_paragraph(&mut self, id: &str, expected_text: &str) -> Result<Applied> {
        let (story_id, paragraph) = self.paragraph(id)?;
        check_expected_paragraph(id, &paragraph.text, expected_text)?;
        let last = self.story_paragraphs(&story_id)?.pop().map(|(id, _)| id);
        if last.as_deref() == Some(id) {
            return Err(Error::edit(
                EditCode::InvalidInput,
                "a shape's last paragraph cannot be deleted; replace its text instead",
            ));
        }
        self.deck
            .remove_paragraph(&EditCtx::local("wasm"), &story_id, id)
            .map_err(refused)?;
        self.stories = None;
        Ok(Applied {
            ids: Vec::new(),
            inverse: None,
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

fn missing_paragraph() -> Error {
    Error::edit(
        EditCode::UnavailableTarget,
        "the paragraph is no longer in the deck",
    )
}

/// A run style read from the deck (none for `undefined`) as the engine takes it.
fn text_style(style: Option<J>) -> Result<TextStyle> {
    match style {
        None => Ok(TextStyle::default()),
        Some(style) => serde_json::to_string(&style)
            .and_then(|json| serde_json::from_str(&json))
            .map_err(Error::engine),
    }
}

/// The deck's refusal to edit an equation's text, or to place a layout
/// placeholder without a size, is an edit refusal with its sentence; any
/// other failure stays the engine's.
fn refused(error: pptx_edit::EditError) -> Error {
    match error {
        pptx_edit::EditError::EquationText(_) | pptx_edit::EditError::InvalidGeometry(_) => {
            Error::edit(EditCode::InvalidInput, error.to_string())
        }
        error => Error::engine(error),
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
