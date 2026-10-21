//! The chat agent's DOCX commands and scoped read (Epo 2026-10-20 to
//! 2026-10-22): insert and delete paragraphs and tables, lists, text
//! formatting and alignment in the body; and a page of the paragraphs, the
//! body's then the headers', footers' and notes' (which only `replace_text`
//! edits), each with its style when not the default, the first page with the
//! file's paragraph styles. A new paragraph continues its anchor as Enter
//! does; a named style takes the place of the anchor style's next style.

use std::collections::{BTreeMap, HashMap};
use std::sync::Arc;

use docx_edit::{
    ColorPatch, EditCtx, EditingDoc, FontFamilyPatch, FormatPolicy, InlineFormatDelta, OpError,
    ParaAttrDelta, ParaSelector, Patch, Position, SegmentContent, StoryRange, StrikePatch,
    TableLocator, UnderlinePatch,
};
use docx_parse::NumberingDefinitions;
use docx_parse::numbering::list_button_num_ids;
use serde_json::{Value, json};
use yrs::Any;

use super::DocxSession;
use super::styles::{Numbering, StyleResolver, map_entries};
use crate::common::{check_expected_paragraph, check_paragraph_text, check_replacement, len16};
use crate::error::{EditCode, Error, Result};
use crate::jsv::V;
use crate::session::Applied;
use crate::touched::{self, Kept, Touched};
use crate::types::{Alignment, Command, InspectQuery, ListKind, TextFormat};

const BODY: &str = "body";
/// Paragraph-mark keys a new paragraph keeps from the one it was split off:
/// its identity and the section it ends.
const KEPT_KEYS: [&str; 4] = ["paraId", "_kind", "sectPr", "sectionBreakType"];
const SECTION_KEYS: [&str; 2] = ["sectPr", "sectionBreakType"];
/// A mark's tracked insertion, deletion or property change.
const TRACKED_KEYS: [&str; 3] = ["pPrIns", "pPrDel", "pPrChange"];
/// Paragraphs a read lists and a `set_list` names.
const MAX_PARAGRAPHS: u32 = 200;
/// Styles the first page lists.
const MAX_STYLES: usize = 50;
/// Characters a new table cell holds.
const MAX_CELL_CHARS: usize = 5_000;
const MAX_TABLE_ROWS: usize = 20;
const MAX_TABLE_COLUMNS: usize = 10;
/// Word's highlight colours (docx-edit `HIGHLIGHT_HEX_TO_NAME`).
const HIGHLIGHTS: [&str; 16] = [
    "yellow",
    "green",
    "cyan",
    "magenta",
    "blue",
    "red",
    "darkBlue",
    "darkCyan",
    "darkGreen",
    "darkMagenta",
    "darkRed",
    "darkYellow",
    "darkGray",
    "lightGray",
    "black",
    "white",
];
const PARAGRAPH_ANCHOR: &str = "a new paragraph goes in the document body: after_id and before_id name a body paragraph, not one of a table, content control, header, footer or note";
const TABLE_ANCHOR: &str = "a new table goes in the document body: after_id names a body paragraph, not one of a table, content control, header, footer or note";
/// The root map of DOCX stories (`crate::document_roots`).
const STORIES: &str = "stories";

fn invalid(message: impl Into<String>) -> Error {
    Error::edit(EditCode::InvalidInput, message)
}

fn ctx() -> EditCtx {
    EditCtx::local(String::new(), String::new())
}

/// An op's refusal of the agent's input, else an engine fault.
fn op(error: OpError) -> Error {
    match error {
        OpError::TextContainsBreak | OpError::InvalidFormatValue(_) => invalid(error.to_string()),
        error => Error::engine(error),
    }
}

/// set_list's and set_alignment's refusal: Word's Reject All would undo
/// the change unseen (Epo 2026-09-25).
fn tracked_refusal() -> Error {
    Error::edit(
        EditCode::UnavailableTarget,
        "the paragraph holds tracked changes; it cannot be changed",
    )
}

fn gone() -> Error {
    Error::edit(
        EditCode::UnavailableTarget,
        "the paragraph is no longer in the document",
    )
}

/// A refusal carrying the paragraph as it reads now.
fn paragraph_refusal(code: EditCode, message: String, id: &str, text: &str) -> Error {
    Error::Edit {
        code,
        message,
        details: Some(json!({ "block": { "id": id, "type": "paragraph", "text": text } })),
    }
}

/// A new cell's text: one paragraph of at most [`MAX_CELL_CHARS`] characters.
fn check_cell(text: &str) -> Result<()> {
    if text.contains(['\n', '\r', '\u{b}', '\u{2028}', '\u{2029}']) {
        return Err(invalid("a cell must be one paragraph without line breaks"));
    }
    let chars = text.chars().count();
    if chars > MAX_CELL_CHARS {
        return Err(invalid(format!(
            "a cell holds {chars} characters; at most {MAX_CELL_CHARS}"
        )));
    }
    Ok(())
}

/// `story` and paraId of a target id (`<story>:paragraph:<paraId>`).
fn split_id(id: &str) -> Result<(&str, &str)> {
    id.rsplit_once(":paragraph:")
        .filter(|(story, _)| !story.is_empty())
        .ok_or_else(|| {
            Error::edit(
                EditCode::UnavailableTarget,
                format!("{id:?} is not a DOCX paragraph id"),
            )
        })
}

/// Header, footer and note paragraphs take `replace_text` only (Epo
/// 2026-10-21); the body's tables and content controls are the body's.
fn in_body(story: &str) -> Result<()> {
    if story == BODY || story.starts_with("body:") {
        return Ok(());
    }
    Err(Error::edit(
        EditCode::UnsupportedOperation,
        "header, footer and note paragraphs take replace_text only",
    ))
}

/// A cell story's table story: `<story>:t<n>:r<row>c<column>` names a cell.
fn table_story(cell: &str) -> Option<&str> {
    let digits = |text: &str| !text.is_empty() && text.bytes().all(|byte| byte.is_ascii_digit());
    let (rest, place) = cell.rsplit_once(':')?;
    let (row, column) = place.strip_prefix('r')?.split_once('c')?;
    let (story, table) = rest.rsplit_once(':')?;
    (digits(row) && digits(column) && table.strip_prefix('t').is_some_and(digits)).then_some(story)
}

/// A table embed's cell stories, row by row.
fn table_cells(payload: &BTreeMap<String, Any>) -> Vec<Vec<String>> {
    let items = |value: Option<&Any>| match value {
        Some(Any::Array(items)) => items.to_vec(),
        _ => Vec::new(),
    };
    let field = |value: &Any, key: &str| match value {
        Any::Map(map) => map.get(key).cloned(),
        _ => None,
    };
    items(payload.get("rows"))
        .iter()
        .map(|row| {
            items(field(row, "cells").as_ref())
                .iter()
                .filter_map(|cell| match field(cell, "story") {
                    Some(Any::String(story)) => Some(story.to_string()),
                    _ => None,
                })
                .collect()
        })
        .collect()
}

/// Why the text commands refuse a paragraph holding an object of `kind`.
fn locked_by(kind: &str) -> String {
    match kind {
        "table" => "holds a table".into(),
        "blockSdt" | "inlineSdt" | "sdt" => "holds a content control".into(),
        "image" => "holds a picture".into(),
        "shape" => "holds a drawing".into(),
        "mathEquation" => "holds an equation".into(),
        "footnoteRef" | "endnoteRef" => "holds a note reference".into(),
        kind => format!("holds a {kind}"),
    }
}

/// A Yjs attribute as JavaScript would test it.
fn truthy(value: Option<&Any>) -> bool {
    !matches!(
        value,
        None | Some(Any::Null | Any::Undefined | Any::Bool(false))
    )
}

/// A paragraph as the agent's commands and read see it (`docxParagraphs`,
/// with the tables and content controls its slot opens with set apart).
struct Para {
    para_id: String,
    /// Where its own content starts: after the page and column breaks,
    /// tables and content controls its slot opens with.
    start: u32,
    /// Its mark's index.
    mark: u32,
    /// Its text, without tracked deletions.
    text: String,
    /// Why the text commands refuse it: an object or a tracked change in it
    /// or on its mark.
    locked: Option<String>,
    /// A tracked change in it or on its mark, which set_list and
    /// set_alignment refuse too (Epo 2026-09-25).
    tracked: bool,
    /// The tables and content controls its slot opens with, in order.
    blocks: Vec<(String, BTreeMap<String, Any>)>,
    properties: BTreeMap<String, Any>,
}

impl Para {
    /// Why `replace_text` and `format_text` refuse it: [`Self::locked`],
    /// else the content control or table before it (a content control first,
    /// as delete_paragraph refuses only that).
    fn refused(&self) -> Option<String> {
        self.locked.clone().or_else(|| {
            if self.follows_control() {
                Some("follows a content control".to_owned())
            } else {
                (!self.blocks.is_empty()).then(|| "follows a table".to_owned())
            }
        })
    }

    fn follows_control(&self) -> bool {
        self.blocks.iter().any(|(kind, _)| kind == "blockSdt")
    }

    fn ends_section(&self) -> bool {
        SECTION_KEYS
            .iter()
            .any(|key| self.properties.contains_key(*key))
    }

    fn style(&self) -> Option<&str> {
        match self.properties.get("pStyle") {
            Some(Any::String(style)) => Some(style),
            _ => None,
        }
    }
}

/// `story`'s paragraphs in order.
fn paras(doc: &EditingDoc, story: &str) -> Result<Vec<Para>> {
    let segments = doc
        .story_segments(story)
        .map_err(|_| Error::edit(EditCode::UnavailableTarget, "unknown DOCX story"))?;
    let mut paras = Vec::new();
    let (mut at, mut start) = (0, 0);
    let mut text = String::new();
    let mut locked: Option<String> = None;
    let mut tracked = false;
    let mut blocks = Vec::new();
    for segment in segments {
        match segment.content {
            SegmentContent::Text(part) => {
                let deleted = truthy(segment.attributes.get("del"));
                if deleted || truthy(segment.attributes.get("ins")) {
                    locked.get_or_insert_with(|| "holds tracked changes".into());
                    tracked = true;
                }
                at += len16(&part);
                if !deleted {
                    text.push_str(&part);
                }
            }
            SegmentContent::Pilcrow(properties) => {
                let mark_tracked = ["ins", "del"]
                    .iter()
                    .any(|key| truthy(segment.attributes.get(*key)))
                    || TRACKED_KEYS
                        .iter()
                        .any(|key| properties.values.contains_key(*key));
                if mark_tracked {
                    locked.get_or_insert_with(|| "holds tracked changes".into());
                }
                paras.push(Para {
                    para_id: properties.para_id,
                    start,
                    mark: at,
                    text: std::mem::take(&mut text),
                    locked: locked.take(),
                    tracked: std::mem::take(&mut tracked) || mark_tracked,
                    blocks: std::mem::take(&mut blocks),
                    properties: properties.values,
                });
                at += 1;
                start = at;
            }
            SegmentContent::OtherEmbed { kind, payload } => {
                let leading = at == start;
                at += 1;
                let block = matches!(kind.as_str(), "table" | "blockSdt");
                if leading && (block || matches!(kind.as_str(), "pageBreak" | "columnBreak")) {
                    start = at;
                } else {
                    locked.get_or_insert_with(|| locked_by(&kind));
                }
                if block {
                    blocks.push((kind, payload));
                }
            }
        }
    }
    Ok(paras)
}

/// The `ordinal`-th table embed of `story` and the format markers at it.
fn table_embed(doc: &EditingDoc, story: &str, ordinal: u32) -> Result<Touched> {
    let segments = doc
        .story_segments(story)
        .map_err(|_| Error::edit(EditCode::UnavailableTarget, "unknown DOCX story"))?;
    let (mut at, mut seen) = (0, 0);
    for segment in segments {
        match segment.content {
            SegmentContent::Text(text) => at += len16(&text),
            SegmentContent::OtherEmbed { kind, .. } if kind == "table" => {
                if seen == ordinal {
                    return Ok(touched::story_range(
                        doc.yrs_doc(),
                        STORIES,
                        story,
                        at,
                        at + 1,
                        Kept::Content,
                    ));
                }
                seen += 1;
                at += 1;
            }
            _ => at += 1,
        }
    }
    Ok(Touched::default())
}

/// Story `story` whole, with the cells' stories of each table in it.
fn story_tree(doc: &EditingDoc, story: &str) -> Result<Touched> {
    let mut sum = touched::story(doc.yrs_doc(), STORIES, story);
    let segments = doc
        .story_segments(story)
        .map_err(|_| Error::edit(EditCode::UnavailableTarget, "unknown DOCX story"))?;
    for segment in segments {
        if let SegmentContent::OtherEmbed { kind, payload } = segment.content
            && kind == "table"
        {
            for cell in table_cells(&payload).iter().flatten() {
                sum += story_tree(doc, cell)?;
            }
        }
    }
    Ok(sum)
}

/// Whether `story` holds a table or a content control (H2: delete_table's
/// expected cells cannot name their text).
fn holds_blocks(doc: &EditingDoc, story: &str) -> Result<bool> {
    Ok(doc
        .story_segments(story)
        .map_err(|_| Error::edit(EditCode::UnavailableTarget, "unknown DOCX story"))?
        .iter()
        .any(|segment| {
            matches!(&segment.content, SegmentContent::OtherEmbed { kind, .. }
                if matches!(kind.as_str(), "table" | "blockSdt" | "inlineSdt" | "sdt"))
        }))
}

/// The header, footer and note stories after the body, labelled for the
/// read, in the order the save's story walk (`reach.rs`) visits them.
fn parts(package: &V) -> Vec<(String, String)> {
    let mut parts = Vec::new();
    for kind in ["headers", "footers"] {
        for (id, part) in map_entries(&package.get(kind)) {
            let label = format!(
                "{} ({})",
                part.get("type").to_js_string(),
                part.get("hdrFtrType").to_js_string()
            );
            parts.push((format!("hf:{id}"), label));
        }
    }
    for (prefix, kind, name) in [
        ("fn:", "footnotes", "footnote"),
        ("en:", "endnotes", "endnote"),
    ] {
        for note in package.get(kind).items() {
            if note
                .get("noteType")
                .as_str()
                .is_some_and(|kind| &*kind != "normal")
            {
                continue;
            }
            let id = note.get("id").to_js_string();
            parts.push((format!("{prefix}{id}"), format!("{name} {id}")));
        }
    }
    parts
}

/// The file's styles and numbering, for the lists paragraph styles give.
struct Lists {
    styles: StyleResolver,
    numbering: Numbering,
}

impl Lists {
    /// A paragraph's list as the read names it: its own `numPr` (numId 0
    /// ends a list its style gives), else its style's.
    fn name(&self, properties: &BTreeMap<String, Any>) -> Option<&'static str> {
        if let Some(Any::Map(num_pr)) = properties.get("numPr")
            && let Some(Any::Number(id)) = num_pr.get("numId")
        {
            return (*id > 0.0).then(|| match properties.get("listIsBullet") {
                Some(Any::Bool(true)) => "bullet",
                _ => "numbered",
            });
        }
        let style = match properties.get("pStyle") {
            Some(Any::String(style)) => V::from(style.as_ref()),
            _ => V::Undef,
        };
        let num_pr = self
            .styles
            .resolve_paragraph_style(&style, &V::Undef)
            .get("numPr");
        let id = num_pr.get("numId").as_num().filter(|id| *id > 0.0)?;
        let ilvl = num_pr.get("ilvl").as_num().unwrap_or(0.0);
        let format = self
            .numbering
            .level(&V::Num(id), &V::Num(ilvl))
            .get("numFmt");
        Some(if format.as_str().as_deref() == Some("bullet") {
            "bullet"
        } else {
            "numbered"
        })
    }
}

/// A paragraph style of the file.
struct ParagraphStyle {
    id: String,
    name: String,
    /// Shown in Word's style lists: neither hidden nor semi-hidden.
    shown: bool,
    default: bool,
    /// The style Enter at a paragraph's end gives the next one (`w:next`).
    next: Option<String>,
}

/// Where a new body paragraph goes.
enum Place {
    /// Split off after the body paragraph at this index.
    After(usize),
    /// Split off before the body paragraph at this index.
    Before(usize),
    /// Into the body's only paragraph, empty, as a new file's is.
    Fill,
}

/// A new body paragraph's place, the body paragraph it continues as Enter
/// does, and the anchor whose last insert it becomes.
struct Placement {
    place: Place,
    continued: usize,
    after: Option<String>,
}

/// A table found by a cell paragraph: its story, its ordinal among that
/// story's tables, its cell stories row by row and its payload.
type Table = (String, u32, Vec<Vec<String>>, BTreeMap<String, Any>);

/// One cell's mark on the paragraphs a read lists.
#[derive(Clone)]
struct CellMark {
    table: Option<String>,
    row: usize,
    column: usize,
}

/// The read's walk: paragraphs counted in order, the page's listed.
struct Walk<'a> {
    doc: &'a EditingDoc,
    lists: Lists,
    /// The file's default paragraph style, which the read does not name.
    default: Option<String>,
    /// The header, footer or note being walked; `None` in the body.
    part: Option<String>,
    start: usize,
    end: usize,
    index: usize,
    listed: Vec<Value>,
}

impl Walk<'_> {
    /// `story`'s paragraphs in order, each table's cells before the
    /// paragraph whose slot the table opens; tracked-deleted text is left
    /// out, and a paragraph the text commands refuse says why (`locked`).
    fn story(&mut self, story: &str, cell: Option<&CellMark>, control: bool) -> Result<()> {
        for para in paras(self.doc, story)? {
            for (kind, payload) in &para.blocks {
                if kind == "table" {
                    self.table(&table_cells(payload), control)?;
                } else if let Some(Any::String(child)) = payload.get("story") {
                    self.story(child, cell, true)?;
                }
            }
            if (self.start..self.end).contains(&self.index) {
                let mut entry = json!({
                    "id": format!("{story}:paragraph:{}", para.para_id),
                    "text": para.text,
                });
                if let Some(style) = para.style()
                    && Some(style) != self.default.as_deref()
                {
                    entry["style"] = Value::from(style);
                }
                if let Some(list) = self.lists.name(&para.properties) {
                    entry["list"] = Value::from(list);
                }
                if let Some(part) = &self.part {
                    entry["part"] = Value::from(part.as_str());
                }
                if let Some(cell) = cell {
                    entry["table"] =
                        json!({ "id": cell.table, "row": cell.row, "column": cell.column });
                }
                if control {
                    entry["contentControl"] = Value::Bool(true);
                }
                if let Some(locked) = para.refused() {
                    entry["locked"] = Value::from(locked);
                }
                self.listed.push(entry);
            }
            self.index += 1;
        }
        Ok(())
    }

    /// A table's cells row by row, each paragraph marked with the table,
    /// which its first cell's first paragraph names.
    fn table(&mut self, cells: &[Vec<String>], control: bool) -> Result<()> {
        let table = match cells.first().and_then(|row| row.first()) {
            Some(first) => paras(self.doc, first)?
                .first()
                .map(|para| format!("{first}:paragraph:{}", para.para_id)),
            None => None,
        };
        for (row, stories) in cells.iter().enumerate() {
            for (column, cell) in stories.iter().enumerate() {
                let mark = CellMark {
                    table: table.clone(),
                    row,
                    column,
                };
                self.story(cell, Some(&mark), control)?;
            }
        }
        Ok(())
    }
}

impl DocxSession {
    /// [`crate::CommandResult::touched`], read before `command` applies:
    /// what its Undo re-creates (the Editor's rule, 2026-10-22). The text a
    /// `replace_text` takes out, the text and format markers a `format_text`
    /// span covers, the mark entries a `set_alignment` or `set_list` writes,
    /// a deleted paragraph's text and mark, a deleted table's embed and
    /// cells; nothing for the rest, nor where the target is not found (the
    /// command's own refusal says why).
    pub(crate) fn touched(&self, command: &Command) -> Touched {
        let doc = self.engine.doc();
        let range = |story: &str, start: u32, end: u32, kept| {
            touched::story_range(doc.yrs_doc(), STORIES, story, start, end, kept)
        };
        let measured = || -> Result<Touched> {
            Ok(match command {
                Command::ReplaceText { target_id, .. } => {
                    let (story, index, paras) = self.find(target_id)?;
                    let para = &paras[index];
                    let (span, _) = check_replacement(&para.text, command, "DOCX")?;
                    range(
                        &story,
                        para.start + span.start as u32,
                        para.start + span.end as u32,
                        Kept::Content,
                    )
                }
                Command::FormatText {
                    target_id,
                    expected_text,
                    ..
                } => {
                    let (story, index, paras) = self.find(target_id)?;
                    let para = &paras[index];
                    let at = para.text.find(expected_text.as_str()).ok_or_else(gone)?;
                    let start = para.start + len16(&para.text[..at]);
                    range(
                        &story,
                        start,
                        start + len16(expected_text),
                        Kept::Formatting,
                    )
                }
                Command::DeleteParagraph { target_id, .. } => {
                    let (story, index, paras) = self.find(target_id)?;
                    let para = &paras[index];
                    range(&story, para.start, para.mark + 1, Kept::Content)
                }
                Command::SetAlignment { target_id, .. } => {
                    let (story, index, paras) = self.find(target_id)?;
                    let mark = paras[index].mark;
                    touched::mark_entries(doc.yrs_doc(), STORIES, &story, mark, &["alignment"])
                }
                Command::SetList { ids, .. } => {
                    // Each story read once: a list may name 200 paragraphs.
                    let mut read: HashMap<&str, Vec<Para>> = HashMap::new();
                    let mut sum = Touched::default();
                    for id in ids {
                        let (story, para_id) = split_id(id)?;
                        if !read.contains_key(story) {
                            read.insert(story, paras(doc, story)?);
                        }
                        let para = read[story]
                            .iter()
                            .find(|para| para.para_id == para_id)
                            .ok_or_else(gone)?;
                        sum += touched::mark_entries(
                            doc.yrs_doc(),
                            STORIES,
                            story,
                            para.mark,
                            &LIST_KEYS,
                        );
                    }
                    sum
                }
                Command::DeleteTable { id, .. } => {
                    let (story, ordinal, cells, _) = self.table(id)?;
                    let mut sum = table_embed(doc, &story, ordinal)?;
                    for cell in cells.iter().flatten() {
                        sum += story_tree(doc, cell)?;
                    }
                    sum
                }
                _ => Touched::default(),
            })
        };
        measured().unwrap_or_default()
    }

    pub(super) fn apply_agent(&mut self, command: &Command) -> Result<Applied> {
        // Tracking keeps deleted items from garbage collection, as for
        // replace_text.
        self.undo.track(self.engine.doc());
        let none = |()| Applied {
            ids: Vec::new(),
            inverse: None,
        };
        let applied = match command {
            Command::InsertParagraph {
                after_id,
                before_id,
                text,
                style,
                list,
            } => self.insert_paragraph(
                after_id.as_deref(),
                before_id.as_deref(),
                text,
                style.as_deref(),
                *list,
            ),
            Command::DeleteParagraph {
                target_id,
                expected_text,
            } => self.delete_paragraph(target_id, expected_text).map(none),
            Command::SetList { ids, list } => self.set_list(ids, *list).map(none),
            Command::InsertTable { after_id, rows } => {
                self.insert_table(after_id.as_deref(), rows).map(none)
            }
            Command::DeleteTable { id, expected } => self.delete_table(id, expected).map(none),
            Command::FormatText {
                target_id,
                expected_text,
                format,
            } => self.format_text(target_id, expected_text, format).map(none),
            Command::SetAlignment {
                target_id,
                alignment,
            } => self.set_alignment(target_id, *alignment).map(none),
            _ => Err(Error::edit(
                EditCode::UnsupportedOperation,
                "this command edits XLSX files",
            )),
        }?;
        self.projected = None;
        Ok(applied)
    }

    /// After a call's commands: the empty paragraph each table the call
    /// inserted was given stays only where Word needs one, after a table
    /// that ends the body or comes before another table (Epo 2026-10-21);
    /// elsewhere the paragraph that follows takes the table.
    pub(crate) fn finish_agent(&mut self, commands: &[Command]) -> Result<()> {
        if !commands
            .iter()
            .any(|command| matches!(command, Command::InsertTable { .. }))
        {
            return Ok(());
        }
        let doc = self.engine.doc();
        let minted = format!("{}:", doc.client_id());
        let body = paras(doc, BODY)?;
        for (index, para) in body.iter().enumerate() {
            let Some(next) = body.get(index + 1) else {
                continue;
            };
            let inserted = match para.blocks.last() {
                Some((kind, payload)) if kind == "table" => {
                    match table_cells(payload).first().and_then(|row| row.first()) {
                        Some(first) => paras(doc, first)?
                            .first()
                            .is_some_and(|cell| cell.para_id.starts_with(&minted)),
                        None => false,
                    }
                }
                _ => false,
            };
            if inserted
                && para.start == para.mark
                && para.locked.is_none()
                && next.blocks.is_empty()
                && !para.ends_section()
            {
                doc.remove_paragraph(&ctx(), BODY, &para.para_id)
                    .map_err(op)?;
            }
        }
        self.projected = None;
        Ok(())
    }

    /// `count` paragraphs from `start`: the body's, each table's cells
    /// listed before the paragraph that follows it and marked with the
    /// table, row and column, then each header's, footer's and note's,
    /// marked with its part. A paragraph names its style when not the
    /// default; the first page also lists the file's shown paragraph
    /// styles, at most [`MAX_STYLES`].
    pub(crate) fn inspect_scoped(&mut self, query: &InspectQuery) -> Result<Value> {
        let start = query.start.unwrap_or(0) as usize;
        let count = query
            .count
            .filter(|count| (1..=MAX_PARAGRAPHS).contains(count))
            .ok_or_else(|| {
                invalid(format!(
                    "a DOCX read names a count of 1 to {MAX_PARAGRAPHS} paragraphs"
                ))
            })? as usize;
        let package = self.package();
        let styles = self.paragraph_styles();
        let mut walk = Walk {
            doc: self.engine.doc(),
            lists: Lists {
                styles: StyleResolver::new(&package.get("styles")),
                numbering: Numbering::new(&package.get("numbering")),
            },
            default: styles
                .iter()
                .find(|style| style.default)
                .map(|style| style.id.clone()),
            part: None,
            start,
            end: start + count,
            index: 0,
            listed: Vec::new(),
        };
        walk.story(BODY, None, false)?;
        for (story, part) in parts(&package) {
            // A part the session never seeded has no story.
            if walk.doc.story_segments(&story).is_err() {
                continue;
            }
            walk.part = Some(part);
            walk.story(&story, None, false)?;
        }
        let total = walk.index;
        let end = (start + count).min(total);
        let mut answer = json!({
            "paragraphs": walk.listed,
            "total": total,
            "nextStart": (end < total).then_some(end),
        });
        if start == 0 {
            let shown: Vec<&ParagraphStyle> = styles.iter().filter(|style| style.shown).collect();
            answer["styles"] = shown
                .iter()
                .take(MAX_STYLES)
                .map(|style| json!({ "id": style.id, "name": style.name }))
                .collect();
            if shown.len() > MAX_STYLES {
                answer["stylesMore"] = json!(shown.len() - MAX_STYLES);
            }
        }
        Ok(answer)
    }

    fn package(&self) -> V {
        self.base_document().get("package")
    }

    /// The file's paragraph styles in its styles part's order.
    fn paragraph_styles(&self) -> Vec<ParagraphStyle> {
        self.package()
            .get("styles")
            .get("styles")
            .items()
            .iter()
            .filter(|style| style.get("type").as_str().as_deref() == Some("paragraph"))
            .filter_map(|style| {
                let id = style.get("styleId").as_str()?.to_string();
                let name = style
                    .get("name")
                    .as_str()
                    .map_or_else(|| id.clone(), |name| name.to_string());
                Some(ParagraphStyle {
                    id,
                    name,
                    shown: !style.get("hidden").truthy() && !style.get("semiHidden").truthy(),
                    default: style.get("default").truthy(),
                    next: style.get("next").as_str().map(|next| next.to_string()),
                })
            })
            .collect()
    }

    /// The style Enter at the end of a paragraph styled `style` (none: the
    /// default) gives the next one: the style's next style when the file has
    /// it, else the style itself; none for the default style.
    fn next_style(styles: &[ParagraphStyle], style: Option<&str>) -> Option<String> {
        let default = styles
            .iter()
            .find(|known| known.default)
            .map(|known| known.id.as_str());
        let style = style.or(default)?;
        let next = styles
            .iter()
            .find(|known| known.id == style)
            .and_then(|known| known.next.as_deref())
            .filter(|next| styles.iter().any(|known| known.id == *next))
            .unwrap_or(style);
        (Some(next) != default).then(|| next.to_owned())
    }

    /// The `w:num` id `list` writes: the toolbar's bullet or number button's
    /// (the two after the file's own), or 0 for `none`.
    fn list_num_id(&self, list: ListKind) -> Result<f64> {
        if list == ListKind::None {
            return Ok(0.0);
        }
        let numbering = self.package().get("numbering");
        let numbering: Option<NumberingDefinitions> = if numbering.is_undef() {
            None
        } else {
            Some(
                serde_json::to_value(&numbering)
                    .and_then(serde_json::from_value)
                    .map_err(Error::engine)?,
            )
        };
        let (bullet, number) = list_button_num_ids(numbering.as_ref());
        Ok(if list == ListKind::Bullet {
            bullet
        } else {
            number
        })
    }

    /// `w:compatSetting`'s mode (documentToYrs.ts `compatibilityModeFromDocument`).
    fn compatibility_mode(&self) -> u8 {
        self.package()
            .get("settings")
            .get("compatibilityFlags")
            .get("compatibilityMode")
            .finite()
            .map_or(12, |mode| mode.trunc().clamp(0.0, 255.0) as u8)
    }

    /// The paragraph `id` names, its story and its story's paragraphs (body
    /// stories only, as every command but `replace_text` edits).
    fn find(&self, id: &str) -> Result<(String, usize, Vec<Para>)> {
        let (story, para_id) = split_id(id)?;
        in_body(story)?;
        let paras = paras(self.engine.doc(), story)?;
        let index = paras
            .iter()
            .position(|para| para.para_id == para_id)
            .ok_or_else(gone)?;
        Ok((story.to_owned(), index, paras))
    }

    /// Where a new body paragraph goes: right after `after_id`, after what
    /// this call already put there, so inserts after one paragraph land in
    /// call order; right before `before_id`; without either at the body's
    /// end, or into its only paragraph when that is empty. It continues the
    /// paragraph `after_id` names, else the one before it, at the body's
    /// start the one it goes before.
    fn place(
        &self,
        body: &[Para],
        after_id: Option<&str>,
        before_id: Option<&str>,
        refusal: &str,
    ) -> Result<Placement> {
        let index = |id: &str| -> Result<usize> {
            let para_id = id
                .strip_prefix("body:paragraph:")
                .ok_or_else(|| invalid(refusal))?;
            body.iter()
                .position(|para| para.para_id == para_id)
                .ok_or_else(gone)
        };
        let after = |named: usize| {
            let at = self
                .after
                .get(&body[named].para_id)
                .and_then(|id| body.iter().position(|para| para.para_id == *id))
                .unwrap_or(named);
            Placement {
                place: Place::After(at),
                continued: named,
                after: Some(body[named].para_id.clone()),
            }
        };
        Ok(match (after_id, before_id) {
            (Some(_), Some(_)) => {
                return Err(invalid(
                    "insert_paragraph names at most one of after_id and before_id",
                ));
            }
            (Some(id), None) => after(index(id)?),
            (None, Some(id)) => {
                let at = index(id)?;
                Placement {
                    place: Place::Before(at),
                    continued: at.saturating_sub(1),
                    after: None,
                }
            }
            (None, None) => match body {
                [only] if only.mark == 0 => Placement {
                    place: Place::Fill,
                    continued: 0,
                    after: None,
                },
                [.., _] => after(body.len() - 1),
                [] => return Err(Error::engine("the document body has no paragraph")),
            },
        })
    }

    /// A new empty body paragraph where `placement` puts it, its mark
    /// holding `kept` besides the identity and section of the mark it
    /// copies; only what differs from that copy is written. Returns its id,
    /// where its text starts and whether deleting it undoes the insert.
    fn new_paragraph(
        &mut self,
        body: &[Para],
        placement: Placement,
        kept: BTreeMap<String, Any>,
    ) -> Result<(String, u32, bool)> {
        let doc = self.engine.doc();
        let (new_id, at, neighbour, undoable) = match placement.place {
            Place::After(index) => {
                let para = &body[index];
                let receipt = doc
                    .split_paragraph(&ctx(), Position::new(BODY, para.mark))
                    .map_err(op)?;
                // A split at a section's end leaves the section mark on the
                // new paragraph, as Enter does; deleting it would join two
                // sections.
                let undoable = index + 1 < body.len() && !para.ends_section();
                (receipt.second_para_id, para.mark + 1, para, undoable)
            }
            Place::Before(index) => {
                let receipt = doc
                    .insert_paragraph_before(&ctx(), BODY, &body[index].para_id)
                    .map_err(op)?;
                let at = index
                    .checked_sub(1)
                    .map_or(0, |previous| body[previous].mark + 1);
                (receipt.first_para_id, at, &body[index], true)
            }
            Place::Fill => (body[0].para_id.clone(), body[0].mark, &body[0], false),
        };
        let mut other: BTreeMap<String, Option<Any>> = BTreeMap::new();
        for (key, value) in &neighbour.properties {
            if !KEPT_KEYS.contains(&key.as_str()) && kept.get(key) != Some(value) {
                other.insert(key.clone(), kept.get(key).cloned());
            }
        }
        for (key, value) in kept {
            if !neighbour.properties.contains_key(&key) {
                other.insert(key, Some(value));
            }
        }
        if !other.is_empty() {
            doc.set_paragraph_attrs(
                &ctx(),
                &ParaSelector::One(new_id.clone()),
                &ParaAttrDelta {
                    other,
                    ..ParaAttrDelta::default()
                },
            )
            .map_err(op)?;
        }
        if let Some(anchor) = placement.after {
            self.after.insert(anchor, new_id.clone());
        }
        Ok((new_id, at, undoable))
    }

    /// A body paragraph placed by [`Self::place`]. Without a style it
    /// continues its anchor as Enter at the anchor's end does (Epo
    /// 2026-10-21): the anchor's paragraph properties (list, indent) and its
    /// style's next style. A named style keeps those properties and takes
    /// the place of the next style (Epo 2026-10-22). `list` sets or ends the
    /// paragraph's list at the anchor's level. The inverse deletes it,
    /// unless delete_paragraph would refuse that.
    fn insert_paragraph(
        &mut self,
        after_id: Option<&str>,
        before_id: Option<&str>,
        text: &str,
        style: Option<&str>,
        list: Option<ListKind>,
    ) -> Result<Applied> {
        check_paragraph_text(text)?;
        let styles = self.paragraph_styles();
        if let Some(style) = style
            && !styles.iter().any(|known| known.id == style)
        {
            return Err(invalid(format!(
                "the document has no paragraph style {style:?}; use a style id inspect_office_document lists"
            )));
        }
        let body = paras(self.engine.doc(), BODY)?;
        let placement = self.place(&body, after_id, before_id, PARAGRAPH_ANCHOR)?;
        let continued = &body[placement.continued];
        let fill = matches!(placement.place, Place::Fill);
        // A new paragraph keeps no source and save bookkeeping (`_` keys) or
        // tracked change of the one it continues; a filled one keeps its own.
        // The style is set below.
        let mut kept: BTreeMap<String, Any> = continued
            .properties
            .iter()
            .filter(|(key, _)| {
                !KEPT_KEYS.contains(&key.as_str())
                    && *key != "pStyle"
                    && (fill || (!TRACKED_KEYS.contains(&key.as_str()) && !key.starts_with('_')))
            })
            .map(|(key, value)| (key.clone(), value.clone()))
            .collect();
        let style = match (style, &placement.place) {
            (Some(style), _) => Some(style.to_owned()),
            // A new file's empty paragraph keeps its own style.
            (None, Place::Fill) => continued.style().map(str::to_owned),
            (None, _) => Self::next_style(&styles, continued.style()),
        };
        if let Some(style) = style {
            kept.insert("pStyle".to_owned(), Any::from(style));
        }
        if let Some(list) = list {
            let ilvl = match continued.properties.get("numPr") {
                Some(Any::Map(num_pr)) => match num_pr.get("ilvl") {
                    Some(Any::Number(ilvl)) => *ilvl,
                    _ => 0.0,
                },
                _ => 0.0,
            };
            for (key, value) in list_attrs(list, self.list_num_id(list)?, ilvl) {
                match value {
                    Some(Any::Null) | None => kept.remove(&key),
                    Some(value) => kept.insert(key, value),
                };
            }
        }
        let (new_id, at, undoable) = self.new_paragraph(&body, placement, kept)?;
        if !text.is_empty() {
            self.engine
                .doc()
                .insert_text(&ctx(), Position::new(BODY, at), text, FormatPolicy::Plain)
                .map_err(op)?;
        }
        let target_id = format!("{BODY}:paragraph:{new_id}");
        Ok(Applied {
            ids: vec![target_id.clone()],
            inverse: undoable.then(|| Command::DeleteParagraph {
                target_id,
                expected_text: text.to_owned(),
            }),
        })
    }

    /// Refused when the paragraph holds an object or a tracked change,
    /// follows a content control, ends its story or a section, or reads
    /// other than `expected_text`. One that follows a table hands the table
    /// to the paragraph after it, unless Word needs it there. No inverse:
    /// putting it back exactly is the live Undo's.
    fn delete_paragraph(&mut self, id: &str, expected_text: &str) -> Result<()> {
        let (story, index, paras) = self.find(id)?;
        let para = &paras[index];
        if let Some(reason) = para.locked.clone().or_else(|| {
            para.follows_control()
                .then(|| "follows a content control".to_owned())
        }) {
            return Err(Error::edit(
                EditCode::UnavailableTarget,
                format!("the paragraph {reason}; it cannot be deleted"),
            ));
        }
        check_expected_paragraph(id, &para.text, expected_text)?;
        if index + 1 == paras.len() {
            return Err(invalid(if para.blocks.is_empty() {
                "a story's last paragraph cannot be deleted; replace its text instead"
            } else {
                "the paragraph after a table that ends the body or a cell cannot be deleted; Word needs it"
            }));
        }
        if !para.blocks.is_empty() && !paras[index + 1].blocks.is_empty() {
            return Err(invalid(
                "the paragraph separates two tables or content controls; Word needs it",
            ));
        }
        if para.ends_section() {
            return Err(invalid(
                "the paragraph ends a section; deleting it would join two sections",
            ));
        }
        // The paragraphs around it keep their own properties and ids.
        self.engine
            .doc()
            .remove_paragraph(&ctx(), &story, &para.para_id)
            .map_err(op)?;
        Ok(())
    }

    /// Each paragraph keeps its list level; `none` ends its list, its
    /// style's too.
    fn set_list(&mut self, ids: &[String], list: ListKind) -> Result<()> {
        if ids.is_empty() || ids.len() > MAX_PARAGRAPHS as usize {
            return Err(invalid(format!(
                "set_list names 1 to {MAX_PARAGRAPHS} paragraphs"
            )));
        }
        let num_id = self.list_num_id(list)?;
        let mut stories: HashMap<&str, Vec<Para>> = HashMap::new();
        let mut targets = Vec::with_capacity(ids.len());
        for id in ids {
            let (story, para_id) = split_id(id)?;
            in_body(story)?;
            if !stories.contains_key(story) {
                stories.insert(story, paras(self.engine.doc(), story)?);
            }
            let para = stories[story]
                .iter()
                .find(|para| para.para_id == para_id)
                .ok_or_else(gone)?;
            if para.tracked {
                return Err(tracked_refusal());
            }
            let ilvl = match para.properties.get("numPr") {
                Some(Any::Map(num_pr)) => match num_pr.get("ilvl") {
                    Some(Any::Number(ilvl)) => *ilvl,
                    _ => 0.0,
                },
                _ => 0.0,
            };
            targets.push((story, para.mark, ilvl));
        }
        for (story, at, ilvl) in targets {
            self.engine
                .doc()
                .set_paragraph_attrs(
                    &ctx(),
                    &ParaSelector::Range(StoryRange::new(story, at, at)),
                    &ParaAttrDelta {
                        other: list_attrs(list, num_id, ilvl).into_iter().collect(),
                        ..ParaAttrDelta::default()
                    },
                )
                .map_err(op)?;
        }
        Ok(())
    }

    /// A table of `rows` in the slot of a new plain empty paragraph placed
    /// as [`Self::place`] places one; [`Self::finish_agent`] removes that
    /// paragraph where Word does not need it.
    fn insert_table(&mut self, after_id: Option<&str>, rows: &[Vec<String>]) -> Result<()> {
        let columns = rows.first().map_or(0, Vec::len);
        if rows.is_empty()
            || rows.len() > MAX_TABLE_ROWS
            || !(1..=MAX_TABLE_COLUMNS).contains(&columns)
            || rows.iter().any(|row| row.len() != columns)
        {
            return Err(invalid(format!(
                "rows must be 1 to {MAX_TABLE_ROWS} rows of the same 1 to {MAX_TABLE_COLUMNS} cells"
            )));
        }
        for cell in rows.iter().flatten() {
            check_cell(cell)?;
        }
        let body = paras(self.engine.doc(), BODY)?;
        let placement = self.place(&body, after_id, None, TABLE_ANCHOR)?;
        let (_, at, _) = self.new_paragraph(&body, placement, BTreeMap::new())?;
        let doc = self.engine.doc();
        let table = doc
            .insert_table(
                &ctx(),
                Position::new(BODY, at),
                rows.len() as u32,
                columns as u32,
                self.compatibility_mode(),
            )
            .map_err(op)?;
        for (story, text) in table.created_story_ids.iter().zip(rows.iter().flatten()) {
            if !text.is_empty() {
                doc.insert_text(
                    &ctx(),
                    Position::new(story.as_str(), 0),
                    text,
                    FormatPolicy::Plain,
                )
                .map_err(op)?;
            }
        }
        Ok(())
    }

    /// The [`Table`] holding cell paragraph `id`.
    fn table(&self, id: &str) -> Result<Table> {
        let (cell, _, _) = self.find(id)?;
        let story = table_story(&cell).ok_or_else(|| {
            invalid("id must be a paragraph of a table cell, as inspect_office_document marks it")
        })?;
        let segments = self
            .engine
            .doc()
            .story_segments(story)
            .map_err(|_| Error::edit(EditCode::UnavailableTarget, "unknown DOCX story"))?;
        let (ordinal, (cells, payload)) = segments
            .into_iter()
            .filter_map(|segment| match segment.content {
                SegmentContent::OtherEmbed { kind, payload } if kind == "table" => {
                    Some((table_cells(&payload), payload))
                }
                _ => None,
            })
            .enumerate()
            .find(|(_, (cells, _))| cells.iter().flatten().any(|story| *story == cell))
            .ok_or_else(|| {
                Error::edit(
                    EditCode::UnavailableTarget,
                    "the table is no longer in the document",
                )
            })?;
        Ok((story.to_owned(), ordinal as u32, cells, payload))
    }

    /// The table holding paragraph `id`, refused unless its cells read
    /// `expected` and hold no table or content control, whose text expected
    /// cannot name.
    fn delete_table(&mut self, id: &str, expected: &[Vec<String>]) -> Result<()> {
        let (story, ordinal, cells, _) = self.table(id)?;
        let doc = self.engine.doc();
        for cell in cells.iter().flatten() {
            if holds_blocks(doc, cell)? {
                return Err(Error::edit(
                    EditCode::UnavailableTarget,
                    "a cell of the table holds a table or content control; the table cannot be deleted",
                ));
            }
        }
        let grid = cells
            .iter()
            .map(|row| {
                row.iter()
                    .map(|cell| {
                        Ok(paras(doc, cell)?
                            .into_iter()
                            .map(|para| para.text)
                            .collect::<Vec<_>>()
                            .join("\n"))
                    })
                    .collect::<Result<Vec<_>>>()
            })
            .collect::<Result<Vec<_>>>()?;
        if grid != expected {
            let tsv = grid
                .iter()
                .map(|row| row.join("\t"))
                .collect::<Vec<_>>()
                .join("\n");
            return Err(Error::Edit {
                code: EditCode::StaleTarget,
                message: "the table's cells differ from expected; read it again".into(),
                details: Some(json!({
                    "block": { "id": id, "type": "table", "text": tsv, "grid": grid }
                })),
            });
        }
        doc.delete_table(&ctx(), &TableLocator::new(story.as_str(), ordinal))
            .map_err(op)?;
        Ok(())
    }

    fn format_text(&mut self, id: &str, expected_text: &str, format: &TextFormat) -> Result<()> {
        let delta = inline_delta(format)?;
        let (story, index, paras) = self.find(id)?;
        let para = &paras[index];
        if let Some(reason) = para.refused() {
            return Err(Error::edit(
                EditCode::UnavailableTarget,
                format!("the paragraph {reason}"),
            ));
        }
        if expected_text.is_empty() {
            return Err(invalid("expected_text names the text to format"));
        }
        let found: Vec<usize> = para
            .text
            .match_indices(expected_text)
            .map(|(at, _)| at)
            .collect();
        let at = match found[..] {
            [at] => at,
            [] => {
                return Err(paragraph_refusal(
                    EditCode::StaleTarget,
                    "expected text was not found in the paragraph".into(),
                    id,
                    &para.text,
                ));
            }
            _ => {
                return Err(paragraph_refusal(
                    EditCode::InvalidInput,
                    format!(
                        "expected text occurs {} times in the paragraph; include more context",
                        found.len()
                    ),
                    id,
                    &para.text,
                ));
            }
        };
        let start = para.start + len16(&para.text[..at]);
        self.engine
            .doc()
            .format_range(
                &ctx(),
                StoryRange::new(story, start, start + len16(expected_text)),
                &delta,
            )
            .map_err(op)?;
        Ok(())
    }

    fn set_alignment(&mut self, id: &str, alignment: Alignment) -> Result<()> {
        let (story, index, paras) = self.find(id)?;
        if paras[index].tracked {
            return Err(tracked_refusal());
        }
        let mark = paras[index].mark;
        let value = match alignment {
            Alignment::Left => "left",
            Alignment::Center => "center",
            Alignment::Right => "right",
            Alignment::Justify => "both",
        };
        self.engine
            .doc()
            .set_paragraph_attrs(
                &ctx(),
                &ParaSelector::Range(StoryRange::new(story, mark, mark)),
                &ParaAttrDelta {
                    alignment: Patch::Set(value.to_owned()),
                    ..ParaAttrDelta::default()
                },
            )
            .map_err(op)?;
        Ok(())
    }
}

/// The mark keys [`list_attrs`] writes.
const LIST_KEYS: [&str; 4] = ["numPr", "listIsBullet", "listNumFmt", "listMarker"];

/// The paragraph properties a list writes, as the toolbar's buttons write
/// them: `num_id` (0 for `none`, which also ends a list the paragraph's
/// style gives it) at level `ilvl`.
fn list_attrs(list: ListKind, num_id: f64, ilvl: f64) -> [(String, Option<Any>); 4] {
    let null = || Some(Any::Null);
    let mut num_pr = HashMap::from([("numId".to_owned(), Any::Number(num_id))]);
    if list != ListKind::None {
        num_pr.insert("ilvl".to_owned(), Any::Number(ilvl));
    }
    let bullet = list == ListKind::Bullet;
    [
        ("numPr".to_owned(), Some(Any::Map(Arc::new(num_pr)))),
        (
            "listIsBullet".to_owned(),
            if list == ListKind::None {
                null()
            } else {
                Some(Any::Bool(bullet))
            },
        ),
        (
            "listNumFmt".to_owned(),
            if list == ListKind::Numbered {
                Some(Any::from("decimal"))
            } else {
                null()
            },
        ),
        ("listMarker".to_owned(), null()),
    ]
}

/// `format_text`'s fields as the editor's formatting delta.
fn inline_delta(format: &TextFormat) -> Result<InlineFormatDelta> {
    let on = |value: Option<bool>| match value {
        None => Patch::Keep,
        Some(value) => Patch::Set(value),
    };
    let mut delta = InlineFormatDelta {
        bold: on(format.bold),
        italic: on(format.italic),
        underline: match format.underline {
            None => Patch::Keep,
            Some(true) => Patch::Set(UnderlinePatch::default()),
            Some(false) => Patch::Clear,
        },
        strike: match format.strikethrough {
            None => Patch::Keep,
            Some(true) => Patch::Set(StrikePatch::default()),
            Some(false) => Patch::Clear,
        },
        ..InlineFormatDelta::default()
    };
    if let Some(color) = &format.color {
        let hex = color.strip_prefix('#').unwrap_or(color);
        if hex.len() != 6 || !hex.bytes().all(|byte| byte.is_ascii_hexdigit()) {
            return Err(invalid(format!(
                "color {color:?} is not a hex colour like \"#1F4E79\""
            )));
        }
        delta.color = Patch::Set(ColorPatch::Rgb(hex.to_ascii_uppercase()));
    }
    if let Some(highlight) = &format.highlight {
        delta.highlight = if highlight.eq_ignore_ascii_case("none") {
            Patch::Clear
        } else {
            let name = HIGHLIGHTS
                .iter()
                .find(|name| name.eq_ignore_ascii_case(highlight))
                .ok_or_else(|| {
                    invalid(format!(
                        "highlight {highlight:?} is not one of none, {}",
                        HIGHLIGHTS.join(", ")
                    ))
                })?;
            Patch::Set((*name).to_owned())
        };
    }
    if let Some(size) = format.font_size {
        if !(1.0..=1638.0).contains(&size) {
            return Err(invalid("font_size is 1 to 1638 points"));
        }
        delta.font_size = Patch::Set(size);
    }
    if let Some(family) = &format.font_family {
        let family = family.trim();
        if family.is_empty() || family.chars().count() > 31 || family.chars().any(char::is_control)
        {
            return Err(invalid("font_family is a font name of 1 to 31 characters"));
        }
        delta.font_family = Patch::Set(FontFamilyPatch {
            ascii: family.to_owned(),
            h_ansi: None,
        });
    }
    if delta.is_empty() {
        return Err(invalid("format_text names no formatting"));
    }
    Ok(delta)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn cell_stories_name_their_table_story() {
        assert_eq!(table_story("body:t0:r1c2"), Some("body"));
        assert_eq!(table_story("body:t3:r0c0:t1:r10c4"), Some("body:t3:r0c0"));
        assert_eq!(table_story("body"), None);
        assert_eq!(table_story("body:t0:rxc1"), None);
        assert_eq!(table_story("hf:rId1"), None);
    }
}
