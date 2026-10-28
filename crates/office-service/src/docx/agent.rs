//! The chat agent's DOCX commands and scoped read (Epo 2026-10-20 to
//! 2026-10-22): insert and delete paragraphs and tables, lists, text
//! formatting and alignment in the body; and a page of the paragraphs, the
//! body's then the headers', footers' and notes' (which only `replace_text`
//! edits), each with its style when not the default, the first page with the
//! file's paragraph styles. A new paragraph continues its anchor as Enter
//! does; a named style takes the place of the anchor style's next style.

use std::borrow::Cow;
use std::collections::{BTreeMap, HashMap};
use std::rc::Rc;
use std::sync::Arc;

use docx_edit::{
    ColorPatch, EditCtx, EditingDoc, FontFamilyPatch, FormatPolicy, InlineFormatDelta, OpError,
    ParaAttrDelta, ParaSelector, Patch, Position, StoryRange, StrikePatch, TableLocator,
    UnderlinePatch,
};
use docx_parse::NumberingDefinitions;
use docx_parse::numbering::list_button_num_ids;
use serde_json::{Value, json};
use yrs::block::ClientID;
use yrs::block::{ItemContent, ItemPtr};
use yrs::branch::Branch;
use yrs::{Any, ID, Map, Out, ReadTxn, TextRef, Transact};

use super::styles::{Numbering, StyleResolver, map_entries};
use super::{DocxSession, Paragraph};
use crate::common::{
    check_expected_paragraph, check_paragraph_text, check_replacement, len16, malformed_id,
};
use crate::error::{EditCode, Error, Result};
use crate::jsv::V;
use crate::objects::{Found, Head, Stored};
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

fn table_gone() -> Error {
    Error::edit(
        EditCode::UnavailableTarget,
        "the table is no longer in the document",
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
        "mathEquation" | "math" => "holds an equation".into(),
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

/// A table or content control embedded in a paragraph's slot.
#[derive(Clone, Debug, PartialEq)]
pub(super) struct Block {
    kind: String,
    payload: BTreeMap<String, Any>,
    /// Its index in the story.
    at: u32,
}

/// A paragraph as the agent's commands and read see it (`docxParagraphs`,
/// with the tables and content controls its slot opens with set apart).
#[derive(Clone, Debug, PartialEq)]
pub(super) struct Para {
    para_id: String,
    /// Where its own content starts: after the page and column breaks,
    /// tables and content controls its slot opens with.
    start: u32,
    /// Where replace_text's paragraph starts (mod.rs `Paragraph`): after
    /// the page and column breaks it opens with only.
    node_start: u32,
    /// No tracked change in its text and no object but those breaks:
    /// replace_text edits it.
    plain: bool,
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
    /// The tables and content controls in its slot, in order: those it
    /// opens with, then any in its text, which lock it.
    blocks: Vec<Block>,
    /// Its mark's properties but its bookmarks.
    properties: BTreeMap<String, Any>,
    /// Its mark's item, where a read of the next paragraph starts.
    mark_item: ID,
    /// The tracked insertion and deletion open right after its mark, which
    /// that read starts with.
    after: Open,
}

/// The tracked insertion and deletion (`ins`, `del`) open at a story place.
type Open = [Option<Any>; 2];

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
        self.blocks.iter().any(|block| block.kind == "blockSdt")
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

    /// The tables in its slot, in order.
    fn tables(&self) -> impl Iterator<Item = &Block> {
        self.blocks.iter().filter(|block| block.kind == "table")
    }

    /// Its places moved by `delta` (an edit before it).
    fn shift(&mut self, delta: i64) {
        let moved = |at: u32| (i64::from(at) + delta) as u32;
        self.start = moved(self.start);
        self.node_start = moved(self.node_start);
        self.mark = moved(self.mark);
        for block in &mut self.blocks {
            block.at = moved(block.at);
        }
    }
}

/// Story `story`'s text.
fn story_text<T: ReadTxn>(txn: &T, story: &str) -> Option<TextRef> {
    txn.get_map(STORIES)?
        .get(txn, story)?
        .cast::<TextRef>()
        .ok()
}

/// A map embed's entries but `skipped`, as the segment read gives them.
fn entries<T: ReadTxn>(map: &yrs::MapRef, txn: &T, skipped: &[&str]) -> BTreeMap<String, Any> {
    map.iter(txn)
        .filter_map(|(key, value)| match value {
            Out::Any(value) if !skipped.contains(&key) => Some((key.to_owned(), value)),
            _ => None,
        })
        .collect()
}

/// A map embed's string entry `key`.
fn entry_string<T: ReadTxn>(map: &yrs::MapRef, txn: &T, key: &str) -> Option<String> {
    match map.get(txn, key) {
        Some(Out::Any(Any::String(value))) => Some(value.to_string()),
        _ => None,
    }
}

/// Paragraphs read item by item, as `docx_edit::EditingDoc::story_segments`
/// gives them, from `next`: the first item of a paragraph slot at story
/// index `at`, with the tracked changes `open` there; through the mark that
/// ends at index `end`, or the story's end.
fn read_paras<T: ReadTxn>(
    txn: &T,
    mut next: Option<ItemPtr>,
    mut at: u32,
    mut open: Open,
    end: Option<u32>,
) -> Vec<Para> {
    let mut paras = Vec::new();
    let (mut start, mut node_start) = (at, at);
    let mut plain = true;
    let mut text = String::new();
    let mut locked: Option<String> = None;
    let mut tracked = false;
    let mut blocks = Vec::new();
    while let Some(item) = next {
        next = item.right();
        if item.is_deleted() {
            continue;
        }
        let mut read_text = |part: &str, len: u32, open: &Open| {
            let deleted = truthy(open[1].as_ref());
            if deleted || truthy(open[0].as_ref()) {
                locked.get_or_insert_with(|| "holds tracked changes".into());
                tracked = true;
                plain = false;
            }
            if !deleted {
                text.push_str(part);
            }
            len
        };
        match item.content() {
            ItemContent::Format(key, value) => {
                if let Some(slot) = ["ins", "del"].iter().position(|name| **key == **name) {
                    open[slot] = (**value != Any::Null).then(|| (**value).clone());
                }
            }
            ItemContent::String(part) => at += read_text(part.as_str(), item.len(), &open),
            // The segment read gives a string embed as text.
            ItemContent::Embed(Any::String(part)) => at += read_text(part, len16(part), &open),
            ItemContent::Type(_) | ItemContent::Embed(_) => {
                let map = match item.content().get_first() {
                    Some(Out::YMap(map)) => Some(map),
                    _ => None,
                };
                let kind = map
                    .as_ref()
                    .and_then(|map| entry_string(map, txn, "_kind"))
                    .unwrap_or_default();
                if let Some(map) = map.as_ref().filter(|_| kind == "pilcrow") {
                    let mut properties = entries(map, txn, &["_kind", "paraId"]);
                    let mark_tracked = open.iter().any(|value| truthy(value.as_ref()))
                        || TRACKED_KEYS.iter().any(|key| properties.contains_key(*key));
                    if mark_tracked {
                        locked.get_or_insert_with(|| "holds tracked changes".into());
                    }
                    // Its bookmarks stay where docx-edit puts them: a new
                    // paragraph copies none, and a replacement before one
                    // moves it, which the index would not follow (review L1).
                    properties.remove("bookmarks");
                    paras.push(Para {
                        para_id: entry_string(map, txn, "paraId").unwrap_or_default(),
                        start,
                        node_start,
                        plain: std::mem::replace(&mut plain, true),
                        mark: at,
                        text: std::mem::take(&mut text),
                        locked: locked.take(),
                        tracked: std::mem::take(&mut tracked) || mark_tracked,
                        blocks: std::mem::take(&mut blocks),
                        properties,
                        mark_item: *item.id(),
                        after: open.clone(),
                    });
                    at += 1;
                    start = at;
                    node_start = at;
                    if end == Some(at) {
                        break;
                    }
                    continue;
                }
                let payload = map
                    .as_ref()
                    .map(|map| entries(map, txn, &["_kind"]))
                    .unwrap_or_default();
                let leading = at == start;
                let breaks =
                    at == node_start && matches!(kind.as_str(), "pageBreak" | "columnBreak");
                let block = matches!(kind.as_str(), "table" | "blockSdt");
                let embedded = Block {
                    kind: kind.clone(),
                    payload,
                    at,
                };
                at += 1;
                if breaks {
                    node_start = at;
                } else {
                    plain = false;
                }
                if leading && (block || matches!(kind.as_str(), "pageBreak" | "columnBreak")) {
                    start = at;
                } else {
                    locked.get_or_insert_with(|| locked_by(&kind));
                }
                if block {
                    blocks.push(embedded);
                }
            }
            _ => {}
        }
    }
    paras
}

/// The item of the first live paragraph mark from `next` on.
fn next_mark<T: ReadTxn>(txn: &T, mut next: Option<ItemPtr>) -> Option<ID> {
    while let Some(item) = next {
        next = item.right();
        if !item.is_deleted()
            && matches!(item.content(), ItemContent::Type(_))
            && let Some(Out::YMap(map)) = item.content().get_first()
            && entry_string(&map, txn, "_kind").as_deref() == Some("pilcrow")
        {
            return Some(*item.id());
        }
    }
    None
}

/// `story`'s paragraphs in order.
fn paras(doc: &EditingDoc, story: &str) -> Result<Vec<Para>> {
    let txn = doc.yrs_doc().transact();
    let text = story_text(&txn, story)
        .ok_or_else(|| Error::edit(EditCode::UnavailableTarget, "unknown DOCX story"))?;
    Ok(read_paras(
        &txn,
        AsRef::<Branch>::as_ref(&text).start(),
        0,
        Open::default(),
        None,
    ))
}

/// Where a walk over paragraph `index` of `paras` starts: after the mark
/// before it, the story's start for the first.
fn seek(paras: &[Para], index: usize) -> touched::Seek {
    index
        .checked_sub(1)
        .map(|previous| (paras[previous].mark_item, paras[previous].mark + 1))
}

/// Story `story` whole, with the cells' stories of each table in it.
fn story_tree(doc: &EditingDoc, story: &str) -> Result<Touched> {
    let mut sum = touched::story(doc.yrs_doc(), STORIES, story);
    for para in paras(doc, story)? {
        for table in para.tables() {
            for cell in table_cells(&table.payload).iter().flatten() {
                sum += story_tree(doc, cell)?;
            }
        }
    }
    Ok(sum)
}

/// One call's paragraphs by story (review r2 L9): a story is read once and
/// kept in step with this session's edits. A replacement shifts the
/// paragraphs after it and a format moves none; any other command reads the
/// paragraphs it changed again ([`reread`]). Any other change to the
/// document (an Undo) is a commit that advances the engine's document epoch
/// and drops the index, so the next lookup reads the story again.
#[derive(Default)]
pub(super) struct ParaIndex {
    stories: HashMap<String, Arc<Vec<Para>>>,
    /// The document epoch `stories` was last in step with.
    at: u64,
}

/// The index after a replacement in paragraph `para_id` of `story` left it
/// reading `text`: its mark and every later paragraph move by the change in
/// length.
pub(super) fn replaced_in_index(
    stories: &mut HashMap<String, Arc<Vec<Para>>>,
    story: &str,
    para_id: &str,
    text: String,
) {
    let Some(paras) = stories.get_mut(story) else {
        return;
    };
    let paras = Arc::make_mut(paras);
    let Some(index) = paras.iter().position(|para| para.para_id == para_id) else {
        // Not where the index has it: read the story again.
        stories.remove(story);
        return;
    };
    let delta = i64::from(len16(&text)) - i64::from(len16(&paras[index].text));
    paras[index].text = text;
    paras[index].mark = (i64::from(paras[index].mark) + delta) as u32;
    for para in &mut paras[index + 1..] {
        para.shift(delta);
    }
}

/// The index after an edit changed paragraphs `first` through `last` of
/// `story` (as the index held them): they are read again from the mark
/// before them, as far as `last`'s mark moved by the story's change in
/// length, and the paragraphs after move by that change. The story drops
/// out of the index, to be read whole again, where that read does not end
/// at a mark there with the same tracked changes open after it, or the
/// next mark the story holds is not the one the index holds next (review
/// L1).
fn reread(
    doc: &EditingDoc,
    stories: &mut HashMap<String, Arc<Vec<Para>>>,
    story: &str,
    first: usize,
    last: usize,
) {
    let read = || -> Option<(Vec<Para>, i64)> {
        let old = stories.get(story)?;
        let delta = i64::from(doc.story_len(story).ok()?) - i64::from(old.last()?.mark + 1);
        let end = (i64::from(old.get(last)?.mark + 1) + delta) as u32;
        let txn = doc.yrs_doc().transact();
        let text = story_text(&txn, story)?;
        let (next, at, open) = match first.checked_sub(1) {
            None => (AsRef::<Branch>::as_ref(&text).start(), 0, Open::default()),
            Some(previous) => {
                let previous = &old[previous];
                let mark = txn.store().get_item(&previous.mark_item)?;
                if mark.is_deleted() {
                    return None;
                }
                (mark.right(), previous.mark + 1, previous.after.clone())
            }
        };
        let read = read_paras(&txn, next, at, open, Some(end));
        let ends = read.last()?;
        let after = next_mark(&txn, txn.store().get_item(&ends.mark_item)?.right());
        (ends.mark + 1 == end
            && ends.after == old[last].after
            && after == old.get(last + 1).map(|para| para.mark_item))
        .then_some((read, delta))
    };
    let Some((read, delta)) = read() else {
        stories.remove(story);
        return;
    };
    let Some(paras) = stories.get_mut(story) else {
        return;
    };
    let paras = Arc::make_mut(paras);
    let added = read.len();
    paras.splice(first..=last, read);
    for para in &mut paras[first + added..] {
        para.shift(delta);
    }
}

/// Why a cell story keeps its table from a delete, as delete_paragraph
/// refuses a paragraph: a table or content control (review H2: expected
/// names only the cells' text), a picture, equation, field or note
/// reference, a tracked change (review r2 L2).
fn cell_refusal(paras: &[Para]) -> Option<String> {
    for para in paras {
        if !para.blocks.is_empty() {
            return Some("holds a table or content control".into());
        }
        if para.locked.is_some() {
            return para.locked.clone();
        }
    }
    None
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

impl Placement {
    /// The body paragraph the new one is split from (the one filled).
    fn split(&self) -> usize {
        match self.place {
            Place::After(index) | Place::Before(index) => index,
            Place::Fill => 0,
        }
    }
}

/// A table found by a cell paragraph.
struct Table {
    story: String,
    /// Its ordinal among the story's tables.
    ordinal: u32,
    /// Its cell stories row by row.
    cells: Vec<Vec<String>>,
    /// The paragraph whose slot holds it, by its place in the story.
    slot: usize,
    /// Its index in the story.
    at: u32,
}

/// One cell's mark on the paragraphs a read lists.
#[derive(Clone)]
struct CellMark {
    table: Option<String>,
    row: usize,
    column: usize,
}

/// The read's walk: paragraphs counted in order, the page's listed.
struct Walk<'a> {
    session: &'a DocxSession,
    /// The page's media and parts, each read once.
    objects: ObjectParts<'a>,
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
            for block in &para.blocks {
                if block.kind == "table" {
                    self.table(&table_cells(&block.payload), control)?;
                } else if let Some(Any::String(child)) = block.payload.get("story") {
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
                // An object locks its paragraph.
                if para.locked.is_some() {
                    let (objects, more) =
                        self.session.object_lines(story, &para, &mut self.objects)?;
                    if !objects.is_empty() {
                        entry["objects"] = Value::Array(objects);
                    }
                    if more > 0 {
                        entry["moreObjects"] = Value::from(more);
                    }
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
        // Each read starts after the mark before the paragraph it measures.
        let range = |story: &str, paras: &[Para], index: usize, start: u32, end: u32, kept| {
            let from = seek(paras, index);
            #[cfg(test)]
            let from = from.filter(|_| !tests::FROM_START.get());
            touched::story_range_from(doc.yrs_doc(), STORIES, story, from, start, end, kept)
        };
        let measured = || -> Result<Touched> {
            Ok(match command {
                Command::ReplaceText { target_id, .. } => {
                    // The lookup apply uses: a header's, footer's or note's
                    // paragraph too (review r2 H1).
                    let (story, paragraph) = self.paragraph(target_id)?;
                    let (span, _) = check_replacement(&paragraph.text, command, "DOCX")?;
                    let paras = self.story_paras(&story)?;
                    let index = paras
                        .iter()
                        .position(|para| para.para_id == paragraph.para_id)
                        .ok_or_else(gone)?;
                    let start = (paragraph.start + span.start) as u32;
                    let end = (paragraph.start + span.end) as u32;
                    range(&story, &paras, index, start, end, Kept::Content)
                }
                Command::FormatText {
                    target_id, text, ..
                } => {
                    let (story, index, paras) = self.find(target_id)?;
                    let para = &paras[index];
                    let at = para.text.find(text.as_str()).ok_or_else(gone)?;
                    let start = para.start + len16(&para.text[..at]);
                    range(
                        &story,
                        &paras,
                        index,
                        start,
                        start + len16(text),
                        Kept::Formatting,
                    )
                }
                Command::DeleteParagraph { target_id, .. } => {
                    let (story, index, paras) = self.find(target_id)?;
                    let para = &paras[index];
                    range(
                        &story,
                        &paras,
                        index,
                        para.start,
                        para.mark + 1,
                        Kept::Content,
                    )
                }
                Command::SetAlignment { target_id, .. } => {
                    let (_, index, paras) = self.find(target_id)?;
                    touched::mark_item_entries(
                        doc.yrs_doc(),
                        &paras[index].mark_item,
                        &["alignment"],
                    )
                }
                Command::SetList { ids, .. } => {
                    // Each story read once: a list may name 200 paragraphs.
                    let mut read: HashMap<Cow<str>, Arc<Vec<Para>>> = HashMap::new();
                    let mut sum = Touched::default();
                    for id in ids {
                        let (story, para_id) = self.split(id, false)?;
                        if !read.contains_key(&story) {
                            read.insert(story.clone(), self.story_paras(&story)?);
                        }
                        let para = read[&story]
                            .iter()
                            .find(|para| para.para_id == para_id)
                            .ok_or_else(gone)?;
                        sum +=
                            touched::mark_item_entries(doc.yrs_doc(), &para.mark_item, &LIST_KEYS);
                    }
                    sum
                }
                Command::DeleteTable { id, .. } => {
                    let table = self.table(id)?;
                    let paras = self.story_paras(&table.story)?;
                    let mut sum = range(
                        &table.story,
                        &paras,
                        table.slot,
                        table.at,
                        table.at + 1,
                        Kept::Content,
                    );
                    for cell in table.cells.iter().flatten() {
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
                text,
                format,
            } => self.format_text(target_id, text, format).map(none),
            Command::SetAlignment {
                target_id,
                alignment,
            } => self.set_alignment(target_id, *alignment).map(none),
            Command::SetShapeText { .. }
            | Command::AddSlide { .. }
            | Command::DeleteSlide { .. }
            | Command::MoveSlide { .. }
            | Command::AddTextBox { .. } => Err(Error::edit(
                EditCode::UnsupportedOperation,
                "this command edits PPTX files",
            )),
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
        let minted = format!("{}:", self.engine.doc().client_id());
        let body = self.story_paras(BODY)?;
        let mut removed = Vec::new();
        for (index, para) in body.iter().enumerate() {
            let Some(next) = body.get(index + 1) else {
                continue;
            };
            let inserted = match para.blocks.last() {
                Some(block) if block.kind == "table" => {
                    match table_cells(&block.payload)
                        .first()
                        .and_then(|row| row.first())
                    {
                        Some(first) => self
                            .story_paras(first)?
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
                removed.push((index, para.para_id.clone()));
            }
        }
        drop(body);
        // From the end, so each index still names its paragraph; the one
        // after each takes its table.
        for (index, para_id) in removed.into_iter().rev() {
            let before = self.index_mark();
            self.engine
                .doc()
                .remove_paragraph(&ctx(), BODY, &para_id)
                .map_err(op)?;
            self.reread(before, BODY, index, index + 1);
        }
        self.projected = None;
        Ok(())
    }

    /// `count` paragraphs from `start`: the body's, each table's cells
    /// listed before the paragraph that follows it and marked with the
    /// table, row and column, then each header's, footer's and note's,
    /// marked with its part. A paragraph names its style when not the
    /// default; the first page also lists the file's shown paragraph
    /// styles, at most [`MAX_STYLES`]. A paragraph holding pictures, charts,
    /// equations or drawn shapes lists them in `objects`: a picture's and a
    /// chart's read lines ([`crate::objects`]), an equation's and a shape's
    /// text. `ids`: those pictures and charts ([`Self::read_ids`]).
    pub(crate) fn inspect_scoped(&mut self, query: &InspectQuery) -> Result<Value> {
        if let Some(ids) = &query.ids {
            return self.read_ids(ids);
        }
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
            session: self,
            objects: ObjectParts::default(),
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

    /// `story`'s paragraphs from the call's index ([`ParaIndex`]).
    pub(super) fn story_paras(&self, story: &str) -> Result<Arc<Vec<Para>>> {
        let doc = self.engine.doc();
        let mut index = self.index.borrow_mut();
        let now = self.engine.doc_epoch();
        if index.at != now {
            index.stories.clear();
            index.at = now;
        }
        if let Some(paras) = index.stories.get(story) {
            return Ok(Arc::clone(paras));
        }
        let read = Arc::new(paras(doc, story)?);
        index.stories.insert(story.to_owned(), Arc::clone(&read));
        Ok(read)
    }

    /// The document epoch before an edit ([`Self::keep_index`]).
    pub(super) fn index_mark(&self) -> u64 {
        self.engine.doc_epoch()
    }

    /// After an edit changed paragraphs `first` through `last` of `story`
    /// (the index's, `before` the edit): the index reads them again
    /// ([`reread`]).
    fn reread(&self, before: u64, story: &str, first: usize, last: usize) {
        let doc = self.engine.doc();
        self.keep_index(before, |stories| reread(doc, stories, story, first, last));
    }

    /// After a text edit the index can follow: `update` brings its stories
    /// in step, when they were in step before the edit (`before`).
    pub(super) fn keep_index(
        &self,
        before: u64,
        update: impl FnOnce(&mut HashMap<String, Arc<Vec<Para>>>),
    ) {
        let mut index = self.index.borrow_mut();
        if index.at == before {
            update(&mut index.stories);
            index.at = self.engine.doc_epoch();
        }
    }

    /// `replace_text`'s paragraph (mod.rs `paragraph`) from the index, in
    /// any story.
    pub(super) fn indexed_paragraph(&self, story: &str, para_id: &str) -> Result<Paragraph> {
        let paras = self.story_paras(story)?;
        let para = paras
            .iter()
            .find(|para| para.para_id == para_id)
            .ok_or_else(gone)?;
        Ok(Paragraph {
            para_id: para.para_id.clone(),
            start: para.node_start as usize,
            length: (para.mark - para.node_start) as usize,
            text: para.text.clone(),
            plain: para.plain,
        })
    }

    /// `story` and paraId of a target id (`<story>:paragraph:<paraId>`).
    /// A short paragraph id, the paraId alone, stands for the one paragraph
    /// of its story it names (Epo 2026-10-24, Q222): a file without
    /// `w14:paraId` seeds paraIds that name their story (`body:p9`, read as
    /// `body:paragraph:body:p9`), and another paraId alone names a body
    /// paragraph. Where a command takes a table (`table`), a table's short id
    /// `<story>:t<n>`, the slot its cells' stories name with or without
    /// paraIds (`body:t0` for `body:t0:r0c0`), stands for the read's id of
    /// the one table of that story it names (Epo 2026-10-23). A short id
    /// naming several is refused naming the read's ids; one naming no table
    /// of a story is gone; any other id is malformed (Q223), its refusal
    /// naming the read's id for what it names, else an id the read shows.
    pub(super) fn split<'a>(
        &self,
        id: &'a str,
        table: bool,
    ) -> Result<(Cow<'a, str>, Cow<'a, str>)> {
        if let Some((story, para_id)) = id
            .rsplit_once(":paragraph:")
            .filter(|(story, _)| !story.is_empty())
        {
            return Ok((story.into(), para_id.into()));
        }
        let story = id.rsplit_once(':').map_or(BODY, |(story, _)| story);
        let paras = self.story_paras(story).ok();
        let named = paras
            .iter()
            .flat_map(|paras| paras.iter())
            .filter(|para| para.para_id == id)
            .count();
        // A story's first paragraph, the read's id for it as (story, paraId).
        let first = |story: &str| {
            let paras = self.story_paras(story).ok()?;
            Some((story.to_owned(), paras.first()?.para_id.clone()))
        };
        // The tables of `story` a short table id names: those with a cell
        // in its slot, each by its first cell's first paragraph as the read
        // names a table.
        let short_table = table_story(&format!("{id}:r0c0")).is_some();
        let mut tables: Vec<(String, String)> = Vec::new();
        if short_table {
            let slot = format!("{id}:");
            for found in paras
                .iter()
                .flat_map(|paras| paras.iter().flat_map(Para::tables))
            {
                let cells = table_cells(&found.payload);
                if cells.iter().flatten().any(|cell| cell.starts_with(&slot))
                    && let Some(cell) = cells.first().and_then(|row| row.first())
                {
                    tables.extend(first(cell));
                }
            }
        }
        if table && tables.len() == 1 {
            let (cell, para_id) = tables.remove(0);
            return Ok((cell.into(), para_id.into()));
        }
        if named == 1 {
            return Ok((story.into(), id.into()));
        }
        let read = |(story, para_id): &(String, String)| {
            format!("{:?}", format!("{story}:paragraph:{para_id}"))
        };
        let reads = |ids: &[(String, String)]| ids.iter().map(read).collect::<Vec<_>>().join(", ");
        if named > 1 || (table && tables.len() > 1) {
            let (count, kind, ids) = if named > 1 {
                (
                    named,
                    "paragraphs",
                    vec![(story.to_owned(), id.to_owned()); named],
                )
            } else {
                (tables.len(), "tables", tables)
            };
            return Err(Error::edit(
                EditCode::UnavailableTarget,
                format!(
                    "{id:?} names {count} {kind}, which the read shows as {}: copy each id whole from inspect_office_document, reading the file first",
                    reads(&ids)
                ),
            ));
        }
        if table && short_table && paras.is_some() {
            return Err(table_gone());
        }
        let shown = match &tables[..] {
            [] => match first(id) {
                Some(para) => format!("the read shows its first paragraph as {}", read(&para)),
                None => self.read_example(table),
            },
            [one] => format!("the read shows that table as {}", read(one)),
            several => format!(
                "it names {} tables, which the read shows as {}",
                several.len(),
                reads(several)
            ),
        };
        let kind = if table {
            "DOCX table"
        } else {
            "DOCX paragraph"
        };
        Err(malformed_id(id, kind, &shown))
    }

    /// Ids the read shows, for a malformed id's refusal: the body's first
    /// paragraph's, with `table` its first table's short id and the read's.
    fn read_example(&self, table: bool) -> String {
        let body = self.story_paras(BODY).ok();
        let mut paras = body.iter().flat_map(|paras| paras.iter());
        if !table {
            return paras.next().map_or_else(
                || "the document body holds no paragraph".to_owned(),
                |para| {
                    let id = format!("{BODY}:paragraph:{}", para.para_id);
                    format!("paragraph ids read like {id:?}, the body's first")
                },
            );
        }
        let read = paras.flat_map(Para::tables).find_map(|found| {
            let cell = table_cells(&found.payload).first()?.first()?.clone();
            let para_id = self.story_paras(&cell).ok()?.first()?.para_id.clone();
            let slot = cell.rsplit_once(':')?.0.to_owned();
            Some((slot, format!("{cell}:paragraph:{para_id}")))
        });
        read.map_or_else(
            || "the document body holds no table".to_owned(),
            |(slot, id)| format!("table ids read like {slot:?} or {id:?}, the body's first table"),
        )
    }

    /// The paragraph `id` names, its story and its story's paragraphs (body
    /// stories only, as every command but `replace_text` edits).
    fn find(&self, id: &str) -> Result<(String, usize, Arc<Vec<Para>>)> {
        let (story, para_id) = self.split(id, false)?;
        self.find_split(story, &para_id)
    }

    /// [`Self::find`] of a split id.
    fn find_split(
        &self,
        story: Cow<'_, str>,
        para_id: &str,
    ) -> Result<(String, usize, Arc<Vec<Para>>)> {
        in_body(&story)?;
        let paras = self.story_paras(&story)?;
        let index = paras
            .iter()
            .position(|para| para.para_id == para_id)
            .ok_or_else(gone)?;
        Ok((story.into_owned(), index, paras))
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
            let (story, para_id) = self.split(id, false)?;
            if story != BODY {
                return Err(invalid(refusal));
            }
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
        let body = self.story_paras(BODY)?;
        let before = self.index_mark();
        let placement = self.place(&body, after_id, before_id, PARAGRAPH_ANCHOR)?;
        let split = placement.split();
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
        drop(body);
        if !text.is_empty() {
            self.engine
                .doc()
                .insert_text(&ctx(), Position::new(BODY, at), text, FormatPolicy::Plain)
                .map_err(op)?;
        }
        self.reread(before, BODY, split, split);
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
        let para_id = para.para_id.clone();
        drop(paras);
        let before = self.index_mark();
        // The paragraphs around it keep their own properties and ids; the
        // next takes the tables and breaks its slot opens with.
        self.engine
            .doc()
            .remove_paragraph(&ctx(), &story, &para_id)
            .map_err(op)?;
        self.reread(before, &story, index, index + 1);
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
        let mut stories: HashMap<Cow<str>, Arc<Vec<Para>>> = HashMap::new();
        let mut targets = Vec::with_capacity(ids.len());
        for id in ids {
            let (story, para_id) = self.split(id, false)?;
            in_body(&story)?;
            if !stories.contains_key(&story) {
                stories.insert(story.clone(), self.story_paras(&story)?);
            }
            let index = stories[&story]
                .iter()
                .position(|para| para.para_id == para_id)
                .ok_or_else(gone)?;
            let para = &stories[&story][index];
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
            targets.push((story, index, para.mark, ilvl));
        }
        drop(stories);
        let before = self.index_mark();
        for &(ref story, _, at, ilvl) in &targets {
            self.engine
                .doc()
                .set_paragraph_attrs(
                    &ctx(),
                    &ParaSelector::Range(StoryRange::new(&**story, at, at)),
                    &ParaAttrDelta {
                        other: list_attrs(list, num_id, ilvl).into_iter().collect(),
                        ..ParaAttrDelta::default()
                    },
                )
                .map_err(op)?;
        }
        let doc = self.engine.doc();
        self.keep_index(before, |stories| {
            for &(ref story, slot, _, _) in &targets {
                reread(doc, stories, story, slot, slot);
            }
        });
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
        let body = self.story_paras(BODY)?;
        let before = self.index_mark();
        let placement = self.place(&body, after_id, None, TABLE_ANCHOR)?;
        let split = placement.split();
        let (_, at, _) = self.new_paragraph(&body, placement, BTreeMap::new())?;
        drop(body);
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
        self.reread(before, BODY, split, split);
        Ok(())
    }

    /// The [`Table`] holding cell paragraph `id`, or the one its short id
    /// names ([`Self::split`]); a cell's story goes with its table.
    fn table(&self, id: &str) -> Result<Table> {
        let (story, para_id) = self.split(id, true)?;
        if table_story(&story).is_some() && self.story_paras(&story).is_err() {
            return Err(table_gone());
        }
        let (cell, _, _) = self.find_split(story, &para_id)?;
        let story = table_story(&cell).ok_or_else(|| {
            invalid("id must be a paragraph of a table cell, as inspect_office_document marks it")
        })?;
        let paras = self.story_paras(story)?;
        let mut ordinal = 0;
        for (slot, para) in paras.iter().enumerate() {
            for table in para.tables() {
                let cells = table_cells(&table.payload);
                if cells.iter().flatten().any(|story| *story == cell) {
                    return Ok(Table {
                        story: story.to_owned(),
                        ordinal,
                        cells,
                        slot,
                        at: table.at,
                    });
                }
                ordinal += 1;
            }
        }
        Err(table_gone())
    }

    /// The table holding paragraph `id`, refused unless its cells read
    /// `expected` and hold nothing delete_paragraph refuses
    /// ([`cell_refusal`]).
    fn delete_table(&mut self, id: &str, expected: &[Vec<String>]) -> Result<()> {
        let table = self.table(id)?;
        for cell in table.cells.iter().flatten() {
            if let Some(reason) = cell_refusal(&self.story_paras(cell)?) {
                return Err(Error::edit(
                    EditCode::UnavailableTarget,
                    format!("a cell of the table {reason}; the table cannot be deleted"),
                ));
            }
        }
        let grid = table
            .cells
            .iter()
            .map(|row| {
                row.iter()
                    .map(|cell| {
                        Ok(self
                            .story_paras(cell)?
                            .iter()
                            .map(|para| para.text.as_str())
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
        let doc = self.engine.doc();
        let before = self.index_mark();
        doc.delete_table(
            &ctx(),
            &TableLocator::new(table.story.as_str(), table.ordinal),
        )
        .map_err(op)?;
        self.keep_index(before, |stories| {
            // Its cells' stories go with it.
            stories.retain(|story, _| doc.story_len(story).is_ok());
            reread(doc, stories, &table.story, table.slot, table.slot);
        });
        Ok(())
    }

    fn format_text(&mut self, id: &str, text: &str, format: &TextFormat) -> Result<()> {
        let delta = inline_delta(format)?;
        let (story, index, paras) = self.find(id)?;
        let para = &paras[index];
        if let Some(reason) = para.refused() {
            return Err(Error::edit(
                EditCode::UnavailableTarget,
                format!("the paragraph {reason}"),
            ));
        }
        if text.is_empty() {
            return Err(invalid("text names the text to format"));
        }
        let found: Vec<usize> = para.text.match_indices(text).map(|(at, _)| at).collect();
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
        let before = self.index_mark();
        self.engine
            .doc()
            .format_range(
                &ctx(),
                StoryRange::new(&story, start, start + len16(text)),
                &delta,
            )
            .map_err(op)?;
        // A format moves no text: the index stays in step.
        self.keep_index(before, |_| {});
        Ok(())
    }

    fn set_alignment(&mut self, id: &str, alignment: Alignment) -> Result<()> {
        let (story, index, paras) = self.find(id)?;
        if paras[index].tracked {
            return Err(tracked_refusal());
        }
        let mark = paras[index].mark;
        drop(paras);
        let before = self.index_mark();
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
                &ParaSelector::Range(StoryRange::new(&story, mark, mark)),
                &ParaAttrDelta {
                    alignment: Patch::Set(value.to_owned()),
                    ..ParaAttrDelta::default()
                },
            )
            .map_err(op)?;
        self.reread(before, &story, index, index);
        Ok(())
    }
}

/// The mark keys [`list_attrs`] writes.
const LIST_KEYS: [&str; 4] = ["numPr", "listIsBullet", "listNumFmt", "listMarker"];

/// A picture, chart, equation or drawn shape embedded in a paragraph.
struct Object {
    id: ID,
    kind: String,
    payload: BTreeMap<String, Any>,
}

fn any_str<'p>(payload: &'p BTreeMap<String, Any>, key: &str) -> Option<&'p str> {
    match payload.get(key) {
        Some(Any::String(text)) => Some(text),
        _ => None,
    }
}

/// A picture's or chart's id: its story and Yjs item, `<story>:picture:`
/// or `<story>:chart:<client>.<clock>`.
fn object_id(story: &str, object: &Object) -> Option<String> {
    let kind = match object.kind.as_str() {
        "image" => "picture",
        "chart" => "chart",
        _ => return None,
    };
    Some(format!(
        "{story}:{kind}:{}.{}",
        object.id.client, object.id.clock
    ))
}

/// The Yjs item an object id names, when the id has the form of one.
fn parse_object_id(id: &str) -> Option<(&'static str, ID)> {
    let (kind, rest) = [("image", ":picture:"), ("chart", ":chart:")]
        .into_iter()
        .find_map(|(kind, marker)| Some((kind, id.rsplit_once(marker)?.1)))?;
    let (client, clock) = rest.split_once('.')?;
    let client = client
        .parse::<u64>()
        .ok()
        .filter(|client| *client < 1 << 53)?;
    Some((kind, ID::new(ClientID::new(client), clock.parse().ok()?)))
}

/// One read's media and parts: each picture's header and each chart part
/// read once however many objects name it (review r1 M3).
#[derive(Default)]
pub(crate) struct ObjectParts<'a> {
    parts: crate::objects::Parts,
    /// The package's media data URLs by part, gathered on first use.
    media: Option<HashMap<String, Rc<str>>>,
    /// The base package, opened on the first chart part read (`None` inside
    /// when it does not open), so a read opens it once (review r3 L1).
    package: Option<Option<ooxml_opc::Package<'a>>>,
}

impl<'a> ObjectParts<'a> {
    /// Chart part `path` of `base`, read once a read.
    fn chart(&mut self, base: &'a [u8], path: &str) -> Option<Rc<crate::objects::Chart>> {
        let package = &mut self.package;
        self.parts.chart(path, || {
            package
                .get_or_insert_with(|| ooxml_opc::Package::open(base).ok())
                .as_mut()?
                .read(path, 32 << 20)
                .ok()
                .flatten()
        })
    }
}

impl DocxSession {
    /// The objects in a paragraph, in order: the embeds between its mark
    /// and the one before.
    fn objects(&self, para: &Para) -> Vec<Object> {
        let txn = self.engine.doc().yrs_doc().transact();
        let mut objects = Vec::new();
        let mut next = txn
            .store()
            .get_item(&para.mark_item)
            .and_then(|mark| mark.left());
        while let Some(item) = next {
            next = item.left();
            if item.is_deleted() || !matches!(item.content(), ItemContent::Type(_)) {
                continue;
            }
            let Some(Out::YMap(map)) = item.content().get_first() else {
                continue;
            };
            let kind = entry_string(&map, &txn, "_kind").unwrap_or_default();
            if kind == "pilcrow" {
                break;
            }
            if matches!(kind.as_str(), "image" | "chart" | "math" | "shape") {
                objects.push(Object {
                    id: *item.id(),
                    kind,
                    payload: entries(&map, &txn, &["_kind"]),
                });
            }
        }
        objects.reverse();
        objects
    }

    /// A picture's header and stored bytes: its package part's
    /// (`media:<part>`), its own data URL, or none for a linked one. Only the
    /// header is read here; the bytes are decoded when a by-id read sends
    /// them.
    fn picture_head(&self, src: &str, cache: &mut ObjectParts<'_>) -> (Head, Stored<'static>) {
        if src.is_empty() {
            return (Head::linked(), Stored::Nothing);
        }
        let Some(part) = src.strip_prefix("media:") else {
            let subtype = src
                .strip_prefix("data:image/")
                .and_then(|rest| rest.split(';').next())
                .unwrap_or_default();
            let part = format!("inline.{}", subtype.trim_start_matches("x-"));
            return (
                Head::of_data_url(src, &part),
                Stored::DataUrl(Rc::from(src)),
            );
        };
        let media = cache.media.get_or_insert_with(|| {
            self.package()
                .get("media")
                .items()
                .iter()
                .filter_map(|pair| {
                    let pair = pair.arr()?;
                    Some((
                        pair.get(0).to_js_string(),
                        pair.get(1).get("dataUrl").as_str()?,
                    ))
                })
                .collect()
        });
        let url = media.get(part).cloned().unwrap_or_else(|| Rc::from(""));
        let head = cache.parts.head(part, || Head::of_data_url(&url, part));
        (head, Stored::DataUrl(url))
    }

    /// What a read finds of a picture or chart object.
    fn object_found<'s>(
        &'s self,
        story: &str,
        object: &Object,
        cache: &mut ObjectParts<'s>,
    ) -> Option<Found<'static>> {
        let id = object_id(story, object)?;
        if object.kind == "chart" {
            let chart =
                any_str(&object.payload, "path").and_then(|path| cache.chart(&self.base, path));
            return Some(Found::Chart(chart));
        }
        let (head, stored) =
            self.picture_head(any_str(&object.payload, "src").unwrap_or_default(), cache);
        let side = |key: &str| match object.payload.get(key) {
            Some(Any::Number(value)) if value.is_finite() => *value,
            _ => 0.0,
        };
        Some(Found::Picture(crate::objects::Picture {
            id: Some(id),
            alt: any_str(&object.payload, "alt").map(str::to_owned),
            head,
            stored,
            crop: [
                side("cropLeft"),
                side("cropTop"),
                side("cropRight"),
                side("cropBottom"),
            ],
        }))
    }

    /// A paragraph's objects as its read lists them, at most
    /// [`crate::objects::MAX_LISTED`], and how many more it holds.
    fn object_lines<'s>(
        &'s self,
        story: &str,
        para: &Para,
        cache: &mut ObjectParts<'s>,
    ) -> Result<(Vec<Value>, usize)> {
        let objects = self.objects(para);
        let more = objects.len().saturating_sub(crate::objects::MAX_LISTED);
        let mut lines = Vec::new();
        for object in objects.into_iter().take(crate::objects::MAX_LISTED) {
            let text = |key: &str| {
                any_str(&object.payload, key)
                    .and_then(|text| crate::objects::shown(text, crate::objects::MAX_TEXT_CHARS))
            };
            let line = match object.kind.as_str() {
                "math" => text("plainText").map(|text| json!({ "kind": "equation", "text": text })),
                // A text box's paragraphs are its text body's content.
                "shape" => any_str(&object.payload, "shapeJson")
                    .and_then(|shape| V::parse(shape).ok())
                    .and_then(|shape| {
                        let text = super::authored_text(&shape.get("textBody").get("content"))
                            + &super::authored_text(&shape);
                        crate::objects::shown(&text, crate::objects::MAX_TEXT_CHARS)
                    })
                    .map(|text| json!({ "kind": "drawing", "text": text })),
                _ => match self.object_found(story, &object, cache) {
                    Some(Found::Picture(picture)) => Some(picture.line()?),
                    Some(Found::Chart(chart)) => Some(crate::objects::chart_line(
                        object_id(story, &object).as_deref(),
                        chart.as_deref(),
                    )),
                    None => None,
                },
            };
            lines.extend(line);
        }
        Ok((lines, more))
    }

    /// The by-id read ([`crate::objects::answer`]): picture and chart ids
    /// from a read; a paragraph id holding exactly one picture or chart
    /// stands for it (a clear repair, Q277); a gone or malformed id is named
    /// in the notice (Q381).
    fn read_ids(&self, ids: &[String]) -> Result<Value> {
        let (ids, past) = crate::objects::requested(ids)?;
        let mut cache = ObjectParts::default();
        let mut answered = Vec::new();
        let mut missing = Vec::new();
        for id in ids {
            let found = match parse_object_id(&id) {
                Some((kind, item)) => self.object_by_item(&id, kind, item, &mut cache),
                None => self.object_of_paragraph(&id, &mut cache),
            };
            match found {
                Ok(found) => answered.push((id, found)),
                Err(why) => missing.push((id, why)),
            }
        }
        crate::objects::answer(answered, &missing, &past)
    }

    fn object_by_item<'s>(
        &'s self,
        id: &str,
        kind: &str,
        item: ID,
        cache: &mut ObjectParts<'s>,
    ) -> std::result::Result<Found<'static>, String> {
        let gone = || "is no longer in the document".to_owned();
        let story = id
            .rsplit_once(if kind == "image" {
                ":picture:"
            } else {
                ":chart:"
            })
            .map_or("", |(story, _)| story);
        let object = {
            let txn = self.engine.doc().yrs_doc().transact();
            let found = txn
                .store()
                .get_item(&item)
                .filter(|found| *found.id() == item && !found.is_deleted());
            let map = found.and_then(|found| match found.content().get_first() {
                Some(Out::YMap(map)) => Some(map),
                _ => None,
            });
            match map {
                Some(map) if entry_string(&map, &txn, "_kind").as_deref() == Some(kind) => Object {
                    id: item,
                    kind: kind.to_owned(),
                    payload: entries(&map, &txn, &["_kind"]),
                },
                _ => return Err(gone()),
            }
        };
        self.object_found(story, &object, cache).ok_or_else(gone)
    }

    /// A paragraph id named in place of its one picture or chart.
    fn object_of_paragraph<'s>(
        &'s self,
        id: &str,
        cache: &mut ObjectParts<'s>,
    ) -> std::result::Result<Found<'static>, String> {
        let malformed = || crate::objects::malformed(&self.object_example());
        let Ok((story, index, paras)) = self.find(id) else {
            return Err(malformed());
        };
        let found: Vec<Found<'static>> = self
            .objects(&paras[index])
            .iter()
            .filter_map(|object| self.object_found(&story, object, cache))
            .collect();
        match <[Found; 1]>::try_from(found) {
            Ok([found]) => Ok(found),
            Err(found) if found.is_empty() => Err(malformed()),
            Err(found) => Err(format!(
                "is a paragraph holding {} pictures and charts; name each by the id its read lists",
                found.len()
            )),
        }
    }

    /// How picture and chart ids read: the document's first, else that it
    /// holds none.
    fn object_example(&self) -> String {
        for story in super::read::story_ids(self.engine.doc()) {
            let Ok(paras) = paras(self.engine.doc(), &story) else {
                continue;
            };
            for para in paras.iter().filter(|para| para.locked.is_some()) {
                if let Some(id) = self
                    .objects(para)
                    .iter()
                    .find_map(|object| object_id(&story, object))
                {
                    return format!("picture and chart ids read like {id:?}");
                }
            }
        }
        "the document holds no picture or chart".to_owned()
    }
}

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

    /// A table's short id naming two tables is refused naming both: a row
    /// added to a table names its cells from the table's place (docx-edit
    /// `table_base`), so once the first table is gone, a row added to the
    /// third (now second) shares the second's slot `body:t1` (Epo
    /// 2026-10-23).
    #[test]
    fn a_short_table_id_naming_two_tables_is_refused_naming_both() {
        let mut session =
            DocxSession::open(crate::starter(crate::Format::Docx).unwrap(), None).unwrap();
        let delete = |id: &str, rows: &[&str]| Command::DeleteTable {
            id: id.into(),
            expected: rows.iter().map(|row| vec![(*row).to_owned()]).collect(),
        };
        for text in ["A", "B", "C"] {
            session
                .apply(&Command::InsertTable {
                    after_id: None,
                    rows: vec![vec![text.into()]],
                })
                .unwrap();
        }
        // Slots t0, t1 and t2; without A, C is the body's second table.
        session.apply(&delete("body:t0", &["A"])).unwrap();
        session
            .engine
            .doc()
            .insert_row(&ctx(), &docx_edit::CellLoc::new(BODY, 1, 0, 0), true)
            .unwrap();
        assert!(session.story_paras("body:t1:r1c0").is_ok());
        let read = |cell: &str| {
            let para_id = session.story_paras(cell).unwrap()[0].para_id.clone();
            format!("{cell}:paragraph:{para_id}")
        };
        let refusal = Error::edit(
            EditCode::UnavailableTarget,
            format!(
                "\"body:t1\" names 2 tables, which the read shows as {:?}, {:?}: copy each id whole from inspect_office_document, reading the file first",
                read("body:t1:r0c0"),
                read("body:t2:r0c0"),
            ),
        );
        assert_eq!(
            session.apply(&delete("body:t1", &["B"])).err(),
            Some(refusal)
        );
        session.apply(&delete("body:t2", &["C", ""])).unwrap();
        session.apply(&delete("body:t1", &["B"])).unwrap();
        assert!(session.story_paras("body:t1:r0c0").is_err());
    }

    /// xorshift64: the commands a seed draws.
    struct Draw(u64);

    impl Draw {
        fn below(&mut self, n: usize) -> usize {
            self.0 ^= self.0 << 13;
            self.0 ^= self.0 >> 7;
            self.0 ^= self.0 << 17;
            (self.0 % n.max(1) as u64) as usize
        }

        fn pick<'a, T>(&mut self, items: &'a [T]) -> Option<&'a T> {
            (!items.is_empty()).then(|| &items[self.below(items.len())])
        }
    }

    /// What a replacement or a new paragraph writes: nothing, ASCII, CJK and
    /// characters of two UTF-16 units.
    const PIECES: [&str; 6] = ["", "x", "a longer piece", "漢字かな", "😀𝄞", "a😀b"];

    /// Each story with its paragraphs as a fresh scan reads them: the body,
    /// its tables' cells, the headers', footers' and notes'.
    fn stories(session: &DocxSession) -> Vec<(String, Vec<Para>)> {
        let doc = session.engine.doc();
        let mut names: Vec<String> = parts(&session.package())
            .into_iter()
            .map(|(story, _)| story)
            .filter(|story| doc.story_segments(story).is_ok())
            .collect();
        names.push(BODY.to_owned());
        let mut read = Vec::new();
        while let Some(story) = names.pop() {
            let paras = paras(doc, &story).unwrap();
            for table in paras.iter().flat_map(Para::tables) {
                names.extend(table_cells(&table.payload).into_iter().flatten());
            }
            read.push((story, paras));
        }
        read
    }

    /// `text` with a random span of its characters replaced by a piece.
    fn edited(text: &str, draw: &mut Draw) -> String {
        let chars: Vec<char> = text.chars().collect();
        let (a, b) = (draw.below(chars.len() + 1), draw.below(chars.len() + 1));
        let (from, to) = (a.min(b), a.max(b));
        let mut edited: String = chars[..from].iter().collect();
        edited.push_str(PIECES[draw.below(PIECES.len())]);
        edited.extend(&chars[to..]);
        edited
    }

    /// A command the session applies: mostly replacements and formats in
    /// every story they edit, sometimes a structural one.
    fn draw_command(session: &DocxSession, draw: &mut Draw) -> Option<Command> {
        let stories = stories(session);
        let id = |story: &str, para: &Para| format!("{story}:paragraph:{}", para.para_id);
        let body_stories: Vec<&(String, Vec<Para>)> = stories
            .iter()
            .filter(|(story, _)| in_body(story).is_ok())
            .collect();
        let body = &stories.iter().find(|(story, _)| story == BODY)?.1;
        Some(match draw.below(16) {
            0..=5 => {
                let (story, paras) = draw.pick(&stories)?;
                let plain: Vec<&Para> = paras.iter().filter(|para| para.plain).collect();
                let para = draw.pick(&plain)?;
                Command::ReplaceText {
                    target_id: id(story, para),
                    expected_text: para.text.clone(),
                    text: edited(&para.text, draw),
                }
            }
            6..=9 => {
                let (story, paras) = draw.pick(&body_stories)?;
                let open: Vec<&Para> = paras
                    .iter()
                    .filter(|para| para.refused().is_none() && !para.text.is_empty())
                    .collect();
                let para = draw.pick(&open)?;
                let chars: Vec<char> = para.text.chars().collect();
                let from = draw.below(chars.len());
                let to = from + 1 + draw.below(chars.len() - from);
                let span: String = chars[from..to].iter().collect();
                let unique = para.text.matches(span.as_str()).count() == 1;
                Command::FormatText {
                    target_id: id(story, para),
                    text: if unique { span } else { para.text.clone() },
                    format: TextFormat {
                        bold: Some(draw.below(2) == 0),
                        italic: (draw.below(2) == 0).then_some(true),
                        ..TextFormat::default()
                    },
                }
            }
            10 => {
                let (story, paras) = draw.pick(&body_stories)?;
                let open: Vec<&Para> = paras.iter().filter(|para| !para.tracked).collect();
                Command::SetAlignment {
                    target_id: id(story, draw.pick(&open)?),
                    alignment: [Alignment::Center, Alignment::Right, Alignment::Left]
                        [draw.below(3)],
                }
            }
            11 => {
                let (story, paras) = draw.pick(&body_stories)?;
                let open: Vec<&Para> = paras.iter().filter(|para| !para.tracked).collect();
                let mut ids: Vec<String> = (0..1 + draw.below(3))
                    .filter_map(|_| draw.pick(&open).map(|para| id(story, para)))
                    .collect();
                ids.dedup();
                Command::SetList {
                    ids,
                    list: [ListKind::Bullet, ListKind::Numbered, ListKind::None][draw.below(3)],
                }
            }
            12 => {
                // After a paragraph, before one, or with neither: at the
                // body's end, or into its only paragraph when that is empty.
                let anchor = draw.below(3);
                let mut named = || draw.pick(body).map(|para| id(BODY, para));
                let (after_id, before_id) = match anchor {
                    0 => (Some(named()?), None),
                    1 => (None, Some(named()?)),
                    _ => (None, None),
                };
                Command::InsertParagraph {
                    after_id,
                    before_id,
                    text: PIECES[draw.below(PIECES.len())].to_owned(),
                    style: None,
                    list: (draw.below(3) == 0).then_some(ListKind::Bullet),
                }
            }
            13 => {
                let (story, paras) = draw.pick(&body_stories)?;
                let deletable: Vec<&Para> = paras
                    .iter()
                    .enumerate()
                    .filter(|(index, para)| {
                        para.locked.is_none()
                            && !para.follows_control()
                            && index + 1 < paras.len()
                            && (para.blocks.is_empty() || paras[index + 1].blocks.is_empty())
                            && !para.ends_section()
                    })
                    .map(|(_, para)| para)
                    .collect();
                let para = draw.pick(&deletable)?;
                Command::DeleteParagraph {
                    target_id: id(story, para),
                    expected_text: para.text.clone(),
                }
            }
            14 => {
                if stories.len() > 30 {
                    return None;
                }
                let columns = 1 + draw.below(2);
                Command::InsertTable {
                    after_id: Some(id(BODY, draw.pick(body)?)),
                    rows: (0..1 + draw.below(2))
                        .map(|_| {
                            (0..columns)
                                .map(|_| PIECES[draw.below(PIECES.len())].to_owned())
                                .collect()
                        })
                        .collect(),
                }
            }
            _ => {
                let doc = session.engine.doc();
                let tables: Vec<Vec<Vec<String>>> = body
                    .iter()
                    .flat_map(Para::tables)
                    .map(|table| table_cells(&table.payload))
                    .filter(|cells| {
                        cells
                            .iter()
                            .flatten()
                            .all(|cell| cell_refusal(&paras(doc, cell).unwrap()).is_none())
                    })
                    .collect();
                let cells = draw.pick(&tables)?;
                let read = |cell: &String| paras(doc, cell).unwrap();
                let first = &cells[0][0];
                Command::DeleteTable {
                    id: id(first, &read(first)[0]),
                    expected: cells
                        .iter()
                        .map(|row| {
                            row.iter()
                                .map(|cell| {
                                    read(cell)
                                        .into_iter()
                                        .map(|para| para.text)
                                        .collect::<Vec<_>>()
                                        .join("\n")
                                })
                                .collect()
                        })
                        .collect(),
                }
            }
        })
    }

    /// Review M2: after each command of seeded random calls, every story the
    /// call's index holds reads as a fresh scan does, field by field, and
    /// each command's `touched` is what a fresh scan of the same document
    /// measures. (Not what a one-command call on the saved state measures:
    /// loading merges a client's adjacent runs that the call's edits split,
    /// so its items differ, and so does what its Undo re-creates.)
    #[test]
    fn the_index_reads_as_a_fresh_scan_after_each_command() {
        let stories: &[u8] = include_bytes!("../../tests/fixtures/stories.docx");
        // The starter's only paragraph is empty: an insert with no anchor
        // fills it.
        let starter = crate::starter(crate::Format::Docx).unwrap();
        for seed in 1..=18_u64 {
            let mut draw = Draw(seed.wrapping_mul(0x9E37_79B9_7F4A_7C15));
            let base = if seed <= 12 { stories } else { starter };
            let mut session = DocxSession::open(base, None).unwrap();
            let mut commands = Vec::new();
            while commands.len() < 40 {
                let Some(command) = draw_command(&session, &mut draw) else {
                    continue;
                };
                let indexed = session.touched(&command);
                let kept = std::mem::take(&mut session.index.borrow_mut().stories);
                let scanned = session.touched(&command);
                session.index.borrow_mut().stories = kept;
                assert_eq!(indexed, scanned, "seed {seed}: {command:?}");
                // A read from a paragraph mark measures what a walk from the
                // story's start does (review L4).
                FROM_START.set(true);
                let walked = session.touched(&command);
                FROM_START.set(false);
                assert_eq!(indexed, walked, "seed {seed}: {command:?}");
                session
                    .apply(&command)
                    .unwrap_or_else(|error| panic!("seed {seed}: {command:?}: {error:?}"));
                commands.push(command);
                assert_in_step(
                    &session,
                    &format!("seed {seed} after {:?}", commands.last()),
                );
            }
            session.finish_agent(&commands).unwrap();
            assert_in_step(&session, &format!("seed {seed} after the call"));
        }
    }

    thread_local! {
        /// `touched` walks each story from its start, not from a mark.
        pub(super) static FROM_START: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
    }

    /// A reread whose slots read as before but past which a mark was
    /// replaced (an edit outside the slots it names) drops the story rather
    /// than keep the gone mark (review L1).
    #[test]
    fn a_reread_drops_a_story_whose_next_mark_changed() {
        let mut session =
            DocxSession::open(crate::starter(crate::Format::Docx).unwrap(), None).unwrap();
        for text in ["One", "Two", "Three"] {
            let command = Command::InsertParagraph {
                after_id: None,
                before_id: None,
                text: text.to_owned(),
                style: None,
                list: None,
            };
            session.apply(&command).unwrap();
        }
        let body = session.story_paras(BODY).unwrap();
        assert!(body.len() >= 3, "{body:?}");
        let before = session.index_mark();
        let doc = session.engine.doc();
        // Paragraph 1's mark goes and a new one takes its place.
        doc.merge_paragraphs(&ctx(), &body[1].para_id, docx_edit::MergeDirection::Forward)
            .unwrap();
        doc.split_paragraph(&ctx(), Position::new(BODY, body[1].mark))
            .unwrap();
        session.reread(before, BODY, 0, 0);
        let index = session.index.borrow();
        if let Some(cached) = index.stories.get(BODY) {
            assert_eq!(**cached, paras(doc, BODY).unwrap());
        }
    }

    /// Every command keeps the index in step, and each story it holds reads
    /// as a fresh read does; that read gives what the story's segments give.
    fn assert_in_step(session: &DocxSession, after: &str) {
        let doc = session.engine.doc();
        let index = session.index.borrow();
        assert_eq!(index.at, session.engine.doc_epoch(), "{after}");
        for (story, cached) in &index.stories {
            let fresh = paras(doc, story);
            assert!(
                fresh.as_ref().is_ok_and(|fresh| **cached == *fresh),
                "story {story} {after}:\n{cached:?}\n{fresh:?}"
            );
        }
        for (story, read) in stories(session) {
            assert_eq!(
                read.into_iter().map(unplaced).collect::<Vec<_>>(),
                segment_paras(doc, &story),
                "story {story} {after}"
            );
        }
    }

    /// A paragraph but where a read of the next one starts, which the
    /// segments do not give.
    fn unplaced(para: Para) -> Para {
        Para {
            mark_item: ID::new(yrs::ClientID::new(0), 0),
            after: Open::default(),
            ..para
        }
    }

    /// `story`'s paragraphs as its segments give them (the read before
    /// paragraphs were read item by item), but where a read of the next one
    /// starts.
    fn segment_paras(doc: &EditingDoc, story: &str) -> Vec<Para> {
        use docx_edit::SegmentContent;
        let mut paras = Vec::new();
        let (mut at, mut start, mut node_start) = (0, 0, 0);
        let mut plain = true;
        let mut text = String::new();
        let mut locked: Option<String> = None;
        let mut tracked = false;
        let mut blocks = Vec::new();
        for segment in doc.story_segments(story).unwrap() {
            match segment.content {
                SegmentContent::Text(part) => {
                    let deleted = truthy(segment.attributes.get("del"));
                    if deleted || truthy(segment.attributes.get("ins")) {
                        locked.get_or_insert_with(|| "holds tracked changes".into());
                        tracked = true;
                        plain = false;
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
                    let mut values = properties.values;
                    values.remove("bookmarks");
                    paras.push(Para {
                        para_id: properties.para_id,
                        start,
                        node_start,
                        plain: std::mem::replace(&mut plain, true),
                        mark: at,
                        text: std::mem::take(&mut text),
                        locked: locked.take(),
                        tracked: std::mem::take(&mut tracked) || mark_tracked,
                        blocks: std::mem::take(&mut blocks),
                        properties: values,
                        mark_item: ID::new(yrs::ClientID::new(0), 0),
                        after: Open::default(),
                    });
                    at += 1;
                    start = at;
                    node_start = at;
                }
                SegmentContent::OtherEmbed { kind, payload } => {
                    let leading = at == start;
                    let breaks =
                        at == node_start && matches!(kind.as_str(), "pageBreak" | "columnBreak");
                    let block = matches!(kind.as_str(), "table" | "blockSdt");
                    if block {
                        blocks.push(Block {
                            kind: kind.clone(),
                            payload,
                            at,
                        });
                    }
                    at += 1;
                    if breaks {
                        node_start = at;
                    } else {
                        plain = false;
                    }
                    if leading && (block || matches!(kind.as_str(), "pageBreak" | "columnBreak")) {
                        start = at;
                    } else {
                        locked.get_or_insert_with(|| locked_by(&kind));
                    }
                }
            }
        }
        paras
    }
}
