//! The chat agent's DOCX commands and scoped read (Epo 2026-10-20): insert
//! and delete paragraphs and tables, lists, text formatting and alignment,
//! and a window of the body's paragraphs with the file's paragraph styles.
//! Paragraphs get a style id and direct formatting only, never resolved
//! style values (fix 9 makes the seed hold only direct formatting). No
//! command has an inverse: the live Undo replaces inverses.

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
use super::styles::{Numbering, StyleResolver};
use crate::common::len16;
use crate::error::{EditCode, Error, Result};
use crate::jsv::V;
use crate::session::Applied;
use crate::types::{Alignment, Command, InspectQuery, ListKind, TextFormat};

const BODY: &str = "body";
/// Paragraphs a call inserts; characters a new paragraph or cell holds.
const MAX_INSERTED_PARAGRAPHS: usize = 20;
const MAX_TEXT_CHARS: usize = 5_000;
const MAX_TABLE_ROWS: usize = 20;
const MAX_TABLE_COLUMNS: usize = 10;
/// Paragraphs a read or a `set_list` covers.
const MAX_PARAGRAPHS: usize = 200;
/// Paragraph-mark keys a new paragraph keeps from the one it was split off:
/// its identity and the section it ends.
const KEPT_KEYS: [&str; 4] = ["paraId", "_kind", "sectPr", "sectionBreakType"];
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

/// A call inserts at most [`MAX_INSERTED_PARAGRAPHS`] paragraphs.
pub(crate) fn check_call(commands: &[Command]) -> Result<()> {
    let inserts = commands
        .iter()
        .filter(|command| matches!(command, Command::InsertParagraph { .. }))
        .count();
    if inserts > MAX_INSERTED_PARAGRAPHS {
        return Err(invalid(format!(
            "the call inserts {inserts} paragraphs; insert at most {MAX_INSERTED_PARAGRAPHS} a call"
        )));
    }
    Ok(())
}

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

/// A refusal carrying the paragraph as it reads now.
fn paragraph_refusal(code: EditCode, message: String, id: &str, text: &str) -> Error {
    Error::Edit {
        code,
        message,
        details: Some(json!({ "block": { "id": id, "type": "paragraph", "text": text } })),
    }
}

/// New text is one paragraph of at most [`MAX_TEXT_CHARS`] characters.
fn check_text(text: &str, what: &str) -> Result<()> {
    if text.contains(['\n', '\r', '\u{b}', '\u{2028}', '\u{2029}']) {
        return Err(invalid(format!(
            "{what} must be one paragraph without line breaks"
        )));
    }
    let chars = text.chars().count();
    if chars > MAX_TEXT_CHARS {
        return Err(invalid(format!(
            "{what} holds {chars} characters; at most {MAX_TEXT_CHARS}"
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

fn gone() -> Error {
    Error::edit(
        EditCode::UnavailableTarget,
        "the paragraph is no longer in the document",
    )
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

/// The reason the read gives for a paragraph that `replace_text`,
/// `delete_paragraph` and `format_text` refuse (`docxParagraphs`' `plain`):
/// an object or tracked change in it, or the table or content control in
/// its slot.
fn locked_by(kind: &str) -> String {
    match kind {
        "table" => "follows a table".into(),
        "blockSdt" => "follows a content control".into(),
        "image" => "holds a picture".into(),
        "shape" => "holds a drawing".into(),
        "mathEquation" => "holds an equation".into(),
        "footnoteRef" | "endnoteRef" => "holds a note reference".into(),
        "inlineSdt" | "sdt" => "holds a content control".into(),
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

/// A paragraph mark: its story index and properties.
struct Mark {
    para_id: String,
    at: u32,
    properties: BTreeMap<String, Any>,
}

/// Where an insert goes.
enum Place {
    /// After this body paragraph, as a new one.
    After(Mark),
    /// Into this paragraph: the body's only one, empty.
    Fill(Mark),
}

/// One cell's mark on the paragraphs a read lists.
#[derive(Clone)]
struct CellMark {
    table: Option<String>,
    row: usize,
    column: usize,
}

/// The read's walk: paragraphs counted in order, the window's listed.
struct Walk<'a> {
    doc: &'a EditingDoc,
    lists: Lists,
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
        let segments = self.doc.story_segments(story).map_err(Error::engine)?;
        let mut text = String::new();
        let mut locked: Option<String> = None;
        // A page or column break a paragraph opens with keeps it editable.
        let mut leading = true;
        for segment in segments {
            match segment.content {
                SegmentContent::Text(part) => {
                    leading = false;
                    let deleted = truthy(segment.attributes.get("del"));
                    if deleted || truthy(segment.attributes.get("ins")) {
                        locked.get_or_insert_with(|| "holds tracked changes".into());
                    }
                    if !deleted {
                        text.push_str(&part);
                    }
                }
                SegmentContent::Pilcrow(properties) => {
                    let text = std::mem::take(&mut text);
                    let locked = locked.take();
                    leading = true;
                    if (self.start..self.end).contains(&self.index) {
                        let mut entry = json!({
                            "id": format!("{story}:paragraph:{}", properties.para_id),
                            "style": match properties.values.get("pStyle") {
                                Some(Any::String(style)) => Value::from(style.to_string()),
                                _ => Value::Null,
                            },
                            "list": self.lists.name(&properties.values),
                            "text": text,
                        });
                        if let Some(cell) = cell {
                            entry["table"] =
                                json!({ "id": cell.table, "row": cell.row, "column": cell.column });
                        }
                        if control {
                            entry["contentControl"] = Value::Bool(true);
                        }
                        if let Some(locked) = locked {
                            entry["locked"] = Value::from(locked);
                        }
                        self.listed.push(entry);
                    }
                    self.index += 1;
                }
                SegmentContent::OtherEmbed { kind, payload } => {
                    if leading && matches!(kind.as_str(), "pageBreak" | "columnBreak") {
                        continue;
                    }
                    leading = false;
                    locked.get_or_insert_with(|| locked_by(&kind));
                    if kind == "table" {
                        self.table(&table_cells(&payload), control)?;
                    } else if kind == "blockSdt"
                        && let Some(Any::String(child)) = payload.get("story")
                    {
                        self.story(child, cell, true)?;
                    }
                }
            }
        }
        Ok(())
    }

    /// A table's cells row by row, each paragraph marked with the table,
    /// which its first cell's first paragraph names.
    fn table(&mut self, cells: &[Vec<String>], control: bool) -> Result<()> {
        let table = match cells.first().and_then(|row| row.first()) {
            Some(first) => self
                .doc
                .story_segments(first)
                .map_err(Error::engine)?
                .into_iter()
                .find_map(|segment| match segment.content {
                    SegmentContent::Pilcrow(properties) => {
                        Some(format!("{first}:paragraph:{}", properties.para_id))
                    }
                    _ => None,
                }),
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
    pub(super) fn apply_agent(&mut self, command: &Command) -> Result<Applied> {
        // No inverse is taken, but tracking keeps deleted items from garbage
        // collection as it does for replace_text, so the state still holds
        // what a delete removed.
        self.undo.track(self.engine.doc());
        match command {
            Command::InsertParagraph {
                after_id,
                text,
                style,
                list,
            } => self.insert_paragraph(after_id.as_deref(), text, style.as_deref(), *list),
            Command::DeleteParagraph { id, expected_text } => {
                self.delete_paragraph(id, expected_text)
            }
            Command::SetList { ids, list } => self.set_list(ids, *list),
            Command::InsertTable { after_id, rows } => self.insert_table(after_id.as_deref(), rows),
            Command::DeleteTable { id, expected } => self.delete_table(id, expected),
            Command::FormatText {
                target_id,
                expected_text,
                format,
            } => self.format_text(target_id, expected_text, format),
            Command::SetAlignment {
                target_id,
                alignment,
            } => self.set_alignment(target_id, *alignment),
            _ => Err(Error::edit(
                EditCode::UnsupportedOperation,
                "this command edits XLSX files",
            )),
        }?;
        self.projected = None;
        Ok(Applied {
            ids: Vec::new(),
            inverse: None,
        })
    }

    /// `count` paragraphs of the body from `start`, each table's cells
    /// listed before the paragraph that follows it and marked with the
    /// table, row and column; the first window also lists the paragraph
    /// styles.
    pub(crate) fn inspect_scoped(&mut self, query: &InspectQuery) -> Result<Value> {
        let start = query.start.unwrap_or(0) as usize;
        let count = query
            .count
            .map(|count| count as usize)
            .filter(|count| (1..=MAX_PARAGRAPHS).contains(count))
            .ok_or_else(|| {
                invalid(format!(
                    "a DOCX read names a count of 1 to {MAX_PARAGRAPHS} paragraphs"
                ))
            })?;
        let package = self.package();
        let mut walk = Walk {
            doc: self.engine.doc(),
            lists: Lists {
                styles: StyleResolver::new(&package.get("styles")),
                numbering: Numbering::new(&package.get("numbering")),
            },
            start,
            end: start + count,
            index: 0,
            listed: Vec::new(),
        };
        walk.story(BODY, None, false)?;
        let total = walk.index;
        let mut answer = json!({
            "paragraphs": walk.listed,
            "total": total,
            "nextStart": (start + count < total).then_some(start + count),
        });
        if start == 0 {
            answer["styles"] = self
                .paragraph_styles()
                .into_iter()
                .map(|(id, name)| json!({ "id": id, "name": name }))
                .collect();
        }
        Ok(answer)
    }

    fn package(&self) -> V {
        self.base_document().get("package")
    }

    /// The file's paragraph styles, id and name, in its styles part's order.
    fn paragraph_styles(&self) -> Vec<(String, String)> {
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
                Some((id, name))
            })
            .collect()
    }

    fn check_style(&self, style: &str) -> Result<()> {
        let styles = self.paragraph_styles();
        if styles.iter().any(|(id, _)| id == style) {
            return Ok(());
        }
        Err(Error::Edit {
            code: EditCode::UnavailableTarget,
            message: format!("the document has no paragraph style {style:?}"),
            details: Some(json!({
                "blocks": styles
                    .iter()
                    .map(|(id, name)| json!({ "id": id, "type": "style", "text": name }))
                    .collect::<Vec<_>>(),
            })),
        })
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

    /// Each paragraph mark of `story`, in order.
    fn marks(&self, story: &str) -> Result<Vec<Mark>> {
        let segments = self
            .engine
            .doc()
            .story_segments(story)
            .map_err(|_| Error::edit(EditCode::UnavailableTarget, "unknown DOCX story"))?;
        let mut at = 0;
        let mut marks = Vec::new();
        for segment in segments {
            match segment.content {
                SegmentContent::Text(text) => at += len16(&text),
                SegmentContent::Pilcrow(properties) => {
                    marks.push(Mark {
                        para_id: properties.para_id,
                        at,
                        properties: properties.values,
                    });
                    at += 1;
                }
                SegmentContent::OtherEmbed { .. } => at += 1,
            }
        }
        Ok(marks)
    }

    /// Where an insert after `after_id` goes: after that body paragraph and
    /// the paragraphs this call already inserted there (their paraIds are
    /// this session's), so inserts after one paragraph keep their order.
    /// Without `after_id`, after the body's last paragraph, or into it when
    /// it is the body's only one and empty (a new document).
    fn place(&self, after_id: Option<&str>) -> Result<Place> {
        let mut marks = self.marks(BODY)?;
        let index = match after_id {
            None => {
                if let [only] = &marks[..]
                    && only.at == 0
                {
                    return Ok(Place::Fill(marks.swap_remove(0)));
                }
                marks
                    .len()
                    .checked_sub(1)
                    .ok_or_else(|| Error::engine("the document body has no paragraph"))?
            }
            Some(id) => {
                let (story, para_id) = split_id(id)?;
                if story != BODY {
                    return Err(invalid(
                        "after_id must be a paragraph of the document body, not of a table, content control, header, footer or note",
                    ));
                }
                let mut index = marks
                    .iter()
                    .position(|mark| mark.para_id == para_id)
                    .ok_or_else(gone)?;
                let minted = format!("{}:", self.engine.doc().client_id());
                while marks
                    .get(index + 1)
                    .is_some_and(|next| next.para_id.starts_with(&minted))
                {
                    index += 1;
                }
                index
            }
        };
        Ok(Place::After(marks.swap_remove(index)))
    }

    /// A new empty body paragraph where `place` puts it, left with its
    /// identity, the section it ends, `style` and `list` only (a split
    /// copies its neighbour's properties, the seed's resolved style values
    /// among them). Returns its mark's index, where its text starts.
    fn new_paragraph(
        &self,
        place: Place,
        style: Option<&str>,
        list: Option<ListKind>,
    ) -> Result<u32> {
        let (at, copied) = match place {
            Place::Fill(mark) => (mark.at, mark.properties),
            Place::After(mark) => {
                self.engine
                    .doc()
                    .split_paragraph(&ctx(), Position::new(BODY, mark.at))
                    .map_err(op)?;
                // A split at a paragraph's end leaves the new, empty
                // paragraph's mark right after the old one's, holding the
                // old one's properties or a copy of them.
                (mark.at + 1, mark.properties)
            }
        };
        let mut other: BTreeMap<String, Option<Any>> = copied
            .into_keys()
            .filter(|key| !KEPT_KEYS.contains(&key.as_str()))
            .map(|key| (key, None))
            .collect();
        if let Some(style) = style {
            other.insert("pStyle".to_owned(), Some(Any::from(style)));
        }
        if let Some(list) = list.filter(|list| *list != ListKind::None) {
            other.extend(list_attrs(list, self.list_num_id(list)?, 0.0));
        }
        self.engine
            .doc()
            .set_paragraph_attrs(
                &ctx(),
                &ParaSelector::Range(StoryRange::new(BODY, at, at)),
                &ParaAttrDelta {
                    other,
                    ..ParaAttrDelta::default()
                },
            )
            .map_err(op)?;
        Ok(at)
    }

    fn insert_paragraph(
        &mut self,
        after_id: Option<&str>,
        text: &str,
        style: Option<&str>,
        list: Option<ListKind>,
    ) -> Result<()> {
        check_text(text, "text")?;
        if let Some(style) = style {
            self.check_style(style)?;
        }
        let place = self.place(after_id)?;
        let at = self.new_paragraph(place, style, list)?;
        if !text.is_empty() {
            self.engine
                .doc()
                .insert_text(&ctx(), Position::new(BODY, at), text, FormatPolicy::Plain)
                .map_err(op)?;
        }
        Ok(())
    }

    /// Refused unless the paragraph is plain, reads `expected_text`, ends no
    /// section and is not its story's last.
    fn delete_paragraph(&mut self, id: &str, expected_text: &str) -> Result<()> {
        let (story, paragraph) = self.paragraph(id)?;
        if !paragraph.plain {
            return Err(Error::edit(
                EditCode::UnavailableTarget,
                "the paragraph holds objects or tracked changes, or follows a table",
            ));
        }
        if paragraph.text != expected_text {
            return Err(paragraph_refusal(
                EditCode::StaleTarget,
                "the paragraph's text differs from expected_text".into(),
                id,
                &paragraph.text,
            ));
        }
        let marks = self.marks(&story)?;
        let index = marks
            .iter()
            .position(|mark| mark.para_id == paragraph.para_id)
            .ok_or_else(gone)?;
        if index + 1 == marks.len() {
            return Err(invalid(
                "a story's last paragraph cannot be deleted; replace its text instead",
            ));
        }
        let mark = &marks[index];
        if ["sectPr", "sectionBreakType"]
            .iter()
            .any(|key| mark.properties.contains_key(*key))
        {
            return Err(invalid(
                "the paragraph ends a section; deleting it would join two sections",
            ));
        }
        self.engine
            .doc()
            .delete_range(
                &ctx(),
                StoryRange::new(story, paragraph.start as u32, mark.at + 1),
            )
            .map_err(op)?;
        Ok(())
    }

    /// Each paragraph keeps its list level; `none` ends its list, its
    /// style's too.
    fn set_list(&mut self, ids: &[String], list: ListKind) -> Result<()> {
        if ids.is_empty() || ids.len() > MAX_PARAGRAPHS {
            return Err(invalid(format!(
                "set_list names 1 to {MAX_PARAGRAPHS} paragraphs"
            )));
        }
        let num_id = self.list_num_id(list)?;
        let mut stories: HashMap<&str, Vec<Mark>> = HashMap::new();
        let mut targets = Vec::with_capacity(ids.len());
        for id in ids {
            let (story, para_id) = split_id(id)?;
            if !stories.contains_key(story) {
                stories.insert(story, self.marks(story)?);
            }
            let mark = stories[story]
                .iter()
                .find(|mark| mark.para_id == para_id)
                .ok_or_else(gone)?;
            let ilvl = match mark.properties.get("numPr") {
                Some(Any::Map(num_pr)) => match num_pr.get("ilvl") {
                    Some(Any::Number(ilvl)) => *ilvl,
                    _ => 0.0,
                },
                _ => 0.0,
            };
            targets.push((story, mark.at, ilvl));
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

    /// A table of `rows` with an empty paragraph after it, placed as
    /// [`Self::insert_paragraph`] places a paragraph.
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
            check_text(cell, "a cell")?;
        }
        let place = self.place(after_id)?;
        let at = self.new_paragraph(place, None, None)?;
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

    /// The table holding paragraph `id`, refused unless its cells read
    /// `expected`.
    fn delete_table(&mut self, id: &str, expected: &[Vec<String>]) -> Result<()> {
        let (cell, _) = self.paragraph(id)?;
        let story = table_story(&cell).ok_or_else(|| {
            invalid("id must be a paragraph of a table cell, as inspect_office_document marks it")
        })?;
        let doc = self.engine.doc();
        let segments = doc
            .story_segments(story)
            .map_err(|_| Error::edit(EditCode::UnavailableTarget, "unknown DOCX story"))?;
        let (ordinal, cells) = segments
            .iter()
            .filter_map(|segment| match &segment.content {
                SegmentContent::OtherEmbed { kind, payload } if kind == "table" => {
                    Some(table_cells(payload))
                }
                _ => None,
            })
            .enumerate()
            .find(|(_, cells)| cells.iter().flatten().any(|story| *story == cell))
            .ok_or_else(|| {
                Error::edit(
                    EditCode::UnavailableTarget,
                    "the table is no longer in the document",
                )
            })?;
        let grid = cells
            .iter()
            .map(|row| {
                row.iter()
                    .map(|cell| {
                        Ok(doc
                            .paragraphs(cell)
                            .map_err(Error::engine)?
                            .into_iter()
                            .map(|paragraph| paragraph.text)
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
        doc.delete_table(&ctx(), &TableLocator::new(story, ordinal as u32))
            .map_err(op)?;
        Ok(())
    }

    fn format_text(&mut self, id: &str, expected_text: &str, format: &TextFormat) -> Result<()> {
        let delta = inline_delta(format)?;
        let (story, paragraph) = self.paragraph(id)?;
        if !paragraph.plain {
            return Err(Error::edit(
                EditCode::UnavailableTarget,
                "the paragraph holds objects or tracked changes",
            ));
        }
        if expected_text.is_empty() {
            return Err(invalid("expected_text names the text to format"));
        }
        let found: Vec<usize> = paragraph
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
                    &paragraph.text,
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
                    &paragraph.text,
                ));
            }
        };
        let start = paragraph.start as u32 + len16(&paragraph.text[..at]);
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
        let (story, paragraph) = self.paragraph(id)?;
        let mark = (paragraph.start + paragraph.length) as u32;
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
    // Highlight is written as the seed stores it, the colour's name, which
    // the save reads; the delta's own `highlight` field writes `{color}`,
    // which the save projection drops.
    if let Some(highlight) = &format.highlight {
        let value = if highlight.eq_ignore_ascii_case("none") {
            None
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
            Some(Any::from(*name))
        };
        delta.other.insert("highlight".to_owned(), value);
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
