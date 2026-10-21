//! The chat agent's DOCX paragraph commands (Epo 2026-10-21): insert a
//! paragraph in the document body with a paragraph style of the file, delete
//! a plain paragraph; and its scoped read, the editable paragraphs with the
//! file's paragraph styles. A new paragraph gets a style id and no direct
//! formatting.

use std::collections::BTreeMap;

use docx_edit::{
    EditCtx, FormatPolicy, OpError, ParaAttrDelta, ParaSelector, Position, SegmentContent,
};
use serde_json::{Value, json};
use yrs::Any;

use super::DocxSession;
use crate::common::{check_expected_paragraph, check_paragraph_text, insert_anchor, len16};
use crate::error::{EditCode, Error, Result};
use crate::session::Applied;
use crate::types::{Command, InspectQuery};

const BODY: &str = "body";
/// Paragraph-mark keys a new paragraph keeps from the one it was split off:
/// its identity and the section it ends.
const KEPT_KEYS: [&str; 4] = ["paraId", "_kind", "sectPr", "sectionBreakType"];
const SECTION_KEYS: [&str; 2] = ["sectPr", "sectionBreakType"];
/// A mark's tracked insertion, deletion or property change.
const TRACKED_KEYS: [&str; 3] = ["pPrIns", "pPrDel", "pPrChange"];
/// Paragraphs a read lists.
const MAX_READ: u32 = 200;

fn invalid(message: impl Into<String>) -> Error {
    Error::edit(EditCode::InvalidInput, message)
}

fn ctx() -> EditCtx {
    EditCtx::local(String::new(), String::new())
}

/// An op's refusal of the agent's input, else an engine fault.
fn op(error: OpError) -> Error {
    match error {
        OpError::TextContainsBreak => invalid(error.to_string()),
        error => Error::engine(error),
    }
}

fn gone() -> Error {
    Error::edit(
        EditCode::UnavailableTarget,
        "the paragraph is no longer in the document",
    )
}

/// A Yjs attribute as JavaScript would test it.
fn truthy(value: Option<&Any>) -> bool {
    !matches!(
        value,
        None | Some(Any::Null | Any::Undefined | Any::Bool(false))
    )
}

/// A paragraph mark: its story index, properties and whether it is tracked.
struct Mark {
    para_id: String,
    at: u32,
    properties: BTreeMap<String, Any>,
    tracked: bool,
}

impl DocxSession {
    pub(super) fn apply_paragraphs(&mut self, command: &Command) -> Result<Applied> {
        // Tracking keeps deleted items from garbage collection, as for
        // replace_text.
        self.undo.track(self.engine.doc());
        let applied = match command {
            Command::InsertParagraph {
                after_id,
                before_id,
                text,
                style,
            } => {
                let (anchor, after) = insert_anchor(after_id.as_deref(), before_id.as_deref())?;
                self.insert_paragraph(anchor, after, text, style.as_deref())
            }
            Command::DeleteParagraph {
                target_id,
                expected_text,
            } => self.delete_paragraph(target_id, expected_text),
            _ => Err(Error::edit(
                EditCode::UnsupportedOperation,
                "DOCX files take replace_text, insert_paragraph and delete_paragraph",
            )),
        }?;
        self.projected = None;
        Ok(applied)
    }

    /// `count` editable paragraphs from `start` (as `inspect` lists them),
    /// the first page with the file's paragraph styles.
    pub(crate) fn inspect_scoped(&mut self, query: &InspectQuery) -> Result<Value> {
        let start = query.start.unwrap_or(0) as usize;
        let count = query
            .count
            .filter(|count| (1..=MAX_READ).contains(count))
            .ok_or_else(|| {
                invalid(format!(
                    "a DOCX read names a count of 1 to {MAX_READ} paragraphs"
                ))
            })? as usize;
        let entries = self.editable()?;
        let total = entries.len();
        let end = (start + count).min(total);
        let mut answer = json!({
            "entries": entries[start.min(total)..end],
            "total": total,
            "nextStart": (end < total).then_some(end),
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

    /// The file's paragraph styles, id and name, in its styles part's order.
    fn paragraph_styles(&self) -> Vec<(String, String)> {
        self.base_document()
            .get("package")
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
                    let tracked = ["ins", "del"]
                        .iter()
                        .any(|key| truthy(segment.attributes.get(*key)))
                        || TRACKED_KEYS
                            .iter()
                            .any(|key| properties.values.contains_key(*key));
                    marks.push(Mark {
                        para_id: properties.para_id,
                        at,
                        properties: properties.values,
                        tracked,
                    });
                    at += 1;
                }
                SegmentContent::OtherEmbed { .. } => at += 1,
            }
        }
        Ok(marks)
    }

    /// A body paragraph after or before `anchor`, its paragraph style `style`
    /// and nothing else of its neighbour's. Several inserts after one
    /// paragraph land in call order. The inverse deletes it, unless it is the
    /// body's last paragraph, which delete_paragraph refuses.
    fn insert_paragraph(
        &mut self,
        anchor: &str,
        after: bool,
        text: &str,
        style: Option<&str>,
    ) -> Result<Applied> {
        check_paragraph_text(text)?;
        if let Some(style) = style
            && !self.paragraph_styles().iter().any(|(id, _)| id == style)
        {
            return Err(invalid(format!(
                "the document has no paragraph style {style:?}; use a style id inspect_office_document lists"
            )));
        }
        let para_id = anchor
            .strip_prefix("body:paragraph:")
            .ok_or_else(|| {
                invalid(
                    "a new paragraph goes in the document body: after_id and before_id name a body paragraph, not one of a table, header, footer or note",
                )
            })?;
        let marks = self.marks(BODY)?;
        let mut index = marks
            .iter()
            .position(|mark| mark.para_id == para_id)
            .ok_or_else(gone)?;
        if after {
            while marks
                .get(index + 1)
                .is_some_and(|next| self.inserted.contains(&next.para_id))
            {
                index += 1;
            }
        }
        let doc = self.engine.doc();
        // A new paragraph after a mark is split off at it; before the first,
        // one goes at the body's start. Either copies its neighbour's mark.
        let (new_id, at, neighbour, last) = match (after, index) {
            (false, 0) => (
                doc.insert_paragraph_at_start(&ctx(), BODY)
                    .map_err(op)?
                    .first_para_id,
                0,
                &marks[0],
                false,
            ),
            _ => {
                let mark = &marks[if after { index } else { index - 1 }];
                let receipt = doc
                    .split_paragraph(&ctx(), Position::new(BODY, mark.at))
                    .map_err(op)?;
                (
                    receipt.second_para_id,
                    mark.at + 1,
                    mark,
                    after && index + 1 == marks.len(),
                )
            }
        };
        let mut other: BTreeMap<String, Option<Any>> = neighbour
            .properties
            .keys()
            .filter(|key| !KEPT_KEYS.contains(&key.as_str()))
            .map(|key| (key.clone(), None))
            .collect();
        if let Some(style) = style {
            other.insert("pStyle".to_owned(), Some(Any::from(style)));
        }
        doc.set_paragraph_attrs(
            &ctx(),
            &ParaSelector::One(new_id.clone()),
            &ParaAttrDelta {
                other,
                ..ParaAttrDelta::default()
            },
        )
        .map_err(op)?;
        if !text.is_empty() {
            doc.insert_text(&ctx(), Position::new(BODY, at), text, FormatPolicy::Plain)
                .map_err(op)?;
        }
        let target_id = format!("{BODY}:paragraph:{new_id}");
        self.inserted.insert(new_id);
        Ok(Applied {
            ids: vec![target_id.clone()],
            inverse: (!last).then(|| Command::DeleteParagraph {
                target_id,
                expected_text: text.to_owned(),
            }),
        })
    }

    /// Refused unless the paragraph is plain, reads `expected_text`, ends no
    /// section and is not its story's last. No inverse: putting it back
    /// exactly is the live Undo's.
    fn delete_paragraph(&mut self, id: &str, expected_text: &str) -> Result<Applied> {
        let (story, paragraph) = self.paragraph(id)?;
        let marks = self.marks(&story)?;
        let index = marks
            .iter()
            .position(|mark| mark.para_id == paragraph.para_id)
            .ok_or_else(gone)?;
        let mark = &marks[index];
        if !paragraph.plain || mark.tracked {
            return Err(Error::edit(
                EditCode::UnavailableTarget,
                "the paragraph holds objects, fields or tracked changes, or follows a table or content control; it cannot be deleted",
            ));
        }
        check_expected_paragraph(id, &paragraph.text, expected_text)?;
        if index + 1 == marks.len() {
            return Err(invalid(
                "a story's last paragraph cannot be deleted; replace its text instead",
            ));
        }
        if SECTION_KEYS
            .iter()
            .any(|key| mark.properties.contains_key(*key))
        {
            return Err(invalid(
                "the paragraph ends a section; deleting it would join two sections",
            ));
        }
        self.engine
            .doc()
            .remove_paragraph(&ctx(), &story, &paragraph.para_id)
            .map_err(op)?;
        Ok(Applied {
            ids: Vec::new(),
            inverse: None,
        })
    }
}
