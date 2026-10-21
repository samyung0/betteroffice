//! The chat agent's DOCX paragraph commands (Epo 2026-10-21): insert a
//! paragraph in the document body, delete a plain paragraph; and its scoped
//! read, the editable paragraphs (each with its style when not the default)
//! and the file's paragraph styles. A new paragraph continues its anchor as
//! Enter does, or takes the paragraph style named and no direct formatting.

use std::collections::{BTreeMap, HashMap};

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
/// Styles the first page lists.
const MAX_STYLES: usize = 50;

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
    /// each label naming the paragraph's style when it is not the default;
    /// the first page also lists the file's shown paragraph styles, at most
    /// [`MAX_STYLES`].
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
        let mut entries = self.editable()?;
        let total = entries.len();
        let end = (start + count).min(total);
        let mut page = entries.split_off(start.min(total));
        page.truncate(end - start.min(end));
        let styles = self.paragraph_styles();
        let default = styles
            .iter()
            .find(|style| style.default)
            .map(|style| &style.id);
        let mut by_story: HashMap<String, HashMap<String, String>> = HashMap::new();
        for entry in &mut page {
            let Some((story, para_id)) = entry.id.rsplit_once(":paragraph:") else {
                continue;
            };
            if !by_story.contains_key(story) {
                let marks = self.marks(story)?;
                let styled = marks
                    .into_iter()
                    .filter_map(|mark| match mark.properties.get("pStyle") {
                        Some(Any::String(style)) => Some((mark.para_id, style.to_string())),
                        _ => None,
                    })
                    .collect();
                by_story.insert(story.to_owned(), styled);
            }
            if let Some(style) = by_story[story]
                .get(para_id)
                .filter(|style| Some(*style) != default)
            {
                entry.label = format!("{} ({style})", entry.label);
            }
        }
        let mut answer = json!({
            "entries": page,
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

    /// The file's paragraph styles in its styles part's order.
    fn paragraph_styles(&self) -> Vec<ParagraphStyle> {
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

    /// The style Enter at the end of a paragraph styled `style` (none: the
    /// default) gives the next one: the style's next style when the file has
    /// it, else the style itself; none for the default style.
    fn next_style(&self, style: Option<&str>) -> Option<String> {
        let styles = self.paragraph_styles();
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

    /// A body paragraph right after or before `anchor`. Without a style it
    /// continues its anchor as Enter at the anchor's end does (Epo
    /// 2026-10-21): the anchor's paragraph properties (list, indent) and its
    /// style's next style; the anchor is the paragraph `after_id` names, else
    /// the one before the new paragraph, at the body's start the one it goes
    /// before. A named style replaces that: the new paragraph takes it and
    /// no direct formatting. Several inserts after one paragraph land in call
    /// order. The inverse deletes it, unless delete_paragraph would refuse
    /// that: a new last paragraph, or one that ends the section its anchor
    /// ended (a split at a section's end leaves the section mark on the new
    /// paragraph, as Enter does).
    fn insert_paragraph(
        &mut self,
        anchor: &str,
        after: bool,
        text: &str,
        style: Option<&str>,
    ) -> Result<Applied> {
        check_paragraph_text(text)?;
        if let Some(style) = style
            && !self
                .paragraph_styles()
                .iter()
                .any(|known| known.id == style)
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
        let position = |id: &str| marks.iter().position(|mark| mark.para_id == id);
        let doc = self.engine.doc();
        // After a mark a new paragraph is split off it, after the paragraph
        // this call last put there; before one it takes that paragraph's
        // mark, which keeps its id and section. Either copies its
        // neighbour's mark.
        let (new_id, at, neighbour, continued, undoable) = if after {
            let named = position(para_id).ok_or_else(gone)?;
            let index = self
                .after
                .get(para_id)
                .and_then(|id| position(id))
                .unwrap_or(named);
            let mark = &marks[index];
            let receipt = doc
                .split_paragraph(&ctx(), Position::new(BODY, mark.at))
                .map_err(op)?;
            let ends_section = SECTION_KEYS
                .iter()
                .any(|key| mark.properties.contains_key(*key));
            (
                receipt.second_para_id,
                mark.at + 1,
                mark,
                &marks[named],
                index + 1 < marks.len() && !ends_section,
            )
        } else {
            let index = position(para_id).ok_or_else(gone)?;
            let receipt = doc
                .insert_paragraph_before(&ctx(), BODY, para_id)
                .map_err(op)?;
            let previous = index.checked_sub(1);
            let at = previous.map_or(0, |previous| marks[previous].at + 1);
            let continued = &marks[previous.unwrap_or(index)];
            (receipt.first_para_id, at, &marks[index], continued, true)
        };
        // The new mark copies `neighbour`'s; it keeps that identity and
        // section and `kept` for the rest. Source and save bookkeeping (`_`
        // keys) and tracked changes are not kept.
        let kept: BTreeMap<String, Any> = match style {
            Some(style) => BTreeMap::from([("pStyle".to_owned(), Any::from(style))]),
            None => {
                let mut kept: BTreeMap<String, Any> = continued
                    .properties
                    .iter()
                    .filter(|(key, _)| {
                        !KEPT_KEYS.contains(&key.as_str())
                            && !TRACKED_KEYS.contains(&key.as_str())
                            && !key.starts_with('_')
                    })
                    .map(|(key, value)| (key.clone(), value.clone()))
                    .collect();
                let style = match kept.remove("pStyle") {
                    Some(Any::String(style)) => Some(style),
                    _ => None,
                };
                if let Some(next) = self.next_style(style.as_deref()) {
                    kept.insert("pStyle".to_owned(), Any::from(next));
                }
                kept
            }
        };
        // Only what differs from the copy is written.
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
        if !text.is_empty() {
            doc.insert_text(&ctx(), Position::new(BODY, at), text, FormatPolicy::Plain)
                .map_err(op)?;
        }
        let target_id = format!("{BODY}:paragraph:{new_id}");
        if after {
            self.after.insert(para_id.to_owned(), new_id);
        }
        Ok(Applied {
            ids: vec![target_id.clone()],
            inverse: undoable.then(|| Command::DeleteParagraph {
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
