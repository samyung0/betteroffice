//! Insert and Update table of contents, as Word builds one: the field
//! `TOC \o "1-3" \h \z \u`, one paragraph per body heading of outline level
//! 1–3 (styles named `toc 1`–`toc 3`), each entry a link to a `_Toc` bookmark
//! around its heading's text, a dot-leader right tab and a `PAGEREF \h` page
//! number the host reads off its layout. The entries are written as Word's
//! markup, read by docx-parse and seeded with the document's styles
//! ([`crate::seed::fragment_ops`]), so the story holds exactly what the saved
//! file seeds to. The field ends after the last entry, in its paragraph.

use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use yrs::{Any, Map, MapRef, Out, ReadTxn, Text, TextRef, TransactionMut};

use crate::op::{OpError, OpResult};
use crate::ops::field_changes::renumber_fields;
use crate::ops::{ChunkKind, snapshot};
use crate::seed::{PackageContext, fragment_ops, style_named, style_outline_level};
use crate::{EditCtx, EditingDoc, KIND_KEY, PARA_ID, Position, RawOp, map_string, story_ref};

/// Word's default field code.
pub const TOC_INSTRUCTION: &str = r#" TOC \o "1-3" \h \z \u "#;
const BODY: &str = "body";
/// Word's `toc 1`–`toc 9` paragraph spacing after and left indent step, in
/// twips, written directly when the document does not define the style.
const TOC_SPACE_AFTER: u32 = 100;
const TOC_INDENT_STEP: u32 = 220;

/// What the host reads off its layout for a table of contents.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct TocLayout {
    /// Each heading's page, as its page shows the number, by paraId.
    pub pages: HashMap<String, String>,
    /// The entries' right tab: the text width less 10 twips, as Word sets it.
    pub tab_twips: u32,
    /// The result when no heading qualifies ("No table of contents entries found.").
    pub empty_text: String,
}

/// A heading a table of contents lists.
#[derive(Clone, Debug, PartialEq)]
pub struct TocHeading {
    pub para_id: String,
    /// 1-based, as `\o "1-3"` counts.
    pub level: u8,
    pub text: String,
}

/// Where a written TOC lies: its first paragraph, and the paragraph after it.
#[derive(Clone, Debug, PartialEq)]
pub struct TocReceipt {
    pub first_para_id: String,
    pub next_para_id: String,
}

/// A TOC field in a story: `start..end` holds its code, result and markers.
#[derive(Clone, Debug, PartialEq)]
pub struct TocField {
    pub start: u32,
    pub end: u32,
    pub instruction: String,
}

/// One paragraph of a story, read in a transaction.
struct Para {
    id: String,
    /// After the leading tables, controls and breaks.
    node_start: u32,
    pilcrow: u32,
    map: MapRef,
    text: String,
}

fn paragraphs<T: ReadTxn>(txn: &T, story: &TextRef) -> Vec<Para> {
    let mut result = Vec::new();
    let mut node_start = None;
    let mut text = String::new();
    for chunk in snapshot(story, txn) {
        match &chunk.kind {
            ChunkKind::Pilcrow(map) => {
                result.push(Para {
                    id: map_string(map, txn, PARA_ID).unwrap_or_default(),
                    node_start: node_start.unwrap_or(chunk.start),
                    pilcrow: chunk.start,
                    map: map.clone(),
                    text: std::mem::take(&mut text),
                });
                node_start = None;
            }
            ChunkKind::Embed(map) => {
                let kind = map.as_ref().and_then(|map| map_string(map, txn, KIND_KEY));
                if node_start.is_none()
                    && kind.as_deref().is_some_and(crate::segments::is_block_embed)
                {
                    continue;
                }
                node_start.get_or_insert(chunk.start);
                if let (Some("field"), Some(map)) = (kind.as_deref(), map)
                    && !chunk.attr_active(crate::DEL)
                {
                    text.push_str(&map_string(map, txn, "displayText").unwrap_or_default());
                }
            }
            ChunkKind::Text(value) => {
                node_start.get_or_insert(chunk.start);
                if !chunk.attr_active(crate::DEL) {
                    text.push_str(value);
                }
            }
        }
    }
    result
}

fn number(value: Option<&Any>) -> Option<f64> {
    match value {
        Some(Any::Number(value)) => Some(*value),
        Some(Any::BigInt(value)) => Some(*value as f64),
        _ => None,
    }
}

/// The outline level (0-based) the saved paragraph has: its own, else its
/// style's. A seeded paragraph's own is its source formatting's; the stored
/// `outlineLevel` merged the style's in at seed and goes stale on a style change.
fn outline_level<T: ReadTxn>(
    txn: &T,
    map: &MapRef,
    package: Option<&PackageContext>,
) -> Option<u8> {
    let own = match map.get(txn, "_originalFormatting") {
        Some(Out::Any(Any::Map(original))) => number(original.get("outlineLevel")),
        _ => match map.get(txn, "outlineLevel") {
            Some(Out::Any(value)) => number(Some(&value)),
            _ => None,
        },
    };
    if let Some(level) = own {
        return (level < 9.0).then_some(level as u8);
    }
    let style = map_string(map, txn, "pStyle")?;
    match package {
        Some(package) => style_outline_level(package, &style),
        None => style
            .strip_prefix("Heading")
            .and_then(|digit| digit.parse::<u8>().ok())
            .filter(|level| (1..=9).contains(level))
            .map(|level| level - 1),
    }
    .filter(|level| *level < 9)
}

/// The levels `\o "a-b"` lists (all nine with a bare `\o`), whether `\h`
/// links the entries and whether `\n` drops the page numbers.
struct Switches {
    levels: (u8, u8),
    links: bool,
    numbers: bool,
}

fn switches(instruction: &str) -> Switches {
    let lower = instruction.to_ascii_lowercase();
    let levels = lower
        .find("\\o")
        .map(|at| {
            let rest = lower[at + 2..].trim_start();
            rest.strip_prefix('"')
                .and_then(|rest| rest.split('"').next())
                .and_then(|range| {
                    let (from, to) = range.split_once('-')?;
                    Some((from.trim().parse().ok()?, to.trim().parse().ok()?))
                })
                .unwrap_or((1, 9))
        })
        .unwrap_or((1, 3));
    Switches {
        levels,
        links: lower.contains("\\h"),
        numbers: !lower.contains("\\n"),
    }
}

/// A `w:t` as the save writes it: `xml:space` only where spaces need it.
fn text_element(text: &str) -> String {
    let preserve = text.starts_with(' ') || text.ends_with(' ') || text.contains("  ");
    format!(
        "<w:t{}>{}</w:t>",
        if preserve {
            r#" xml:space="preserve""#
        } else {
            ""
        },
        escape(text)
    )
}

fn escape(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

/// A number from `seed` (FNV-1a), below 2^31 as Word's ids are.
fn hash31(seed: &str) -> u32 {
    let mut hash = 0x811c_9dc5u32;
    for byte in seed.bytes() {
        hash = (hash ^ u32::from(byte)).wrapping_mul(0x0100_0193);
    }
    hash & 0x7fff_ffff
}

/// One entry: its heading's bookmark, level and text, and its page.
struct Entry<'a> {
    bookmark: &'a str,
    level: u8,
    text: &'a str,
    page: &'a str,
}

/// Word's markup for a TOC: one `w:p` per entry (paraIds `para_ids`), the
/// field opening the first and closing the last; with no entry, one
/// paragraph whose result is `empty_text`.
fn toc_xml(
    instruction: &str,
    entries: &[Entry<'_>],
    para_ids: &[String],
    tab_twips: u32,
    empty_text: &str,
    package: Option<&PackageContext>,
) -> String {
    let flags = switches(instruction);
    let fld = |kind: &str| format!(r#"<w:r><w:fldChar w:fldCharType="{kind}"/></w:r>"#);
    let instr = |code: &str| {
        format!(
            r#"<w:r><w:instrText xml:space="preserve">{}</w:instrText></w:r>"#,
            escape(code)
        )
    };
    let begin = format!("{}{}{}", fld("begin"), instr(instruction), fld("separate"));
    let mut xml = String::from(
        r#"<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">"#,
    );
    if entries.is_empty() {
        xml.push_str(&format!(
            r#"<w:p w14:paraId="{}">{begin}<w:r><w:rPr><w:b/><w:bCs/></w:rPr>{}</w:r>{}</w:p>"#,
            para_ids[0],
            text_element(empty_text),
            fld("end")
        ));
    }
    for (index, entry) in entries.iter().enumerate() {
        let name = format!("toc {}", entry.level);
        let defined = package.and_then(|package| style_named(package, &name));
        let style = defined
            .clone()
            .unwrap_or_else(|| format!("TOC{}", entry.level));
        let direct = if defined.is_some() {
            String::new()
        } else {
            let indent = TOC_INDENT_STEP * u32::from(entry.level - 1);
            format!(
                r#"<w:spacing w:after="{TOC_SPACE_AFTER}"/>{}"#,
                if indent > 0 {
                    format!(r#"<w:ind w:left="{indent}"/>"#)
                } else {
                    String::new()
                }
            )
        };
        let number = if flags.numbers {
            format!(
                r#"<w:r><w:tab/></w:r>{}{}{}<w:r>{}</w:r>{}"#,
                fld("begin"),
                instr(&format!(" PAGEREF {} \\h ", entry.bookmark)),
                fld("separate"),
                text_element(entry.page),
                fld("end")
            )
        } else {
            String::new()
        };
        let content = format!("<w:r>{}</w:r>{number}", text_element(entry.text));
        let content = if flags.links {
            format!(
                r#"<w:hyperlink w:anchor="{}">{content}</w:hyperlink>"#,
                entry.bookmark
            )
        } else {
            content
        };
        xml.push_str(&format!(
            r#"<w:p w14:paraId="{}"><w:pPr><w:pStyle w:val="{}"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="{tab_twips}"/></w:tabs>{direct}</w:pPr>{}{content}{}</w:p>"#,
            para_ids[index],
            escape(&style),
            if index == 0 { begin.as_str() } else { "" },
            if index + 1 == entries.len() {
                fld("end")
            } else {
                String::new()
            },
        ));
    }
    xml.push_str("</w:body>");
    xml
}

/// The TOC fields of `story`, in order. A field starts at its first
/// projected entry (or its own embed) and ends at its end marker, or right
/// after its embed when it closes in its own paragraph.
fn fields<T: ReadTxn>(txn: &T, story: &TextRef, story_id: &str) -> Vec<TocField> {
    let chunks = snapshot(story, txn);
    let mut markers: Option<Vec<(u32, Any)>> = None;
    let mut result = Vec::new();
    for (position, chunk) in chunks.iter().enumerate() {
        let ChunkKind::Embed(Some(map)) = &chunk.kind else {
            continue;
        };
        if map_string(map, txn, KIND_KEY).as_deref() != Some("field")
            || map_string(map, txn, "fieldType").as_deref() != Some("TOC")
            || chunk.attr_active(crate::format::FIELD_RESULT)
        {
            continue;
        }
        let mut start = chunk.start;
        if let Some((id, _)) = crate::ops::field_changes::projection(txn, chunk) {
            for before in chunks[..position].iter().rev() {
                match before.attrs.get(crate::format::FIELD_RESULT) {
                    Some(Any::Map(marker)) if number(marker.get("id")) == Some(id as f64) => {
                        start = before.start;
                    }
                    _ => break,
                }
            }
        }
        let end = match map_string(map, txn, "continuationId") {
            Some(id) => markers
                .get_or_insert_with(|| crate::bookmarks::positions(txn, story_id))
                .iter()
                .find(|(_, data)| {
                    matches!(data, Any::Map(data)
                        if matches!(data.get("id"), Some(Any::String(own)) if own.as_ref() == id)
                        && matches!(data.get("kind"), Some(Any::String(kind)) if kind.as_ref() == "fieldend"))
                })
                .map_or(chunk.end(), |(at, _)| *at),
            None => chunk.end(),
        };
        result.push(TocField {
            start,
            end: end.max(chunk.end()),
            instruction: map_string(map, txn, "instruction").unwrap_or_default(),
        });
    }
    result
}

impl EditingDoc {
    /// The body's headings a table of contents lists, in order.
    pub fn toc_headings(&self) -> OpResult<Vec<TocHeading>> {
        let package = self.package();
        let txn = yrs::Transact::transact(self.yrs_doc());
        let story = story_ref(&txn, BODY)?;
        Ok(headings(&txn, &story, package.as_deref(), (1, 3)))
    }

    /// The TOC fields of `story_id`.
    pub fn toc_fields(&self, story_id: &str) -> OpResult<Vec<TocField>> {
        let txn = yrs::Transact::transact(self.yrs_doc());
        let story = story_ref(&txn, story_id)?;
        Ok(fields(&txn, &story, story_id))
    }

    /// Inserts a table of contents at `at` in the body. A paragraph with text
    /// before `at` is split there first, so the TOC takes paragraphs of its
    /// own: what came before stays above it, the rest follows it.
    pub fn insert_toc(
        &self,
        ctx: &EditCtx,
        at: Position,
        layout: &TocLayout,
    ) -> OpResult<TocReceipt> {
        if at.story != BODY {
            return Err(OpError::NotTopLevel);
        }
        let para = self
            .segment_index(BODY)?
            .para_at(at.index)
            .map(|para| (para.node_start, para.pilcrow))
            .ok_or(OpError::OutOfBounds {
                index: at.index,
                len: self.story_len(BODY)?,
            })?;
        let index = if at.index > para.0 {
            self.split_paragraph(ctx, at.clone(), None)?;
            at.index + 1
        } else {
            para.0
        };
        let mut txn = self.transact_for(ctx);
        let story = story_ref(&txn, BODY)?;
        self.write_toc(&mut txn, &story, index, TOC_INSTRUCTION, layout)
    }

    /// Rebuilds the TOC holding `at`, else the body's first, from the
    /// current headings and `layout`, keeping its field code: Word's
    /// "Update entire table", so edits made inside it are replaced. Returns
    /// `None` when the body has no TOC. A first paragraph with text before
    /// the field is split there first.
    pub fn update_toc(
        &self,
        ctx: &EditCtx,
        at: Option<&Position>,
        layout: &TocLayout,
    ) -> OpResult<Option<TocReceipt>> {
        let pick = |doc: &Self| -> OpResult<Option<TocField>> {
            let found = doc.toc_fields(BODY)?;
            let inside = at.filter(|at| at.story == BODY).and_then(|at| {
                found
                    .iter()
                    .find(|field| field.start <= at.index && at.index <= field.end)
            });
            Ok(inside.or(found.first()).cloned())
        };
        let Some(field) = pick(self)? else {
            return Ok(None);
        };
        let head = self
            .segment_index(BODY)?
            .para_at(field.start)
            .map(|para| para.node_start)
            .unwrap_or(field.start);
        let field = if head < field.start {
            self.split_paragraph(ctx, Position::new(BODY, field.start), None)?;
            // The new mark went in ahead of the field, which now opens the next paragraph.
            let Some(field) = self
                .toc_fields(BODY)?
                .into_iter()
                .find(|moved| moved.start == field.start + 1)
            else {
                return Ok(None);
            };
            field
        } else {
            field
        };
        let mut txn = self.transact_for(ctx);
        let story = story_ref(&txn, BODY)?;
        let paras = paragraphs(&txn, &story);
        let last = paras.iter().find(|para| para.pilcrow >= field.end).ok_or(
            OpError::ExpectedPilcrow {
                story: BODY.to_owned(),
                index: field.end,
            },
        )?;
        // The field's paragraphs go whole; text after its end stays in the
        // last one's own paragraph, now first after the new entries. An end
        // that opened that paragraph (Word's shape) ran ahead of its fields in
        // the saved paragraph; one later in it is left unnumbered.
        let (remove_to, kept) = if field.end == last.pilcrow {
            (last.pilcrow + 1, None)
        } else {
            (
                field.end,
                (field.end == last.node_start).then_some(last.pilcrow),
            )
        };
        for continuation in fields_continuations(&txn, &story, field.start, field.end) {
            remove_field_markers(&mut txn, &continuation);
        }
        story.remove_range(&mut txn, field.start, remove_to - field.start);
        if let Some(pilcrow) = kept {
            let pilcrow = pilcrow - (remove_to - field.start);
            renumber_fields(&mut txn, &story, field.start, pilcrow, -1);
        }
        self.write_toc(
            &mut txn,
            &story,
            field.start,
            &field.instruction_code(),
            layout,
        )
        .map(Some)
    }

    /// Writes a TOC's paragraphs at `index` (a paragraph's content start)
    /// and bookmarks every heading it lists.
    fn write_toc(
        &self,
        txn: &mut TransactionMut<'_>,
        story: &TextRef,
        index: u32,
        instruction: &str,
        layout: &TocLayout,
    ) -> OpResult<TocReceipt> {
        let package = self.package();
        let package = package.as_deref();
        let flags = switches(instruction);
        let listed = headings(txn, story, package, flags.levels);
        let paras = paragraphs(txn, story);
        let marked = crate::bookmarks::paragraph_properties(txn, BODY, story);
        let mut taken: HashSet<u32> = crate::bookmarks::positions(txn, BODY)
            .into_iter()
            .filter_map(|(_, data)| match data {
                Any::Map(data) => number(data.get("id")).map(|id| id as u32),
                _ => None,
            })
            .collect();
        let mut fresh = |doc: &Self| loop {
            let id = hash31(&doc.next_id());
            if id > 0 && taken.insert(id) {
                return id;
            }
        };
        let mut bookmarks = Vec::with_capacity(listed.len());
        for heading in &listed {
            let existing = marked.get(&heading.para_id).and_then(|marks| match marks {
                Any::Array(marks) => marks.iter().find_map(|mark| match mark {
                    Any::Map(mark)
                        if matches!(mark.get("kind"), Some(Any::String(kind)) if kind.as_ref() == "start") =>
                    {
                        match mark.get("name") {
                            Some(Any::String(name)) if name.starts_with("_Toc") => {
                                Some(name.to_string())
                            }
                            _ => None,
                        }
                    }
                    _ => None,
                }),
                _ => None,
            });
            if let Some(name) = existing {
                bookmarks.push(name);
                continue;
            }
            let para = paras
                .iter()
                .find(|para| para.id == heading.para_id)
                .ok_or_else(|| OpError::UnknownPara(heading.para_id.clone()))?;
            let id = fresh(self);
            let name = format!("_Toc{:09}", id % 1_000_000_000);
            // The start runs ahead of the heading's fields in its saved paragraph.
            renumber_fields(txn, story, para.node_start, para.pilcrow, 1);
            for (kind, at) in [("start", para.node_start), ("end", para.pilcrow)] {
                let mut data = HashMap::from([
                    ("id".to_owned(), Any::Number(f64::from(id))),
                    ("kind".to_owned(), Any::from(kind)),
                ]);
                if kind == "start" {
                    data.insert("name".to_owned(), Any::from(name.as_str()));
                }
                crate::bookmarks::set(txn, story, BODY, at, Any::Map(Arc::new(data)))?;
            }
            bookmarks.push(name);
        }
        let existing_ids: HashSet<String> = paras.iter().map(|para| para.id.clone()).collect();
        let para_ids: Vec<String> = (0..listed.len().max(1))
            .map(|_| {
                loop {
                    let id = format!("{:08X}", hash31(&self.next_id()));
                    if !existing_ids.contains(&id) && id != "00000000" {
                        return id;
                    }
                }
            })
            .collect();
        let entries: Vec<Entry<'_>> = listed
            .iter()
            .zip(&bookmarks)
            .map(|(heading, bookmark)| Entry {
                bookmark,
                level: heading.level,
                text: &heading.text,
                page: layout
                    .pages
                    .get(&heading.para_id)
                    .map_or("", String::as_str),
            })
            .collect();
        let xml = toc_xml(
            instruction,
            &entries,
            &para_ids,
            layout.tab_twips,
            &layout.empty_text,
            package,
        );
        let blocks = docx_parse::parse_body_fragment(&xml, package.map(|package| &package.theme))
            .map_err(|error| OpError::InvalidUpdate(error.to_string()))?;
        let ops = fragment_ops(&blocks, package).map_err(OpError::InvalidUpdate)?;
        let shifted = ops
            .into_iter()
            .map(|op| match op {
                RawOp::Insert {
                    index: at,
                    text,
                    attrs,
                } => RawOp::Insert {
                    index: at + index,
                    text,
                    attrs,
                },
                RawOp::InsertEmbed {
                    index: at,
                    kind,
                    payload,
                    attrs,
                } => RawOp::InsertEmbed {
                    index: at + index,
                    kind,
                    payload,
                    attrs,
                },
                RawOp::SetBookmark { index: at, data } => RawOp::SetBookmark {
                    index: at + index,
                    data,
                },
                other => other,
            })
            .collect();
        let next_para_id = paras
            .iter()
            .find(|para| para.pilcrow >= index)
            .map(|para| para.id.clone())
            .ok_or(OpError::ExpectedPilcrow {
                story: BODY.to_owned(),
                index,
            })?;
        crate::raw::apply_raw_ops_to_story(txn, BODY, shifted, false)?;
        Ok(TocReceipt {
            first_para_id: para_ids[0].clone(),
            next_para_id,
        })
    }
}

impl TocField {
    /// The field code as written, with Word's spacing around it.
    fn instruction_code(&self) -> String {
        format!(" {} ", self.instruction.trim())
    }
}

fn headings<T: ReadTxn>(
    txn: &T,
    story: &TextRef,
    package: Option<&PackageContext>,
    (from, to): (u8, u8),
) -> Vec<TocHeading> {
    paragraphs(txn, story)
        .into_iter()
        .filter_map(|para| {
            let level = outline_level(txn, &para.map, package)? + 1;
            let text = para
                .text
                .chars()
                .map(|ch| if ch.is_control() { ' ' } else { ch })
                .collect::<String>()
                .trim()
                .to_owned();
            ((from..=to).contains(&level) && !text.is_empty()).then_some(TocHeading {
                para_id: para.id,
                level,
                text,
            })
        })
        .collect()
}

/// The continuation ids of the continued fields embedded in `start..end`.
fn fields_continuations<T: ReadTxn>(txn: &T, story: &TextRef, start: u32, end: u32) -> Vec<String> {
    crate::ops::snapshot_range(story, txn, start, end)
        .into_iter()
        .filter_map(|chunk| match &chunk.kind {
            ChunkKind::Embed(Some(map)) => map_string(map, txn, "continuationId"),
            _ => None,
        })
        .collect()
}

/// Drops a removed continued field's separate and end markers.
fn remove_field_markers(txn: &mut TransactionMut<'_>, continuation: &str) {
    let Some(root) = txn.get_map(crate::bookmarks::ROOT) else {
        return;
    };
    for kind in ["fieldseparate", "fieldend"] {
        root.remove(txn, &format!("{BODY}:{continuation}:{kind}"));
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{EditCtx, SegmentContent, StoryRange, seed_from_docx};

    const NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#;

    fn styles(toc_styles: bool) -> String {
        let heading = |n: u8| {
            format!(
                r#"<w:style w:type="paragraph" w:styleId="Heading{n}"><w:name w:val="heading {n}"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="{}"/></w:pPr></w:style>"#,
                n - 1
            )
        };
        let toc = |n: u8| {
            format!(
                r#"<w:style w:type="paragraph" w:styleId="TM{n}"><w:name w:val="toc {n}"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="100"/></w:pPr></w:style>"#
            )
        };
        format!(
            r#"<w:styles {NS}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>{}{}</w:styles>"#,
            (1..=4).map(heading).collect::<String>(),
            if toc_styles {
                (1..=3).map(toc).collect::<String>()
            } else {
                String::new()
            }
        )
    }

    fn package(body: &str, toc_styles: bool) -> Vec<u8> {
        const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        let parts = [
            ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"#.to_owned()),
            ("_rels/.rels", format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="word/document.xml"/></Relationships>"#)),
            ("word/_rels/document.xml.rels", format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="{REL}/styles" Target="styles.xml"/></Relationships>"#)),
            ("word/styles.xml", styles(toc_styles)),
            ("word/document.xml", format!(r#"<w:document {NS}><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>"#)),
        ];
        ooxml_opc::rezip_parts(
            &parts
                .into_iter()
                .map(|(name, value)| (name.to_owned(), value.into_bytes()))
                .collect::<Vec<_>>(),
        )
        .unwrap()
    }

    fn p(id: &str, style: Option<&str>, text: &str) -> String {
        let ppr = style.map_or(String::new(), |style| {
            format!(r#"<w:pPr><w:pStyle w:val="{style}"/></w:pPr>"#)
        });
        format!(
            r#"<w:p w14:paraId="{id}">{ppr}<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>"#
        )
    }

    /// An empty first paragraph, then H1 Introduction, text, H2 Background,
    /// H4 Deep (not listed), H3 Details.
    fn body() -> String {
        [
            p("10000001", None, ""),
            p("10000002", Some("Heading1"), "Introduction"),
            p("10000003", None, "Some text."),
            p("10000004", Some("Heading2"), "Background"),
            p("10000005", Some("Heading4"), "Deep"),
            p("10000006", Some("Heading3"), "Details"),
        ]
        .concat()
    }

    fn open(bytes: &[u8], client: u64) -> EditingDoc {
        let doc = EditingDoc::new(client);
        seed_from_docx(&doc, bytes).unwrap();
        doc
    }

    fn ctx() -> EditCtx {
        EditCtx::local("Tester", "2026-10-06T00:00:00Z")
    }

    fn layout(pages: &[(&str, &str)]) -> TocLayout {
        TocLayout {
            pages: pages
                .iter()
                .map(|(id, page)| ((*id).to_owned(), (*page).to_owned()))
                .collect(),
            tab_twips: 9350,
            empty_text: "No table of contents entries found.".to_owned(),
        }
    }

    const PAGES: [(&str, &str); 3] = [("10000002", "2"), ("10000004", "3"), ("10000006", "4")];

    /// The body as text, `[code|shown]` per field, `<` after a field's projected text, ¶ per paragraph.
    fn units(doc: &EditingDoc) -> String {
        doc.story_segments(BODY)
            .unwrap()
            .into_iter()
            .map(|segment| match segment.content {
                SegmentContent::Text(text) => {
                    if segment.attributes.contains_key("fieldResult") {
                        format!("{text}<")
                    } else {
                        text
                    }
                }
                SegmentContent::Pilcrow(_) => "¶".to_owned(),
                SegmentContent::OtherEmbed { kind, payload } if kind == "field" => format!(
                    "[{}|{}]",
                    match payload.get("instruction") {
                        Some(Any::String(code)) => code.to_string(),
                        _ => String::new(),
                    },
                    match payload.get("displayText") {
                        Some(Any::String(text)) => text.to_string(),
                        _ => String::new(),
                    }
                ),
                SegmentContent::OtherEmbed { kind, .. } => format!("[{kind}]"),
            })
            .collect()
    }

    /// Where the `nth` (0-based) whole-paragraph-text or in-text match of `needle` starts.
    fn nth(doc: &EditingDoc, needle: &str, nth: usize) -> u32 {
        let mut at = 0u32;
        let mut seen = 0;
        for segment in doc.story_segments(BODY).unwrap() {
            match segment.content {
                SegmentContent::Text(text) => {
                    for (found, _) in text.match_indices(needle) {
                        if seen == nth {
                            return at + text[..found].encode_utf16().count() as u32;
                        }
                        seen += 1;
                    }
                    at += text.encode_utf16().count() as u32;
                }
                _ => at += 1,
            }
        }
        panic!("{needle} #{nth} not found");
    }

    fn index_of(doc: &EditingDoc, needle: &str) -> u32 {
        nth(doc, needle, 0)
    }

    /// The body's segments as sorted JSON, so map order never differs.
    fn canonical(doc: &EditingDoc) -> Vec<String> {
        let json = |value: &Any| {
            let mut text = String::new();
            value.to_json(&mut text);
            serde_json::from_str::<serde_json::Value>(&text)
                .unwrap()
                .to_string()
        };
        doc.story_segments(BODY)
            .unwrap()
            .into_iter()
            .map(|segment| {
                let attributes: String = segment
                    .attributes
                    .iter()
                    .map(|(key, value)| format!("{key}={}", json(value)))
                    .collect();
                let content = match segment.content {
                    SegmentContent::Text(text) => text,
                    SegmentContent::Pilcrow(properties) => format!(
                        "¶{}{:?}",
                        properties.para_id,
                        properties
                            .values
                            .iter()
                            .map(|(key, value)| format!("{key}={}", json(value)))
                            .collect::<Vec<_>>()
                    ),
                    SegmentContent::OtherEmbed { kind, payload } => format!(
                        "[{kind}{:?}]",
                        payload
                            .iter()
                            .map(|(key, value)| format!("{key}={}", json(value)))
                            .collect::<Vec<_>>()
                    ),
                };
                format!("{content} {attributes}")
            })
            .collect()
    }

    /// Each heading's `_Toc` bookmark name, by paraId.
    fn toc_bookmarks(doc: &EditingDoc) -> HashMap<String, String> {
        doc.paragraphs(BODY)
            .unwrap()
            .into_iter()
            .filter_map(|para| {
                let Some(Any::Array(marks)) = para.properties.get("bookmarks") else {
                    return None;
                };
                marks.iter().find_map(|mark| match mark {
                    Any::Map(mark) => match mark.get("name") {
                        Some(Any::String(name)) if name.starts_with("_Toc") => {
                            Some((para.para_id.clone(), name.to_string()))
                        }
                        _ => None,
                    },
                    _ => None,
                })
            })
            .collect()
    }

    #[test]
    fn word_markup_for_two_entries() {
        let entries = [
            Entry {
                bookmark: "_Toc1",
                level: 1,
                text: "A & B",
                page: "1",
            },
            Entry {
                bookmark: "_Toc2",
                level: 2,
                text: "C",
                page: "ii",
            },
        ];
        let ids = ["0000000A".to_owned(), "0000000B".to_owned()];
        let xml = toc_xml(TOC_INSTRUCTION, &entries, &ids, 9350, "", None);
        let begin = r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#;
        let separate = r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#;
        let end = r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#;
        let tabs = r#"<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs>"#;
        let entry = |bm: &str, text: &str, page: &str| {
            format!(
                r#"<w:hyperlink w:anchor="{bm}"><w:r><w:t>{text}</w:t></w:r><w:r><w:tab/></w:r>{begin}<w:r><w:instrText xml:space="preserve"> PAGEREF {bm} \h </w:instrText></w:r>{separate}<w:r><w:t>{page}</w:t></w:r>{end}</w:hyperlink>"#
            )
        };
        assert_eq!(
            xml,
            format!(
                r#"<w:body {NS}><w:p w14:paraId="0000000A"><w:pPr><w:pStyle w:val="TOC1"/>{tabs}<w:spacing w:after="100"/></w:pPr>{begin}<w:r><w:instrText xml:space="preserve"> TOC \o &quot;1-3&quot; \h \z \u </w:instrText></w:r>{separate}{}</w:p><w:p w14:paraId="0000000B"><w:pPr><w:pStyle w:val="TOC2"/>{tabs}<w:spacing w:after="100"/><w:ind w:left="220"/></w:pPr>{}{end}</w:p></w:body>"#,
                entry("_Toc1", "A &amp; B", "1"),
                entry("_Toc2", "C", "ii"),
            )
        );
    }

    /// The body's units with each `_Toc` bookmark named after its heading's paraId.
    fn named(doc: &EditingDoc) -> String {
        let mut text = units(doc);
        for (id, name) in toc_bookmarks(doc) {
            text = text.replace(&format!("{name} "), &format!("_Toc{id} "));
        }
        text
    }

    /// `body()` with each heading wrapped in the bookmark `doc` gave it, after `toc`.
    fn expected_package(doc: &EditingDoc, toc: &str, toc_styles: bool) -> Vec<u8> {
        let marks = toc_bookmarks(doc);
        let ids: HashMap<String, f64> = doc
            .paragraphs(BODY)
            .unwrap()
            .into_iter()
            .filter_map(|para| match para.properties.get("bookmarks") {
                Some(Any::Array(list)) => list.iter().find_map(|mark| match mark {
                    Any::Map(mark) => number(mark.get("id")).map(|id| (para.para_id.clone(), id)),
                    _ => None,
                }),
                _ => None,
            })
            .collect();
        let mut body = body();
        for (para, name) in &marks {
            let id = ids[para];
            let open = format!(r#"<w:p w14:paraId="{para}">"#);
            let at = body.find(&open).unwrap();
            let run = body[at..].find("<w:r>").unwrap() + at;
            let close = body[at..].find("</w:p>").unwrap() + at;
            body.insert_str(close, &format!(r#"<w:bookmarkEnd w:id="{id}"/>"#));
            body.insert_str(
                run,
                &format!(r#"<w:bookmarkStart w:id="{id}" w:name="{name}"/>"#),
            );
        }
        package(&format!("{toc}{body}"), toc_styles)
    }

    fn toc_markup(
        doc: &EditingDoc,
        toc_styles: bool,
        entries: &[(&str, u8, &str, &str)],
    ) -> String {
        let marks = toc_bookmarks(doc);
        let ids: Vec<String> = doc
            .paragraphs(BODY)
            .unwrap()
            .into_iter()
            .take(entries.len().max(1))
            .map(|para| para.para_id)
            .collect();
        let entries: Vec<Entry<'_>> = entries
            .iter()
            .map(|(para, level, text, page)| Entry {
                bookmark: &marks[*para],
                level: *level,
                text,
                page,
            })
            .collect();
        let package = doc.package();
        let xml = toc_xml(
            TOC_INSTRUCTION,
            &entries,
            &ids,
            9350,
            "No table of contents entries found.",
            package.as_deref().filter(|_| toc_styles),
        );
        xml[xml.find('>').unwrap() + 1..xml.len() - "</w:body>".len()].to_owned()
    }

    const ENTRIES: [(&str, u8, &str, &str); 3] = [
        ("10000002", 1, "Introduction", "2"),
        ("10000004", 2, "Background", "3"),
        ("10000006", 3, "Details", "4"),
    ];

    /// The inserted TOC is what the seed of a file holding Word's markup for
    /// it gives (the save writes that markup back), listing headings 1–3
    /// linked to `_Toc` bookmarks around their text.
    #[test]
    fn insert_seeds_as_its_markup() {
        for toc_styles in [false, true] {
            let bytes = package(&body(), toc_styles);
            let doc = open(&bytes, 11);
            doc.insert_toc(&ctx(), Position::new(BODY, 0), &layout(&PAGES))
                .unwrap();
            assert_eq!(
                named(&doc),
                r#"Introduction	<[PAGEREF _Toc10000002 \h|2][TOC \o "1-3" \h \z \u|]¶Background	[PAGEREF _Toc10000004 \h|3]¶Details	[PAGEREF _Toc10000006 \h|4]¶¶Introduction¶Some text.¶Background¶Deep¶Details¶"#
            );
            let styles: Vec<_> = doc
                .paragraphs(BODY)
                .unwrap()
                .into_iter()
                .take(3)
                .map(|para| match para.properties.get("pStyle") {
                    Some(Any::String(style)) => style.to_string(),
                    _ => String::new(),
                })
                .collect();
            let prefix = if toc_styles { "TM" } else { "TOC" };
            assert_eq!(styles, [1, 2, 3].map(|n| format!("{prefix}{n}")));
            let reopened = open(
                &expected_package(&doc, &toc_markup(&doc, toc_styles, &ENTRIES), toc_styles),
                12,
            );
            assert_eq!(canonical(&doc), canonical(&reopened));
            assert_eq!(
                doc.toc_fields(BODY).unwrap(),
                vec![TocField {
                    start: 0,
                    end: index_of(&doc, "Details\t") + 9,
                    instruction: r#"TOC \o "1-3" \h \z \u"#.to_owned(),
                }]
            );
        }
    }

    /// Insert in a paragraph with text before the caret splits it there.
    #[test]
    fn insert_mid_paragraph_takes_its_own_paragraphs() {
        let bytes = package(
            &format!("{}{}", p("20000001", None, "Hello world"), body()),
            false,
        );
        let doc = open(&bytes, 13);
        let at = index_of(&doc, "world");
        doc.insert_toc(&ctx(), Position::new(BODY, at), &layout(&PAGES))
            .unwrap();
        assert!(named(&doc).starts_with(
            r#"Hello ¶Introduction	<[PAGEREF _Toc10000002 \h|2][TOC \o "1-3" \h \z \u|]¶Background"#
        ));
        assert!(named(&doc).contains("[PAGEREF _Toc10000006 \\h|4]¶world¶¶Introduction"));
    }

    #[test]
    fn insert_without_headings_shows_words_message() {
        let bytes = package(&p("10000001", None, ""), false);
        let doc = open(&bytes, 14);
        doc.insert_toc(&ctx(), Position::new(BODY, 0), &layout(&[]))
            .unwrap();
        assert_eq!(
            units(&doc),
            r#"[TOC \o "1-3" \h \z \u|No table of contents entries found.]¶¶"#
        );
        assert_eq!(doc.toc_fields(BODY).unwrap().len(), 1);
    }

    /// Update rebuilds every entry from the headings as they are now: a
    /// renamed heading, new page numbers, a hand edit inside the TOC gone.
    #[test]
    fn update_rebuilds_the_entries() {
        let bytes = package(&body(), false);
        let doc = open(&bytes, 15);
        doc.insert_toc(&ctx(), Position::new(BODY, 0), &layout(&PAGES))
            .unwrap();
        let marks = toc_bookmarks(&doc);
        let typed = index_of(&doc, "Backg");
        doc.insert_text(&ctx(), Position::new(BODY, typed), "XX", Default::default())
            .unwrap();
        let heading = nth(&doc, "Background", 1);
        doc.replace_range(
            &ctx(),
            StoryRange::new(BODY, heading, heading + 10),
            "Context",
        )
        .unwrap();
        let caret = Position::new(BODY, index_of(&doc, "Some text"));
        assert!(
            doc.update_toc(
                &ctx(),
                Some(&caret),
                &layout(&[("10000002", "3"), ("10000004", "4"), ("10000006", "5")])
            )
            .unwrap()
            .is_some()
        );
        assert_eq!(
            named(&doc),
            r#"Introduction	<[PAGEREF _Toc10000002 \h|3][TOC \o "1-3" \h \z \u|]¶Context	[PAGEREF _Toc10000004 \h|4]¶Details	[PAGEREF _Toc10000006 \h|5]¶¶Introduction¶Some text.¶Context¶Deep¶Details¶"#
        );
        assert_eq!(toc_bookmarks(&doc), marks, "bookmarks are reused");
        assert_eq!(doc.toc_fields(BODY).unwrap().len(), 1);
    }

    /// Update keeps a Word TOC's own field code and levels, and its shape
    /// (the end opening the paragraph after the entries) becomes Insert's.
    #[test]
    fn update_keeps_the_field_code() {
        let fld = |kind: &str| format!(r#"<w:r><w:fldChar w:fldCharType="{kind}"/></w:r>"#);
        let code =
            r#"<w:r><w:instrText xml:space="preserve"> TOC \o "1-2" \h \z \u </w:instrText></w:r>"#;
        let entry = |bm: &str, text: &str| {
            format!(
                r#"<w:hyperlink w:anchor="{bm}"><w:r><w:t>{text}</w:t></w:r><w:r><w:tab/></w:r>{}<w:r><w:instrText xml:space="preserve"> PAGEREF {bm} \h </w:instrText></w:r>{}<w:r><w:t>9</w:t></w:r>{}</w:hyperlink>"#,
                fld("begin"),
                fld("separate"),
                fld("end")
            )
        };
        let toc = format!(
            r#"<w:p w14:paraId="30000001"><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>{}{code}{}{}</w:p><w:p w14:paraId="30000002"><w:pPr><w:pStyle w:val="TOC2"/></w:pPr>{}</w:p><w:p w14:paraId="30000003">{}<w:r><w:t>after</w:t></w:r></w:p>"#,
            fld("begin"),
            fld("separate"),
            entry("_Toc1", "Old one"),
            entry("_Toc2", "Old two"),
            fld("end"),
        );
        let headings = format!(
            r#"<w:p w14:paraId="10000002"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="1" w:name="_Toc1"/><w:r><w:t>Introduction</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>{}{}"#,
            p("10000004", Some("Heading2"), "Background"),
            p("10000006", Some("Heading3"), "Details"),
        );
        let doc = open(&package(&format!("{toc}{headings}"), false), 16);
        assert!(
            doc.update_toc(&ctx(), None, &layout(&PAGES))
                .unwrap()
                .is_some()
        );
        assert_eq!(
            named(&doc),
            r#"Introduction	<[PAGEREF _Toc10000002 \h|2][TOC \o "1-2" \h \z \u|]¶Background	[PAGEREF _Toc10000004 \h|3]¶after¶Introduction¶Background¶Details¶"#
        );
        assert_eq!(toc_bookmarks(&doc)["10000002"], "_Toc1");
    }

    /// Text before the field (the matrix's "Contents" shape, the end inside
    /// the last entry) is split off; the new field opens its own paragraph.
    #[test]
    fn update_splits_off_text_before_the_field() {
        let fld = |kind: &str| format!(r#"<w:r><w:fldChar w:fldCharType="{kind}"/></w:r>"#);
        let entry = |bm: &str, text: &str| {
            format!(
                r#"<w:hyperlink w:anchor="{bm}"><w:r><w:t>{text}</w:t></w:r><w:r><w:tab/></w:r>{}<w:r><w:instrText xml:space="preserve"> PAGEREF {bm} \h </w:instrText></w:r>{}<w:r><w:t>9</w:t></w:r>{}</w:hyperlink>"#,
                fld("begin"),
                fld("separate"),
                fld("end")
            )
        };
        let toc = format!(
            r#"<w:p w14:paraId="30000001"><w:r><w:t>Contents</w:t><w:tab/></w:r>{}<w:r><w:instrText xml:space="preserve"> TOC \o "1-3" \h \z \u </w:instrText></w:r>{}{}</w:p><w:p w14:paraId="30000002">{}{}</w:p>"#,
            fld("begin"),
            fld("separate"),
            entry("_Toc1", "Old one"),
            entry("_Toc2", "Old two"),
            fld("end"),
        );
        let doc = open(&package(&format!("{toc}{}", body()), false), 17);
        let caret = Position::new(BODY, index_of(&doc, "Old two"));
        assert!(
            doc.update_toc(&ctx(), Some(&caret), &layout(&PAGES))
                .unwrap()
                .is_some()
        );
        assert_eq!(
            named(&doc),
            r#"Contents	¶Introduction	<[PAGEREF _Toc10000002 \h|2][TOC \o "1-3" \h \z \u|]¶Background	[PAGEREF _Toc10000004 \h|3]¶Details	[PAGEREF _Toc10000006 \h|4]¶¶Introduction¶Some text.¶Background¶Deep¶Details¶"#
        );
        assert_eq!(doc.toc_fields(BODY).unwrap().len(), 1);
    }

    fn sync(a: &EditingDoc, b: &EditingDoc) {
        let to_b = a.encode_diff_v1(&b.encode_state_vector_v1()).unwrap();
        let to_a = b.encode_diff_v1(&a.encode_state_vector_v1()).unwrap();
        a.apply_update_v1(&to_a).unwrap();
        b.apply_update_v1(&to_b).unwrap();
    }

    /// One peer updates while another renames a heading or adds one: both
    /// converge on the update's entries and the other peer's heading.
    #[test]
    fn two_peers_converge() {
        let bytes = package(&body(), false);
        for rename in [true, false] {
            let a = open(&bytes, 21);
            a.insert_toc(&ctx(), Position::new(BODY, 0), &layout(&PAGES))
                .unwrap();
            let b = open(&bytes, 22);
            sync(&a, &b);
            if rename {
                let at = nth(&b, "Details", 1);
                b.replace_range(&ctx(), StoryRange::new(BODY, at, at + 7), "Summary")
                    .unwrap();
            } else {
                let at = index_of(&b, "Some text.") + 10;
                b.split_paragraph(&ctx(), Position::new(BODY, at), None)
                    .unwrap();
                b.insert_text(
                    &ctx(),
                    Position::new(BODY, at + 1),
                    "Method",
                    Default::default(),
                )
                .unwrap();
                let id = b.paragraphs(BODY).unwrap()[5].para_id.clone();
                b.set_paragraph_attrs(&ctx(), &crate::ParaSelector::One(id), &{
                    let mut delta = crate::ParaAttrDelta::default();
                    delta
                        .other
                        .insert("pStyle".to_owned(), Some(Any::from("Heading2")));
                    delta
                })
                .unwrap();
            }
            a.update_toc(&ctx(), None, &layout(&PAGES)).unwrap();
            sync(&a, &b);
            assert_eq!(canonical(&a), canonical(&b));
            let text = named(&a);
            if rename {
                assert!(text.ends_with("Background¶Deep¶Summary¶"), "{text}");
                assert!(text.contains("Details	[PAGEREF"), "{text}");
            } else {
                assert!(text.contains("Some text.¶Method¶Background"), "{text}");
                assert!(
                    !text[..text.find("¶¶").unwrap()].contains("Method"),
                    "{text}"
                );
            }
            assert_eq!(a.toc_fields(BODY).unwrap().len(), 1);
        }
    }

    /// Insert, and Update, are one Undo step.
    #[test]
    fn undo_restores_in_one_step() {
        let bytes = package(
            &format!("{}{}", p("20000001", None, "Hello world"), body()),
            false,
        );
        let doc = open(&bytes, 31);
        let before = units(&doc);
        let mut undo = doc.undo_manager();
        let at = index_of(&doc, "world");
        doc.insert_toc(&ctx(), Position::new(BODY, at), &layout(&PAGES))
            .unwrap();
        undo.add_undo_barrier();
        let inserted = units(&doc);
        doc.update_toc(&ctx(), None, &layout(&[("10000002", "7")]))
            .unwrap();
        let fields = doc.toc_fields(BODY).unwrap();
        assert!(undo.undo());
        assert_eq!(units(&doc), inserted);
        assert!(undo.undo());
        assert_eq!(units(&doc), before);
        assert!(doc.toc_fields(BODY).unwrap().is_empty());
        // Redo brings back the field with its end where it was.
        assert!(undo.redo());
        assert!(undo.redo());
        assert_eq!(doc.toc_fields(BODY).unwrap(), fields);
    }
}
