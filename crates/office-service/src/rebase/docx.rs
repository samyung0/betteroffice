//! office-rebase.ts `DOCX_LINEAGE` (anchor rewrite, `settle`, `check`),
//! `fieldAttributes`, `keepContinuation`, `docxIds` and
//! `docxWrittenParagraphs` on yrs.

use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use std::sync::Arc;

use yrs::types::ToJson;
use yrs::types::text::YChange;
use yrs::updates::decoder::Decode;
use yrs::updates::encoder::Encode;
use yrs::{
    Any, Assoc, IndexedSequence, Map, MapRef, Out, ReadTxn, StickyIndex, Text, TextRef,
    TransactionMut,
};

use super::align::{
    Alignment, Key, PILCROW, align, any_string, child_at_offset, embeds, items, past, point_at,
    unit_attributes, units,
};
use super::fail;
use super::transplant::{Attrs, Copied, DOCX_LINEAGE, open, ordered_keys, root_map};
use crate::error::Result;
use crate::jsv::{V, stringify};

fn get_string<T: ReadTxn>(txn: &T, map: &MapRef, key: &str) -> Option<String> {
    match map.get(txn, key) {
        Some(Out::Any(Any::String(text))) => Some(text.to_string()),
        _ => None,
    }
}

fn any_field<'a>(value: &'a Any, key: &str) -> Option<&'a Any> {
    match value {
        Any::Map(map) => map.get(key),
        _ => None,
    }
}

fn any_str(value: Option<&Any>) -> Option<&str> {
    match value {
        Some(Any::String(text)) => Some(text),
        _ => None,
    }
}

fn any_number(value: Option<&Any>) -> Option<f64> {
    match value {
        Some(Any::Number(number)) => Some(*number),
        Some(Any::BigInt(number)) => Some(*number as f64),
        _ => None,
    }
}

fn any_bytes(value: Option<&Any>) -> Option<&[u8]> {
    match value {
        Some(Any::Buffer(bytes)) => Some(bytes),
        _ => None,
    }
}

/// `Y.createAbsolutePositionFromRelativePosition(decode(bytes), doc)`.
fn resolve<T: ReadTxn>(txn: &T, bytes: &[u8]) -> Result<Option<yrs::Offset>> {
    let index = StickyIndex::decode_v1(bytes).map_err(|error| fail(error.to_string()))?;
    Ok(index.get_offset(txn))
}

/// The stories, by key in their entries' order.
fn stories<T: ReadTxn>(txn: &T) -> Vec<(String, Out)> {
    let stories = root_map(txn, "stories");
    ordered_keys(txn, &stories)
        .into_iter()
        .filter_map(|key| stories.get(txn, &key).map(|value| (key, value)))
        .collect()
}

fn story<T: ReadTxn>(txn: &T, key: &str) -> Option<TextRef> {
    match root_map(txn, "stories").get(txn, key) {
        Some(Out::YText(text)) => Some(text),
        _ => None,
    }
}

/// Alignments made once per rebased story for the anchors (`anchorAlignments`).
#[derive(Default)]
pub(crate) struct AnchorAlignments(HashMap<String, Rc<Alignment>>);

/// `DOCX_LINEAGE.positions.rewrite`: the anchors of `value` (captured
/// positions in `from`) at their places in `to`.
pub(crate) fn rewrite_anchors<T: ReadTxn>(
    value: &Any,
    from: &T,
    to: &TransactionMut,
    id: &dyn Fn(&str) -> String,
    cache: &mut AnchorAlignments,
) -> Result<Copied> {
    let Any::Array(anchors) = value else {
        return Err(fail("comment anchors are not a list"));
    };
    let mut rewritten = Vec::new();
    for anchor in anchors.iter() {
        let story_key = any_string(
            any_field(anchor, "story")
                .map(|value| Out::Any(value.clone()))
                .as_ref(),
        );
        let (Some(source), Some(target)) = (story(from, &story_key), story(to, &id(&story_key)))
        else {
            return Err(fail(format!("comment anchor story {story_key} is missing")));
        };
        let position = |key: &str| -> Result<yrs::Offset> {
            let bytes = any_bytes(any_field(anchor, key)).unwrap_or_default();
            resolve(from, bytes)?.ok_or_else(|| fail("a comment anchor no longer resolves"))
        };
        let (first, last) = (position("start")?, position("end")?);
        let target_key = id(&story_key);
        let f = match cache.0.get(&target_key) {
            Some(f) => f.clone(),
            None => {
                let f = Rc::new(align(units(from, &source), units(to, &target))?);
                cache.0.insert(target_key.clone(), f.clone());
                f
            }
        };
        let (first_index, last_index) = (first.index as usize, last.index as usize);
        let mut start_at = point_at(&f, first_index);
        let mut end_at = point_at(&f, last_index);
        if last_index > first_index {
            let covered = &f.map[first_index..last_index];
            if covered[0] < 0
                || covered
                    .iter()
                    .enumerate()
                    .any(|(index, unit)| index > 0 && *unit != past(&f, covered[index - 1] + 1))
            {
                return Err(fail(
                    "a comment anchor covers content the export wrote differently",
                ));
            }
            start_at = Some(covered[0]);
            end_at = Some(covered[covered.len() - 1] + 1);
        }
        let (Some(start_at), Some(end_at)) = (start_at, end_at) else {
            return Err(fail(
                "a comment anchor lands in content the export wrote differently",
            ));
        };
        let at = |index: i64, assoc: Assoc| -> Result<Any> {
            let sticky = target
                .sticky_index(to, index as u32, assoc)
                .ok_or_else(|| {
                    fail("a comment anchor lands in content the export wrote differently")
                })?;
            Ok(Any::Buffer(sticky.encode_v1().into()))
        };
        let mut entry = HashMap::new();
        entry.insert(Arc::<str>::from("story"), Any::String(target_key.into()));
        entry.insert(Arc::<str>::from("start"), at(start_at, first.assoc)?);
        entry.insert(Arc::<str>::from("end"), at(end_at, last.assoc)?);
        rewritten.push(Any::Map(Arc::new(
            entry
                .into_iter()
                .map(|(key, value)| (key.to_string(), value))
                .collect(),
        )));
    }
    Ok(Copied::Any(Any::Array(rewritten.into())))
}

/// `DOCX_LINEAGE.settle`: a comment the later edits removed takes its
/// reference field with it.
pub(crate) fn settle(txn: &mut TransactionMut) {
    let comments = root_map(txn, "comments");
    for (_, value) in stories(txn) {
        let Out::YText(text) = value else { continue };
        let mut found = embeds(txn, &text);
        found.reverse();
        for (offset, embed) in found {
            let reference =
                get_string(txn, &embed, "modelKind").as_deref() == Some("commentReference");
            if reference
                && !comments.contains_key(txn, &any_string(embed.get(txn, "commentId").as_ref()))
            {
                text.remove_range(txn, offset, 1);
            }
        }
    }
}

/// `DOCX_LINEAGE.check`.
pub(crate) fn check<T: ReadTxn>(
    doc: &TransactionMut,
    later: &T,
    id: &dyn Fn(&str) -> String,
    captured: &T,
) -> Result<()> {
    for (key, source) in stories(later) {
        let (Out::YText(source), Some(target)) = (source, story(doc, &id(&key))) else {
            continue;
        };
        assert_children_landed(later, &source, doc, &target)?;
        assert_breaks_lead(later, &source, doc, &target, &key)?;
        assert_references_before_bookmarks(later, &source, doc, &target, &id(&key))?;
        assert_comments_cover(later, &key, &source, doc, &id(&key), &target, id, captured)?;
        assert_continuations_kept(doc, &id(&key), &target)?;
    }
    Ok(())
}

/// Refuses when a continued field would lose its separate or end.
fn assert_continuations_kept(doc: &TransactionMut, story: &str, target: &TextRef) -> Result<()> {
    let bookmarks = root_map(doc, "bookmarks");
    fn visit(value: &V, story: &str, bookmarks: &MapRef, doc: &TransactionMut) -> Result<()> {
        if let Some(items) = value.arr() {
            for item in items.items() {
                visit(&item, story, bookmarks, doc)?;
            }
            return Ok(());
        }
        let Some(node) = value.obj() else {
            return Ok(());
        };
        let continuation = node.get("continuation");
        let continued = continuation.get("id");
        if node.get("type").as_str().as_deref() == Some("complexField") && continued.truthy() {
            for kind in ["separate", "end"] {
                let key = format!("{story}:{}:field{kind}", continued.to_js_string());
                if continuation.get(kind).truthy() && !bookmarks.contains_key(doc, &key) {
                    return Err(fail("a continued field would lose its separate or end"));
                }
            }
        }
        for child in node.values() {
            visit(&child, story, bookmarks, doc)?;
        }
        Ok(())
    }
    for (_, embed) in embeds(doc, target) {
        if let Some(data) =
            get_string(doc, &embed, "fieldData").filter(|data| data.contains("\"continuation\""))
        {
            let parsed = V::parse(&data).map_err(fail)?;
            visit(&parsed, story, &bookmarks, doc)?;
        }
    }
    Ok(())
}

/// Comment ids with their ranges.
type Ranges = Vec<(String, Vec<(i64, i64)>)>;

/// Each comment's ranges in `story`, as absolute offsets, in entry order.
fn comment_ranges<T: ReadTxn>(txn: &T, story: &str) -> Result<Ranges> {
    let comments = root_map(txn, "comments");
    let mut ranges: Ranges = Vec::new();
    for key in ordered_keys(txn, &comments) {
        let Some(Out::YMap(comment)) = comments.get(txn, &key) else {
            continue;
        };
        let Some(Out::Any(Any::Array(anchors))) = comment.get(txn, "anchors") else {
            continue;
        };
        for anchor in anchors.iter() {
            if any_str(any_field(anchor, "story")) != Some(story) {
                continue;
            }
            let at = |name: &str| -> Result<Option<i64>> {
                Ok(
                    resolve(txn, any_bytes(any_field(anchor, name)).unwrap_or_default())?
                        .map(|offset| offset.index as i64),
                )
            };
            if let (Some(start), Some(end)) = (at("start")?, at("end")?) {
                match ranges.iter_mut().find(|(id, _)| *id == key) {
                    Some((_, list)) => list.push((start, end)),
                    None => ranges.push((key.clone(), vec![(start, end)])),
                }
            }
        }
    }
    Ok(ranges)
}

/// `units(text).ids.slice(a, b).join(",")`.
fn ids_between(units: &super::align::Units, start: usize, end: usize) -> String {
    units.ids[start.min(units.ids.len())..end.min(units.ids.len())]
        .iter()
        .map(|(client, clock)| format!("{client}:{clock}"))
        .collect::<Vec<_>>()
        .join(",")
}

/// Refuses when a comment would cover other content of the latest state in
/// the rebased one, but where the export settles a range.
#[allow(clippy::too_many_arguments)]
fn assert_comments_cover<T: ReadTxn>(
    later: &T,
    story: &str,
    source: &TextRef,
    doc: &TransactionMut,
    target_story: &str,
    target: &TextRef,
    id: &dyn Fn(&str) -> String,
    captured: &T,
) -> Result<()> {
    let from = comment_ranges(later, story)?;
    if from.is_empty() {
        return Ok(());
    }
    // Only a comment the later edits touched: its range, or the content it covers.
    let was = comment_ranges(captured, story)?;
    let captured_units = story_units(captured, story);
    let source_units = units(later, source);
    let covers =
        |text: Option<&super::align::Units>, ranges: Option<&Vec<(i64, i64)>>| -> Option<String> {
            let (text, ranges) = (text?, ranges?);
            Some(
                ranges
                    .iter()
                    .map(|&(start, end)| {
                        // A reversed range holds what lies between its ends.
                        if start > end {
                            ids_between(text, end as usize, start as usize)
                        } else {
                            ids_between(text, start as usize, end as usize)
                        }
                    })
                    .collect::<Vec<_>>()
                    .join("|"),
            )
        };
    let get = |list: &Ranges, key: &str| {
        list.iter()
            .find(|(id, _)| id == key)
            .map(|(_, ranges)| ranges.clone())
    };
    let reversed = |ranges: Option<&Vec<(i64, i64)>>| {
        ranges.is_some_and(|ranges| ranges.iter().any(|(start, end)| start > end))
    };
    let touched = |key: &str| {
        let before = covers(captured_units.as_ref(), get(&was, key).as_ref());
        before.is_none()
            || before != covers(Some(&source_units), get(&from, key).as_ref())
            || (reversed(get(&from, key).as_ref()) && !reversed(get(&was, key).as_ref()))
    };
    let to = comment_ranges(doc, target_story)?;
    let attributes = unit_attributes(later, source);
    let field = |unit: usize, key: &str| -> Option<Any> {
        attributes
            .get(unit)
            .and_then(|attrs| attrs.get(key))
            .filter(|value| !matches!(value, Any::Null | Any::Undefined))
            .cloned()
    };
    let same = |a: usize, b: usize, key: &str| field(a, key) == field(b, key);
    // A unit a settled boundary may move over: a projected child, or a link
    // unit in the same link (and child) as the unit across a boundary.
    let settles = |unit: usize, edges: &[i64]| {
        field(unit, "fieldResult").is_some()
            || (field(unit, "hyperlink").is_some()
                && edges.iter().any(|&edge| {
                    edge >= 1
                        && field(edge as usize - 1, "hyperlink").is_some()
                        && (edge as usize) < attributes.len()
                        && same(edge as usize - 1, edge as usize, "hyperlink")
                        && same(edge as usize - 1, edge as usize, "fieldResult")
                }))
    };
    let Ok(f) = align(source_units.clone(), units(doc, target)) else {
        return Ok(());
    };
    for (key, ranges) in &from {
        if !touched(key) {
            continue;
        }
        let landed = get(&to, &id(key)).or_else(|| get(&to, key));
        let Some(landed) = landed else {
            // An emptied range the export kept as a reference-only comment settles.
            let referenced = embeds(doc, target).iter().any(|(_, embed)| {
                get_string(doc, embed, "modelKind").as_deref() == Some("commentReference")
                    && [key.clone(), id(key)]
                        .contains(&any_string(embed.get(doc, "commentId").as_ref()))
            });
            if !referenced || ranges.iter().any(|(start, end)| start != end) {
                return Err(fail("a comment would lose its range and reference"));
            }
            continue;
        };
        let edges: Vec<i64> = ranges
            .iter()
            .flat_map(|&(start, end)| [start, end])
            .collect();
        let inside = |at: i64, spans: &[(i64, i64)]| {
            spans.iter().any(|&(start, end)| start <= at && at < end)
        };
        for (unit, &at) in f.map.iter().enumerate() {
            if at < 0 {
                continue;
            }
            if inside(unit as i64, ranges) != inside(at, &landed) && !settles(unit, &edges) {
                return Err(fail(
                    "a comment would cover other content than the latest state's",
                ));
            }
        }
    }
    Ok(())
}

fn story_units<T: ReadTxn>(txn: &T, key: &str) -> Option<super::align::Units> {
    story(txn, key).map(|text| units(txn, &text))
}

/// A break unit opening a paragraph, with its map.
type Slot = Vec<(u32, MapRef)>;

/// The page and column break units opening each paragraph that follows
/// another paragraph and holds something after them (`breakSlots`).
fn break_slots<T: ReadTxn>(txn: &T, text: &TextRef) -> Vec<Slot> {
    let mut slots = Vec::new();
    let mut slot: Slot = Vec::new();
    let mut after_paragraph = false;
    let mut content = false;
    for (offset, values, _) in items(txn, text) {
        if values.is_empty() {
            // Text units are content.
            content = true;
            continue;
        }
        for (index, unit) in values.into_iter().enumerate() {
            let map = match &unit {
                Out::YMap(map) => Some(map.clone()),
                _ => None,
            };
            let kind = map.as_ref().and_then(|map| get_string(txn, map, "_kind"));
            match kind.as_deref() {
                Some(PILCROW) => {
                    if content && !slot.is_empty() {
                        slots.push(std::mem::take(&mut slot));
                    }
                    slot.clear();
                    after_paragraph = true;
                    content = false;
                }
                Some("table" | "blockSdt") => {
                    slot.clear();
                    after_paragraph = false;
                }
                Some("pageBreak" | "columnBreak") if !content && after_paragraph => {
                    if let Some(map) = map {
                        slot.push((offset + index as u32, map));
                    }
                }
                _ => {
                    let reference = map.as_ref().is_some_and(|map| {
                        get_string(txn, map, "modelKind").as_deref() == Some("commentReference")
                    });
                    if !reference {
                        content = true;
                    }
                }
            }
        }
    }
    slots
}

/// Offsets of the bookmark boundaries (not continued field characters) in `story`.
fn bookmark_offsets<T: ReadTxn>(txn: &T, story: &str) -> Result<Vec<i64>> {
    let bookmarks = root_map(txn, "bookmarks");
    let mut found = Vec::new();
    for key in ordered_keys(txn, &bookmarks) {
        let Some(Out::YMap(entry)) = bookmarks.get(txn, &key) else {
            continue;
        };
        let data = match entry.get(txn, "data") {
            Some(Out::Any(any)) => any,
            _ => Any::Undefined,
        };
        if any_str(any_field(&data, "id")).is_some() {
            continue;
        }
        let Some(Out::Any(Any::Array(anchors))) = entry.get(txn, "anchors") else {
            continue;
        };
        for anchor in anchors.iter() {
            if any_str(any_field(anchor, "story")) != Some(story) {
                continue;
            }
            if let Some(offset) = resolve(
                txn,
                any_bytes(any_field(anchor, "start")).unwrap_or_default(),
            )? {
                found.push(offset.index as i64);
            }
        }
    }
    Ok(found)
}

/// Refuses when a bookmark boundary sits right before the export's
/// reference for a comment that opens a paragraph.
fn assert_references_before_bookmarks<T: ReadTxn>(
    later: &T,
    source: &TextRef,
    doc: &TransactionMut,
    target: &TextRef,
    story: &str,
) -> Result<()> {
    let bookmarks = bookmark_offsets(doc, story)?;
    if bookmarks.is_empty() {
        return Ok(());
    }
    let Ok(f) = align(units(later, source), units(doc, target)) else {
        return Ok(());
    };
    for (unit, key) in f.to.keys.iter().enumerate() {
        let opens = unit == 0 || f.to.keys[unit - 1] == Key::Embed(PILCROW.into());
        if *key == Key::Reference && !f.matched[unit] && opens && bookmarks.contains(&(unit as i64))
        {
            return Err(fail(
                "the export wrote a comment's reference after bookmarks a direct save writes it ahead of",
            ));
        }
    }
    Ok(())
}

/// Refuses when the seed reads a break opening a paragraph as leading its
/// text where the latest state does not.
fn assert_breaks_lead<T: ReadTxn>(
    later: &T,
    source: &TextRef,
    doc: &TransactionMut,
    target: &TextRef,
    story: &str,
) -> Result<()> {
    let leads = |txn: &dyn Fn(&MapRef) -> bool, map: &MapRef| txn(map);
    let leads_later =
        |map: &MapRef| matches!(map.get(later, "leading"), Some(Out::Any(Any::Bool(true))));
    let slots: Vec<Slot> = break_slots(later, source)
        .into_iter()
        .filter(|slot| slot.iter().any(|(_, unit)| !leads(&leads_later, unit)))
        .collect();
    if slots.is_empty() {
        return Ok(());
    }
    let Ok(f) = align(units(later, source), units(doc, target)) else {
        return Ok(());
    };
    let bookmarks = bookmark_offsets(later, story)?;
    let boundaries: Vec<i64> = comment_ranges(later, story)?
        .into_iter()
        .flat_map(|(_, ranges)| ranges.into_iter().flat_map(|(start, end)| [start, end]))
        .collect();
    let kind = |map: &MapRef| get_string(later, map, "_kind");
    for slot in &slots {
        let first = slot[0].0 as i64;
        let last = slot[slot.len() - 1].0 as i64 + 1;
        // The capture's save wrote a new comment's reference ahead of breaks it left without text.
        let mut unit = f.map[first as usize] - 1;
        while unit >= 0 && f.to.keys[unit as usize] == Key::Reference {
            if !f.matched[unit as usize] {
                return Err(fail(
                    "the export wrote a reference ahead of breaks the latest state opens a paragraph with",
                ));
            }
            unit -= 1;
        }
        let owned = kind(&slot[slot.len() - 1].1).as_deref() == Some("columnBreak")
            && !slot.iter().any(|(_, unit)| {
                matches!(unit.get(later, "trailing"), Some(Out::Any(Any::Bool(true))))
            })
            && !bookmarks
                .iter()
                .any(|offset| *offset >= first && *offset <= last);
        let marked = slot.iter().position(|(offset, unit)| {
            kind(unit).as_deref() == Some("columnBreak")
                && boundaries
                    .iter()
                    .any(|boundary| *boundary >= first && *boundary <= *offset as i64)
        });
        let flagged = slot.iter().position(|(_, unit)| leads_later(unit));
        let leading: Vec<usize> = [flagged, owned.then_some(0), marked]
            .into_iter()
            .flatten()
            .collect();
        let latest = leading.iter().min().copied();
        let rebased = slot.iter().position(|(at, _)| {
            let mapped = f.map[*at as usize];
            mapped >= 0
                && matches!(
                    child_at_offset(doc, target, mapped as u32),
                    Some(Out::YMap(map)) if matches!(map.get(doc, "leading"), Some(Out::Any(Any::Bool(true))))
                )
        });
        if let Some(rebased) = rebased
            && latest.is_none_or(|latest| rebased < latest)
        {
            return Err(fail(
                "text follows breaks the export reads as leading it and the latest state does not",
            ));
        }
    }
    Ok(())
}

#[derive(Clone, Copy, PartialEq)]
struct Marker {
    id: f64,
    index: f64,
}

fn marker_of(value: Option<&Any>) -> Option<Marker> {
    let value = value.filter(|value| !matches!(value, Any::Null | Any::Undefined))?;
    Some(Marker {
        id: any_number(any_field(value, "id")).unwrap_or(f64::NAN),
        index: any_number(any_field(value, "index")).unwrap_or(f64::NAN),
    })
}

struct Children {
    markers: Vec<Option<Marker>>,
    owners: Vec<i64>,
}

/// A text's projected field children as the export pairs them with their
/// fields (`projectedChildren`).
fn projected_children<T: ReadTxn>(txn: &T, text: &TextRef) -> Children {
    let mut markers: Vec<Option<Marker>> = Vec::new();
    let mut owners: Vec<i64> = Vec::new();
    let mut run = 0_usize;
    for diff in text.diff(txn, YChange::identity) {
        let length = match &diff.insert {
            Out::Any(Any::String(text)) => text.encode_utf16().count(),
            _ => 1,
        };
        let marker = diff
            .attributes
            .as_ref()
            .and_then(|attrs| marker_of(attrs.get("fieldResult")));
        for _ in 0..length {
            markers.push(marker);
            owners.push(-1);
        }
        if marker.is_some() {
            run += length;
            continue;
        }
        let projection = match &diff.insert {
            Out::YMap(map) => match map.get(txn, "resultProjection") {
                Some(Out::Any(any)) => Some(any),
                _ => None,
            },
            _ => None,
        };
        let at = markers.len() as i64 - 1;
        if let Some(projection) = &projection
            && let Some(Any::Number(projection_id)) = any_field(projection, "id")
        {
            let indices: Vec<Option<f64>> = match any_field(projection, "children") {
                Some(Any::Array(children)) => children
                    .iter()
                    .map(|child| any_number(any_field(child, "index")))
                    .collect(),
                _ => Vec::new(),
            };
            let mut child = at - 1;
            while child >= at - run as i64 && child >= 0 {
                let Some(Marker { id, index }) = markers[child as usize] else {
                    break;
                };
                if id != *projection_id || !indices.contains(&Some(index)) {
                    break;
                }
                owners[child as usize] = at;
                child -= 1;
            }
        }
        run = 0;
    }
    Children { markers, owners }
}

/// Refuses unless the rebased story holds each projected child of the
/// latest state in the field and result slot it has there.
fn assert_children_landed<T: ReadTxn>(
    later: &T,
    source: &TextRef,
    doc: &TransactionMut,
    target: &TextRef,
) -> Result<()> {
    let from = projected_children(later, source);
    let to = projected_children(doc, target);
    if !from.markers.iter().any(Option::is_some) && !to.markers.iter().any(Option::is_some) {
        return Ok(());
    }
    let Ok(f) = align(units(later, source), units(doc, target)) else {
        return Ok(());
    };
    let mut slots: HashMap<String, String> = HashMap::new();
    let mut taken: HashMap<String, String> = HashMap::new();
    for (unit, &owner) in from.owners.iter().enumerate() {
        let at = f.map[unit];
        let field = if owner < 0 { -1 } else { f.map[owner as usize] };
        if owner >= 0 && field < 0 {
            return Err(fail("a field result's child lands where its field did not"));
        }
        if at < 0 {
            if from.markers[unit].is_some() && owner < 0 {
                return Err(fail("a field result's child has no place in the export"));
            }
            continue;
        }
        if to.owners[at as usize] != field {
            return Err(fail("a field result's child would not export in its field"));
        }
        if owner < 0 {
            continue;
        }
        let number = |value: f64| crate::js::number_to_string(value);
        let slot = format!(
            "{owner}:{}",
            number(from.markers[unit].map_or(f64::NAN, |m| m.index))
        );
        let place = format!(
            "{field}:{}",
            number(to.markers[at as usize].map_or(f64::NAN, |m| m.index))
        );
        if slots.get(&slot).unwrap_or(&place) != &place
            || taken.get(&place).unwrap_or(&slot) != &slot
        {
            return Err(fail(
                "two field result children would land in one result slot",
            ));
        }
        slots.insert(slot.clone(), place.clone());
        taken.insert(place, slot);
    }
    Ok(())
}

/// `fieldAttributes`: a lone surviving child can change index when an
/// earlier sibling was deleted.
pub(crate) struct FieldAttributes {
    from: Children,
    slots: HashMap<String, Marker>,
    ambiguous: HashSet<String>,
    siblings: HashMap<i64, HashSet<u64>>,
}

pub(crate) fn field_attributes<T: ReadTxn>(
    captured: &T,
    source: &TextRef,
    result: &TransactionMut,
    target: &TextRef,
    f: &Alignment,
) -> FieldAttributes {
    let from = projected_children(captured, source);
    let to = projected_children(result, target);
    let key = |owner: i64, marker: Marker| {
        format!("{owner}:{}", crate::js::number_to_string(marker.index))
    };
    let mut slots: HashMap<String, Marker> = HashMap::new();
    let mut taken: HashMap<String, String> = HashMap::new();
    let mut ambiguous: HashSet<String> = HashSet::new();
    let mut siblings: HashMap<i64, HashSet<u64>> = HashMap::new();
    for (unit, &owner) in from.owners.iter().enumerate() {
        if owner >= 0
            && let Some(marker) = from.markers[unit]
        {
            siblings
                .entry(owner)
                .or_default()
                .insert(marker.index.to_bits());
        }
        if owner < 0 || f.map[unit] < 0 {
            continue;
        }
        let Some(own) = from.markers[unit] else {
            continue;
        };
        let at = f.map[unit] as usize;
        let slot = key(owner, own);
        let marker = to.markers[at];
        let Some(marker) = marker.filter(|_| to.owners[at] == f.map[owner as usize]) else {
            ambiguous.insert(slot);
            continue;
        };
        let place = key(to.owners[at], marker);
        if let Some(prior) = slots.get(&slot)
            && (prior.id != marker.id || prior.index != marker.index)
        {
            ambiguous.insert(slot.clone());
        }
        if let Some(other) = taken.get(&place)
            && *other != slot
        {
            ambiguous.insert(slot.clone());
            ambiguous.insert(other.clone());
        }
        slots.insert(slot.clone(), marker);
        taken.insert(place, slot);
    }
    FieldAttributes {
        from,
        slots,
        ambiguous,
        siblings,
    }
}

impl FieldAttributes {
    pub(crate) fn rewrite(&self, attributes: &Attrs, at: usize) -> Attrs {
        let Some(marker) = marker_of(attributes.get("fieldResult")) else {
            return attributes.clone();
        };
        let neighbour = [at.checked_sub(1), Some(at)]
            .into_iter()
            .flatten()
            .find(|&unit| {
                self.from.owners.get(unit).is_some_and(|owner| *owner >= 0)
                    && self.from.markers[unit]
                        .is_some_and(|own| own.id == marker.id && own.index == marker.index)
            });
        let Some(neighbour) = neighbour else {
            return attributes.clone();
        };
        let owner = self.from.owners[neighbour];
        if self.siblings.get(&owner).map(HashSet::len) != Some(1) {
            return attributes.clone();
        }
        let slot = format!("{owner}:{}", crate::js::number_to_string(marker.index));
        let Some(saved) = self.slots.get(&slot) else {
            return attributes.clone();
        };
        if saved.index == marker.index || self.ambiguous.contains(&slot) {
            return attributes.clone();
        }
        let mut rewritten = attributes.clone();
        let mut value = HashMap::new();
        value.insert("id".to_owned(), Any::Number(saved.id));
        value.insert("index".to_owned(), Any::Number(saved.index));
        rewritten.insert(Arc::from("fieldResult"), Any::Map(Arc::new(value)));
        rewritten
    }
}

/// A later edit to a continued field's embed keeps the continuation id the
/// export's seed gave it (`keepContinuation`).
pub(crate) fn keep_continuation(key: &str, old: Option<&Out>, next: Copied) -> Copied {
    let old = match old {
        Some(Out::Any(Any::String(text))) => Some(text.to_string()),
        _ => None,
    };
    let next_text = match &next {
        Copied::Any(Any::String(text)) => Some(text.to_string()),
        _ => None,
    };
    if key == "continuationId" {
        return match (old, next_text) {
            (Some(old), Some(_)) => Copied::Any(Any::String(old.into())),
            _ => next,
        };
    }
    let (Some(old), Some(text)) = (old, next_text) else {
        return next;
    };
    if key != "fieldData" || !old.contains("\"continuation\"") || !text.contains("\"continuation\"")
    {
        return next;
    }
    let (Ok(before), Ok(after)) = (V::parse(&old), V::parse(&text)) else {
        return next;
    };
    let (kept, changed) = (
        before.get("continuation").get("id"),
        after.get("continuation").get("id"),
    );
    if !kept.truthy() || !changed.truthy() || kept.same(&changed) {
        return next;
    }
    if let Some(continuation) = after.get("continuation").obj() {
        continuation.set("id", kept);
    }
    Copied::Any(Any::String(stringify(&after).unwrap_or_default().into()))
}

/// `Y.Map.toJSON()` as a JS value.
fn map_json<T: ReadTxn>(txn: &T, map: &MapRef) -> V {
    V::from(&serde_json::to_value(map.to_json(txn)).unwrap_or_default())
}

/// The continued fields an embed holds: (continuation id, instruction).
fn continued_fields(value: &V) -> Vec<(String, String)> {
    let mut found: Vec<(String, String)> = Vec::new();
    fn visit(value: &V, found: &mut Vec<(String, String)>) {
        if let Some(items) = value.arr() {
            for item in items.items() {
                visit(&item, found);
            }
            return;
        }
        let Some(payload) = value.obj() else { return };
        let mut id = if payload.get("modelKind").as_str().as_deref() == Some("field") {
            payload.get("continuationId")
        } else {
            V::Undef
        };
        let continuation = payload.get("continuation");
        if payload.get("type").as_str().as_deref() == Some("complexField")
            && continuation.obj().is_some()
        {
            id = continuation.get("id");
        }
        if let (Some(id), Some(instruction)) = (id.as_str(), payload.get("instruction").as_str()) {
            match found.iter_mut().find(|(key, _)| **key == *id) {
                Some(entry) => entry.1 = instruction.to_string(),
                None => found.push((id.to_string(), instruction.to_string())),
            }
        }
        if let Some(data) = payload
            .get("fieldData")
            .as_str()
            .filter(|data| data.contains("\"continuation\""))
            && let Ok(parsed) = V::parse(&data)
        {
            visit(&parsed, found);
        }
        for (key, child) in payload.entries() {
            if &*key != "blocks" && &*key != "fieldTree" {
                visit(&child, found);
            }
        }
    }
    visit(value, &mut found);
    found
}

/// Captured DOCX story, paragraph and comment ids to seed(export)'s (`docxIds`).
pub(crate) fn docx_ids(captured: &[u8], seed: &[u8]) -> Result<Vec<(String, String)>> {
    use yrs::Transact;
    let from_doc = open(&DOCX_LINEAGE, captured, None)?;
    let to_doc = open(&DOCX_LINEAGE, seed, None)?;
    let from = from_doc.transact();
    let to = to_doc.transact();
    let mut ids: Vec<(String, String)> = Vec::new();
    fn set(ids: &mut Vec<(String, String)>, key: String, value: String) {
        match ids.iter_mut().find(|(from, _)| *from == key) {
            Some(entry) => entry.1 = value,
            None => ids.push((key, value)),
        }
    }
    fn pair<T: ReadTxn>(
        from: &T,
        to: &T,
        ids: &mut Vec<(String, String)>,
        a: Option<String>,
        b: Option<String>,
    ) {
        let (Some(a), Some(b)) = (a, b) else { return };
        if ids.iter().any(|(key, _)| *key == a) {
            return;
        }
        set(ids, a.clone(), b.clone());
        let (Some(left), Some(right)) = (story(from, &a), story(to, &b)) else {
            return;
        };
        let Ok(f) = align(units(from, &left), units(to, &right)) else {
            return;
        };
        let theirs: HashMap<u32, MapRef> = embeds(to, &right).into_iter().collect();
        for (index, embed) in embeds(from, &left) {
            let mapped = f.map[index as usize];
            let Some(other) = (mapped >= 0)
                .then(|| theirs.get(&(mapped as u32)))
                .flatten()
            else {
                continue;
            };
            let kind = get_string(from, &embed, "_kind");
            if get_string(to, other, "_kind") != kind {
                continue;
            }
            if matches!(kind.as_deref(), Some("field" | "sdt")) {
                let original = continued_fields(&map_json(from, &embed));
                let saved = continued_fields(&map_json(to, other));
                if original.len() == saved.len()
                    && original.iter().zip(&saved).all(|((_, a), (_, b))| a == b)
                {
                    for ((id, _), (saved, _)) in original.iter().zip(&saved) {
                        set(ids, id.clone(), saved.clone());
                    }
                }
            }
            match kind.as_deref() {
                Some(PILCROW) => pair(
                    from,
                    to,
                    ids,
                    get_string(from, &embed, "paraId"),
                    get_string(to, other, "paraId"),
                ),
                Some("blockSdt") => pair(
                    from,
                    to,
                    ids,
                    get_string(from, &embed, "story"),
                    get_string(to, other, "story"),
                ),
                Some("table") => {
                    let rows = |txn: &T, map: &MapRef| match map.get(txn, "rows") {
                        Some(Out::Any(Any::Array(rows))) => rows.to_vec(),
                        _ => Vec::new(),
                    };
                    let (mine, others) = (rows(from, &embed), rows(to, other));
                    for (r, row) in mine.iter().enumerate() {
                        let Some(Any::Array(cells)) = any_field(row, "cells") else {
                            continue;
                        };
                        for (c, cell) in cells.iter().enumerate() {
                            let theirs = others
                                .get(r)
                                .and_then(|row| match any_field(row, "cells") {
                                    Some(Any::Array(cells)) => cells.get(c).cloned(),
                                    _ => None,
                                })
                                .and_then(|cell| {
                                    any_str(any_field(&cell, "story")).map(str::to_owned)
                                });
                            pair(
                                from,
                                to,
                                ids,
                                any_str(any_field(cell, "story")).map(str::to_owned),
                                theirs,
                            );
                        }
                    }
                }
                _ => {}
            }
        }
    }
    let to_stories = root_map(&to, "stories");
    for (key, _) in stories(&from) {
        let top = key == "body"
            || (["hf:", "fn:", "en:"]
                .iter()
                .any(|prefix| key.starts_with(prefix))
                && key[3..].find(':').is_none()
                && key.len() > 3);
        if top && to_stories.contains_key(&to, &key) {
            pair(&from, &to, &mut ids, Some(key.clone()), Some(key));
        }
    }
    let comments = root_map(&from, "comments");
    let keys: Vec<String> = ordered_keys(&from, &comments)
        .into_iter()
        .filter(|key| matches!(comments.get(&from, key), Some(Out::YMap(_))))
        .collect();
    let saved_ids = crate::docx::comment_ooxml_ids(&keys);
    let to_comments = root_map(&to, "comments");
    for key in &keys {
        let saved = saved_ids[key].to_string();
        if to_comments.contains_key(&to, &saved) {
            set(&mut ids, key.clone(), saved);
        }
    }
    let bookmarks = root_map(&to, "bookmarks");
    let from_bookmarks = root_map(&from, "bookmarks");
    for key in ordered_keys(&from, &from_bookmarks) {
        let Some(Out::YMap(value)) = from_bookmarks.get(&from, &key) else {
            continue;
        };
        let data = match value.get(&from, "data") {
            Some(Out::Any(any)) => any,
            _ => continue,
        };
        let anchors = match value.get(&from, "anchors") {
            Some(Out::Any(Any::Array(anchors))) => anchors,
            _ => continue,
        };
        let Some(story) = anchors
            .first()
            .and_then(|anchor| any_str(any_field(anchor, "story")))
            .filter(|s| !s.is_empty())
        else {
            continue;
        };
        let lookup = |key: &str| {
            ids.iter()
                .find(|(from, _)| from == key)
                .map(|(_, to)| to.clone())
        };
        let Some(story_id) = lookup(story) else {
            continue;
        };
        let marker = match any_field(&data, "id") {
            Some(Any::String(id)) => lookup(id).unwrap_or_else(|| id.to_string()),
            other => any_string(other.map(|value| Out::Any(value.clone())).as_ref()),
        };
        let kind = any_string(
            any_field(&data, "kind")
                .map(|value| Out::Any(value.clone()))
                .as_ref(),
        );
        let saved = format!("{story_id}:{marker}:{kind}");
        if bookmarks.contains_key(&to, &saved) {
            set(&mut ids, key, saved);
        }
    }
    Ok(ids)
}

/// Story and id of each DOCX paragraph whose mark or id `client` wrote
/// (`docxWrittenParagraphs`).
pub(crate) fn docx_written_paragraphs(state: &[u8], client: u64) -> Result<Vec<(String, String)>> {
    use yrs::Transact;
    let doc = open(&DOCX_LINEAGE, state, None)?;
    let txn = doc.transact();
    let mut found = Vec::new();
    for (key, value) in stories(&txn) {
        let Out::YText(text) = value else { continue };
        for (_, embed) in embeds(&txn, &text) {
            if get_string(&txn, &embed, "_kind").as_deref() != Some(PILCROW) {
                continue;
            }
            let branch: &yrs::branch::Branch = embed.as_ref();
            let written = matches!(branch.id(), yrs::BranchID::Nested(item) if item.client.get() == client)
                || branch.map_items().any(|(key, item)| {
                    key.as_ref() == "paraId" && item.id().client.get() == client
                });
            if written {
                found.push((key.clone(), any_string(embed.get(&txn, "paraId").as_ref())));
            }
        }
    }
    Ok(found)
}
