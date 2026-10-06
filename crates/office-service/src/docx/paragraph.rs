//! yrsToDocument.ts, part two: comment, bookmark and raw-inline boundaries,
//! the original-run cache, and the page and column breaks of paragraph slots.

use std::cmp::Ordering;
use std::collections::{HashMap, HashSet};

use crate::error::Result;
use crate::jsv::{Obj, V, text};
use crate::{arr, obj};

use super::item::{
    BOOLEAN_MARKS, attrs_to_text_formatting, break_run, build_paragraph_content, drop_nulls,
    formatting_attrs, is, is_kind, is_unit_break, o, run_content_for_text, st, stable_stringify,
    tracked_content_for_item, tracked_info, ty,
};

fn marks_key_to_yrs_attrs(marks_key: &V) -> Option<V> {
    let Some(key) = marks_key.as_str().filter(|key| !key.is_empty()) else {
        return Some(obj! {});
    };
    let attrs = Obj::new();
    // Mark attribute JSON does not normally contain `|`; if a custom string
    // does, declining restoration merely coalesces runs and never loses text.
    for part in key.split('|') {
        let colon = part.find(':').filter(|colon| *colon > 0)?;
        let name = &part[..colon];
        let value = V::parse(&part[colon + 1..]).ok()?;
        if name == "comment" || name == "footnoteRef" {
            continue;
        }
        if BOOLEAN_MARKS.contains(&name) {
            attrs.set(name, V::Bool(true));
        } else if name == "highlight" {
            attrs.set("highlight", value.get("color"));
        } else if name == "insertion" || name == "deletion" {
            attrs.set(
                if name == "insertion" { "ins" } else { "del" },
                drop_nulls(&obj! {
                    "id": value.get("revisionId"),
                    "author": value.get("author"),
                    "date": value.get("date"),
                    "isMovePair": value.get("isMovePair"),
                }),
            );
        } else {
            attrs.set(name, drop_nulls(&value));
        }
    }
    Some(V::Obj(attrs))
}

/// Rebuilds the note number marks a run held; they occupy no story unit.
fn note_mark_content(boundary: &V) -> Option<Vec<V>> {
    let marks = boundary.get("noteMarks").items();
    if marks.is_empty() {
        return None;
    }
    Some(
        marks
            .iter()
            .map(|mark| {
                let kind = if mark.as_str().as_deref() == Some("endnote") {
                    "endnoteRefMark"
                } else {
                    "footnoteRefMark"
                };
                obj! { "type": kind }
            })
            .collect(),
    )
}

/// Rebuilds a run from its recorded boundary, flow breaks at their offsets.
fn boundary_content(boundary: &V, formatting: &Obj) -> Vec<V> {
    if let Some(notes) = note_mark_content(boundary) {
        return notes;
    }
    let source = boundary.get("text").to_js_string();
    let breaks = boundary.get("breaks").items();
    if breaks.is_empty() {
        return run_content_for_text(&source, formatting);
    }
    let length = text::len(&source) as f64;
    let mut content = Vec::new();
    let mut cursor = 0.0;
    for entry in breaks {
        let at = entry.get("offset").to_number().max(cursor).min(length);
        if at > cursor {
            content.extend(run_content_for_text(
                &text::slice(&source, cursor as usize, at as usize),
                formatting,
            ));
        }
        content.push(obj! { "type": "break", "breakType": entry.get("type") });
        cursor = at;
    }
    if cursor < length {
        content.extend(run_content_for_text(
            &text::slice_from(&source, cursor as usize),
            formatting,
        ));
    }
    content
}

/// An element as `Array.prototype.join` writes it.
fn joined_text(value: &V) -> String {
    if value.nullish() {
        String::new()
    } else {
        value.to_js_string()
    }
}

fn restore_original_runs(content: Vec<V>, items: &[V], boundaries: Option<Vec<V>>) -> Vec<V> {
    let Some(boundaries) = boundaries.filter(|boundaries| !boundaries.is_empty()) else {
        return content;
    };
    if !content.iter().all(|child| {
        is(child, "run")
            && child
                .get("content")
                .items()
                .iter()
                .all(|entry| is(entry, "text"))
    }) || items
        .iter()
        .any(|item| !is_kind(item, "text") || item.get("attributes").get("hyperlink").truthy())
        || boundaries.iter().any(|boundary| {
            !boundary.get("noteMarks").items().is_empty()
                && text::len(&boundary.get("text").to_js_string()) > 0
        })
    {
        return content;
    }
    let full: String = items
        .iter()
        .map(|item| {
            if is_kind(item, "text") {
                joined_text(&item.get("text"))
            } else {
                String::new()
            }
        })
        .collect();
    let joined: String = boundaries
        .iter()
        .map(|boundary| joined_text(&boundary.get("text")))
        .collect();
    if full != joined {
        return content;
    }
    let mut item_index = 0;
    let mut item_offset = 0;
    let mut restored = Vec::new();
    for boundary in &boundaries {
        let Some(expected) = marks_key_to_yrs_attrs(&boundary.get("marksKey")) else {
            return content;
        };
        let expected_key = stable_stringify(&expected);
        restored.push(expected);
        let mut remaining = text::len(&boundary.get("text").to_js_string());
        while remaining > 0 {
            let Some(item) = items.get(item_index).filter(|item| is_kind(item, "text")) else {
                return content;
            };
            if stable_stringify(&formatting_attrs(&item.get("attributes"))) != expected_key {
                return content;
            }
            let length = text::len(&item.get("text").to_js_string());
            let consumed = remaining.min(length - item_offset);
            remaining -= consumed;
            item_offset += consumed;
            if item_offset == length {
                item_index += 1;
                item_offset = 0;
            }
        }
    }
    if item_index != items.len() || item_offset != 0 {
        return content;
    }
    boundaries
        .iter()
        .enumerate()
        .map(|(index, boundary)| {
            // Non-empty runs keep formatting rebuilt from their live marks;
            // only empty runs take the cached formatting.
            let formatting = if text::len(&boundary.get("text").to_js_string()) == 0 {
                boundary.get("formatting")
            } else {
                V::Obj(attrs_to_text_formatting(&restored[index]))
            };
            let formatting_object = o(&formatting).unwrap_or_default();
            let run = obj! {
                "type": "run",
                "content": V::array(boundary_content(boundary, &formatting_object)),
            };
            let object = o(&run).unwrap_or_default();
            if formatting.truthy() && !formatting_object.is_empty() {
                object.set("formatting", formatting);
            }
            let changes = boundary.get("propertyChanges");
            if changes.arr().is_some_and(|changes| !changes.is_empty()) {
                object.set("propertyChanges", changes);
            }
            run
        })
        .collect()
}

fn run_text_length(run: &V) -> f64 {
    run.get("content")
        .items()
        .iter()
        .map(|content| match &*ty(content) {
            "text" | "instrText" => text::len(&content.get("text").to_js_string()) as f64,
            "symbol" => text::len(&content.get("char").to_js_string()) as f64,
            "tab" | "ptab" | "softHyphen" | "noBreakHyphen" | "footnoteRef" | "endnoteRef"
            | "horizontalRule" => 1.0,
            _ => 0.0,
        })
        .sum()
}

pub(crate) fn paragraph_content_length(content: &V) -> f64 {
    let runs = |key: &str| -> f64 {
        content
            .get(key)
            .items()
            .iter()
            .map(|child| {
                if is(child, "run") {
                    run_text_length(child)
                } else {
                    0.0
                }
            })
            .sum()
    };
    match &*ty(content) {
        "run" => run_text_length(content),
        "hyperlink" => runs("children"),
        "simpleField" | "inlineSdt" => content
            .get("content")
            .items()
            .iter()
            .map(paragraph_content_length)
            .sum(),
        "complexField" => content
            .get("fieldResult")
            .items()
            .iter()
            .map(run_text_length)
            .sum(),
        "insertion" | "deletion" | "moveFrom" | "moveTo" => runs("content"),
        "mathEquation" => {
            let plain = content.get("plainText");
            if plain.nullish() {
                0.0
            } else {
                text::len(&plain.to_js_string()) as f64
            }
        }
        _ => 0.0,
    }
}

pub(crate) fn run_content_units(content: &V) -> f64 {
    match &*ty(content) {
        "text" => text::len(&content.get("text").to_js_string()) as f64,
        "symbol" => {
            let digits: String = content
                .get("char")
                .to_js_string()
                .trim_start()
                .chars()
                .take_while(char::is_ascii_hexdigit)
                .collect();
            match u32::from_str_radix(&digits, 16) {
                Ok(code) if code <= 0x10ffff => {
                    if code > 0xffff {
                        2.0
                    } else {
                        1.0
                    }
                }
                _ => 0.0,
            }
        }
        "break" => {
            let kind = content.get("breakType");
            if kind.is_undef()
                || kind.as_str().as_deref() == Some("textWrapping")
                || is_unit_break(content)
            {
                1.0
            } else {
                0.0
            }
        }
        "tab" | "ptab" | "softHyphen" | "noBreakHyphen" | "horizontalRule" | "commentReference"
        | "drawing" | "shape" | "chart" | "opaqueDrawing" | "footnoteRef" | "endnoteRef" => 1.0,
        _ => 0.0,
    }
}

/// Story units `content` seeds as (seed.rs `paragraph_units`).
pub(crate) fn story_units(content: &V) -> f64 {
    match &*ty(content) {
        "run" => content
            .get("content")
            .items()
            .iter()
            .map(run_content_units)
            .sum(),
        "hyperlink" => {
            let children = content.get("structuredChildren");
            let children = if children.nullish() {
                content.get("children")
            } else {
                children
            };
            children
                .items()
                .iter()
                .map(|child| match &*ty(child) {
                    "run" => story_units(child),
                    "simpleField" | "complexField" | "mathEquation" => 1.0,
                    _ => 0.0,
                })
                .sum()
        }
        "complexField" => {
            // seed.rs `field_to_units`: result hyperlinks and nested simple fields seed their own units.
            let mut projected = content.get("structuredCode").get("inline").items();
            projected.extend(content.get("structuredResult").get("inline").items());
            let projected: Vec<V> = projected
                .into_iter()
                .filter(|child| is(child, "hyperlink") || is(child, "simpleField"))
                .collect();
            let instruction = content.get("instruction").to_js_string();
            let trimmed = instruction.trim();
            if (!trimmed.is_empty() && trimmed.bytes().all(|byte| byte.is_ascii_digit()))
                || projected.is_empty()
            {
                return 1.0;
            }
            projected.iter().map(story_units).sum::<f64>() + 1.0
        }
        "simpleField" | "inlineSdt" | "mathEquation" => 1.0,
        "insertion" | "deletion" | "moveFrom" | "moveTo" => {
            content.get("content").items().iter().map(story_units).sum()
        }
        _ => 0.0,
    }
}

/// `content` split `offset` story units in, or none where it cannot split.
fn split_content(content: &V, offset: f64) -> Option<(V, V)> {
    if matches!(
        &*ty(content),
        "insertion" | "deletion" | "moveFrom" | "moveTo"
    ) {
        let children = content.get("content").items();
        if children.len() == 1 && is(&children[0], "run") {
            let (left, right) = split_content(&children[0], offset)?;
            let object = o(content).unwrap_or_default();
            return Some((
                V::Obj(object.with(&[("content", arr![left])])),
                V::Obj(object.with(&[("content", arr![right])])),
            ));
        }
    }
    if !is(content, "run") {
        return None;
    }
    let mut left = Vec::new();
    let mut right = Vec::new();
    let mut units = 0.0;
    for entry in content.get("content").items() {
        let width = run_content_units(&entry);
        if units >= offset {
            right.push(entry);
        } else if units + width <= offset {
            left.push(entry);
        } else if is(&entry, "text") {
            let source = entry.get("text").to_js_string();
            let at = (offset - units) as usize;
            left.push(obj! { "type": "text", "text": text::slice(&source, 0, at) });
            right.push(obj! { "type": "text", "text": text::slice_from(&source, at) });
        } else {
            return None;
        }
        units += width;
    }
    let half = |entries: Vec<V>| {
        let run = Obj::new();
        run.set("type", V::str("run"));
        let formatting = content.get("formatting");
        if formatting.truthy() {
            run.set("formatting", formatting);
        }
        let changes = content.get("propertyChanges");
        if changes.truthy() {
            run.set("propertyChanges", changes);
        }
        run.set("content", V::array(entries));
        V::Obj(run)
    };
    Some((half(left), half(right)))
}

fn boundary_kind(boundary: &V) -> String {
    boundary.get("kind").to_js_string()
}

/// Boundaries by offset; at one offset, ranges closing there first, then
/// empty ranges, then ranges opening there.
pub(crate) fn sort_boundaries(boundaries: &[V]) -> Vec<V> {
    let key = |id: &V, kind: &str, offset: &V| {
        format!("{}:{kind}:{}", id.to_js_string(), offset.to_js_string())
    };
    let at: HashSet<String> = boundaries
        .iter()
        .map(|boundary| {
            key(
                &boundary.get("id"),
                &boundary_kind(boundary),
                &boundary.get("offset"),
            )
        })
        .collect();
    let place = |boundary: &V| {
        let kind = boundary_kind(boundary);
        let other = if kind == "start" { "end" } else { "start" };
        if at.contains(&key(&boundary.get("id"), other, &boundary.get("offset"))) {
            1
        } else if kind == "end" {
            0
        } else {
            2
        }
    };
    let mut sorted = boundaries.to_vec();
    sorted.sort_by(|left, right| {
        let number = |value: &V, key: &str| value.get(key).to_number();
        let difference = number(left, "offset") - number(right, "offset");
        let ordered = |difference: f64| {
            if difference < 0.0 {
                Ordering::Less
            } else if difference > 0.0 {
                Ordering::Greater
            } else {
                Ordering::Equal
            }
        };
        ordered(difference)
            .then_with(|| place(left).cmp(&place(right)))
            .then_with(|| ordered(number(left, "id") - number(right, "id")))
            .then_with(|| {
                u8::from(boundary_kind(left) == "end").cmp(&u8::from(boundary_kind(right) == "end"))
            })
    });
    sorted
}

pub(crate) fn comment_marker(boundary: &V) -> V {
    let kind = if boundary_kind(boundary) == "start" {
        "commentRangeStart"
    } else {
        "commentRangeEnd"
    };
    obj! { "type": kind, "id": boundary.get("id") }
}

/// Places each boundary `offset` units (by `measure`) into `content`; one
/// inside content that cannot split moves to that content's edge.
pub(crate) fn insert_boundaries(
    content: Vec<V>,
    boundaries: &[V],
    measure: &dyn Fn(&V) -> f64,
    make_marker: &dyn Fn(&V) -> V,
    after_markers: bool,
) -> Vec<V> {
    if boundaries.is_empty() {
        return content;
    }
    let sorted = sort_boundaries(boundaries);
    let offset_of = |index: usize| sorted[index].get("offset").to_number();
    let mut result = Vec::new();
    let mut cursor = 0.0;
    let mut next = 0;
    let emit = |result: &mut Vec<V>, next: &mut usize, offset: f64| {
        while *next < sorted.len() && offset_of(*next) == offset {
            result.push(make_marker(&sorted[*next]));
            *next += 1;
        }
    };
    if !after_markers {
        emit(&mut result, &mut next, 0.0);
    }
    for item in content {
        if after_markers && measure(&item) > 0.0 {
            emit(&mut result, &mut next, cursor);
        }
        let end = cursor + measure(&item);
        let mut rest = item;
        let mut rest_start = cursor;
        let mut widened = Vec::new();
        while next < sorted.len()
            && (offset_of(next) < end
                || (offset_of(next) == end
                    && flow_tokens(std::slice::from_ref(&rest))
                        .last()
                        .is_some_and(FlowToken::is_break)))
        {
            let boundary = sorted[next].clone();
            let offset = offset_of(next);
            match split_content(&rest, offset - rest_start) {
                Some((left, right)) => {
                    result.push(left);
                    emit(&mut result, &mut next, offset);
                    rest = right;
                    rest_start = offset;
                }
                None => {
                    next += 1;
                    if boundary_kind(&boundary) == "start" {
                        result.push(make_marker(&boundary));
                    } else {
                        widened.push(make_marker(&boundary));
                    }
                }
            }
        }
        result.push(rest);
        result.extend(widened);
        cursor = end;
        if !after_markers {
            emit(&mut result, &mut next, cursor);
        }
    }
    emit(&mut result, &mut next, cursor);
    result
}

/// `properties` without the bookmarks at its content's start when `slot`
/// holds breaks, and those bookmarks.
pub(crate) fn split_slot_bookmarks(properties: &V, slot: bool) -> (V, Vec<V>) {
    let Some(bookmarks) = properties.get("bookmarks").arr().cloned().filter(|_| slot) else {
        return (properties.clone(), Vec::new());
    };
    let at_start = |raw: &V| {
        let kind = raw.get("kind").as_str();
        o(raw).is_some()
            && matches!(kind.as_deref(), Some("start" | "end"))
            && raw.get("offset").finite().unwrap_or(0.0) == 0.0
    };
    let slotted: Vec<V> = bookmarks
        .items()
        .into_iter()
        .filter(|raw| at_start(raw))
        .collect();
    if slotted.is_empty() {
        return (properties.clone(), Vec::new());
    }
    let kept: Vec<V> = bookmarks
        .items()
        .into_iter()
        .filter(|raw| !at_start(raw))
        .collect();
    let rest = o(properties)
        .unwrap_or_default()
        .with(&[("bookmarks", V::array(kept))]);
    (
        V::Obj(rest),
        bookmark_boundaries(&obj! { "bookmarks": V::array(slotted) }),
    )
}

pub(crate) fn bookmark_node(boundary: &V) -> V {
    if boundary_kind(boundary) == "start" {
        let node = Obj::new();
        node.set("type", V::str("bookmarkStart"));
        node.set("id", boundary.get("id"));
        node.set("name", boundary.get("name").or(|| V::str("")));
        if !boundary.get("colFirst").is_undef() {
            node.set("colFirst", boundary.get("colFirst"));
        }
        if !boundary.get("colLast").is_undef() {
            node.set("colLast", boundary.get("colLast"));
        }
        node.set("position", obj! { "offset": boundary.get("offset") });
        V::Obj(node)
    } else {
        obj! {
            "type": "bookmarkEnd",
            "id": boundary.get("id"),
            "position": obj! { "offset": boundary.get("offset") },
        }
    }
}

pub(crate) fn bookmark_boundaries(properties: &V) -> Vec<V> {
    let Some(bookmarks) = properties.get("bookmarks").arr().cloned() else {
        return Vec::new();
    };
    let mut result = Vec::new();
    for raw in bookmarks.items() {
        let Some(bookmark) = o(&raw) else { continue };
        let Some(id) = bookmark.get("id").finite() else {
            continue;
        };
        let offset = bookmark.get("offset").finite().unwrap_or(0.0);
        let entry = |kind: &str, offset: f64| {
            obj! {
                "id": id,
                "kind": kind,
                "offset": offset,
                "name": st(&bookmark.get("name")),
                "colFirst": bookmark.get("colFirst").finite(),
                "colLast": bookmark.get("colLast").finite(),
                "breaksAfter": bookmark.get("breaksAfter").finite(),
                "inHyperlink": bookmark.get("inHyperlink").same(&V::Bool(true)),
            }
        };
        match bookmark.get("kind").as_str().as_deref() {
            Some("start") => result.push(entry("start", offset)),
            Some("end") => result.push(entry("end", offset)),
            _ => {
                result.push(entry("start", 0.0));
                result.push(entry("end", 9_007_199_254_740_991.0));
            }
        }
    }
    result
}

/// Insert zero-width markers before folding projected children back into fields.
fn insert_bookmark_items(items: Vec<V>, bookmarks: Vec<V>) -> Vec<V> {
    if bookmarks.is_empty() {
        return items;
    }
    let width = |item: &V| -> f64 {
        if is_kind(item, "text") {
            text::len(&item.get("text").to_js_string()) as f64
        } else if item.get("embedKind").as_str().as_deref() == Some("sdt")
            && !item.get("attributes").get("hyperlink").truthy()
        {
            2.0
        } else {
            1.0
        }
    };
    let length: f64 = items.iter().map(width).sum();
    let mut sorted: Vec<V> = bookmarks
        .iter()
        .map(|marker| {
            let offset = marker.get("offset").to_number();
            V::Obj(o(marker).unwrap_or_default().with(&[(
                "offset",
                V::Num(if offset == 9_007_199_254_740_991.0 {
                    length
                } else {
                    offset
                }),
            )]))
        })
        .collect();
    sorted.sort_by(|a, b| {
        a.get("offset")
            .to_number()
            .partial_cmp(&b.get("offset").to_number())
            .unwrap_or(Ordering::Equal)
    });
    let mut result = Vec::new();
    let mut cursor = 0.0;
    let mut next = 0;
    let emit =
        |result: &mut Vec<V>, next: &mut usize, at: f64, before: Option<&V>, after: Option<&V>| {
            while *next < sorted.len() && sorted[*next].get("offset").to_number() == at {
                let marker = o(&sorted[*next]).unwrap_or_default().spread();
                *next += 1;
                marker.delete("offset");
                if marker.has("run") {
                    let attributes = marker.get("attributes");
                    marker.delete("attributes");
                    result.push(obj! {
                        "kind": "embed",
                        "embedKind": "bookmark",
                        "payload": V::Obj(marker),
                        "attributes": attributes,
                    });
                    continue;
                }
                marker.delete("breaksAfter");
                let candidates = if marker.get("kind").as_str().as_deref() == Some("end") {
                    [before, after]
                } else {
                    [after, before]
                };
                let container = if marker.get("inHyperlink").truthy() {
                    candidates
                        .into_iter()
                        .flatten()
                        .find(|item| item.get("attributes").get("hyperlink").truthy())
                } else {
                    None
                };
                let attributes = Obj::new();
                if let Some(container) = container {
                    for key in ["hyperlink", "fieldResult"] {
                        let value = container.get("attributes").get(key);
                        if value.truthy() {
                            attributes.set(key, value);
                        }
                    }
                }
                result.push(obj! {
                    "kind": "embed",
                    "embedKind": "bookmark",
                    "payload": V::Obj(marker),
                    "attributes": V::Obj(attributes),
                });
            }
        };
    for (index, item) in items.iter().enumerate() {
        let before = index.checked_sub(1).map(|previous| &items[previous]);
        emit(&mut result, &mut next, cursor, before, Some(item));
        let end = cursor + width(item);
        if is_kind(item, "text") {
            let source = item.get("text").to_js_string();
            let mut start = cursor;
            while next < sorted.len() && sorted[next].get("offset").to_number() < end {
                let at = sorted[next].get("offset").to_number();
                result.push(V::Obj(o(item).unwrap_or_default().with(&[(
                    "text",
                    V::from(text::slice(
                        &source,
                        (start - cursor) as usize,
                        (at - cursor) as usize,
                    )),
                )])));
                emit(&mut result, &mut next, at, Some(item), Some(item));
                start = at;
            }
            if start < end {
                result.push(V::Obj(o(item).unwrap_or_default().with(&[(
                    "text",
                    V::from(text::slice_from(&source, (start - cursor) as usize)),
                )])));
            }
        } else {
            result.push(item.clone());
        }
        cursor = end;
    }
    emit(&mut result, &mut next, cursor, items.last(), None);
    result
}

fn restore_raw_inlines(content: Vec<V>, base: Option<&V>) -> Vec<V> {
    let Some(base) = base else { return content };
    let length: f64 = content.iter().map(paragraph_content_length).sum();
    let base_content = base.get("content").items();
    let mut boundaries = Vec::new();
    let mut offset: f64 = 0.0;
    for (index, child) in base_content.iter().enumerate() {
        if is(child, "rawXml") {
            boundaries
                .push(obj! { "id": index as f64, "kind": "start", "offset": offset.min(length) });
        }
        offset += paragraph_content_length(child);
    }
    insert_boundaries(
        content,
        &boundaries,
        &paragraph_content_length,
        &|boundary| base_content[boundary.get("id").to_number() as usize].clone(),
        false,
    )
}

/// The keys a paragraph's save attributes start from (null, `hangingIndent` false).
const PARAGRAPH_ATTR_DEFAULTS: &[&str] = &[
    "paraId",
    "textId",
    "alignment",
    "spaceBefore",
    "spaceAfter",
    "lineSpacing",
    "lineSpacingRule",
    "spacingExplicit",
    "indentLeft",
    "indentRight",
    "indentFirstLine",
    "hangingIndent",
    "numPr",
    "numPrFromStyle",
    "listNumFmt",
    "listIsBullet",
    "listMarker",
    "listMarkerHidden",
    "listMarkerFontFamily",
    "listMarkerFontSize",
    "listMarkerSuffix",
    "listLevelNumFmts",
    "listAbstractNumId",
    "listStartOverride",
    "styleId",
    "borders",
    "shading",
    "tabs",
    "pageBreakBefore",
    "renderedPageBreakBefore",
    "keepNext",
    "keepLines",
    "widowControl",
    "contextualSpacing",
    "snapToGrid",
    "autoSpaceDE",
    "autoSpaceDN",
    "defaultTextFormatting",
    "sectionBreakType",
    "bidi",
    "outlineLevel",
    "bookmarks",
    "_originalFormatting",
    "_originalRunBoundaries",
    "_sectionProperties",
    "pPrIns",
    "pPrDel",
    "pPrChange",
];

fn paragraph_attrs(properties: &V) -> V {
    let attrs = Obj::new();
    for key in PARAGRAPH_ATTR_DEFAULTS {
        attrs.set(
            key,
            if *key == "hangingIndent" {
                V::Bool(false)
            } else {
                V::Null
            },
        );
    }
    if let Some(properties) = properties.obj() {
        attrs.assign(properties);
    }
    attrs.set("styleId", properties.get("pStyle").or_else(|| V::Null));
    attrs.set(
        "_sectionProperties",
        properties.get("sectPr").or_else(|| V::Null),
    );
    attrs.delete("pStyle");
    attrs.delete("sectPr");
    V::Obj(attrs)
}

pub(crate) fn paragraph_from_story(
    para_id: &str,
    properties: &V,
    items: Vec<V>,
    comment_boundaries: &[V],
    base_paragraph: Option<&V>,
    seeded: &mut dyn FnMut(&V) -> Result<V>,
) -> Result<V> {
    let attrs = paragraph_attrs(properties);
    // The engine lists markers in their order at each point.
    let mut markers = Vec::new();
    for raw in properties.get("bookmarks").items() {
        let kind = raw.get("kind").as_str();
        if !matches!(kind.as_deref(), Some("fieldend" | "fieldseparate")) || o(&raw).is_none() {
            markers.extend(bookmark_boundaries(
                &obj! { "bookmarks": arr![raw.clone()] },
            ));
            continue;
        }
        let run = raw.get("run");
        let offset = raw.get("offset").finite();
        if raw.get("id").as_str().is_some()
            && let Some(offset) = offset
            && is(&run, "run")
            && run.get("content").arr().is_some()
        {
            markers.push(obj! {
                "id": raw.get("id"),
                "kind": raw.get("kind"),
                "offset": offset,
                "order": raw.get("order").finite().unwrap_or(0.0),
                "run": run,
                "attributes": o(&raw.get("attributes")).map(V::Obj).unwrap_or_else(|| obj! {}),
            });
        }
    }
    let mut content = build_paragraph_content(insert_bookmark_items(items.clone(), markers))?;
    content = restore_original_runs(
        content,
        &items,
        attrs
            .get("_originalRunBoundaries")
            .arr()
            .map(|runs| runs.items()),
    );
    content = restore_raw_inlines(content, base_paragraph);
    content = insert_boundaries(
        content,
        comment_boundaries,
        &story_units,
        &comment_marker,
        false,
    );
    let formatting = super::format::paragraph_attrs_to_formatting(&attrs, &seeded(&attrs)?);
    let paragraph = Obj::new();
    paragraph.set("type", V::str("paragraph"));
    paragraph.set(
        "paraId",
        if para_id.is_empty() {
            V::Undef
        } else {
            V::str(para_id)
        },
    );
    paragraph.set(
        "textId",
        base_paragraph
            .map(|base| base.get("textId"))
            .unwrap_or_default(),
    );
    paragraph.set("formatting", formatting);
    paragraph.set("content", V::array(content));
    if base_paragraph.is_some_and(|base| base.get("renderedPageBreakBefore").truthy()) {
        paragraph.set("renderedPageBreakBefore", V::Bool(true));
    }
    if let Some(info) = tracked_info(&properties.get("pPrIns")) {
        paragraph.set("pPrIns", info);
    }
    if let Some(info) = tracked_info(&properties.get("pPrDel")) {
        paragraph.set("pPrDel", info);
    }
    if let Some(changes) = properties
        .get("pPrChange")
        .arr()
        .filter(|changes| !changes.is_empty())
    {
        // A split's copy carries an editor revision id; the file takes numbers.
        paragraph.set(
            "propertyChanges",
            V::array(
                changes
                    .items()
                    .iter()
                    .map(|change| {
                        let info =
                            tracked_info(&change.get("info")).unwrap_or_else(|| change.get("info"));
                        V::Obj(o(change).unwrap_or_default().with(&[("info", info)]))
                    })
                    .collect(),
            ),
        );
    }
    if properties.get("sectPr").truthy() {
        paragraph.set("sectionProperties", properties.get("sectPr"));
    } else if properties.get("sectionBreakType").truthy() {
        paragraph.set(
            "sectionProperties",
            obj! { "sectionStart": properties.get("sectionBreakType") },
        );
    }
    Ok(V::Obj(paragraph))
}

pub(crate) fn page_break_paragraph() -> V {
    obj! {
        "type": "paragraph",
        "content": arr![obj! {
            "type": "run",
            "content": arr![obj! { "type": "break", "breakType": "page" }],
        }],
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum TokenKind {
    Page,
    Column,
    Visible,
    Mark,
    Start,
    End,
    Reference,
}

/// A paragraph's break, text, bookmark or comment boundary as the seed reads it.
#[derive(Clone)]
pub(crate) struct FlowToken {
    pub kind: TokenKind,
    pub run: Option<V>,
    pub index: Option<usize>,
    pub tracked: Option<&'static str>,
    pub id: Option<f64>,
}

impl FlowToken {
    fn new(kind: TokenKind) -> Self {
        Self {
            kind,
            run: None,
            index: None,
            tracked: None,
            id: None,
        }
    }

    pub fn is_break(&self) -> bool {
        matches!(self.kind, TokenKind::Page | TokenKind::Column)
    }
}

/// A break unit opening a paragraph slot.
#[derive(Clone)]
pub(crate) struct SlotBreak {
    pub kind: &'static str,
    pub at: f64,
    pub leading: bool,
    pub trailing: bool,
    pub attributes: V,
}

pub(crate) fn tracked_kind(attributes: &V) -> Option<&'static str> {
    if attributes.get("ins").truthy() {
        Some("ins")
    } else if attributes.get("del").truthy() {
        Some("del")
    } else {
        None
    }
}

/// A slot's break as a run, inside the tracked change it carries.
pub(crate) fn slot_break_content(entry: &SlotBreak) -> Result<V> {
    let attributes = &entry.attributes;
    let Some(info) = tracked_info(&attributes.get("ins").or_else(|| attributes.get("del"))) else {
        return Ok(break_run(entry.kind));
    };
    let item = obj! {
        "kind": "embed",
        "embedKind": if entry.kind == "page" { "pageBreak" } else { "columnBreak" },
        "payload": obj! {},
        "attributes": attributes.clone(),
    };
    tracked_content_for_item(&item, &info, true)
}

/// A paragraph's breaks and visible content in the order the seed reads them.
pub(crate) fn flow_tokens(content: &[V]) -> Vec<FlowToken> {
    let mut tokens = Vec::new();
    fn run(tokens: &mut Vec<FlowToken>, item: &V, tracked: Option<&'static str>) {
        for (index, entry) in item.get("content").items().iter().enumerate() {
            let kind = entry.get("breakType").as_str();
            if is(entry, "break")
                && !is_unit_break(entry)
                && matches!(kind.as_deref(), Some("page" | "column"))
            {
                tokens.push(FlowToken {
                    kind: if kind.as_deref() == Some("page") {
                        TokenKind::Page
                    } else {
                        TokenKind::Column
                    },
                    run: Some(item.clone()),
                    index: Some(index),
                    tracked,
                    id: None,
                });
            } else if is(entry, "commentReference") {
                tokens.push(FlowToken::new(TokenKind::Reference));
            } else if !is(entry, "text") || entry.get("text").to_js_string() != "" {
                tokens.push(FlowToken::new(TokenKind::Visible));
            }
        }
    }
    fn inline(tokens: &mut Vec<FlowToken>, items: &[V], tracked: Option<&'static str>) {
        for item in items {
            match &*ty(item) {
                "run" => run(tokens, item, tracked),
                "hyperlink" => {
                    let children = item.get("structuredChildren");
                    let children = if children.nullish() {
                        item.get("children")
                    } else {
                        children
                    };
                    inline(tokens, &children.items(), tracked);
                }
                "simpleField" => {
                    tokens.push(FlowToken::new(TokenKind::Visible));
                    let result = item.get("structuredResult").get("inline");
                    let result = if result.nullish() {
                        item.get("content")
                    } else {
                        result
                    };
                    inline(tokens, &result.items(), tracked);
                }
                "bookmarkStart" | "bookmarkEnd" => tokens.push(FlowToken::new(TokenKind::Mark)),
                "commentRangeEnd" => tokens.push(FlowToken {
                    id: item.get("id").as_num(),
                    ..FlowToken::new(TokenKind::End)
                }),
                "commentRangeStart" => tokens.push(FlowToken {
                    id: item.get("id").as_num(),
                    ..FlowToken::new(TokenKind::Start)
                }),
                "complexField" => {
                    let code = item.get("structuredCode").get("inline");
                    let code = if code.nullish() {
                        item.get("fieldCode")
                    } else {
                        code
                    };
                    let result = item.get("structuredResult").get("inline");
                    let result = if result.nullish() {
                        item.get("fieldResult")
                    } else {
                        result
                    };
                    let mut children = code.items();
                    children.extend(result.items());
                    inline(tokens, &children, tracked);
                }
                "inlineSdt" => inline(tokens, &item.get("content").items(), tracked),
                "insertion" | "moveTo" => inline(tokens, &item.get("content").items(), Some("ins")),
                "deletion" | "moveFrom" => {
                    inline(tokens, &item.get("content").items(), Some("del"))
                }
                "mathEquation" => tokens.push(FlowToken::new(TokenKind::Visible)),
                _ => {}
            }
        }
    }
    inline(&mut tokens, content, None);
    tokens
}

/// Splits a paragraph's breaks as the seed places them (`splitFlow`).
pub(crate) fn split_flow(
    tokens: &[FlowToken],
    referenced: &mut dyn FnMut(f64) -> bool,
) -> (Vec<FlowToken>, Vec<FlowToken>) {
    if !tokens.iter().any(FlowToken::is_break) {
        return (Vec::new(), Vec::new());
    }
    let last_break = tokens.iter().rposition(FlowToken::is_break);
    let mut split = None;
    for (index, token) in tokens.iter().enumerate() {
        let found = match token.kind {
            TokenKind::Visible => true,
            TokenKind::Reference => last_break.is_some_and(|last| index as isize > last as isize),
            TokenKind::End => {
                !referenced(token.id.unwrap_or(f64::NAN))
                    && !tokens[index + 1..].iter().any(FlowToken::is_break)
            }
            _ => false,
        };
        if found {
            split = Some(index);
            break;
        }
    }
    let split = split
        .or_else(|| {
            tokens
                .iter()
                .rposition(|token| token.kind == TokenKind::Mark)
        })
        .unwrap_or_else(|| {
            tokens
                .iter()
                .rposition(|token| token.kind == TokenKind::Column)
                .map_or(0, |index| index + 1)
        });
    let breaks = |part: &[FlowToken]| {
        part.iter()
            .filter(|token| token.is_break())
            .cloned()
            .collect()
    };
    (breaks(&tokens[..split]), breaks(&tokens[split..]))
}

/// The text a paragraph's content shows, and its story units.
fn text_signature(content: &[V]) -> String {
    let mut texts = String::new();
    fn walk(texts: &mut String, items: &[V]) {
        for item in items {
            match &*ty(item) {
                "run" => {
                    for entry in item.get("content").items() {
                        if is(&entry, "text") {
                            texts.push_str(&entry.get("text").to_js_string());
                        }
                    }
                }
                "hyperlink" => {
                    let children = item.get("structuredChildren");
                    let children = if children.nullish() {
                        item.get("children")
                    } else {
                        children
                    };
                    walk(texts, &children.items());
                }
                "simpleField" | "inlineSdt" => walk(texts, &item.get("content").items()),
                "complexField" => walk(texts, &item.get("fieldResult").items()),
                "insertion" | "deletion" | "moveFrom" | "moveTo" => {
                    walk(texts, &item.get("content").items())
                }
                _ => {}
            }
        }
    }
    walk(&mut texts, content);
    let units: f64 = content.iter().map(story_units).sum();
    crate::jsv::stringify(&arr![units, texts]).unwrap_or_default()
}

/// The paragraph's trailing breaks where its base paragraph held them while
/// its text is unchanged, otherwise at its end.
pub(crate) fn place_trailing(
    content: Vec<V>,
    entries: &[SlotBreak],
    base: Option<&V>,
) -> Result<Vec<V>> {
    let mut appended = content.clone();
    for entry in entries {
        appended.push(slot_break_content(entry)?);
    }
    let Some(base) = base.filter(|_| !entries.is_empty()) else {
        return Ok(appended);
    };
    let mut places = Vec::new();
    let mut offset = 0.0;
    let mut visible = false;
    for item in base.get("content").items() {
        if !is(&item, "run") {
            let tokens = flow_tokens(std::slice::from_ref(&item));
            if tokens.iter().any(FlowToken::is_break) {
                return Ok(appended);
            }
            visible |= tokens.iter().any(|token| token.kind == TokenKind::Visible);
            offset += story_units(&item);
            continue;
        }
        for entry in item.get("content").items() {
            let kind = entry.get("breakType").as_str();
            if is(&entry, "break") && matches!(kind.as_deref(), Some("page" | "column")) {
                let index = places.len();
                if !visible {
                    continue;
                }
                let Some(expected) = entries.get(index) else {
                    return Ok(appended);
                };
                if Some(expected.kind) != kind.as_deref()
                    || tracked_kind(&expected.attributes).is_some()
                {
                    return Ok(appended);
                }
                places.push(obj! { "id": index as f64, "kind": "start", "offset": offset });
                continue;
            }
            visible |= !is(&entry, "text") || entry.get("text").to_js_string() != "";
            offset += run_content_units(&entry);
        }
    }
    if places.len() != entries.len()
        || text_signature(&content) != text_signature(&base.get("content").items())
    {
        return Ok(appended);
    }
    Ok(insert_boundaries(
        content,
        &places,
        &story_units,
        &|boundary| {
            // A break's content never fails to build.
            slot_break_content(&entries[boundary.get("id").to_number() as usize])
                .unwrap_or_default()
        },
        true,
    ))
}

/// `content` without the given break tokens; runs left empty are dropped.
pub(crate) fn without_breaks(content: &[V], tokens: &[FlowToken]) -> Vec<V> {
    let mut drop: HashMap<usize, HashSet<usize>> = HashMap::new();
    for token in tokens {
        let (Some(run), Some(index)) = (&token.run, token.index) else {
            continue;
        };
        if let Some(run) = run.obj() {
            drop.entry(run.id()).or_default().insert(index);
        }
    }
    if drop.is_empty() {
        return content.to_vec();
    }
    let run = |item: &V| -> Option<V> {
        let Some(indexes) = item.obj().and_then(|object| drop.get(&object.id())) else {
            return Some(item.clone());
        };
        let kept: Vec<V> = item
            .get("content")
            .items()
            .into_iter()
            .enumerate()
            .filter(|(index, _)| !indexes.contains(index))
            .map(|(_, entry)| entry)
            .collect();
        (!kept.is_empty()).then(|| {
            V::Obj(
                o(item)
                    .unwrap_or_default()
                    .with(&[("content", V::array(kept))]),
            )
        })
    };
    let runs = |items: &[V]| -> Vec<V> {
        items
            .iter()
            .flat_map(|item| {
                if is(item, "run") {
                    run(item)
                } else {
                    Some(item.clone())
                }
            })
            .collect()
    };
    fn inline(items: &[V], runs: &dyn Fn(&[V]) -> Vec<V>) -> Vec<V> {
        let mut out = Vec::new();
        for item in items {
            match &*ty(item) {
                "run" => out.extend(runs(std::slice::from_ref(item))),
                "hyperlink" => {
                    let structured = item.get("structuredChildren");
                    let copy = o(item).unwrap_or_default().spread();
                    if structured.truthy() {
                        let kept = inline(&structured.items(), runs);
                        copy.set("children", V::array(super::item::shown_runs(&kept)));
                        copy.set(
                            "structuredChildren",
                            V::array(inline(&structured.items(), runs)),
                        );
                    } else {
                        copy.set("children", V::array(runs(&item.get("children").items())));
                    }
                    out.push(V::Obj(copy));
                }
                "simpleField" | "complexField" => {
                    let simple = is(item, "simpleField");
                    let structured_result = item.get("structuredResult");
                    let result = structured_result.truthy().then(|| {
                        let copy = o(&structured_result).unwrap_or_default().spread();
                        copy.set(
                            "inline",
                            V::array(inline(&structured_result.get("inline").items(), runs)),
                        );
                        V::Obj(copy)
                    });
                    let structured_code = item.get("structuredCode");
                    let code = (!simple && structured_code.truthy()).then(|| {
                        let copy = o(&structured_code).unwrap_or_default().spread();
                        copy.set(
                            "inline",
                            V::array(inline(&structured_code.get("inline").items(), runs)),
                        );
                        V::Obj(copy)
                    });
                    let kept = o(item).unwrap_or_default().spread();
                    if let Some(result) = &result {
                        kept.set("structuredResult", result.clone());
                    }
                    if let Some(code) = &code {
                        kept.set("structuredCode", code.clone());
                    }
                    let tree = item.get("fieldTree");
                    if tree.truthy() {
                        let copy = o(&tree).unwrap_or_default().spread();
                        if let Some(result) = &result {
                            copy.set("result", result.clone());
                        }
                        if let Some(code) = &code {
                            copy.set("code", code.clone());
                        }
                        kept.set("fieldTree", V::Obj(copy));
                    }
                    if simple {
                        kept.set(
                            "content",
                            V::array(match &result {
                                Some(result) => {
                                    super::item::shown_runs(&result.get("inline").items())
                                }
                                None => runs(&item.get("content").items()),
                            }),
                        );
                    } else {
                        kept.set(
                            "fieldCode",
                            V::array(match &code {
                                Some(code) => super::item::shown_runs(&code.get("inline").items()),
                                None => runs(&item.get("fieldCode").items()),
                            }),
                        );
                        kept.set(
                            "fieldResult",
                            V::array(match &result {
                                Some(result) => {
                                    super::item::shown_runs(&result.get("inline").items())
                                }
                                None => runs(&item.get("fieldResult").items()),
                            }),
                        );
                    }
                    out.push(V::Obj(kept));
                }
                "inlineSdt" => out.push(V::Obj(o(item).unwrap_or_default().with(&[(
                    "content",
                    V::array(inline(&item.get("content").items(), runs)),
                )]))),
                "insertion" | "deletion" | "moveFrom" | "moveTo" => {
                    let children = item.get("content").items();
                    let kept = inline(&children, runs);
                    if !kept.is_empty() || children.is_empty() {
                        out.push(V::Obj(
                            o(item)
                                .unwrap_or_default()
                                .with(&[("content", V::array(kept))]),
                        ));
                    }
                }
                _ => out.push(item.clone()),
            }
        }
        out
    }
    inline(content, &runs)
}

/// A note number mark run, or a tracked-change wrapper holding only those.
fn is_note_mark(content: &V) -> bool {
    if matches!(
        &*ty(content),
        "insertion" | "deletion" | "moveFrom" | "moveTo"
    ) {
        let children = content.get("content").items();
        return !children.is_empty() && children.iter().all(is_note_mark);
    }
    let entries = content.get("content").items();
    is(content, "run")
        && !entries.is_empty()
        && entries
            .iter()
            .all(|entry| matches!(&*ty(entry), "footnoteRefMark" | "endnoteRefMark"))
}

fn has_note_mark(blocks: &[V]) -> bool {
    blocks.iter().any(|block| {
        is(block, "paragraph") && block.get("content").items().iter().any(is_note_mark)
    })
}

/// Reinstates the note number mark on a projected note.
pub(crate) fn restore_note_marks(projected: Vec<V>, base: &[V]) -> Vec<V> {
    if has_note_mark(&projected) {
        return projected;
    }
    let opening = base
        .iter()
        .find(|block| is(block, "paragraph"))
        .map(|block| block.get("content").items())
        .unwrap_or_default();
    let end = opening
        .iter()
        .position(|child| !is_note_mark(child))
        .unwrap_or(opening.len());
    let marks = &opening[..end];
    if marks.is_empty() {
        return projected;
    }
    let Some(index) = projected.iter().position(|block| is(block, "paragraph")) else {
        return projected;
    };
    let mut restored = projected.clone();
    let mut content = marks.to_vec();
    content.extend(projected[index].get("content").items());
    restored[index] = V::Obj(
        o(&projected[index])
            .unwrap_or_default()
            .with(&[("content", V::array(content))]),
    );
    restored
}
