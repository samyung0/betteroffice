//! Bookmark and continued-field boundaries follow text identities.

use std::cmp::Ordering;
use std::collections::{BTreeMap, HashMap, HashSet};
use std::sync::Arc;

use yrs::{Any, Assoc, IndexedSequence, Map, MapPrelim, Out, ReadTxn, TextRef, TransactionMut};

use crate::op::{OpError, OpResult};
use crate::ops::{ChunkKind, snapshot};
use crate::{KIND_KEY, PARA_ID, anchor_value, decode_anchor, map_string};

pub(crate) const ROOT: &str = "bookmarks";

pub(crate) fn set(
    txn: &mut TransactionMut<'_>,
    story: &TextRef,
    story_id: &str,
    index: u32,
    data: Any,
) -> OpResult<()> {
    let Any::Map(metadata) = &data else {
        return Err(OpError::InvalidUpdate(
            "bookmark metadata must be an object".into(),
        ));
    };
    let id = match metadata.get("id") {
        Some(Any::Number(id)) => id.to_string(),
        Some(Any::String(id)) => id.to_string(),
        _ => return Err(OpError::InvalidUpdate("inline boundary needs an id".into())),
    };
    let Some(Any::String(kind)) = metadata.get("kind") else {
        return Err(OpError::InvalidUpdate(
            "bookmark needs a boundary kind".into(),
        ));
    };
    let mut assoc = match kind.as_ref() {
        "start" => Assoc::After,
        "end" | "fieldend" | "fieldseparate" => Assoc::Before,
        _ => {
            return Err(OpError::InvalidUpdate(
                "invalid bookmark boundary kind".into(),
            ));
        }
    };
    let root = txn.get_map(ROOT).expect("bookmarks root is declared");
    let opposite = if kind.as_ref() == "start" {
        "end"
    } else {
        "start"
    };
    let other = match root.get(txn, &format!("{story_id}:{id}:{opposite}")) {
        Some(Out::YMap(entry)) => Some(entry),
        _ => None,
    };
    let collapsed = other.as_ref().is_some_and(|entry| {
        let Some(Out::Any(Any::Array(anchors))) = entry.get(txn, "anchors") else {
            return false;
        };
        anchors
            .first()
            .and_then(|value| decode_anchor(value).ok())
            .and_then(|anchor| anchor.start.get_offset(txn))
            .is_some_and(|offset| offset.index == index)
    });
    if collapsed {
        assoc = Assoc::Before;
    }
    let anchor = story
        .sticky_index(txn, index, assoc)
        .ok_or_else(|| OpError::InvalidUpdate("bookmark position does not resolve".into()))?;
    let anchors = Any::Array(Arc::from([anchor_value(story_id, &anchor, &anchor)]));
    if collapsed && let Some(other) = other {
        other.insert(txn, "anchors", anchors.clone());
    }
    let key = format!("{story_id}:{id}:{kind}");
    let entry = root.insert(txn, key, MapPrelim::default());
    entry.insert(txn, "data", data);
    entry.insert(txn, "anchors", anchors);
    Ok(())
}

pub(crate) fn positions<T: ReadTxn>(txn: &T, story_id: &str) -> Vec<(u32, Any)> {
    let Some(root) = txn.get_map(ROOT) else {
        return Vec::new();
    };
    let mut result = Vec::new();
    for (_, value) in root.iter(txn) {
        let Out::YMap(entry) = value else {
            continue;
        };
        let (Some(Out::Any(data)), Some(Out::Any(Any::Array(anchors)))) =
            (entry.get(txn, "data"), entry.get(txn, "anchors"))
        else {
            continue;
        };
        for encoded in anchors.iter() {
            if let Ok(anchor) = decode_anchor(encoded)
                && anchor.story == story_id
                && let Some(at) = anchor.start.get_offset(txn)
            {
                result.push((at.index, data.clone()));
            }
        }
    }
    // Deleting a range can collapse its opposite-facing anchors. Typing in
    // that gap must keep an empty bookmark, never invert its start and end.
    let ends: HashMap<i64, u32> = result
        .iter()
        .filter_map(|(at, data)| {
            let Any::Map(data) = data else {
                return None;
            };
            let (Some(Any::Number(id)), Some(Any::String(kind))) =
                (data.get("id"), data.get("kind"))
            else {
                return None;
            };
            (kind.as_ref() == "end").then_some((*id as i64, *at))
        })
        .collect();
    for (at, data) in &mut result {
        if let Any::Map(data) = data
            && let (Some(Any::Number(id)), Some(Any::String(kind))) =
                (data.get("id"), data.get("kind"))
            && kind.as_ref() == "start"
            && let Some(end) = ends.get(&(*id as i64))
        {
            *at = (*at).min(*end);
        }
    }
    let ranges: HashMap<(i64, bool), (u32, f64)> = result
        .iter()
        .filter_map(
            |(at, data)| match (marker_field(data, "id"), marker_field(data, "kind")) {
                (Some(Any::Number(id)), Some(Any::String(kind)))
                    if matches!(kind.as_ref(), "start" | "end") =>
                {
                    Some((
                        (id as i64, kind.as_ref() == "start"),
                        (*at, marker_order(data)),
                    ))
                }
                _ => None,
            },
        )
        .collect();
    let mut keyed: Vec<_> = result
        .into_iter()
        .map(|marker| ((marker.0, marker_key(&marker, &ranges)), marker))
        .collect();
    keyed.sort_by(|(a, _), (b, _)| a.partial_cmp(b).unwrap_or(Ordering::Equal));
    keyed.into_iter().map(|(_, marker)| marker).collect()
}

fn marker_field(data: &Any, key: &str) -> Option<Any> {
    match data {
        Any::Map(data) => data.get(key).cloned(),
        _ => None,
    }
}

fn marker_order(data: &Any) -> f64 {
    match marker_field(data, "order") {
        Some(Any::Number(order)) => order,
        _ => f64::INFINITY,
    }
}

/// A total order at one point: seeded markers by source `order` (an end never
/// before its own start), then the rest with ends of earlier ranges first,
/// empty ranges, and starts of later ranges last.
fn marker_key(
    (at, data): &(u32, Any),
    ranges: &HashMap<(i64, bool), (u32, f64)>,
) -> (f64, u8, f64, String, u8) {
    let kind = match marker_field(data, "kind") {
        Some(Any::String(kind)) => kind.to_string(),
        _ => String::new(),
    };
    let (number, name) = match marker_field(data, "id") {
        Some(Any::Number(id)) => (id, String::new()),
        Some(Any::String(id)) => (0.0, id.to_string()),
        _ => (0.0, String::new()),
    };
    let partner = matches!(kind.as_str(), "start" | "end")
        .then(|| ranges.get(&(number as i64, kind == "end")))
        .flatten();
    let mut order = marker_order(data);
    if kind == "end"
        && let Some((_, start)) = partner
    {
        order = order.max(*start);
    }
    let place = match (kind.as_str(), partner) {
        ("start" | "end", Some((other, _))) if other == at => 1,
        ("end", _) => 0,
        ("start", _) => 2,
        _ => 1,
    };
    let kind = match kind.as_str() {
        "start" => 0,
        "fieldseparate" => 1,
        "end" => 2,
        _ => 3,
    };
    (order, place, number, name, kind)
}

/// Removing an empty paragraph before a block hands its markers to the
/// surviving paragraph, on the same side of its breaks as the save.
pub(crate) fn move_range(
    txn: &mut TransactionMut<'_>,
    story: &TextRef,
    story_id: &str,
    from: u32,
    to: u32,
    target: u32,
) -> OpResult<()> {
    for (_, data) in positions(txn, story_id)
        .into_iter()
        .filter(|(at, _)| *at >= from && *at < to)
    {
        set(txn, story, story_id, target, data)?;
    }
    Ok(())
}

/// The save's paragraph-relative offsets, computed from live anchors. A
/// bookmark among leading breaks records how many breaks still follow it.
pub(crate) fn paragraph_properties<T: ReadTxn>(
    txn: &T,
    story_id: &str,
    story: &TextRef,
) -> HashMap<String, Any> {
    let markers = positions(txn, story_id);
    if markers.is_empty() {
        return HashMap::new();
    }
    let chunks = snapshot(story, txn);
    fn nested_fields(value: &Any, fields: &mut HashSet<String>) {
        match value {
            Any::Array(values) => {
                for value in values.iter() {
                    nested_fields(value, fields);
                }
            }
            Any::Map(values) => {
                if matches!(values.get("modelKind"), Some(Any::String(kind)) if kind.as_ref() == "field")
                    && let Some(Any::String(id)) = values.get("continuationId")
                {
                    fields.insert(id.to_string());
                }
                if matches!(values.get("type"), Some(Any::String(kind)) if kind.as_ref() == "complexField")
                    && let Some(Any::Map(continuation)) = values.get("continuation")
                    && let Some(Any::String(id)) = continuation.get("id")
                {
                    fields.insert(id.to_string());
                }
                if let Some(Any::String(data)) = values.get("fieldData") {
                    field_data_continuations(data, fields);
                }
                for (key, value) in values.iter() {
                    if !matches!(key.as_str(), "blocks" | "fieldTree") {
                        nested_fields(value, fields);
                    }
                }
            }
            _ => {}
        }
    }
    fn field_data_continuations(data: &str, fields: &mut HashSet<String>) {
        if data.contains("\"continuation\"")
            && let Ok(value) = Any::from_json(data)
        {
            nested_fields(&value, fields);
        }
    }
    let mut fields = HashSet::new();
    for chunk in &chunks {
        if let ChunkKind::Embed(Some(map)) = &chunk.kind {
            if let Some(id) = map_string(map, txn, "continuationId") {
                fields.insert(id);
            }
            if let Some(data) = map_string(map, txn, "fieldData") {
                field_data_continuations(&data, &mut fields);
            }
            if let Some(Out::Any(content)) = map.get(txn, "content") {
                nested_fields(&content, &mut fields);
            }
        }
    }
    let mut result = HashMap::new();
    let mut first = 0;
    let mut marker_index = 0;
    for (last, chunk) in chunks.iter().enumerate() {
        let ChunkKind::Pilcrow(map) = &chunk.kind else {
            continue;
        };
        let mut head = first;
        while head < last
            && matches!(&chunks[head].kind, ChunkKind::Embed(Some(map))
            if map_string(map, txn, KIND_KEY).is_some_and(|kind| crate::segments::is_block_embed(&kind)))
        {
            head += 1;
        }
        let content_start = chunks[head].start;
        let mut bookmarks = Vec::new();
        while let Some((at, Any::Map(data))) = markers.get(marker_index)
            && *at <= chunk.start
        {
            let mut data = data.as_ref().clone();
            if let Some(Any::String(id)) = data.get("id")
                && !fields.contains(id.as_ref())
            {
                marker_index += 1;
                continue;
            }
            let extra = chunks[head..last].iter().filter(|unit| unit.start < *at && !unit.attr_active("hyperlink")
                && matches!(&unit.kind, ChunkKind::Embed(Some(map)) if map_string(map, txn, KIND_KEY).as_deref() == Some("sdt"))).count() as u32;
            data.insert(
                "offset".into(),
                Any::Number((at.saturating_sub(content_start) + extra).into()),
            );
            let after = chunks[first..head].iter().filter(|unit| unit.start >= *at
                && matches!(&unit.kind, ChunkKind::Embed(Some(map)) if map_string(map, txn, KIND_KEY).is_some_and(|kind| kind.ends_with("Break")))).count();
            data.remove("breaksAfter");
            if after > 0 {
                data.insert("breaksAfter".into(), Any::Number(after as f64));
            }
            bookmarks.push(Any::Map(Arc::new(data)));
            marker_index += 1;
        }
        if !bookmarks.is_empty()
            && let Some(id) = map_string(map, txn, PARA_ID)
        {
            result.insert(id, Any::Array(Arc::from(bookmarks)));
        }
        first = last + 1;
    }
    result
}

pub(crate) fn project(
    properties: &mut BTreeMap<String, Any>,
    id: &str,
    bookmarks: &HashMap<String, Any>,
) {
    if let Some(value) = bookmarks.get(id) {
        properties.insert("bookmarks".into(), value.clone());
    }
}
