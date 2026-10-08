//! Bookmark and continued-field boundaries follow text identities.

use std::cmp::Ordering;
use std::collections::{BTreeMap, HashMap, HashSet};
use std::sync::Arc;

use yrs::{
    Any, Assoc, IdSet, IndexedSequence, Map, MapPrelim, Out, ReadTxn, StickyIndex, TextRef,
    Transact, TransactionMut,
};

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

/// Re-anchors the bookmarks and comment ranges whose own text this replica's
/// Undo or Redo restored onto that text. It restores text as new items,
/// which markers anchored to the old ones follow only on this replica,
/// through yrs's redone links: another replica and the save would place them
/// elsewhere. A marker in text that stays deleted, or was deleted again,
/// keeps its anchor, even beside restored text, so the Undo that restores
/// its own text brings it back. Runs after every Undo or Redo step and after
/// a peer's update that writes markers, which may name text this replica
/// restored earlier (a comment another peer removed and restored).
///
/// A continued field's end is re-anchored too, so every peer and the save end
/// the field where the peer that pressed Undo or Redo shows it (decided
/// 2026-10-02). Its separate keeps its anchor: re-anchoring it made a rebase
/// of Undo and Redo of a join that removed a nested continued field refuse
/// (the matrix's `join nested continued field, undo, redo`).
pub(crate) fn rebind(doc: &yrs::Doc) {
    let mut txn = doc.transact_mut_with("system");
    let entries: Vec<_> = [ROOT, crate::COMMENTS]
        .into_iter()
        .filter_map(|name| txn.get_map(name))
        .flat_map(|root| {
            root.iter(&txn)
                .filter_map(|(_, value)| match value {
                    Out::YMap(entry)
                        if !matches!(entry.get(&txn, "data"), Some(Out::Any(Any::Map(data)))
                            if matches!(data.get("kind"), Some(Any::String(kind)) if kind.as_ref() == "fieldseparate")) =>
                    {
                        Some(entry)
                    }
                    _ => None,
                })
                .collect::<Vec<_>>()
        })
        .collect();
    for entry in entries {
        let Some(Out::Any(Any::Array(anchors))) = entry.get(&txn, "anchors") else {
            continue;
        };
        let mut changed = false;
        let mut rebound = Vec::with_capacity(anchors.len());
        for encoded in anchors.iter() {
            let fresh = decode_anchor(encoded).ok().and_then(|anchor| {
                let story = crate::story_ref(&txn, &anchor.story).ok()?;
                let restore = |sticky: &yrs::StickyIndex| {
                    restored_copy(&txn, sticky.id()?)?;
                    let index = sticky.get_offset(&txn)?.index;
                    let fresh = story.sticky_index(&txn, index, sticky.assoc)?;
                    (fresh != *sticky).then_some(fresh)
                };
                let (start, end) = (restore(&anchor.start), restore(&anchor.end));
                (start.is_some() || end.is_some()).then(|| {
                    anchor_value(
                        &anchor.story,
                        start.as_ref().unwrap_or(&anchor.start),
                        end.as_ref().unwrap_or(&anchor.end),
                    )
                })
            });
            changed |= fresh.is_some();
            rebound.push(fresh.unwrap_or_else(|| encoded.clone()));
        }
        if changed {
            entry.insert(&mut txn, "anchors", Any::Array(Arc::from(rebound)));
        }
    }
}

/// The live copy this replica's Undo or Redo made of the item `id` names,
/// following its redone links; `None` when there is none or it was deleted
/// again.
fn restored_copy<T: ReadTxn>(txn: &T, id: &yrs::ID) -> Option<yrs::block::ItemPtr> {
    let mut item = txn.store().get_item(id)?;
    let mut offset = id.clock - item.id().clock;
    let mut redone = false;
    while let Some(next) = item.redone() {
        let next = yrs::ID::new(next.client, next.clock + offset);
        item = txn.store().get_item(&next)?;
        offset = next.clock - item.id().clock;
        redone = true;
    }
    (redone && !item.is_deleted()).then_some(item)
}

/// Whether `inserted` (what a transaction inserted) writes a bookmark or a
/// comment: an entry of either root, or a key of such an entry.
pub(crate) fn writes_markers<T: ReadTxn>(txn: &T, inserted: &IdSet) -> bool {
    use yrs::branch::{Branch, BranchPtr};
    let roots: Vec<BranchPtr> = [ROOT, crate::COMMENTS]
        .into_iter()
        .filter_map(|name| txn.get_map(name))
        .map(|root| BranchPtr::from(<yrs::MapRef as AsRef<Branch>>::as_ref(&root)))
        .collect();
    let in_roots = |branch: Option<BranchPtr>| branch.is_some_and(|branch| roots.contains(&branch));
    inserted.iter().any(|(client, ranges)| {
        ranges.iter().any(|range| {
            let mut clock = range.start;
            while clock < range.end {
                let Some(item) = txn.store().get_item(&yrs::ID::new(*client, clock)) else {
                    return false;
                };
                clock = item.id().clock + item.len();
                let parent = item.parent_branch();
                if in_roots(parent)
                    || in_roots(
                        parent
                            .and_then(|entry| entry.item())
                            .and_then(|entry| entry.parent_branch()),
                    )
                {
                    return true;
                }
            }
            false
        })
    })
}

pub(crate) fn positions<T: ReadTxn>(txn: &T, story_id: &str) -> Vec<(u32, Any)> {
    positions_by_story_where(txn, |story| story == story_id)
        .remove(story_id)
        .unwrap_or_default()
}

/// `positions` of every story, reading the anchors once.
pub(crate) fn positions_by_story<T: ReadTxn>(txn: &T) -> HashMap<String, Vec<(u32, Any)>> {
    positions_by_story_where(txn, |_| true)
}

/// Markers of the stories `wanted` picks, every anchor resolved in one walk
/// of its story rather than one walk per anchor.
fn positions_by_story_where<T: ReadTxn>(
    txn: &T,
    wanted: impl Fn(&str) -> bool,
) -> HashMap<String, Vec<(u32, Any)>> {
    let Some(root) = txn.get_map(ROOT) else {
        return HashMap::new();
    };
    let mut found: Vec<(String, Any)> = Vec::new();
    let mut starts: Vec<StickyIndex> = Vec::new();
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
                && wanted(&anchor.story)
            {
                found.push((anchor.story, data.clone()));
                starts.push(anchor.start);
            }
        }
    }
    #[cfg(test)]
    tests::RESOLVED_BATCHES.with(|batches| batches.borrow_mut().push(starts.len()));
    let mut stories: HashMap<String, Vec<(u32, Any)>> = HashMap::new();
    for ((story, data), at) in found
        .into_iter()
        .zip(StickyIndex::get_offsets(txn, &starts))
    {
        if let Some(at) = at {
            stories.entry(story).or_default().push((at.index, data));
        }
    }
    stories
        .into_iter()
        .map(|(story, markers)| (story, ordered(markers)))
        .collect()
}

/// One story's markers in save order.
fn ordered(mut result: Vec<(u32, Any)>) -> Vec<(u32, Any)> {
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
/// surviving paragraph, on the same side of its breaks as the save, and
/// returns how many moved.
pub(crate) fn move_range(
    txn: &mut TransactionMut<'_>,
    story: &TextRef,
    story_id: &str,
    from: u32,
    to: u32,
    target: u32,
) -> OpResult<usize> {
    let moving: Vec<_> = positions(txn, story_id)
        .into_iter()
        .filter(|(at, _)| *at >= from && *at < to)
        .collect();
    for (_, data) in &moving {
        set(txn, story, story_id, target, data.clone())?;
    }
    Ok(moving.len())
}

/// The save's paragraph-relative offsets, computed from live anchors. A
/// bookmark among leading breaks records how many breaks still follow it.
pub(crate) fn paragraph_properties<T: ReadTxn>(
    txn: &T,
    story_id: &str,
    story: &TextRef,
) -> HashMap<String, Any> {
    paragraph_properties_at(txn, story, &positions(txn, story_id))
}

/// `paragraph_properties` with the story's `positions`.
pub(crate) fn paragraph_properties_at<T: ReadTxn>(
    txn: &T,
    story: &TextRef,
    markers: &[(u32, Any)],
) -> HashMap<String, Any> {
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

#[cfg(test)]
mod tests {
    use std::cell::RefCell;

    use super::*;
    use crate::{EditCtx, EditingDoc, FormatPolicy, Position, RawOp};

    thread_local! {
        /// The anchor count of every batch a bookmark read resolved.
        pub(super) static RESOLVED_BATCHES: RefCell<Vec<usize>> =
            const { RefCell::new(Vec::new()) };
    }

    fn resolved_batches() -> Vec<usize> {
        RESOLVED_BATCHES.with(|batches| batches.take())
    }

    fn bookmark(doc: &EditingDoc, story: &str, id: usize, start: u32, end: u32) {
        let boundary = |index, kind: &str| RawOp::SetBookmark {
            index,
            data: Any::Map(Arc::new(
                [
                    ("id".to_owned(), Any::Number(id as f64)),
                    ("kind".to_owned(), Any::from(kind)),
                ]
                .into(),
            )),
        };
        doc.apply_raw_ops(
            story,
            vec![boundary(start, "start"), boundary(end, "end")],
            &EditCtx::local("", ""),
        )
        .unwrap();
    }

    fn comment(doc: &EditingDoc, story: &str, id: String, start: u32, end: u32) {
        doc.apply_raw_ops(
            story,
            vec![RawOp::SetComment {
                id,
                ranges: vec![(start, end)],
                author: "Ada".into(),
                date: "2026-10-08T00:00:00Z".into(),
                body: Any::Null,
            }],
            &EditCtx::local("", ""),
        )
        .unwrap();
    }

    /// The batch resolution answers exactly what resolving each anchor on
    /// its own does, across two peers' inserts, deletes and undos over the
    /// anchors of several stories. Undo tracking keeps deleted text, as in
    /// an editing session, so anchors sit on deleted and redone items. (The
    /// yrs test beside `get_offsets` is outside the workspace, so this is
    /// the copy CI runs.)
    #[test]
    fn batch_anchor_resolution_matches_one_by_one() {
        let a = EditingDoc::new(1);
        let b = EditingDoc::new(2);
        a.create_story("body", "The handbook body text.", "Normal", "left")
            .unwrap();
        a.create_story("fn:1", "A footnote.", "Normal", "left")
            .unwrap();
        let sync = |from: &EditingDoc, to: &EditingDoc| {
            to.apply_update_v1(&from.encode_state_as_update_v1())
                .unwrap();
        };
        sync(&a, &b);
        let (undo_a, undo_b) = (crate::UndoSession::new(), crate::UndoSession::new());
        undo_a.track(&a);
        undo_b.track(&b);
        let mut seed = 7_u64;
        let mut next = |below: u32| {
            seed = seed
                .wrapping_mul(6_364_136_223_846_793_005)
                .wrapping_add(1_442_695_040_888_963_407);
            u32::try_from(seed >> 33).unwrap() % below.max(1)
        };
        for round in 0..24 {
            for story in ["body", "fn:1"] {
                let len = a.story_len(story).unwrap();
                // An `After` boundary needs content to its right.
                let start = next(len);
                let end = (start + 3).min(len);
                bookmark(&a, story, round, start, end);
                if start < end {
                    comment(&a, story, format!("{story}-{round}"), start, end);
                }
                let len = b.story_len(story).unwrap();
                if round % 3 == 2 && len > 4 {
                    let index = next(len - 3);
                    b.apply_raw_ops(
                        story,
                        vec![RawOp::Delete { index, len: 3 }],
                        &EditCtx::local("", ""),
                    )
                    .unwrap();
                } else {
                    b.insert_text(
                        &EditCtx::local("", ""),
                        Position::new(story, next(len + 1)),
                        "\u{1F600}x",
                        FormatPolicy::Inherit,
                    )
                    .unwrap();
                }
                undo_b.add_undo_barrier();
                if round % 5 == 4 {
                    assert!(undo_b.undo());
                }
            }
            sync(&b, &a);
            sync(&a, &b);
        }

        // `b` keeps the text it deleted for undo; `a` collected it.
        for doc in [&a, &b] {
            let txn = doc.yrs_doc().transact();
            let mut anchors = Vec::new();
            for root in [ROOT, crate::COMMENTS] {
                for (_, entry) in txn.get_map(root).unwrap().iter(&txn) {
                    let Out::YMap(entry) = entry else { continue };
                    let Some(Out::Any(Any::Array(encoded))) = entry.get(&txn, "anchors") else {
                        continue;
                    };
                    for anchor in encoded.iter().map(|value| decode_anchor(value).unwrap()) {
                        anchors.extend([anchor.start, anchor.end]);
                    }
                }
            }
            let resolve =
                |offset: Option<yrs::Offset>| offset.map(|offset| (offset.index, offset.assoc));
            let one_by_one: Vec<_> = anchors
                .iter()
                .map(|anchor| resolve(anchor.get_offset(&txn)))
                .collect();
            let batch: Vec<_> = StickyIndex::get_offsets(&txn, &anchors)
                .into_iter()
                .map(resolve)
                .collect();
            assert_eq!(batch, one_by_one);
            assert!(anchors.len() > 150);
            for assoc in [Assoc::After, Assoc::Before] {
                assert!(anchors.iter().any(|anchor| anchor.assoc == assoc));
            }
        }
    }

    /// Reading a story's segments resolves all of its bookmark anchors in one
    /// batch however many paragraphs carry them, and leaves other stories'
    /// anchors alone; reading every story is one batch as well.
    #[test]
    fn a_story_read_resolves_its_bookmarks_in_one_batch() {
        const PARAGRAPHS: u32 = 300;
        let ctx = EditCtx::local("", "");
        let doc = EditingDoc::new(9);
        let text: String = (0..PARAGRAPHS)
            .map(|index| format!("Paragraph {index:>3} here "))
            .collect();
        doc.create_story("body", &text, "Normal", "left").unwrap();
        // Each paragraph is 19 units, so its break sits at 20k - 1.
        for paragraph in 1..PARAGRAPHS {
            doc.split_paragraph(&ctx, Position::new("body", paragraph * 20 - 1))
                .unwrap();
        }
        for paragraph in 0..PARAGRAPHS {
            let start = paragraph * 20;
            bookmark(&doc, "body", paragraph as usize, start, start + 9);
        }
        doc.create_story("fn:1", "A footnote.", "Normal", "left")
            .unwrap();
        bookmark(&doc, "fn:1", 1000, 0, 1);
        resolved_batches();

        doc.story_segments("body").unwrap();
        assert_eq!(resolved_batches(), [2 * PARAGRAPHS as usize]);
        doc.all_story_segments();
        assert_eq!(resolved_batches(), [2 * PARAGRAPHS as usize + 2]);
        let txn = doc.yrs_doc().transact();
        let markers = positions(&txn, "body");
        assert_eq!(markers.len(), 2 * PARAGRAPHS as usize);
        assert_eq!(markers.last().unwrap().0, (PARAGRAPHS - 1) * 20 + 9);
    }

    /// A peer sets a bookmark's start again on text another peer deleted and
    /// restored with Undo: the restoring peer re-anchors it to its copy, so
    /// every peer reads the bookmark at the same place.
    #[test]
    fn a_boundary_naming_text_an_undo_restored_resolves_alike_on_every_peer() {
        let base = EditingDoc::new(1);
        base.create_story(
            "body",
            "Hello brave new world, and then some more words to read.",
            "Normal",
            "left",
        )
        .unwrap();
        bookmark(&base, "body", 7, 6, 21);
        let state = base.encode_state_as_update_v1();
        let peers: Vec<EditingDoc> = [701, 702, 703]
            .into_iter()
            .map(|client| {
                let doc = EditingDoc::new(client);
                doc.apply_update_v1(&state).unwrap();
                doc
            })
            .collect();
        let deleter = crate::UndoSession::new();
        deleter.track(&peers[2]);
        peers[2]
            .delete_range(
                &EditCtx::local("", ""),
                crate::StoryRange::new("body", 0, 30),
            )
            .unwrap();
        deleter.add_undo_barrier();
        assert!(deleter.undo());
        peers[1]
            .apply_raw_ops(
                "body",
                vec![RawOp::SetBookmark {
                    index: 6,
                    data: Any::Map(Arc::new(
                        [
                            ("id".to_owned(), Any::Number(7.0)),
                            ("kind".to_owned(), Any::from("start")),
                        ]
                        .into(),
                    )),
                }],
                &EditCtx::local("", ""),
            )
            .unwrap();
        for _ in 0..3 {
            for from in &peers {
                for to in &peers {
                    if !std::ptr::eq(from, to) {
                        let update = from.encode_diff_v1(&to.encode_state_vector_v1()).unwrap();
                        to.applying_peer_update(|| to.apply_update_v1(&update).unwrap());
                    }
                }
            }
        }
        let read = |doc: &EditingDoc| {
            let txn = doc.yrs_doc().transact();
            positions(&txn, "body")
                .into_iter()
                .map(|(at, _)| at)
                .collect::<Vec<_>>()
        };
        assert_eq!(read(&peers[0]), [6, 21]);
        for peer in &peers[1..] {
            assert_eq!(read(peer), read(&peers[0]));
        }
    }
}
