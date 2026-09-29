//! Tracked changes a field keeps as markup in its code or result. The
//! revision list leaves them out; Accept All and Reject All resolve them and
//! leave each field as the seed makes it of its export.

use std::collections::{BTreeMap, HashMap};
use std::sync::Arc;

use serde_json::Value;
use yrs::types::Attrs;
use yrs::{Any, Map, Out, ReadTxn, Text, TextRef, TransactionMut};

use crate::op::{OpError, OpResult};
use crate::ops::{Chunk, ChunkKind, snapshot};
use crate::seed::{
    JsonObject, PackageContext, any_from_value, field_units, payload, shown_runs, yrs_attrs,
};
use crate::{KIND_KEY, RawOp, map_string};

/// Whether a field's data keeps a tracked change as markup where resolving
/// reaches it: its code and result, and the fields inside them.
pub(crate) fn field_data_keeps_changes(data: &str) -> bool {
    fn field_keeps(field: &Value) -> bool {
        ["structuredCode", "structuredResult"]
            .into_iter()
            .any(|key| {
                field[key]["inline"].as_array().is_some_and(|nodes| {
                    nodes.iter().any(|node| match node["xml"].as_str() {
                        Some(xml) if node["type"] == "rawXml" => {
                            docx_parse::paragraph::raw_holds_changes(xml)
                        }
                        _ => nested_keeps(node),
                    })
                })
            })
    }
    fn nested_keeps(value: &Value) -> bool {
        match value {
            Value::Object(entries)
                if matches!(
                    entries.get("type").and_then(Value::as_str),
                    Some("simpleField" | "complexField")
                ) =>
            {
                field_keeps(value)
            }
            Value::Array(items) => items.iter().any(nested_keeps),
            Value::Object(entries) => entries.values().any(nested_keeps),
            _ => false,
        }
    }
    data.contains("rawXml")
        && serde_json::from_str(data).is_ok_and(|field: Value| field_keeps(&field))
}

/// Whether a field inside an inline content control's content keeps a tracked change.
pub(crate) fn sdt_keeps_changes(content: &Any) -> bool {
    let Any::Array(items) = content else {
        return false;
    };
    items.iter().any(|item| {
        let Any::Map(entry) = item else {
            return false;
        };
        let Some(Any::Map(payload)) = entry.get("payload") else {
            return false;
        };
        match (payload.get("fieldData"), payload.get("content")) {
            (Some(Any::String(data)), _) => field_data_keeps_changes(data),
            (_, Some(inner)) => sdt_keeps_changes(inner),
            _ => false,
        }
    })
}

/// How a resolve goes: accepting or rejecting, reading markup with the
/// source package's theme and the relationships of the story's part.
#[derive(Clone, Copy)]
struct Resolve<'a> {
    accept: bool,
    package: Option<&'a PackageContext>,
    relationships: Option<&'a docx_parse::RelationshipMap>,
}

/// Resolves the changes `field` (a field node) and the fields inside it keep,
/// re-reading what each shows as the parser reads its export. Returns each
/// old node's new place in the field's own code and result and whether a
/// field inside it resolved (`None` for a resolved change), or `None` when it
/// kept no change.
fn resolve_field(field: &mut Value, how: Resolve<'_>) -> Option<[Places; 2]> {
    let mut moved = [Vec::new(), Vec::new()];
    let mut changed = false;
    for (slot, key) in ["structuredCode", "structuredResult"]
        .into_iter()
        .enumerate()
    {
        if let Some(nodes) = field
            .get_mut(key)
            .and_then(|content| content.get_mut("inline"))
            .and_then(Value::as_array_mut)
        {
            let (places, resolved) = resolve_nodes(nodes, how);
            moved[slot] = places;
            changed |= resolved;
        }
    }
    if !changed {
        return None;
    }
    if let Value::Object(entries) = field {
        refresh_field(entries);
    }
    Some(moved)
}

type Places = Vec<Option<(usize, bool)>>;

/// Replaces each kept change among `nodes` by what it resolves to.
fn resolve_nodes(nodes: &mut Vec<Value>, how: Resolve<'_>) -> (Places, bool) {
    let mut places = Vec::with_capacity(nodes.len());
    let mut resolved = Vec::with_capacity(nodes.len());
    let mut changed = false;
    for mut node in nodes.drain(..) {
        let parsed = (node["type"] == "rawXml")
            .then(|| node["xml"].as_str())
            .flatten()
            .and_then(|xml| {
                docx_parse::paragraph::resolve_raw_inline(
                    xml,
                    how.accept,
                    how.relationships,
                    how.package.map(|package| &package.theme),
                )
            });
        match parsed {
            Some(parsed) => {
                changed = true;
                places.push(None);
                resolved.extend(
                    parsed
                        .iter()
                        .filter_map(|node| serde_json::to_value(node).ok()),
                );
            }
            None => {
                let nested = resolve_nested(&mut node, how);
                changed |= nested;
                places.push(Some((resolved.len(), nested)));
                resolved.push(node);
            }
        }
    }
    *nodes = resolved;
    (places, changed)
}

/// Resolves the changes kept by the fields `value` holds.
fn resolve_nested(value: &mut Value, how: Resolve<'_>) -> bool {
    let items: Vec<&mut Value> = match value {
        Value::Object(entries)
            if matches!(
                entries.get("type").and_then(Value::as_str),
                Some("simpleField" | "complexField")
            ) =>
        {
            return resolve_field(value, how).is_some();
        }
        Value::Array(items) => items.iter_mut().collect(),
        Value::Object(entries) => entries.values_mut().collect(),
        _ => return false,
    };
    let mut changed = false;
    for item in items {
        changed |= resolve_nested(item, how);
    }
    changed
}

/// Re-reads what a field shows from its structured code and result, as
/// docx-parse's `OpenComplexField` and simple field parse do: the runs its
/// result shows, its code runs and instruction, and the formatting its first
/// formatted run lends a field that had none.
fn refresh_field(field: &mut serde_json::Map<String, Value>) {
    let nodes = |field: &serde_json::Map<String, Value>, key: &str| -> Vec<Value> {
        field
            .get(key)
            .and_then(|content| content["inline"].as_array())
            .cloned()
            .unwrap_or_default()
    };
    let result = nodes(field, "structuredResult");
    let code = nodes(field, "structuredCode");
    if field["type"] == "simpleField" {
        field.insert("content".to_owned(), Value::Array(shown_runs(&result)));
    } else {
        field.insert("fieldResult".to_owned(), Value::Array(shown_runs(&result)));
        // A nested field in code is code of its own, not this field's.
        let code_runs = shown_runs(
            &code
                .iter()
                .filter(|node| node["type"] != "complexField")
                .cloned()
                .collect::<Vec<_>>(),
        );
        let instruction: String = code_runs
            .iter()
            .flat_map(|run| run["content"].as_array().into_iter().flatten())
            .filter(|content| content["type"] == "instrText")
            .filter_map(|content| content["text"].as_str())
            .collect();
        let instruction = instruction.trim();
        field.insert(
            "fieldType".to_owned(),
            Value::String(docx_parse::inline::parse_field_type(instruction)),
        );
        field.insert(
            "instruction".to_owned(),
            Value::String(instruction.to_owned()),
        );
        field.insert("fieldCode".to_owned(), Value::Array(code_runs));
        if field.get("formatting").is_none_or(Value::is_null)
            && let Some(formatting) = code
                .iter()
                .chain(&result)
                .filter(|node| node["type"] == "run")
                .find_map(|run| run.get("formatting").filter(|value| !value.is_null()))
        {
            field.insert("formatting".to_owned(), formatting.clone());
        }
    }
    if let Some(Value::Object(mut tree)) = field.get("fieldTree").cloned() {
        for (content, key) in [("structuredCode", "code"), ("structuredResult", "result")] {
            if let Some(value) = field.get(content) {
                tree.insert(key.to_owned(), value.clone());
            }
        }
        if field["type"] == "complexField" {
            let children: Vec<Value> = code
                .iter()
                .chain(&result)
                .filter(|node| node["type"] == "complexField")
                .filter_map(|node| node.get("fieldTree").cloned())
                .collect();
            if children.is_empty() {
                tree.remove("children");
            } else {
                tree.insert("children".to_owned(), Value::Array(children));
            }
        }
        field.insert("fieldTree".to_owned(), Value::Object(tree));
    }
}

fn any_value(value: &Any) -> Value {
    let mut json = String::new();
    value.to_json(&mut json);
    serde_json::from_str(&json).unwrap_or(Value::Null)
}

fn field_result_attr(chunk: &Chunk) -> Option<(i64, i64)> {
    let Any::Map(marker) = chunk.attrs.get("fieldResult")? else {
        return None;
    };
    let number = |key: &str| match marker.get(key) {
        Some(Any::Number(value)) => Some(*value as i64),
        Some(Any::BigInt(value)) => Some(*value),
        _ => None,
    };
    Some((number("id")?, number("index")?))
}

/// Clears the `fieldResult` markers of the children of each projecting field
/// embed among `chunks` in `start..end` of `story`, which a delete is about
/// to remove: the children right before it that carry its number. A child
/// whose field is deleted is plain content written in place, which no other
/// field pairs, even one that comes to share its number after a join or
/// Accept All.
pub(crate) fn release_children(
    txn: &mut TransactionMut<'_>,
    story: &TextRef,
    chunks: &[Chunk],
    start: u32,
    end: u32,
) {
    let owners: Vec<(u32, i64)> = chunks
        .iter()
        .filter(|chunk| (start..end).contains(&chunk.start) && field_result_attr(chunk).is_none())
        .filter_map(|chunk| {
            let ChunkKind::Embed(Some(map)) = &chunk.kind else {
                return None;
            };
            match map.get(txn, "resultProjection") {
                Some(Out::Any(value)) => Some((chunk.start, any_value(&value)["id"].as_i64()?)),
                _ => None,
            }
        })
        .collect();
    if owners.is_empty() {
        return;
    }
    let chunks = snapshot(story, txn);
    for (at, id) in owners {
        let Some(position) = chunks.iter().position(|chunk| chunk.start == at) else {
            continue;
        };
        let children = chunks[..position]
            .iter()
            .rev()
            .take_while(|chunk| field_result_attr(chunk).is_some_and(|(child, _)| child == id))
            .count();
        if children > 0 {
            let from = chunks[position - children].start;
            story.format(
                txn,
                from,
                at - from,
                Attrs::from([(Arc::from("fieldResult"), Any::Null)]),
            );
        }
    }
}

/// Formatting attributes: those a field's marks set, less what its owner
/// keeps besides (a hyperlink or tracked-change stamp stays).
fn marks_diff(old: &JsonObject, new: &JsonObject) -> Result<yrs::types::Attrs, String> {
    let mut diff = JsonObject::new();
    for key in old.keys().filter(|key| !new.contains_key(*key)) {
        diff.insert(key.clone(), Value::Null);
    }
    diff.extend(new.iter().map(|(key, value)| (key.clone(), value.clone())));
    yrs_attrs(diff)
}

/// The style of the paragraph the chunk at `position` sits in.
fn paragraph_style<T: ReadTxn>(txn: &T, chunks: &[Chunk], position: usize) -> Option<String> {
    chunks[position..]
        .iter()
        .find_map(|chunk| match &chunk.kind {
            ChunkKind::Pilcrow(map) => Some(map_string(map, txn, "pStyle")),
            _ => None,
        })?
}

type Units = Vec<(Result<String, (String, JsonObject)>, JsonObject)>;

fn unit_index(attrs: &JsonObject) -> Option<i64> {
    attrs.get("fieldResult")?["index"].as_i64()
}

fn unit_op(
    index: u32,
    unit: (Result<String, (String, JsonObject)>, JsonObject),
) -> OpResult<(u32, RawOp)> {
    let error = |error: String| OpError::InvalidUpdate(error);
    let (content, attrs) = unit;
    let attrs = yrs_attrs(attrs).map_err(error)?;
    Ok(match content {
        Ok(text) => {
            let len = text.encode_utf16().count() as u32;
            (len, RawOp::Insert { index, text, attrs })
        }
        Err((kind, values)) => (
            1,
            RawOp::InsertEmbed {
                index,
                kind,
                payload: payload(values).map_err(error)?,
                attrs,
            },
        ),
    })
}

/// The index the seed gives the field at `position` among its paragraph's
/// content once saved, which names its projection (`field_to_units`): a run,
/// link, tracked change or embed per chunk the export writes before it (as
/// `buildParagraphContent` groups them; a projected child joins its field),
/// and a marker per bookmark or comment boundary there (`comments`, the
/// story's comment boundaries): a comment marker splits the run it falls in,
/// and the export writes a reference after the end of a comment with none.
fn content_index<T: ReadTxn>(
    txn: &T,
    chunks: &[Chunk],
    first: usize,
    position: usize,
    last: usize,
    comments: &[(u32, bool)],
) -> usize {
    let paragraph = chunks[first].start..=chunks[position].start;
    let comments: usize = comments
        .iter()
        .filter(|(at, _)| paragraph.contains(at))
        .map(|&(at, adds_reference)| {
            let splits = chunks[first..position].iter().any(|chunk| {
                matches!(chunk.kind, ChunkKind::Text(_))
                    && field_result_attr(chunk).is_none()
                    && !chunk.attr_active("hyperlink")
                    && chunk.start < at
                    && at < chunk.end()
            });
            1 + usize::from(adds_reference) + usize::from(splits)
        })
        .sum();
    let mut count = 0;
    let mut offset = 0;
    let mut link: Option<&Any> = None;
    // The blocks opening the paragraph's slot are no content of it, save the
    // breaks the export writes as its first runs: from a leading one on, or
    // all of them with no paragraph before (as `settle` in yrsToDocument.ts).
    let mut head = first;
    let mut after_paragraph = first > 0;
    let mut leads = false;
    while head < position {
        let ChunkKind::Embed(Some(map)) = &chunks[head].kind else {
            break;
        };
        match map_string(map, txn, KIND_KEY).as_deref() {
            Some("table" | "blockSdt") => {
                after_paragraph = false;
                leads = false;
            }
            Some("pageBreak" | "columnBreak") => {
                leads |= !after_paragraph
                    || matches!(map.get(txn, "leading"), Some(Out::Any(Any::Bool(true))));
                if leads {
                    count += 1;
                }
            }
            _ => break,
        }
        head += 1;
    }
    for chunk in &chunks[head..position] {
        offset += match &chunk.kind {
            ChunkKind::Embed(Some(map))
                if map_string(map, txn, KIND_KEY).as_deref() == Some("sdt") =>
            {
                2
            }
            _ => chunk.len,
        };
        if field_result_attr(chunk).is_some() {
            continue;
        }
        let href = chunk
            .attrs
            .get("hyperlink")
            .filter(|value| **value != Any::Null);
        let tracked = chunk.attr_active("ins") || chunk.attr_active("del");
        if tracked || href.is_none() || href != link {
            count += 1;
        }
        link = if tracked { None } else { href };
    }
    let bookmarks = match chunks.get(last).map(|chunk| &chunk.kind) {
        Some(ChunkKind::Pilcrow(map)) => match map.get(txn, "bookmarks") {
            Some(Out::Any(value)) => any_value(&value),
            _ => Value::Null,
        },
        _ => Value::Null,
    };
    count
        + comments
        + bookmarks
            .as_array()
            .into_iter()
            .flatten()
            .filter(|bookmark| {
                bookmark["offset"]
                    .as_u64()
                    .is_some_and(|at| at <= u64::from(offset))
                    && bookmark["kind"] != Value::Null
            })
            .count()
}

/// Resolves the changes the fields in `story` keep, last field first so the
/// offsets of those before stay put. Only the changes resolve: each field
/// embed becomes what the seed makes of the resolved field, where a result
/// hyperlink or simple field a change uncovered is projected before it, while
/// the projected children the user edited keep their edits, those the user
/// deleted stay deleted, and a field inside a kept child resolves in place.
pub(crate) fn resolve_field_changes(
    txn: &mut TransactionMut<'_>,
    story: &TextRef,
    story_id: &str,
    accept: bool,
    package: Option<&PackageContext>,
    comments: &[(u32, bool)],
) -> OpResult<()> {
    let how = Resolve {
        accept,
        package,
        relationships: package.map(|package| package.relationships(story_id)),
    };
    let chunks = snapshot(story, txn);
    for position in (0..chunks.len()).rev() {
        let ChunkKind::Embed(Some(map)) = &chunks[position].kind else {
            continue;
        };
        match map_string(map, txn, KIND_KEY).as_deref() {
            Some("field") if field_result_attr(&chunks[position]).is_none() => {
                resolve_owner(txn, story, &chunks, position, how, comments)?
            }
            Some("sdt") => {
                let style = paragraph_style(txn, &chunks, position);
                if let Some(Out::Any(content)) = map.get(txn, "content")
                    && let Some(content) = resolved_sdt_content(&content, how, style.as_deref())
                {
                    map.insert(txn, "content", content);
                }
            }
            _ => {}
        }
    }
    Ok(())
}

/// A field embed's payload and marks once its kept changes resolve, as an
/// embed that is not projected (a projected child, or one in a content
/// control); `None` when it kept none.
fn resolved_embed(
    data: &str,
    how: Resolve<'_>,
    style: Option<&str>,
) -> Option<(JsonObject, JsonObject, JsonObject)> {
    let old: Value = serde_json::from_str(data).ok()?;
    let mut field = old.clone();
    resolve_field(&mut field, how)?;
    let unit = |field: &Value| field_units(field, how.package, style, 0, false).pop();
    let (Some((_, old_marks)), Some((Err((_, payload)), marks))) = (unit(&old), unit(&field))
    else {
        return None;
    };
    Some((payload, old_marks, marks))
}

/// Applies `edits` (offset, order, op) last first; at one offset the lower
/// order goes first.
fn apply(txn: &mut TransactionMut<'_>, story: &TextRef, mut edits: Vec<(u32, usize, RawOp)>) {
    edits.sort_by(|left, right| right.0.cmp(&left.0).then(left.1.cmp(&right.1)));
    for (_, _, op) in edits {
        match op {
            RawOp::Insert { index, text, attrs } => {
                story.insert_with_attributes(txn, index, &text, attrs)
            }
            RawOp::InsertEmbed {
                index,
                kind,
                payload,
                attrs,
            } => {
                let embed = story.insert_embed_with_attributes(
                    txn,
                    index,
                    yrs::MapPrelim::default(),
                    attrs,
                );
                embed.insert(txn, KIND_KEY, kind.as_str());
                for (key, value) in payload {
                    embed.insert(txn, key, value);
                }
            }
            RawOp::Format { index, len, attrs } => story.format(txn, index, len, attrs),
            RawOp::Delete { index, len } => story.remove_range(txn, index, len),
            _ => {}
        }
    }
}

/// The embed's payload entries replaced by `payload`, keeping its kind.
fn set_payload(
    txn: &mut TransactionMut<'_>,
    map: &yrs::MapRef,
    payload: JsonObject,
) -> OpResult<()> {
    let keys: Vec<String> = map.keys(txn).map(str::to_owned).collect();
    for key in keys
        .iter()
        .filter(|key| *key != KIND_KEY && !payload.contains_key(*key))
    {
        map.remove(txn, key);
    }
    for (key, value) in payload {
        map.insert(
            txn,
            key,
            any_from_value(value).map_err(OpError::InvalidUpdate)?,
        );
    }
    Ok(())
}

/// Resolves the changes the field embed at `position` keeps (see `resolve_field_changes`).
fn resolve_owner(
    txn: &mut TransactionMut<'_>,
    story: &TextRef,
    chunks: &[Chunk],
    position: usize,
    how: Resolve<'_>,
    comments: &[(u32, bool)],
) -> OpResult<()> {
    let error = OpError::InvalidUpdate;
    let chunk = &chunks[position];
    let ChunkKind::Embed(Some(map)) = &chunk.kind else {
        return Ok(());
    };
    let Some(Ok(old)) =
        map_string(map, txn, "fieldData").map(|data| serde_json::from_str::<Value>(&data))
    else {
        return Ok(());
    };
    let mut field = old.clone();
    let Some([code_moved, result_moved]) = resolve_field(&mut field, how) else {
        return Ok(());
    };

    // The paragraph's chunks, and the children this field projected.
    let first = chunks[..position]
        .iter()
        .rposition(|chunk| matches!(chunk.kind, ChunkKind::Pilcrow(_)))
        .map_or(0, |index| index + 1);
    let last = chunks[position..]
        .iter()
        .position(|chunk| matches!(chunk.kind, ChunkKind::Pilcrow(_)))
        .map_or(chunks.len(), |offset| position + offset);
    let old_projection = match map.get(txn, "resultProjection") {
        Some(Out::Any(value)) => Some(any_value(&value)),
        _ => None,
    };
    let old_id = old_projection
        .as_ref()
        .and_then(|value| value["id"].as_i64());
    let id = content_index(txn, chunks, first, position, last, comments) as i64;
    let mut old_items: HashMap<i64, Value> = old_projection
        .as_ref()
        .and_then(|value| value["children"].as_array())
        .into_iter()
        .flatten()
        .filter_map(|child| Some((child["index"].as_i64()?, child["items"].clone())))
        .collect();
    // Its children: the chunks right before it that carry its number at an
    // index it records, as the export pairs them (`restoreProjectedFieldResults`).
    let start = position
        - chunks[first..position]
            .iter()
            .rev()
            .take_while(|chunk| {
                field_result_attr(chunk).is_some_and(|(chunk_id, index)| {
                    Some(chunk_id) == old_id && old_items.contains_key(&index)
                })
            })
            .count();
    let children: Vec<&Chunk> = chunks[start..position].iter().collect();
    let mut spans: BTreeMap<i64, (u32, u32)> = BTreeMap::new();
    for chunk in &children {
        let (_, index) = field_result_attr(chunk).unwrap_or_default();
        let span = spans.entry(index).or_insert((chunk.start, chunk.end()));
        span.0 = span.0.min(chunk.start);
        span.1 = span.1.max(chunk.end());
    }
    // Where an old child now sits, and whether a field inside it resolved.
    let moved = |index: i64| -> Option<(i64, bool)> {
        if index >= 0 {
            let (at, nested) = result_moved.get(index as usize).copied().flatten()?;
            Some((at as i64, nested))
        } else {
            let (at, nested) = code_moved.get((-index - 1) as usize).copied().flatten()?;
            Some((-(at as i64) - 1, nested))
        }
    };

    // The units the seed makes of the resolved field: children, then the field.
    let style = paragraph_style(txn, chunks, position);
    let old_units = field_units(
        &old,
        how.package,
        style.as_deref(),
        old_id.unwrap_or(0) as usize,
        true,
    );
    let mut units = field_units(&field, how.package, style.as_deref(), id as usize, true);
    let (Some((_, old_marks)), Some((Err((_, mut owner_payload)), owner_marks))) =
        (old_units.last().cloned(), units.pop())
    else {
        return Ok(());
    };
    let mut groups: Vec<(i64, Units)> = Vec::new();
    for unit in units {
        let index = unit_index(&unit.1).unwrap_or_default();
        match groups.last_mut() {
            Some((last, group)) if *last == index => group.push(unit),
            _ => groups.push((index, vec![unit])),
        }
    }
    // Old children by new index. One with a span stays as it is; one the
    // user deleted keeps its recorded items, so the export keeps it deleted.
    let from_old: BTreeMap<i64, (i64, bool)> = old_items
        .keys()
        .chain(spans.keys())
        .filter_map(|&old_index| {
            let (new_index, nested) = moved(old_index)?;
            groups
                .iter()
                .any(|(index, _)| *index == new_index)
                .then_some((new_index, (old_index, nested)))
        })
        .collect();

    let mut edits: Vec<(u32, usize, RawOp)> = Vec::new();
    let marks = marks_diff(&old_marks, &owner_marks).map_err(error)?;
    if !marks.is_empty() {
        edits.push((
            chunk.start,
            0,
            RawOp::Format {
                index: chunk.start,
                len: 1,
                attrs: marks,
            },
        ));
    }
    for (&old_index, &(from, to)) in &spans {
        let Some((&new_index, &(_, nested))) =
            from_old.iter().find(|(_, (kept, _))| *kept == old_index)
        else {
            edits.push((
                from,
                0,
                RawOp::Delete {
                    index: from,
                    len: to - from,
                },
            ));
            continue;
        };
        if new_index != old_index || Some(id) != old_id {
            let attrs = yrs_attrs(JsonObject::from([(
                "fieldResult".to_owned(),
                serde_json::json!({ "id": id, "index": new_index }),
            )]))
            .map_err(error)?;
            edits.push((
                from,
                0,
                RawOp::Format {
                    index: from,
                    len: to - from,
                    attrs,
                },
            ));
        }
        if !nested {
            continue;
        }
        // A field inside the kept child resolves in place, in its recorded items too.
        for inner in children
            .iter()
            .filter(|chunk| field_result_attr(chunk).map(|(_, index)| index) == Some(old_index))
        {
            let ChunkKind::Embed(Some(inner_map)) = &inner.kind else {
                continue;
            };
            let Some(data) = map_string(inner_map, txn, "fieldData") else {
                continue;
            };
            let Some((payload, old_marks, marks)) = resolved_embed(&data, how, style.as_deref())
            else {
                continue;
            };
            if let Some(Value::Array(items)) = old_items.get_mut(&old_index) {
                for item in items
                    .iter_mut()
                    .filter(|item| item["payload"]["fieldData"] == data.as_str())
                {
                    item["payload"] = Value::Object(payload.clone().into_iter().collect());
                    if let Some(Value::Object(attributes)) = item.get_mut("attributes") {
                        for key in old_marks.keys() {
                            attributes.remove(key);
                        }
                        attributes.extend(marks.clone());
                    }
                }
            }
            let diff = marks_diff(&old_marks, &marks).map_err(error)?;
            if !diff.is_empty() {
                edits.push((
                    inner.start,
                    0,
                    RawOp::Format {
                        index: inner.start,
                        len: 1,
                        attrs: diff,
                    },
                ));
            }
            set_payload(txn, inner_map, payload)?;
        }
    }
    if let Some(Value::Array(projected)) = owner_payload
        .get_mut("resultProjection")
        .and_then(|value| value.get_mut("children"))
    {
        for child in projected {
            let items = child["index"]
                .as_i64()
                .and_then(|index| from_old.get(&index))
                .and_then(|(old_index, _)| old_items.get(old_index));
            if let Some(items) = items {
                child["items"] = items.clone();
            }
        }
    }
    let count = groups.len();
    let anchors: Vec<u32> = (0..count)
        .map(|order| {
            groups[order + 1..]
                .iter()
                .find_map(|(index, _)| {
                    let (old_index, _) = from_old.get(index)?;
                    spans.get(old_index).map(|span| span.0)
                })
                .unwrap_or(chunk.start)
        })
        .collect();
    for (order, (index, group)) in groups.into_iter().enumerate() {
        // Only what a change uncovered is new; the user's children stay as they are.
        if from_old.contains_key(&index) {
            continue;
        }
        let mut at = anchors[order];
        for unit in group {
            let (len, op) = unit_op(at, unit)?;
            edits.push((anchors[order], 1 + count - order, op));
            at += len;
        }
    }
    apply(txn, story, edits);
    set_payload(txn, map, owner_payload)
}

/// An inline content control's content with its fields' kept changes resolved.
fn resolved_sdt_content(content: &Any, how: Resolve<'_>, style: Option<&str>) -> Option<Any> {
    let Any::Array(items) = content else {
        return None;
    };
    let mut changed = false;
    let items: Vec<Any> = items
        .iter()
        .map(|item| {
            let Any::Map(entry) = item else {
                return item.clone();
            };
            let Some(Any::Map(payload)) = entry.get("payload") else {
                return item.clone();
            };
            let mut entry = entry.as_ref().clone();
            match entry.get("kind") {
                Some(Any::String(kind)) if kind.as_ref() == "sdt" => {
                    let Some(inner) = payload
                        .get("content")
                        .and_then(|inner| resolved_sdt_content(inner, how, style))
                    else {
                        return item.clone();
                    };
                    let mut payload = payload.as_ref().clone();
                    payload.insert("content".to_owned(), inner);
                    entry.insert("payload".to_owned(), Any::Map(Arc::new(payload)));
                }
                Some(Any::String(kind)) if kind.as_ref() == "field" => {
                    let Some((values, old_marks, marks)) = (match payload.get("fieldData") {
                        Some(Any::String(data)) => resolved_embed(data, how, style),
                        _ => None,
                    }) else {
                        return item.clone();
                    };
                    let mut attrs = match entry.get("attrs") {
                        Some(Any::Map(attrs)) => attrs.as_ref().clone(),
                        _ => HashMap::new(),
                    };
                    for key in old_marks.keys() {
                        attrs.remove(key);
                    }
                    for (key, value) in marks {
                        if let Ok(value) = any_from_value(value) {
                            attrs.insert(key, value);
                        }
                    }
                    let Ok(values) = any_from_value(Value::Object(values.into_iter().collect()))
                    else {
                        return item.clone();
                    };
                    entry.insert("payload".to_owned(), values);
                    entry.insert("attrs".to_owned(), Any::Map(Arc::new(attrs)));
                }
                _ => return item.clone(),
            }
            changed = true;
            Any::Map(Arc::new(entry))
        })
        .collect();
    changed.then(|| Any::Array(Arc::from(items)))
}
