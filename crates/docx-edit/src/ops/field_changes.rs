//! Tracked changes a field keeps as markup in its code or result. The
//! revision list leaves them out; Accept All and Reject All resolve them and
//! leave each field as the seed makes it of its export.

use std::collections::{BTreeMap, HashMap};
use std::sync::Arc;

use serde_json::Value;
use yrs::{Any, Map, Out, ReadTxn, TextRef, TransactionMut};

use crate::op::{OpError, OpResult};
use crate::ops::{Chunk, ChunkKind, snapshot};
use crate::raw::{RawOp, apply_raw_ops_to_story};
use crate::seed::{JsonObject, any_from_value, field_units, payload, shown_runs, yrs_attrs};
use crate::{KIND_KEY, map_string};

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

/// Resolves the changes `field` (a field node) and the fields inside it keep,
/// re-reading what each shows as the parser reads its export. Returns each
/// old node's new place in the field's own code and result and whether a
/// field inside it resolved (`None` for a resolved change), or `None` when it
/// kept no change.
fn resolve_field(field: &mut Value, accept: bool) -> Option<[Places; 2]> {
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
            let (places, resolved) = resolve_nodes(nodes, accept);
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
fn resolve_nodes(nodes: &mut Vec<Value>, accept: bool) -> (Places, bool) {
    let mut places = Vec::with_capacity(nodes.len());
    let mut resolved = Vec::with_capacity(nodes.len());
    let mut changed = false;
    for mut node in nodes.drain(..) {
        let parsed = (node["type"] == "rawXml")
            .then(|| node["xml"].as_str())
            .flatten()
            .and_then(|xml| docx_parse::paragraph::resolve_raw_inline(xml, accept));
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
                let nested = resolve_nested(&mut node, accept);
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
fn resolve_nested(value: &mut Value, accept: bool) -> bool {
    let items: Vec<&mut Value> = match value {
        Value::Object(entries)
            if matches!(
                entries.get("type").and_then(Value::as_str),
                Some("simpleField" | "complexField")
            ) =>
        {
            return resolve_field(value, accept).is_some();
        }
        Value::Array(items) => items.iter_mut().collect(),
        Value::Object(entries) => entries.values_mut().collect(),
        _ => return false,
    };
    let mut changed = false;
    for item in items {
        changed |= resolve_nested(item, accept);
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
    // A result spanning paragraphs that keeps no markup saves its runs
    // (`ComplexField::written_result`), so its first paragraph reads back as them.
    let spans_paragraphs = field
        .get("structuredResult")
        .is_some_and(|content| !content["blocks"].is_null());
    let keeps_markup = result
        .iter()
        .any(|node| matches!(node["type"].as_str(), Some("rawXml" | "inlineSdt")));
    if field["type"] == "complexField" && spans_paragraphs && !keeps_markup {
        let runs = field["fieldResult"].clone();
        if let Some(content) = field.get_mut("structuredResult") {
            content["inline"] = runs;
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

/// Resolves the changes the fields in `story` keep, last field first so the
/// offsets of those before stay put. A field embed becomes the units the seed
/// makes of the resolved field: a result hyperlink or simple field the change
/// uncovered is projected before it, and projected children it kept stay as
/// they are (edited ones included), renumbered.
pub(crate) fn resolve_field_changes(
    txn: &mut TransactionMut<'_>,
    story_id: &str,
    story: &TextRef,
    accept: bool,
    styles: Option<&Value>,
) -> OpResult<()> {
    let chunks = snapshot(story, txn);
    for position in (0..chunks.len()).rev() {
        let ChunkKind::Embed(Some(map)) = &chunks[position].kind else {
            continue;
        };
        match map_string(map, txn, KIND_KEY).as_deref() {
            Some("field") if field_result_attr(&chunks[position]).is_none() => {
                resolve_owner(txn, story_id, &chunks, position, accept, styles)?
            }
            Some("sdt") => {
                let style = paragraph_style(txn, &chunks, position);
                if let Some(Out::Any(content)) = map.get(txn, "content")
                    && let Some(content) =
                        resolved_sdt_content(&content, accept, styles, style.as_deref())
                {
                    map.insert(txn, "content", content);
                }
            }
            _ => {}
        }
    }
    Ok(())
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

fn projection_id<T: ReadTxn>(txn: &T, chunk: &Chunk) -> Option<i64> {
    let ChunkKind::Embed(Some(map)) = &chunk.kind else {
        return None;
    };
    match map.get(txn, "resultProjection") {
        Some(Out::Any(value)) => any_value(&value)["id"].as_i64(),
        _ => None,
    }
}

/// The index the seed gives the field at `position` among its paragraph's
/// content once saved, which names its projection (`field_to_units`): a run,
/// link, tracked change or embed per chunk the export writes before it (as
/// `buildParagraphContent` groups them; a projected child joins its field),
/// and a marker per bookmark boundary there.
fn content_index<T: ReadTxn>(
    txn: &T,
    chunks: &[Chunk],
    first: usize,
    position: usize,
    last: usize,
) -> usize {
    let mut count = 0;
    let mut offset = 0;
    let mut link: Option<&Any> = None;
    for chunk in &chunks[first..position] {
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

/// Resolves the changes the field embed at `position` keeps (see `resolve_field_changes`).
fn resolve_owner(
    txn: &mut TransactionMut<'_>,
    story_id: &str,
    chunks: &[Chunk],
    position: usize,
    accept: bool,
    styles: Option<&Value>,
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
    let Some([code_moved, result_moved]) = resolve_field(&mut field, accept) else {
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
    let id = content_index(txn, chunks, first, position, last) as i64;
    let start = chunks[first..position]
        .iter()
        .rposition(|chunk| old_id.is_some() && projection_id(txn, chunk) == old_id)
        .map_or(first, |offset| first + offset + 1);
    let mut spans: BTreeMap<i64, (u32, u32)> = BTreeMap::new();
    for chunk in &chunks[start..position] {
        if let Some((chunk_id, index)) = field_result_attr(chunk)
            && Some(chunk_id) == old_id
        {
            let span = spans.entry(index).or_insert((chunk.start, chunk.end()));
            span.0 = span.0.min(chunk.start);
            span.1 = span.1.max(chunk.end());
        }
    }
    // Where an old child now sits, unless a field inside it resolved too.
    let moved = |index: i64| -> Option<i64> {
        let (at, nested) = if index >= 0 {
            let (at, nested) = result_moved.get(index as usize).copied().flatten()?;
            (at as i64, nested)
        } else {
            let (at, nested) = code_moved.get((-index - 1) as usize).copied().flatten()?;
            (-(at as i64) - 1, nested)
        };
        (!nested).then_some(at)
    };

    // The units the seed makes of the resolved field: children, then the field.
    let style = paragraph_style(txn, chunks, position);
    let old_units = field_units(
        &old,
        styles,
        style.as_deref(),
        old_id.unwrap_or(0) as usize,
        true,
    );
    let mut units = field_units(&field, styles, style.as_deref(), id as usize, true);
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
    // Projected children the field keeps, by new index: they stay as they are.
    let kept: BTreeMap<i64, i64> = spans
        .keys()
        .filter_map(|&old_index| {
            let new_index = moved(old_index)?;
            groups
                .iter()
                .any(|(index, _)| *index == new_index)
                .then_some((new_index, old_index))
        })
        .collect();
    // A kept child keeps its recorded items, so an edit to it still shows on export.
    let old_items: HashMap<i64, Value> = old_projection
        .as_ref()
        .and_then(|value| value["children"].as_array())
        .into_iter()
        .flatten()
        .filter_map(|child| Some((child["index"].as_i64()?, child["items"].clone())))
        .collect();
    if let Some(Value::Array(children)) = owner_payload
        .get_mut("resultProjection")
        .and_then(|value| value.get_mut("children"))
    {
        for child in children {
            let items = child["index"]
                .as_i64()
                .and_then(|index| kept.get(&index))
                .and_then(|old_index| old_items.get(old_index));
            if let Some(items) = items {
                child["items"] = items.clone();
            }
        }
    }

    // Edits at their offsets now, applied last first; at one offset a format
    // goes before the inserts, and a later child before an earlier one.
    let mut edits: Vec<(u32, usize, Vec<RawOp>)> = Vec::new();
    let marks = marks_diff(&old_marks, &owner_marks).map_err(error)?;
    if !marks.is_empty() {
        edits.push((
            chunk.start,
            0,
            vec![RawOp::Format {
                index: chunk.start,
                len: 1,
                attrs: marks,
            }],
        ));
    }
    for (&old_index, &(from, to)) in &spans {
        let op = match kept.iter().find(|(_, kept)| **kept == old_index) {
            Some((&new_index, _)) if new_index == old_index && Some(id) == old_id => continue,
            Some((&new_index, _)) => RawOp::Format {
                index: from,
                len: to - from,
                attrs: yrs_attrs(JsonObject::from([(
                    "fieldResult".to_owned(),
                    serde_json::json!({ "id": id, "index": new_index }),
                )]))
                .map_err(error)?,
            },
            None => RawOp::Delete {
                index: from,
                len: to - from,
            },
        };
        edits.push((from, 0, vec![op]));
    }
    let count = groups.len();
    let anchors: Vec<u32> = (0..count)
        .map(|order| {
            groups[order + 1..]
                .iter()
                .find_map(|(index, _)| kept.get(index).map(|old_index| spans[old_index].0))
                .unwrap_or(chunk.start)
        })
        .collect();
    for (order, (index, group)) in groups.into_iter().enumerate() {
        if kept.contains_key(&index) {
            continue;
        }
        let anchor = anchors[order];
        let mut at = anchor;
        let mut ops = Vec::new();
        for unit in group {
            let (len, op) = unit_op(at, unit)?;
            at += len;
            ops.push(op);
        }
        edits.push((anchor, 1 + count - order, ops));
    }
    edits.sort_by(|left, right| right.0.cmp(&left.0).then(left.1.cmp(&right.1)));
    let ops: Vec<RawOp> = edits.into_iter().flat_map(|(_, _, ops)| ops).collect();
    if !ops.is_empty() {
        apply_raw_ops_to_story(txn, story_id, ops, false)?;
    }

    // The field embed takes the resolved field's payload.
    let keys: Vec<String> = map.keys(txn).map(str::to_owned).collect();
    for key in keys
        .iter()
        .filter(|key| *key != KIND_KEY && !owner_payload.contains_key(*key))
    {
        map.remove(txn, key);
    }
    for (key, value) in owner_payload {
        map.insert(txn, key, any_from_value(value).map_err(error)?);
    }
    Ok(())
}

/// An inline content control's content with its fields' kept changes resolved.
fn resolved_sdt_content(
    content: &Any,
    accept: bool,
    styles: Option<&Value>,
    style: Option<&str>,
) -> Option<Any> {
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
            let kind = match entry.get("kind") {
                Some(Any::String(kind)) => kind.to_string(),
                _ => return item.clone(),
            };
            if kind == "sdt" {
                let Some(inner) = payload
                    .get("content")
                    .and_then(|inner| resolved_sdt_content(inner, accept, styles, style))
                else {
                    return item.clone();
                };
                let mut payload = payload.as_ref().clone();
                payload.insert("content".to_owned(), inner);
                entry.insert("payload".to_owned(), Any::Map(Arc::new(payload)));
            } else {
                let Some(Any::String(data)) = payload.get("fieldData") else {
                    return item.clone();
                };
                let Ok(old) = serde_json::from_str::<Value>(data) else {
                    return item.clone();
                };
                let mut field = old.clone();
                if kind != "field" || resolve_field(&mut field, accept).is_none() {
                    return item.clone();
                }
                let unit = |field: &Value| field_units(field, styles, style, 0, false).pop();
                let (Some((_, old_marks)), Some((Err((_, values)), marks))) =
                    (unit(&old), unit(&field))
                else {
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
                let Ok(values) = any_from_value(Value::Object(values.into_iter().collect())) else {
                    return item.clone();
                };
                entry.insert("payload".to_owned(), values);
                entry.insert("attrs".to_owned(), Any::Map(Arc::new(attrs)));
            }
            changed = true;
            Any::Map(Arc::new(entry))
        })
        .collect();
    changed.then(|| Any::Array(Arc::from(items)))
}
