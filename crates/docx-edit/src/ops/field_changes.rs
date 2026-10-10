//! Tracked changes a field keeps as markup in its code or result. The
//! revision list leaves them out; Accept All and Reject All resolve them and
//! leave each field as the seed makes it of its export.

use std::collections::{BTreeMap, HashMap};
use std::sync::Arc;

use serde_json::Value;
use yrs::types::Attrs;
use yrs::{Any, Map, Out, ReadTxn, Text, TextRef, TransactionMut};

use crate::format::{FIELD_RESULT, HYPERLINK};
use crate::op::{OpError, OpResult};
use crate::ops::{Chunk, ChunkKind, snapshot};
use crate::seed::{
    JsonObject, PackageContext, any_from_value, field_units, own_shown_runs, payload, run_units,
    shown_runs, yrs_attrs,
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
        if value["type"] == "run"
            && value["propertyChanges"]
                .as_array()
                .is_some_and(|changes| !changes.is_empty())
        {
            return true;
        }
        if matches!(
            value["type"].as_str(),
            Some("insertion" | "deletion" | "moveFrom" | "moveTo")
        ) {
            return true;
        }
        if value["type"] == "rawXml" {
            return value["xml"]
                .as_str()
                .is_some_and(docx_parse::paragraph::raw_holds_changes);
        }
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
    serde_json::from_str(data).is_ok_and(|field: Value| field_keeps(&field))
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
        if let Some(Any::Map(attrs)) = entry.get("attrs")
            && ["ins", "del", "rPrChange"]
                .iter()
                .any(|key| attrs.get(*key).is_some_and(|value| *value != Any::Null))
        {
            return true;
        }
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
        if let Some(kind @ ("insertion" | "deletion" | "moveFrom" | "moveTo")) =
            node["type"].as_str()
        {
            changed = true;
            places.push(None);
            if matches!(kind, "deletion" | "moveFrom") != how.accept {
                let mut content = node["content"].as_array().cloned().unwrap_or_default();
                resolve_nodes(&mut content, how);
                resolved.extend(content);
            }
            continue;
        }
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
    if value["type"] == "run" {
        let Some(changes) = value
            .as_object_mut()
            .and_then(|run| run.remove("propertyChanges"))
        else {
            return false;
        };
        if !how.accept {
            for change in changes.as_array().into_iter().flatten() {
                value["formatting"] = change
                    .get("previousFormatting")
                    .cloned()
                    .unwrap_or(Value::Null);
            }
        }
        return true;
    }
    let items: Vec<&mut Value> = match value {
        Value::Object(entries)
            if matches!(
                entries.get("type").and_then(Value::as_str),
                Some("simpleField" | "complexField")
            ) =>
        {
            return resolve_field(value, how).is_some();
        }
        Value::Array(items) => return resolve_nodes(items, how).1,
        Value::Object(entries) => entries.values_mut().collect(),
        _ => return false,
    };
    let mut changed = false;
    for item in items {
        changed |= resolve_nested(item, how);
    }
    if changed
        && value["type"] == "hyperlink"
        && let Some(nodes) = value["structuredChildren"].as_array()
    {
        let children = nodes
            .iter()
            .flat_map(|node| {
                if matches!(node["type"].as_str(), Some("bookmarkStart" | "bookmarkEnd")) {
                    vec![node.clone()]
                } else {
                    shown_runs(std::slice::from_ref(node))
                }
            })
            .collect();
        value["children"] = Value::Array(children);
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
    let Any::Map(marker) = chunk.attrs.get(FIELD_RESULT)? else {
        return None;
    };
    let number = |key: &str| match marker.get(key) {
        Some(Any::Number(value)) => Some(*value as i64),
        Some(Any::BigInt(value)) => Some(*value),
        _ => None,
    };
    Some((number("id")?, number("index")?))
}

/// A projecting field embed's number and the child indices it records.
pub(crate) fn projection<T: ReadTxn>(txn: &T, chunk: &Chunk) -> Option<(i64, Vec<i64>)> {
    if field_result_attr(chunk).is_some() {
        return None;
    }
    let ChunkKind::Embed(Some(map)) = &chunk.kind else {
        return None;
    };
    let Some(Out::Any(value)) = map.get(txn, "resultProjection") else {
        return None;
    };
    let value = any_value(&value);
    let indices = value["children"]
        .as_array()
        .into_iter()
        .flatten()
        .filter_map(|child| child["index"].as_i64())
        .collect();
    Some((value["id"].as_i64()?, indices))
}

/// Whether `start..end` removes the field whose projected child
/// `chunks[child]` is: the embed its run of children ends at. A unit that is
/// no child, or whose field lies past `chunks`, has none there.
pub(crate) fn removes_owner<T: ReadTxn>(
    txn: &T,
    chunks: &[Chunk],
    child: usize,
    start: u32,
    end: u32,
) -> bool {
    let Some((id, _)) = field_result_attr(&chunks[child]) else {
        return false;
    };
    chunks[child..]
        .iter()
        .find(|chunk| field_result_attr(chunk).is_none_or(|(child_id, _)| child_id != id))
        .filter(|chunk| projection(txn, chunk).is_some_and(|(owner, _)| owner == id))
        .is_some_and(|owner| (start..end).contains(&owner.start))
}

/// Adds `by` to the numbers of the projecting fields in `start..end` and of
/// their children: the seed numbers a field by its place in its saved
/// paragraph, so markers written ahead of it move it on.
pub(crate) fn renumber_fields(
    txn: &mut TransactionMut<'_>,
    story: &TextRef,
    start: u32,
    end: u32,
    by: i64,
) {
    if by == 0 {
        return;
    }
    for chunk in crate::ops::snapshot_range(story, txn, start, end) {
        if let Some((id, index)) = field_result_attr(&chunk) {
            let marker = Any::from(HashMap::from([
                ("id".to_owned(), Any::Number((id + by) as f64)),
                ("index".to_owned(), Any::Number(index as f64)),
            ]));
            story.format(
                txn,
                chunk.start,
                chunk.len,
                Attrs::from([(Arc::from(FIELD_RESULT), marker)]),
            );
        } else if projection(txn, &chunk).is_some()
            && let ChunkKind::Embed(Some(map)) = &chunk.kind
            && let Some(Out::Any(Any::Map(value))) = map.get(txn, "resultProjection")
            && let Some(Any::Number(id)) = value.get("id")
        {
            let mut value = (*value).clone();
            value.insert("id".to_owned(), Any::from(*id + by as f64));
            map.insert(txn, "resultProjection", Any::Map(Arc::new(value)));
        }
    }
}

/// Enter at `at` inside a field result's projected children: the field's
/// begin stays after the children before `at`, those after it become plain
/// content, and the field ends where its embed stood, across the paragraphs.
/// A field holding a kept tracked insertion in its result stays as it is.
pub(crate) fn split_field(
    txn: &mut TransactionMut<'_>,
    story: &TextRef,
    story_id: &str,
    at: u32,
    continuation: &str,
    package: Option<&PackageContext>,
) -> OpResult<()> {
    let chunks = snapshot(story, txn);
    let Some(before) = chunks.iter().position(|chunk| chunk.end() == at) else {
        return Ok(());
    };
    let Some((id, split)) = field_result_attr(&chunks[before]).filter(|(_, index)| *index >= 0)
    else {
        return Ok(());
    };
    let owner = chunks
        .get(before + 2..)
        .unwrap_or_default()
        .iter()
        .find(|chunk| field_result_attr(chunk).is_none_or(|(child, _)| child != id));
    let Some(owner) = owner.filter(|chunk| projection(txn, chunk).is_some_and(|(of, _)| of == id))
    else {
        return Ok(());
    };
    let ChunkKind::Embed(Some(map)) = &owner.kind else {
        return Ok(());
    };
    let Some(Ok(mut data)) =
        map_string(map, txn, "fieldData").map(|data| serde_json::from_str::<Value>(&data))
    else {
        return Ok(());
    };
    fn holds_insertion(value: &Value) -> bool {
        match value {
            Value::Object(entries) => {
                matches!(
                    entries.get("type").and_then(Value::as_str),
                    Some("insertion" | "moveTo")
                ) || entries
                    .get("xml")
                    .and_then(Value::as_str)
                    .is_some_and(|xml| {
                        xml.trim_start().starts_with("<w:ins")
                            || xml.trim_start().starts_with("<w:moveTo")
                    })
                    || entries.values().any(holds_insertion)
            }
            Value::Array(items) => items.iter().any(holds_insertion),
            _ => false,
        }
    }
    if holds_insertion(&data["structuredResult"]) {
        // Enter stays as it was; a field left with only result runs for
        // children shows them itself, as its seed would.
        let children = &chunks[before + 2
            ..chunks
                .iter()
                .position(|chunk| chunk.start == owner.start)
                .expect("owner chunk")];
        if children.is_empty()
            || children.iter().any(|chunk| {
                !matches!(chunk.kind, ChunkKind::Text(_)) || chunk.attr_active(HYPERLINK)
            })
        {
            return Ok(());
        }
        let (_, recorded) = projection(txn, owner).expect("found as a projecting field");
        let folded: Vec<i64> = children
            .iter()
            .filter_map(|chunk| field_result_attr(chunk).map(|(_, index)| index))
            .collect();
        let inline = data["structuredResult"]["inline"]
            .as_array()
            .cloned()
            .unwrap_or_default();
        let shown = own_shown_runs(
            &inline
                .iter()
                .enumerate()
                .filter(|(index, _)| {
                    let index = *index as i64;
                    !recorded.contains(&index) || folded.contains(&index)
                })
                .map(|(_, node)| node.clone())
                .collect::<Vec<_>>(),
        );
        map.insert(txn, "displayText", display_text(&shown));
        if let Some(Out::Any(Any::Map(projection))) = map.get(txn, "resultProjection") {
            let mut projection = (*projection).clone();
            if let Some(Any::Array(entries)) = projection.get("children") {
                let kept: Vec<Any> = entries
                    .iter()
                    .filter(|entry| match entry {
                        Any::Map(entry) => !matches!(entry.get("index"), Some(Any::Number(index)) if folded.contains(&(*index as i64))),
                        _ => true,
                    })
                    .cloned()
                    .collect();
                projection.insert("children".to_owned(), Any::Array(Arc::from(kept)));
            }
            map.insert(txn, "resultProjection", Any::Map(Arc::new(projection)));
        }
        let from = children[0].start;
        story.remove_range(txn, from, owner.start - from);
        return Ok(());
    }
    // The result after the split goes to the second paragraph: its children as
    // plain content there, its own runs as text among them. Other content the
    // field shows after the split keeps the field whole.
    let (_, recorded) = projection(txn, owner).expect("found as a projecting field");
    let inline = data["structuredResult"]["inline"]
        .as_array()
        .cloned()
        .unwrap_or_default();
    let mut runs = Vec::new();
    for (index, node) in inline.iter().enumerate().skip(split as usize + 1) {
        if recorded.contains(&(index as i64)) {
            continue;
        }
        if let Some(deletion) = tracked_deletion(node) {
            // Moves with the text around it (decided 2026-10-05).
            runs.push((index as i64, deletion));
            continue;
        }
        let raw = node["xml"].as_str().map(str::trim_start);
        match node["type"].as_str() {
            Some("run") => runs.push((index as i64, node.clone())),
            // A nested field that projects nothing moves as its own embed.
            Some("complexField") if field_units(node, None, None, 0, true).len() == 1 => {
                runs.push((index as i64, node.clone()))
            }
            // A deletion holding more than text keeps the old Enter, which
            // keeps it in its text position (review round 2).
            Some("deletion") => return Ok(()),
            Some("rawXml")
                if raw.is_some_and(|xml| {
                    xml.starts_with("<w:del ") || xml.starts_with("<w:del>")
                }) =>
            {
                return Ok(());
            }
            // Shown nowhere, they stay with the field.
            Some("moveFrom" | "bookmarkStart" | "bookmarkEnd") => {}
            Some("rawXml")
                if raw.is_some_and(|xml| {
                    xml.starts_with("<w:moveFrom") || xml.starts_with("<w:bookmark")
                }) => {}
            // A content control or foreign markup keeps the old Enter.
            _ => return Ok(()),
        }
    }
    // Written only to the new embed, so Undo restores the old one as it was.
    let mut overrides: Vec<(String, Any)> = Vec::new();
    if !runs.is_empty() {
        // The moved runs read as children the field lost, so its save leaves
        // them out of the result; it shows what stays.
        let shown = own_shown_runs(
            &inline
                .iter()
                .enumerate()
                .filter(|(index, _)| {
                    !recorded.contains(&(*index as i64))
                        && !runs.iter().any(|(moved, _)| *moved == *index as i64)
                })
                .map(|(_, node)| node.clone())
                .collect::<Vec<_>>(),
        );
        overrides.push(("displayText".to_owned(), Any::from(display_text(&shown))));
        if let Some(Out::Any(Any::Map(projection))) = map.get(txn, "resultProjection") {
            let mut projection = (*projection).clone();
            let mut children = match projection.get("children") {
                Some(Any::Array(children)) => children.to_vec(),
                _ => Vec::new(),
            };
            for (index, run) in &runs {
                let text = display_text(&shown_runs(std::slice::from_ref(run)));
                let item = Any::from(HashMap::from([
                    ("kind".to_owned(), Any::from("text")),
                    ("text".to_owned(), Any::from(text)),
                ]));
                children.push(Any::from(HashMap::from([
                    ("index".to_owned(), Any::Number(*index as f64)),
                    ("items".to_owned(), Any::Array(Arc::from([item]))),
                ])));
            }
            projection.insert("children".to_owned(), Any::Array(Arc::from(children)));
            overrides.push((
                "resultProjection".to_owned(),
                Any::Map(Arc::new(projection)),
            ));
        }
    }
    let end = data["continuation"]["end"] != Value::Bool(true);
    if end {
        data["continuation"] =
            serde_json::json!({"id": continuation, "separate": false, "end": true});
    }
    let entries: Vec<(String, Any)> = map
        .iter(txn)
        .filter_map(|(key, value)| match value {
            Out::Any(value) => Some((key.to_owned(), value)),
            _ => None,
        })
        .collect();
    let attrs: Attrs = owner
        .attrs
        .iter()
        .map(|(key, value)| (Arc::from(key.as_str()), value.clone()))
        .collect();
    let (from, to) = (at + 1, owner.start);
    if from < to {
        story.format(
            txn,
            from,
            to - from,
            Attrs::from([(Arc::from(FIELD_RESULT), Any::Null)]),
        );
    }
    let style = paragraph_style(txn, &chunks, before);
    let children: Vec<(u32, i64)> = chunks[before + 2..]
        .iter()
        .take_while(|chunk| chunk.start < owner.start)
        .filter_map(|chunk| field_result_attr(chunk).map(|(_, index)| (chunk.start, index)))
        .collect();
    story.remove_range(txn, to, 1);
    let embed = story.insert_embed_with_attributes(txn, at, yrs::MapPrelim::default(), attrs);
    for (key, value) in entries.into_iter().chain(overrides) {
        embed.insert(txn, key, value);
    }
    embed.insert(txn, "fieldData", data.to_string());
    // Each run before the first child after it, else where the field ends;
    // the last first, so earlier positions hold.
    let mut moved = 0;
    for (index, run) in runs.iter().rev() {
        let mut at = children
            .iter()
            .find(|(_, child)| child > index)
            .map_or(to + 1, |(start, _)| start + 1);
        let units = if run["type"] == "complexField" {
            field_units(run, package, style.as_deref(), 0, false)
        } else {
            run_units(run, package, style.as_deref())
        };
        for unit in units {
            let (len, op) = unit_op(at, unit)?;
            apply(txn, story, story_id, vec![(at, 0, op)])?;
            at += len;
            moved += len;
        }
    }
    let to = to + moved;
    if end {
        embed.insert(txn, "continuationId", continuation);
        let marker = serde_json::json!({
            "id": continuation,
            "kind": "fieldend",
            "run": {"type": "run", "content": [
                {"type": "fieldChar", "charType": "end", "continuationId": continuation}
            ]},
        });
        crate::bookmarks::set(
            txn,
            story,
            story_id,
            to + 1,
            any_from_value(marker).map_err(OpError::InvalidUpdate)?,
        )?;
    }
    Ok(())
}

/// What the seed shows of a projecting field once its export keeps, of the
/// children it `recorded`, only those in `live`
/// (`restoreProjectedFieldResults` drops the rest): its own result runs while
/// a link or simple field is still projected, else its whole result, as
/// `field_payload` reads an unprojected field.
pub(crate) fn seeded_display(data: &Value, recorded: &[i64], live: &[i64]) -> String {
    let kept = |index: i64| !recorded.contains(&index) || live.contains(&index);
    let nodes = |key: &str, index: fn(usize) -> i64| -> Vec<(i64, Value)> {
        data[key]["inline"]
            .as_array()
            .into_iter()
            .flatten()
            .enumerate()
            .map(|(at, node)| (index(at), node.clone()))
            .filter(|(index, _)| kept(*index))
            .collect()
    };
    let result = nodes("structuredResult", |at| at as i64);
    let code = nodes("structuredCode", |at| -(at as i64) - 1);
    let projects = data["type"] == "complexField"
        && !crate::seed::numeric_field_instruction(
            data["instruction"].as_str().unwrap_or_default(),
        )
        && result
            .iter()
            .chain(&code)
            .any(|(_, node)| matches!(node["type"].as_str(), Some("hyperlink" | "simpleField")));
    let result: Vec<Value> = result
        .into_iter()
        .filter(|(index, _)| !projects || !recorded.contains(index))
        .map(|(_, node)| node)
        .collect();
    display_text(&if projects {
        own_shown_runs(&result)
    } else {
        shown_runs(&result)
    })
}

/// The fields whose projected children lie in `start..end` of `chunks`.
pub(crate) fn child_owners(chunks: &[Chunk], start: u32, end: u32) -> Vec<i64> {
    chunks
        .iter()
        .filter(|chunk| chunk.start < end && chunk.end() > start)
        .filter_map(|chunk| field_result_attr(chunk).map(|(id, _)| id))
        .collect()
}

/// Re-reads what the projecting fields numbered `ids` show, after an edit
/// removed some of their projected children ([`seeded_display`]). Only the
/// shown text changes, so peers refreshing at once converge; text left ending
/// a continued result stays in the field until the next publication, which
/// reads it as text after the field (decided 2026-10-04).
pub(crate) fn refresh_shown(txn: &mut TransactionMut<'_>, story: &TextRef, ids: &[i64]) {
    if ids.is_empty() {
        return;
    }
    let chunks = snapshot(story, txn);
    for (position, chunk) in chunks.iter().enumerate() {
        let Some((id, recorded)) = projection(txn, chunk).filter(|(id, _)| ids.contains(id)) else {
            continue;
        };
        let ChunkKind::Embed(Some(map)) = &chunk.kind else {
            continue;
        };
        let Some(Ok(data)) =
            map_string(map, txn, "fieldData").map(|data| serde_json::from_str::<Value>(&data))
        else {
            continue;
        };
        let live: Vec<i64> = chunks[..position]
            .iter()
            .rev()
            .map_while(|chunk| {
                field_result_attr(chunk)
                    .filter(|(child, index)| *child == id && recorded.contains(index))
            })
            .map(|(_, index)| index)
            .collect();
        let shown = seeded_display(&data, &recorded, &live);
        if map_string(map, txn, "displayText").as_deref() != Some(shown.as_str()) {
            map.insert(txn, "displayText", shown);
        }
    }
}

/// Runs `apply` (a peer's update, or an Undo or Redo step), then re-reads
/// what each projecting field a deleted or inserted range now ends at shows
/// ([`refresh_shown`]), in a system transaction peers receive and Undo skips:
/// two peers each deleting part of a field's last link remove it only once
/// each has the other's update, and an Undo that brings the link back
/// restores what the field shows (Epo 2026-10-10). Each deleted range costs
/// one step right from its last item, each inserted item one step right
/// from it; only a field found there reads its story.
pub(crate) fn refreshing_fields<R>(doc: &yrs::Doc, apply: impl FnOnce() -> R) -> R {
    use yrs::block::{ItemContent, ItemPtr};
    use yrs::branch::BranchPtr;
    use yrs::types::TypeRef;
    use yrs::{ID, IdSet, MapRef, Transact};
    let touched = Arc::new(std::sync::Mutex::new((IdSet::new(), IdSet::new())));
    let subscription = {
        let touched = Arc::clone(&touched);
        doc.observe_after_transaction(move |txn| {
            let (deleted, inserted) = &mut *touched.lock().unwrap();
            deleted.merge_with(txn.delete_set().clone());
            inserted.merge_with(txn.insert_set().clone());
        })
        .ok()
    };
    let result = apply();
    drop(subscription);
    let (deleted, inserted) = std::mem::take(&mut *touched.lock().unwrap());
    let mut owners: Vec<(BranchPtr, i64)> = Vec::new();
    {
        let txn = doc.transact();
        // The first live, countable item right of `item`, as
        // `Store::next_live_item` finds it, without looking `item` up again.
        // An insert's walk stops at another inserted item that is deleted or
        // not countable (a format item): that item's own walk goes on from
        // there by the same rule, so a run of k such items in one update
        // costs k steps, not k²/2.
        let next_live = |item: ItemPtr, restores: bool| {
            let mut next = item.right();
            while let Some(right) = next {
                #[cfg(test)]
                tests::PROBE_STEPS.with(|steps| steps.set(steps.get() + 1));
                if !right.is_deleted() && right.is_countable() {
                    return Some(right);
                }
                if restores && inserted.contains(right.id()) {
                    return None;
                }
                next = right.right();
            }
            None
        };
        let deleted = deleted
            .iter()
            .flat_map(|(client, ranges)| ranges.iter().map(|range| ID::new(*client, range.end - 1)))
            .filter_map(|end| txn.store().get_item(&end))
            .map(|item| (item, false))
            .collect::<Vec<_>>();
        // Every inserted item: a restored copy and a peer's system write can
        // share one range of clocks.
        let inserted_items = crate::items_in(&txn, &inserted)
            .into_iter()
            .map(|item| (item, true));
        for (item, restores) in deleted.into_iter().chain(inserted_items) {
            let Some(next) = next_live(item, restores) else {
                continue;
            };
            let Some(story) = next
                .parent_branch()
                .filter(|story| matches!(story.type_ref(), TypeRef::Text))
            else {
                continue;
            };
            let ItemContent::Type(branch) = next.content() else {
                continue;
            };
            if !matches!(branch.type_ref(), TypeRef::Map) {
                continue;
            }
            let map = MapRef::from(BranchPtr::from(branch.as_ref()));
            // Content put back changes what a field shows only from its whole
            // result (none of its projected children left) to its own runs, a
            // part of it; a field showing nothing stays so, which keeps typing
            // at the end of a cross-reference's link from reading its story.
            if restores && map_string(&map, &txn, "displayText").is_none_or(|text| text.is_empty())
            {
                continue;
            }
            let id = match map.get(&txn, "resultProjection") {
                Some(Out::Any(projection)) => any_value(&projection)["id"].as_i64(),
                _ => None,
            };
            if let Some(id) = id
                && !owners.contains(&(story, id))
            {
                owners.push((story, id));
            }
        }
    }
    if owners.is_empty() {
        return result;
    }
    let mut txn = doc.transact_mut_with("system");
    while let Some(&(story, _)) = owners.first() {
        let ids: Vec<i64> = owners
            .iter()
            .filter(|(of, _)| *of == story)
            .map(|(_, id)| *id)
            .collect();
        owners.retain(|(of, _)| *of != story);
        refresh_shown(&mut txn, &TextRef::from(story), &ids);
    }
    result
}

/// What a field shows of `runs`, as the seed's `displayText` reads them: their
/// text, without tabs.
fn display_text(runs: &[Value]) -> String {
    runs.iter()
        .flat_map(|run| run["content"].as_array().cloned().unwrap_or_default())
        .filter(|content| content["type"] == "text")
        .filter_map(|content| content["text"].as_str().map(str::to_owned))
        .collect()
}

/// The result nodes Enter moved out of the projecting field at `position`:
/// runs, nested fields and tracked deletions it records as children but no
/// child shows. They are content of the next paragraph now.
fn moved_out<T: ReadTxn>(txn: &T, chunks: &[Chunk], position: usize, data: &Value) -> Vec<usize> {
    let Some((id, recorded)) = projection(txn, &chunks[position]) else {
        return Vec::new();
    };
    let live: Vec<i64> = chunks[..position]
        .iter()
        .rev()
        .map_while(|chunk| {
            field_result_attr(chunk)
                .filter(|(child, index)| *child == id && recorded.contains(index))
        })
        .map(|(_, index)| index)
        .collect();
    let mut gone: Vec<usize> = recorded
        .iter()
        .filter(|index| **index >= 0 && !live.contains(index))
        .map(|index| *index as usize)
        .filter(|index| {
            data["structuredResult"]["inline"][*index]["type"]
                .as_str()
                .is_some_and(|kind| !matches!(kind, "hyperlink" | "simpleField"))
        })
        .collect();
    gone.sort();
    gone.dedup();
    gone
}

/// Whether the field embed at `position` keeps a tracked change only Accept
/// or Reject All resolves: not one Enter moved out, which is listed in the
/// next paragraph.
pub(crate) fn embed_keeps_changes<T: ReadTxn>(
    txn: &T,
    chunks: &[Chunk],
    position: usize,
    data: &str,
) -> bool {
    if !field_data_keeps_changes(data) {
        return false;
    }
    let Ok(value) = serde_json::from_str::<Value>(data) else {
        return true;
    };
    let gone = moved_out(txn, chunks, position, &value);
    gone.is_empty()
        || field_data_keeps_changes(&crate::seed::without_result_nodes(&value, &gone).to_string())
}

/// A field result node that is one tracked deletion, as the seed reads it
/// outside a field (a field keeps it as raw markup).
fn tracked_deletion(node: &Value) -> Option<Value> {
    let deletion = if node["type"] == "deletion" {
        node.clone()
    } else {
        let xml = node["xml"]
            .as_str()
            .filter(|xml| xml.trim_start().starts_with("<w:del "))?;
        match docx_parse::paragraph::parse_raw_inline(xml)?.as_slice() {
            [deletion] => serde_json::to_value(deletion).ok()?,
            _ => return None,
        }
    };
    // Only plain text runs (text and tabs) showing some text: a field
    // result, note reference, bookmark, break, control or symbol inside it
    // does not move intact, nor does a deletion that shows nothing.
    let plain = deletion["content"].as_array().is_some_and(|nodes| {
        !runs_text(nodes).is_empty()
            && nodes.iter().all(|node| {
                node["type"] == "run"
                    && node["content"].as_array().is_some_and(|content| {
                        content
                            .iter()
                            .all(|item| matches!(item["type"].as_str(), Some("text" | "tab")))
                    })
            })
    });
    (deletion["type"] == "deletion" && plain).then_some(deletion)
}

/// The text runs show, tabs as tab characters.
fn runs_text(runs: &[Value]) -> String {
    runs.iter()
        .flat_map(|run| run["content"].as_array().cloned().unwrap_or_default())
        .filter_map(|content| match content["type"].as_str() {
            Some("text") => content["text"].as_str().map(str::to_owned),
            Some("tab") => Some("\t".to_owned()),
            _ => None,
        })
        .collect()
}

/// Joining a field's paragraphs back undoes `split_field`: the result content
/// after a continued field's embed in its paragraph becomes its projected
/// children again, up to its end when that is now in the same paragraph (the
/// field then ends there again), or to the paragraph's end. Children take
/// their link's slot: the one before them, else the next one the field
/// records. Text after them that the split did not move out stays as the
/// result's tail when the field continues past the paragraph; anything else
/// after the embed leaves the field as it is.
pub(crate) fn rejoin_fields(
    txn: &mut TransactionMut<'_>,
    story: &TextRef,
    story_id: &str,
    at: u32,
) -> OpResult<()> {
    let chunks = snapshot(story, txn);
    let start = chunks
        .iter()
        .rev()
        .find(|chunk| chunk.start < at && matches!(chunk.kind, ChunkKind::Pilcrow(_)))
        .map_or(0, Chunk::end);
    let Some(end) = chunks
        .iter()
        .find(|chunk| chunk.start >= at && matches!(chunk.kind, ChunkKind::Pilcrow(_)))
        .map(|chunk| chunk.start)
    else {
        return Ok(());
    };
    let Some((position, owner)) = chunks
        .iter()
        .enumerate()
        .find(|(_, chunk)| (start..end).contains(&chunk.start) && projection(txn, chunk).is_some())
    else {
        return Ok(());
    };
    let ChunkKind::Embed(Some(map)) = &owner.kind else {
        return Ok(());
    };
    let (id, mut recorded) = projection(txn, owner).expect("found as a projecting field");
    let Some(Ok(mut data)) =
        map_string(map, txn, "fieldData").map(|data| serde_json::from_str::<Value>(&data))
    else {
        return Ok(());
    };
    let Some(continuation) = data["continuation"]["id"].as_str().map(str::to_owned) else {
        return Ok(());
    };
    if data["continuation"]["end"] != Value::Bool(true) {
        return Ok(());
    }
    let marker = crate::bookmarks::positions(txn, story_id)
        .into_iter()
        .find_map(|(at, data)| match &data {
            Any::Map(data)
                if matches!(data.get("kind"), Some(Any::String(kind)) if kind.as_ref() == "fieldend")
                    && matches!(data.get("id"), Some(Any::String(of)) if of.as_ref() == continuation) =>
            {
                Some(at)
            }
            _ => None,
        });
    let ends_here = marker.filter(|marker| (owner.start + 1..=end).contains(marker));
    let target = ends_here.unwrap_or(end);
    if target == owner.start + 1 && ends_here.is_none() {
        return Ok(());
    }
    // The children before the embed, then the units after it, each with its
    // slot: a link's run keeps the one before it, a new link takes the next
    // slot the field records.
    recorded.sort();
    let first = chunks[..position]
        .iter()
        .rev()
        .take_while(|chunk| field_result_attr(chunk).is_some_and(|(child, _)| child == id))
        .count();
    // Each unit with its slot and its length up to `target` (text there can run on past it).
    let mut units: Vec<(&Chunk, i64, u32)> = chunks[position - first..position]
        .iter()
        .map(|chunk| {
            (
                chunk,
                field_result_attr(chunk).expect("a child").1,
                chunk.len,
            )
        })
        .collect();
    let inline = data["structuredResult"]["inline"]
        .as_array()
        .cloned()
        .unwrap_or_default();
    let projected_run = |index: i64| {
        inline
            .get(index as usize)
            .is_some_and(|node| node["type"] == "run")
            && inline[..(index.max(0) as usize).min(inline.len())]
                .iter()
                .rev()
                .find_map(|node| match node["type"].as_str() {
                    Some("simpleField") => Some(true),
                    Some("hyperlink") => Some(false),
                    _ => None,
                })
                .unwrap_or(false)
    };
    // The text a moved run or tracked deletion went out as, and whether it is
    // a deletion (struck text).
    // The text a moved run or tracked deletion went out as, and the
    // deletion's revision id (struck text).
    let moved_text = |node: &Value| match tracked_deletion(node) {
        Some(deletion) => {
            let text: String = run_units(&deletion, None, None)
                .into_iter()
                .map(|(unit, _)| unit.unwrap_or_else(|_| "\u{fffc}".to_owned()))
                .collect();
            (text, Some(deletion["info"]["id"].clone()))
        }
        None if node["type"] == "run" => (runs_text(std::slice::from_ref(node)), None),
        None => (String::new(), None),
    };
    // A run that seeds nothing (an empty or formatting-only run) has no
    // place in the editor: it takes no slot. A projected one stays the
    // child it was; a moved one goes back in order (`back` below).
    let silent = |index: &i64| {
        inline
            .get(*index as usize)
            .is_some_and(|node| node["type"] == "run" && run_units(node, None, None).is_empty())
    };
    let next_slot = |from: i64| {
        recorded
            .iter()
            .copied()
            .find(|next| *next >= from && !silent(next))
    };
    let mut slot = units
        .last()
        .map(|(chunk, index, _)| (*index, chunk.attrs.get(HYPERLINK).cloned()));
    // Result text the split moved out goes back as the runs it came from, in
    // order (start, length and slot of each); anything else (text typed
    // between the halves) keeps the split.
    let mut folded: Vec<(u32, u32, i64)> = Vec::new();
    let mut runs: Vec<(u32, i64, u32)> = Vec::new();
    let mut dropped: Vec<i64> = Vec::new();
    // Where a field that continues past this paragraph has its result's tail
    // here: text after its embed the split did not move out, which the seed
    // reads as text after the embed (`continued_result_tail`).
    let mut tail = None;
    'scan: for chunk in chunks[position + 1..]
        .iter()
        .take_while(|chunk| chunk.start < target)
    {
        // A projected simple field takes a slot of its own.
        let field = matches!(&chunk.kind, ChunkKind::Embed(Some(map))
            if map_string(map, txn, KIND_KEY).as_deref() == Some("field"));
        let link = chunk
            .attrs
            .get(HYPERLINK)
            .filter(|link| **link != Any::Null)
            .cloned()
            .or_else(|| field.then(|| Any::from(format!("field@{}", chunk.start))));
        let len = chunk.len.min(target - chunk.start);
        if let (None, ChunkKind::Text(text)) = (&link, &chunk.kind) {
            let mut left = String::from_utf16_lossy(
                &text.encode_utf16().take(len as usize).collect::<Vec<_>>(),
            );
            let struck = chunk
                .attrs
                .get(crate::DEL)
                .filter(|mark| **mark != Any::Null)
                .map(|mark| any_value(mark)["id"].clone());
            let mut last = slot.as_ref().map_or(-1, |(index, _)| *index);
            let mut taken = Vec::new();
            let mut lengths = Vec::new();
            while !left.is_empty() {
                let next = next_slot(last + 1);
                let node = next
                    .and_then(|next| inline.get(next as usize).cloned())
                    .unwrap_or(Value::Null);
                // Struck text matches only the moved deletion of its own
                // revision, so a deletion suggested on moved text stays that
                // deletion (review F7, R3-2).
                let (own, deletion) = moved_text(&node);
                let same_revision = match (&deletion, &struck) {
                    (None, None) => true,
                    (Some(moved), Some(struck)) => match (moved.as_f64(), struck.as_f64()) {
                        (Some(moved), Some(struck)) => moved == struck,
                        _ => moved == struck,
                    },
                    _ => false,
                };
                let Some(next) =
                    next.filter(|_| !own.is_empty() && left.starts_with(&own) && same_revision)
                else {
                    if ends_here.is_none() && taken.is_empty() {
                        tail = Some(chunk.start);
                        break 'scan;
                    }
                    return Ok(());
                };
                left = left[own.len()..].to_owned();
                taken.push(next);
                lengths.push(own.encode_utf16().count() as u32);
                last = next;
            }
            if taken.iter().any(|index| projected_run(*index)) {
                // Runs the seed projects after a simple field are children,
                // each in its own slot.
                if !taken.iter().all(|index| projected_run(*index)) {
                    return Ok(());
                }
                let mut at = chunk.start;
                for (index, run) in taken.iter().zip(&lengths) {
                    runs.push((at, *index, *run));
                    at += run;
                }
            } else {
                let mut at = chunk.start;
                for (index, run) in taken.iter().zip(&lengths) {
                    folded.push((at, *run, *index));
                    at += run;
                }
                dropped.extend(&taken);
            }
            slot = Some((last, None));
            continue;
        }
        // A moved run holding one embed (a line break, a positional tab)
        // goes back as that run; the seed never projects it.
        if let (None, ChunkKind::Embed(Some(map))) = (&link, &chunk.kind) {
            let next = next_slot(slot.as_ref().map_or(0, |(index, _)| index + 1));
            let kind = map_string(map, txn, KIND_KEY);
            let Some(next) = next.filter(|next| {
                let node = &inline[*next as usize];
                let units = run_units(node, None, None);
                node["type"] == "run"
                    && !chunk.attr_active(crate::DEL)
                    && matches!(units.as_slice(), [(Err((own, _)), _)] if Some(own) == kind.as_ref())
            }) else {
                if ends_here.is_none() {
                    tail = Some(chunk.start);
                    break 'scan;
                }
                return Ok(());
            };
            folded.push((chunk.start, 1, next));
            dropped.push(next);
            slot = Some((next, None));
            continue;
        }
        // A nested field the split moved out goes back to its slot.
        if field
            && let Some(next) = next_slot(slot.as_ref().map_or(0, |(index, _)| index + 1))
                .filter(|next| inline[*next as usize]["type"] == "complexField")
        {
            folded.push((chunk.start, 1, next));
            dropped.push(next);
            slot = Some((next, None));
            continue;
        }
        if link.is_none() || !matches!(chunk.kind, ChunkKind::Text(_) | ChunkKind::Embed(Some(_))) {
            return Ok(());
        }
        let next = match &slot {
            Some((index, previous)) if *previous == link => Some(*index),
            Some((index, _)) => next_slot(index + 1),
            None => next_slot(0),
        };
        let Some(index) = next else {
            // A comment's reference or other content the split did not move
            // out begins the tail of a field that continues.
            if ends_here.is_none() {
                tail = Some(chunk.start);
                break 'scan;
            }
            return Ok(());
        };
        slot = Some((index, link));
        units.push((chunk, index, len));
    }
    // Moved runs left ending this paragraph's part of a result that continues
    // are its tail, which the seed of the save reads as text after the field
    // (`continued_result_tail`): they stay where they are. A run the field
    // keeps after them (a tracked deletion) keeps them in, in their order.
    let plain = |from: u32| {
        chunks
            .iter()
            .filter(|chunk| chunk.start >= from && chunk.start < end)
            .all(|chunk| match &chunk.kind {
                ChunkKind::Text(_) => !chunk.attr_active(HYPERLINK),
                ChunkKind::Embed(Some(map)) => matches!(
                    map_string(map, txn, KIND_KEY).as_deref(),
                    Some("tab" | "break")
                ),
                _ => false,
            })
    };
    if ends_here.is_none() && tail.is_none_or(plain) {
        let after = units
            .iter()
            .filter(|(chunk, _, _)| chunk.start > owner.start)
            .map(|(chunk, _, len)| chunk.start + len)
            .chain(runs.iter().map(|(start, _, len)| start + len))
            .max()
            .unwrap_or(0);
        while let Some(&(start, _, index)) = folded.last() {
            if start < after
                || !crate::seed::tail_run(&inline[index as usize])
                || (index + 1..inline.len() as i64).any(|later| !recorded.contains(&later))
            {
                break;
            }
            folded.pop();
            dropped.retain(|dropped| *dropped != index);
            tail = Some(start);
        }
    }
    // A moved silent run goes back with the moved run after it, or, last,
    // with the one before it, so the result keeps its order.
    let back: Vec<i64> = units
        .iter()
        .map(|(_, index, _)| *index)
        .chain(runs.iter().map(|(_, index, _)| *index))
        .chain(dropped.iter().copied())
        .collect();
    let quiet: Vec<i64> = recorded
        .iter()
        .copied()
        .filter(|index| silent(index) && !projected_run(*index))
        .filter(|index| {
            let after = recorded
                .iter()
                .copied()
                .find(|next| next > index && !silent(next));
            let before = recorded
                .iter()
                .copied()
                .rev()
                .find(|next| next < index && !silent(next));
            match (after, before) {
                (Some(after), _) => back.contains(&after),
                (None, Some(before)) => back.contains(&before),
                (None, None) => true,
            }
        })
        .collect();
    dropped.extend(quiet);
    let target = tail.unwrap_or(target);
    if tail.is_some()
        && folded.is_empty()
        && runs.is_empty()
        && !units.iter().any(|(chunk, _, _)| chunk.start > owner.start)
    {
        return Ok(());
    }
    if ends_here.is_some() {
        if data["continuation"]["separate"] == Value::Bool(true) {
            data["continuation"]["end"] = Value::Bool(false);
        } else if let Some(field) = data.as_object_mut() {
            field.remove("continuation");
        }
        if let Some(root) = txn.get_map(crate::bookmarks::ROOT) {
            root.remove(txn, &format!("{story_id}:{continuation}:fieldend"));
        }
    }
    let mut overrides: Vec<(String, Any)> = Vec::new();
    if !dropped.is_empty() {
        recorded.retain(|index| !dropped.contains(index));
        let shown = own_shown_runs(
            &inline
                .iter()
                .enumerate()
                .filter(|(index, _)| !recorded.contains(&(*index as i64)))
                .map(|(_, node)| node.clone())
                .collect::<Vec<_>>(),
        );
        overrides.push(("displayText".to_owned(), Any::from(display_text(&shown))));
        if let Some(Out::Any(Any::Map(projection))) = map.get(txn, "resultProjection") {
            let mut projection = (*projection).clone();
            if let Some(Any::Array(children)) = projection.get("children") {
                let kept: Vec<Any> = children
                    .iter()
                    .filter(|child| match child {
                        Any::Map(child) => !matches!(child.get("index"), Some(Any::Number(index)) if dropped.contains(&(*index as i64))),
                        _ => true,
                    })
                    .cloned()
                    .collect();
                projection.insert("children".to_owned(), Any::Array(Arc::from(kept)));
            }
            overrides.push((
                "resultProjection".to_owned(),
                Any::Map(Arc::new(projection)),
            ));
        }
    }
    for (start, index, len) in units
        .iter()
        .filter(|(chunk, _, _)| chunk.start > owner.start)
        .map(|(chunk, index, len)| (chunk.start, *index, *len))
        .chain(runs)
    {
        let marker = Any::from(HashMap::from([
            ("id".to_owned(), Any::Number(id as f64)),
            ("index".to_owned(), Any::Number(index as f64)),
        ]));
        story.format(
            txn,
            start,
            len,
            Attrs::from([(Arc::from(FIELD_RESULT), marker)]),
        );
    }
    let mut entries: Vec<(String, Any)> = map
        .iter(txn)
        .filter_map(|(key, value)| match value {
            Out::Any(value) => Some((key.to_owned(), value)),
            _ => None,
        })
        .filter(|(key, _)| key != "fieldData")
        .filter(|(key, _)| key != "continuationId" || data.get("continuation").is_some())
        .collect();
    entries.push(("fieldData".to_owned(), Any::from(data.to_string())));
    entries.extend(overrides);
    let attrs: Attrs = owner
        .attrs
        .iter()
        .map(|(key, value)| (Arc::from(key.as_str()), value.clone()))
        .collect();
    let mut removed = 0;
    for (start, len, _) in folded.iter().rev() {
        story.remove_range(txn, *start, *len);
        removed += len;
    }
    story.remove_range(txn, owner.start, 1);
    let embed = story.insert_embed_with_attributes(
        txn,
        target - removed - 1,
        yrs::MapPrelim::default(),
        attrs,
    );
    for (key, value) in entries {
        embed.insert(txn, key, value);
    }
    Ok(())
}

/// Clears the `fieldResult` markers of the children of the projecting field
/// embeds in `start..end` of `story`, which a delete is about to remove: the
/// children right before each that carry its number. A child whose field is
/// deleted is plain content written in place, which no other field pairs,
/// even one that comes to share its number after a join or Accept All.
/// `chunks` covers the range and may start before it.
pub(crate) fn release_children(
    txn: &mut TransactionMut<'_>,
    story: &TextRef,
    chunks: &[Chunk],
    start: u32,
    end: u32,
) {
    // Only the first field in the range can have children before it; the
    // others' lie in the range and go with it.
    let Some((first, id)) = chunks.iter().enumerate().find_map(|(index, chunk)| {
        let (id, _) = projection(txn, chunk).filter(|_| (start..end).contains(&chunk.start))?;
        Some((index, id))
    }) else {
        return;
    };
    let at = chunks[first].start;
    let child = |chunk: &&Chunk| field_result_attr(chunk).is_some_and(|(child, _)| child == id);
    let run = chunks[..first].iter().rev().take_while(child).count();
    let from = if run < first {
        chunks[first - run].start
    } else {
        // The children may begin before the chunks at hand.
        let all = snapshot(story, txn);
        let position = all.partition_point(|chunk| chunk.start < at);
        all[..position]
            .iter()
            .rev()
            .take_while(child)
            .last()
            .map_or(at, |chunk| chunk.start)
    };
    if from < start {
        story.format(
            txn,
            from,
            start - from,
            Attrs::from([(Arc::from(FIELD_RESULT), Any::Null)]),
        );
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
    attrs.get(FIELD_RESULT)?["index"].as_i64()
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
        Err((kind, values)) if kind == "bookmark" => (
            0,
            RawOp::SetBookmark {
                index,
                data: any_from_value(serde_json::json!(values)).map_err(error)?,
            },
        ),
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
    // The chunks the export writes inside a field's result: the children
    // right before each field embed that carry its number at an index it
    // records. Any other chunk is written in place.
    let mut paired = vec![false; position - first];
    for owner in first..=position {
        let Some((id, indices)) = projection(txn, &chunks[owner]) else {
            continue;
        };
        for child in (first..owner).rev() {
            match field_result_attr(&chunks[child]) {
                Some((child_id, index)) if child_id == id && indices.contains(&index) => {
                    paired[child - first] = true;
                }
                _ => break,
            }
        }
    }
    let paragraph = chunks[first].start..=chunks[position].start;
    let comments: usize = comments
        .iter()
        .filter(|(at, _)| paragraph.contains(at))
        .map(|&(at, adds_reference)| {
            let splits = chunks[first..position]
                .iter()
                .zip(&paired)
                .any(|(chunk, &paired)| {
                    matches!(chunk.kind, ChunkKind::Text(_))
                        && !paired
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
    // Breaks a text-less paragraph owned (ending in a column break, none taken
    // from the paragraph before) lead the text after them too.
    let (mut breaks, mut trailing, mut column) = (0, false, false);
    while head < position {
        let ChunkKind::Embed(Some(map)) = &chunks[head].kind else {
            break;
        };
        match map_string(map, txn, KIND_KEY).as_deref() {
            Some("table" | "blockSdt") => {
                after_paragraph = false;
                leads = false;
                (breaks, trailing, column) = (0, false, false);
            }
            Some(kind @ ("pageBreak" | "columnBreak")) => {
                leads |= !after_paragraph
                    || matches!(map.get(txn, "leading"), Some(Out::Any(Any::Bool(true))));
                if leads {
                    count += 1;
                }
                breaks += 1;
                trailing |= matches!(map.get(txn, "trailing"), Some(Out::Any(Any::Bool(true))));
                column = kind == "columnBreak";
            }
            _ => break,
        }
        head += 1;
    }
    if !leads && column && !trailing {
        count += breaks;
    }
    for (chunk, &paired) in chunks[head..position].iter().zip(&paired[head - first..]) {
        offset += match &chunk.kind {
            ChunkKind::Embed(Some(map))
                if map_string(map, txn, KIND_KEY).as_deref() == Some("sdt") =>
            {
                2
            }
            _ => chunk.len,
        };
        if paired {
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
    let mut boundaries = comments.to_vec();
    boundaries.extend(crate::bookmarks::positions(txn, story_id).into_iter().filter_map(|(at, data)| {
        (!matches!(data, Any::Map(data) if data.get("inHyperlink") == Some(&Any::Bool(true)) || matches!(data.get("id"), Some(Any::String(_))))).then_some((at, false))
    }));
    let chunks = snapshot(story, txn);
    for position in (0..chunks.len()).rev() {
        let ChunkKind::Embed(Some(map)) = &chunks[position].kind else {
            continue;
        };
        match map_string(map, txn, KIND_KEY).as_deref() {
            Some("field") if field_result_attr(&chunks[position]).is_none() => {
                resolve_owner(txn, story, story_id, &chunks, position, how, &boundaries)?
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
fn apply(
    txn: &mut TransactionMut<'_>,
    story: &TextRef,
    story_id: &str,
    mut edits: Vec<(u32, usize, RawOp)>,
) -> OpResult<()> {
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
            RawOp::SetBookmark { index, data } => {
                crate::bookmarks::set(txn, story, story_id, index, data)?
            }
            _ => {}
        }
    }
    Ok(())
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
    story_id: &str,
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
    let Some(mut places) = resolve_field(&mut field, how) else {
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
    // What Enter moved out is content in the next paragraph, which resolves
    // there: the field resolves without it.
    let gone = moved_out(txn, chunks, position, &old);
    if !gone.is_empty() {
        field = crate::seed::without_result_nodes(&old, &gone);
        match resolve_field(&mut field, how) {
            Some(resolved) => places = resolved,
            None => return Ok(()),
        }
    }
    let [code_moved, result_moved] = places;
    // Where an old child now sits, and whether a field inside it resolved.
    let moved = |index: i64| -> Option<(i64, bool)> {
        if index >= 0 {
            if gone.contains(&(index as usize)) {
                return None;
            }
            let at = index as usize - gone.iter().filter(|gone| (**gone as i64) < index).count();
            let (at, nested) = result_moved.get(at).copied().flatten()?;
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
                FIELD_RESULT.to_owned(),
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
        let mut bookmarks = Vec::new();
        for unit in group {
            let (len, op) = unit_op(at, unit)?;
            let edit = (anchors[order], 1 + count - order, op);
            if matches!(edit.2, RawOp::SetBookmark { .. }) {
                bookmarks.push(edit);
            } else {
                edits.push(edit);
            }
            at += len;
        }
        // Anchor only after the group's text exists. Later groups inserted
        // before this one then move these boundaries with that text.
        edits.extend(bookmarks);
    }
    apply(txn, story, story_id, edits)?;
    set_payload(txn, map, owner_payload)
}

/// An inline content control's content with its revisions and fields resolved.
fn resolved_sdt_content(content: &Any, how: Resolve<'_>, style: Option<&str>) -> Option<Any> {
    let Any::Array(items) = content else {
        return None;
    };
    let mut changed = false;
    let items: Vec<Any> = items
        .iter()
        .filter_map(|item| {
            let Any::Map(entry) = item else {
                return Some(item.clone());
            };
            let mut entry = entry.as_ref().clone();
            let mut attrs = match entry.get("attrs") {
                Some(Any::Map(attrs)) => attrs.as_ref().clone(),
                _ => HashMap::new(),
            };
            let active = |key: &str| attrs.get(key).is_some_and(|value| *value != Any::Null);
            if active(if how.accept { "del" } else { "ins" }) {
                changed = true;
                return None;
            }
            if active("ins") || active("del") {
                changed = true;
                attrs.remove("ins");
                attrs.remove("del");
                entry.insert("attrs".to_owned(), Any::Map(Arc::new(attrs.clone())));
            }
            if let Some((delta, _)) =
                super::resolve::run_property_delta(attrs.get("rPrChange"), how.accept, None)
            {
                for (key, value) in delta {
                    if value == Any::Null {
                        attrs.remove(&key);
                    } else {
                        attrs.insert(key, value);
                    }
                }
                entry.insert("attrs".into(), Any::Map(Arc::new(attrs.clone())));
                changed = true;
            }
            if let Some(Any::Map(payload)) = entry.get("payload") {
                match entry.get("kind") {
                    Some(Any::String(kind)) if kind.as_ref() == "sdt" => {
                        if let Some(inner) = payload
                            .get("content")
                            .and_then(|inner| resolved_sdt_content(inner, how, style))
                        {
                            let mut payload = payload.as_ref().clone();
                            payload.insert("content".to_owned(), inner);
                            entry.insert("payload".to_owned(), Any::Map(Arc::new(payload)));
                            changed = true;
                        }
                    }
                    Some(Any::String(kind)) if kind.as_ref() == "field" => {
                        if let Some(Any::String(data)) = payload.get("fieldData")
                            && let Some((values, old_marks, marks)) =
                                resolved_embed(data, how, style)
                            && let Ok(values) =
                                any_from_value(Value::Object(values.into_iter().collect()))
                        {
                            for key in old_marks.keys() {
                                attrs.remove(key);
                            }
                            for (key, value) in marks {
                                if let Ok(value) = any_from_value(value) {
                                    attrs.insert(key, value);
                                }
                            }
                            entry.insert("payload".to_owned(), values);
                            entry.insert("attrs".to_owned(), Any::Map(Arc::new(attrs)));
                            changed = true;
                        }
                    }
                    _ => {}
                }
            }
            Some(Any::Map(Arc::new(entry)))
        })
        .collect();
    changed.then(|| Any::Array(Arc::from(items)))
}

#[cfg(test)]
mod tests {
    use std::cell::Cell;

    use super::*;
    use serde_json::json;

    thread_local! {
        /// The items the re-read probes' walks stepped over.
        pub(super) static PROBE_STEPS: Cell<usize> = const { Cell::new(0) };
    }

    /// The vendored yrs's `Store::next_live_item` (the peer re-read's lookup):
    /// tested here, since CI does not run the vendored crate's own tests.
    #[test]
    fn next_live_item_skips_deleted_items() {
        use yrs::{ClientID, Doc, GetString, ID, Transact};
        let doc = Doc::with_client_id(1);
        let text = doc.get_or_insert_text("test");
        text.insert(&mut doc.transact_mut(), 0, "ab");
        text.insert(&mut doc.transact_mut(), 2, "cd");
        text.remove_range(&mut doc.transact_mut(), 1, 2);
        assert_eq!(text.get_string(&doc.transact()), "ad");
        let txn = doc.transact();
        let client = ClientID::new(1);
        // "b" (clock 1) and "c" (clock 2) are deleted: right of either is "d".
        for clock in [1, 2] {
            let (next, parent) = txn.store().next_live_item(&ID::new(client, clock)).unwrap();
            assert_eq!(next, ID::new(client, 3));
            let branch: &yrs::branch::Branch = text.as_ref();
            assert!(parent == yrs::branch::BranchPtr::from(branch));
        }
        assert!(txn.store().next_live_item(&ID::new(client, 3)).is_none());
    }

    /// A peer typing at the end of a cross-reference's link, the projected
    /// child right before its field, makes the peers that receive it re-read
    /// nothing: the field shows nothing, and content put back cannot change
    /// that. A peer deleting the link makes them re-read the field. Counted
    /// in the system transactions a re-read opens.
    #[test]
    fn a_peer_typing_at_a_cross_references_end_re_reads_no_field() {
        use crate::{EditCtx, EditingDoc, FormatPolicy, Position, StoryRange};
        use std::sync::atomic::{AtomicUsize, Ordering};
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="11111111"><w:r><w:t>a</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF a \h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:hyperlink w:anchor="target"><w:r><w:t>AA</w:t></w:r></w:hyperlink><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:body></w:document>"#;
        let bytes = ooxml_opc::rezip_parts(&[("word/document.xml".into(), xml.to_vec())]).unwrap();
        let base = EditingDoc::new(7);
        crate::seed::seed_from_docx(&base, &bytes).unwrap();
        let state = base.encode_state_as_update_v1();
        let (receiver, typist) = (EditingDoc::new(701), EditingDoc::new(702));
        for doc in [&receiver, &typist] {
            doc.apply_update_v1(&state).unwrap();
        }
        let rereads = Arc::new(AtomicUsize::new(0));
        let _counting = {
            let rereads = Arc::clone(&rereads);
            receiver
                .yrs_doc()
                .observe_after_transaction(move |txn| {
                    if txn.origin() == Some(&yrs::Origin::from("system")) {
                        rereads.fetch_add(1, Ordering::Relaxed);
                    }
                })
                .unwrap()
        };
        let ctx = EditCtx::local("Ada", "2026-10-10T00:00:00Z");
        let deliver = |edit: &dyn Fn(&EditingDoc)| {
            let before = typist.encode_state_vector_v1();
            edit(&typist);
            let update = typist.encode_diff_v1(&before).unwrap();
            rereads.store(0, Ordering::Relaxed);
            refreshing_fields(receiver.yrs_doc(), || {
                receiver.apply_update_v1(&update).unwrap()
            });
            rereads.load(Ordering::Relaxed)
        };
        // "a", the link's "AA" at 1..3, then the field.
        let typed = deliver(&|doc| {
            doc.insert_text(&ctx, Position::new("body", 3), "X", FormatPolicy::Inherit)
                .unwrap();
        });
        assert_eq!(typed, 0, "typing at the link's end re-read the field");
        let deleted = deliver(&|doc| {
            doc.delete_range(&ctx, StoryRange::new("body", 1, 4))
                .unwrap();
        });
        assert_eq!(deleted, 1, "deleting the link re-read no field");
    }

    /// A catch-up update holds a garbage-collected run (the paragraph mark of
    /// a split the peer joined back) before the link the peer's Undo
    /// restored, right before a field that shows text: the receiver re-reads
    /// that field. The receiver deleted the link itself, so the peer's delete
    /// of it changes nothing there and only the restored copy can say so.
    #[test]
    fn a_link_restored_after_collected_content_is_re_read() {
        use crate::{EditCtx, EditingDoc, Position, StoryRange, UndoSession};
        use std::sync::atomic::{AtomicUsize, Ordering};
        use yrs::Transact;
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="11111111"><w:r><w:t>a</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF a \h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>x</w:t></w:r><w:hyperlink w:anchor="target"><w:r><w:t>AA</w:t></w:r></w:hyperlink><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p><w:p w14:paraId="22222222"><w:r><w:t>tail text</w:t></w:r></w:p></w:body></w:document>"#;
        let bytes = ooxml_opc::rezip_parts(&[("word/document.xml".into(), xml.to_vec())]).unwrap();
        let base = EditingDoc::new(7);
        crate::seed::seed_from_docx(&base, &bytes).unwrap();
        let state = base.encode_state_as_update_v1();
        let (peer, receiver) = (EditingDoc::new(702), EditingDoc::new(703));
        for doc in [&peer, &receiver] {
            doc.apply_update_v1(&state).unwrap();
        }
        let shown = |doc: &EditingDoc| {
            let txn = doc.yrs_doc().transact();
            let story = crate::story_ref(&txn, "body").unwrap();
            snapshot(&story, &txn)
                .iter()
                .filter_map(|chunk| match &chunk.kind {
                    ChunkKind::Embed(Some(map)) => map_string(map, &txn, "displayText"),
                    _ => None,
                })
                .collect::<Vec<_>>()
        };
        let ctx = EditCtx::local("Ada", "2026-10-10T00:00:00Z");
        // "a", the link's "AA" at 1..3, the field (showing "x") at 3, then
        // "tail text" from 5.
        receiver
            .delete_range(&ctx, StoryRange::new("body", 1, 3))
            .unwrap();
        assert_eq!(shown(&receiver), ["x"]);
        peer.split_paragraph(&ctx, Position::new("body", 7))
            .unwrap();
        peer.delete_range(&ctx, StoryRange::new("body", 7, 8))
            .unwrap();
        let undo = UndoSession::new();
        undo.track(&peer);
        peer.delete_range(&ctx, StoryRange::new("body", 1, 3))
            .unwrap();
        undo.add_undo_barrier();
        assert!(undo.undo());
        let update = peer
            .encode_diff_v1(&receiver.encode_state_vector_v1())
            .unwrap();
        let rereads = Arc::new(AtomicUsize::new(0));
        let inserted = Arc::new(std::sync::Mutex::new(yrs::IdSet::new()));
        let _observing = {
            let (rereads, inserted) = (Arc::clone(&rereads), Arc::clone(&inserted));
            receiver
                .yrs_doc()
                .observe_after_transaction(move |txn| {
                    if txn.origin() == Some(&yrs::Origin::from("system")) {
                        rereads.fetch_add(1, Ordering::Relaxed);
                    } else {
                        inserted
                            .lock()
                            .unwrap()
                            .merge_with(txn.insert_set().clone());
                    }
                })
                .unwrap()
        };
        refreshing_fields(receiver.yrs_doc(), || {
            receiver.apply_update_v1(&update).unwrap()
        });
        let txn = receiver.yrs_doc().transact();
        assert!(
            crate::collected_runs(&txn, &inserted.lock().unwrap()) > 0,
            "no collected run"
        );
        assert_eq!(
            rereads.load(Ordering::Relaxed),
            1,
            "the field was not re-read"
        );
    }

    /// One update holding a long run of deleted items (each typed at the
    /// start, so none merges with the next, then all deleted) costs the
    /// re-read probes steps linear in the run, as a reconnect or the runtime's
    /// queued updates bring: an insert's walk stops at the next inserted
    /// tombstone, whose own walk goes on.
    #[test]
    fn a_long_deleted_run_in_one_update_costs_linear_probe_steps() {
        use yrs::updates::decoder::Decode;
        use yrs::{Doc, StateVector, Text, Transact, Update};
        let run = 2_000;
        let writer = Doc::with_client_id(1);
        let text = writer.get_or_insert_text("story");
        for _ in 0..run {
            text.insert(&mut writer.transact_mut(), 0, "x");
        }
        text.remove_range(&mut writer.transact_mut(), 0, run);
        let state = writer
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        let fresh = Doc::with_client_id(2);
        fresh.get_or_insert_text("story");
        PROBE_STEPS.with(|steps| steps.set(0));
        refreshing_fields(&fresh, || {
            fresh
                .transact_mut()
                .apply_update(Update::decode_v1(&state).unwrap())
                .unwrap()
        });
        // One step per inserted item, and the deleted range's own walk over
        // the run; a walk to the run's end from every item takes 2 million.
        let steps = PROBE_STEPS.with(Cell::get);
        assert!(
            steps <= 3 * run as usize,
            "{steps} probe steps for a run of {run}"
        );
    }

    #[test]
    fn nested_revision_nodes_and_control_units_resolve() {
        for accept in [false, true] {
            let how = Resolve {
                accept,
                package: None,
                relationships: None,
            };
            let mut field = json!({"type":"complexField", "structuredResult":{"inline":[
                {"type":"hyperlink","structuredChildren":[{"type":"insertion","content":[{"type":"run","content":[{"type":"text","text":"new"}]}]}]},
                {"type":"inlineSdt","content":[{"type":"deletion","content":[{"type":"run","content":[{"type":"text","text":"old"}]}]}]}
            ]}});
            assert!(field_data_keeps_changes(&field.to_string()));
            assert!(resolve_field(&mut field, how).is_some());
            let saved = field.to_string();
            assert!(!saved.contains("insertion") && !saved.contains("deletion"));
            assert_eq!(saved.contains("new"), accept);
            assert_eq!(saved.contains("old"), !accept);

            let content = any_from_value(json!([
                {"text":"new","attrs":{"ins":{"revisionId":1}}},
                {"text":"old","attrs":{"del":{"revisionId":2}}},
                {"text":"both","attrs":{"ins":{"revisionId":1},"del":{"revisionId":2}}}
            ]))
            .unwrap();
            assert!(sdt_keeps_changes(&content));
            let resolved = resolved_sdt_content(&content, how, None).unwrap();
            assert_eq!(
                any_value(&resolved),
                json!([{ "text": if accept { "new" } else { "old" }, "attrs": {} }])
            );
            assert!(!sdt_keeps_changes(&resolved));
        }
    }
}
