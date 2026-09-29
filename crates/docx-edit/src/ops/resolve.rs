//! Tracked-change resolution: `accept_change` and `reject_change`.
//!
//! Each stamp class resolves two ways:
//!
//! | stamp    | accept                           | reject                           |
//! |----------|----------------------------------|----------------------------------|
//! | `ins`    | drop the stamp, text stays       | remove the text                  |
//! | `del`    | remove the text                  | drop the stamp, text stays       |
//! | `pPrIns` | clear the marker, split stays    | remove the pilcrow, paragraphs join |
//! | `pPrDel` | remove the pilcrow, paragraphs join | clear the marker, split stays |
//!
//! A unit carrying BOTH `ins` and `del` — one author suggesting over another's
//! suggestion — is removed either way: the remove class wins over the keep
//! class on the same content.
//!
//! Removing a boundary pilcrow joins two paragraphs, and it is the FOLLOWING
//! paragraph's mark that survives, so the merged paragraph keeps the SECOND
//! paragraph's properties and paraId. That is the OOXML rule — the surviving
//! `w:p` owns the properties — and it deliberately differs from a plain
//! delete, which models a user removing a paragraph mark rather than a
//! revision being applied. A story's FINAL pilcrow is never removed, because a
//! document always keeps its last paragraph mark; a join that would remove it
//! clears the markers instead. Neither is a mark whose paragraph still holds
//! content when the next paragraph opens with a table, block content control
//! or break, whether its deletion is accepted or its insertion rejected: the
//! editor never puts content ahead of such a block in one paragraph, so the
//! mark stays and only its markers go.
//!
//! Resolving APPLIES a revision, it does not author one: no new revision is
//! ever stamped and the context's suggesting mode is ignored.
//!
//! Structural table-row revisions (`trIns`/`trDel`) live in each row's `trPr`
//! bag and resolve in the same transaction as story-unit revisions; removing a
//! row also removes the cell stories it made unreachable. Paragraph-property
//! revisions (`pPrChange`) resolve alongside them: accepting drops the record,
//! rejecting restores the properties it captured.

use std::sync::Arc;

use yrs::types::Attrs;
use yrs::{Any, Map, MapRef, Out, ReadTxn, Text, TextRef, TransactionMut};

use crate::op::{OpError, OpResult, Receipt, loc_range_in_txn};
use crate::ops::table::resolve_table_row_revisions;
use crate::ops::{Chunk, ChunkKind, last_pilcrow, snapshot, snapshot_range};
use crate::queries::revision_parts;
use crate::segments::is_block_embed;
use crate::{
    DEL, EditCtx, EditingDoc, INS, KIND_KEY, PARA_ID, PPR_CHANGE, PPR_DEL, PPR_INS, RevisionId,
    StoryRange, check_range, story_ref,
};

/// What a resolve op targets.
#[derive(Clone, Debug, PartialEq)]
pub enum ChangeTarget {
    /// Resolve every tracked change overlapping the range (no id filtering).
    Range(StoryRange),
    /// Resolve every unit stamped with this revision id, in any story.
    Revision(RevisionId),
    /// Resolve every tracked change in every story, with those fields keep
    /// as markup in their code or result (which the revision list leaves out).
    All,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum ResolveMode {
    Accept,
    Reject,
}

/// Returns the stamp value when it is active (non-null) and — under a revision-id
/// filter — carries that id. Stamps without a parseable revision id never match a filter
/// (they remain resolvable by range).
fn active_stamp(value: Option<Any>, filter: Option<&str>) -> Option<Any> {
    let value = value?;
    if value == Any::Null {
        return None;
    }
    match filter {
        None => Some(value),
        Some(id) => match revision_parts(&value) {
            Some((stamp_id, ..)) if stamp_id == id => Some(value),
            _ => None,
        },
    }
}

fn map_stamp<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> Option<Any> {
    match map.get(txn, key) {
        Some(Out::Any(value)) => Some(value),
        _ => None,
    }
}

/// Records the revision id carried by a resolved stamp (deduplicated, resolution order).
fn record(resolved: &mut Vec<String>, stamp: Option<&Any>) {
    let Some(stamp) = stamp else {
        return;
    };
    if let Some((id, ..)) = revision_parts(stamp)
        && !resolved.contains(&id)
    {
        resolved.push(id);
    }
}

fn clear_attr(txn: &mut TransactionMut<'_>, story: &TextRef, start: u32, len: u32, key: &str) {
    story.format(txn, start, len, Attrs::from([(Arc::from(key), Any::Null)]));
}

fn property_map<'a>(
    change: &'a Any,
    key: &str,
) -> Option<&'a std::collections::HashMap<String, Any>> {
    let Any::Map(change) = change else {
        return None;
    };
    match change.get(key) {
        Some(Any::Map(value)) => Some(value.as_ref()),
        _ => None,
    }
}

/// Rewinds a pilcrow to the `previousFormatting` a `pPrChange` captured:
/// properties the change introduced are removed, then the previous values are
/// written back. Introducing `numPr` also clears the derived list attributes.
/// Schema-managed keys are never touched.
fn restore_paragraph_properties(txn: &mut TransactionMut<'_>, map: &MapRef, change: &Any) {
    let previous = property_map(change, "previousFormatting");
    let current = property_map(change, "currentFormatting");
    if let Some(current) = current {
        for key in current.keys() {
            if previous.is_none_or(|prior| !prior.contains_key(key))
                && !matches!(
                    key.as_str(),
                    KIND_KEY | PARA_ID | PPR_INS | PPR_DEL | PPR_CHANGE
                )
            {
                map.remove(txn, key);
            }
        }
        if current.contains_key("numPr")
            && previous.is_none_or(|prior| !prior.contains_key("numPr"))
        {
            for key in [
                "numPr",
                "listIsBullet",
                "listNumFmt",
                "listMarker",
                "listLevel",
                "listStart",
            ] {
                map.remove(txn, key);
            }
        }
    }
    if let Some(previous) = previous {
        for (key, value) in previous {
            if !matches!(
                key.as_str(),
                KIND_KEY | PARA_ID | PPR_INS | PPR_DEL | PPR_CHANGE
            ) {
                map.insert(txn, key.clone(), value.clone());
            }
        }
    }
}

/// Resolves the `pPrChange` records on one pilcrow. Matching records are
/// consumed — rejecting restores their captured properties, accepting only
/// drops them — and non-matching records stay for a later resolve.
fn resolve_paragraph_property_changes(
    txn: &mut TransactionMut<'_>,
    map: &MapRef,
    mode: ResolveMode,
    filter: Option<&str>,
    resolved: &mut Vec<String>,
) {
    let changes = match map.get(txn, PPR_CHANGE) {
        Some(Out::Any(Any::Array(changes))) => changes.to_vec(),
        _ => return,
    };
    let mut remaining = Vec::new();
    for change in changes {
        if active_stamp(Some(change.clone()), filter).is_some() {
            record(resolved, Some(&change));
            if mode == ResolveMode::Reject {
                restore_paragraph_properties(txn, map, &change);
            }
        } else {
            remaining.push(change);
        }
    }
    if remaining.is_empty() {
        map.remove(txn, PPR_CHANGE);
    } else {
        map.insert(txn, PPR_CHANGE, Any::Array(Arc::from(remaining)));
    }
}

/// Whether the unit at `index` is a block embed that opens a paragraph slot.
fn opens_with_block<T: ReadTxn>(story: &TextRef, txn: &T, index: u32) -> bool {
    snapshot_range(story, txn, index, index + 1)
        .first()
        .is_some_and(|chunk| match &chunk.kind {
            ChunkKind::Embed(Some(map)) => {
                crate::map_string(map, txn, KIND_KEY).is_some_and(|kind| is_block_embed(&kind))
            }
            _ => false,
        })
}

/// Whether the paragraph ending after `before` keeps inline content once
/// resolving in `mode` under `filter` has removed its units (deleted ones on
/// accept, inserted ones on reject); `None` when `before` ends without
/// reaching the paragraph's start.
fn holds_content<T: ReadTxn>(
    before: &[Chunk],
    txn: &T,
    mode: ResolveMode,
    filter: Option<&str>,
) -> Option<bool> {
    let removed = match mode {
        ResolveMode::Accept => DEL,
        ResolveMode::Reject => INS,
    };
    for chunk in before.iter().rev() {
        match &chunk.kind {
            ChunkKind::Pilcrow(_) => return Some(false),
            ChunkKind::Embed(Some(map))
                if crate::map_string(map, txn, KIND_KEY)
                    .is_some_and(|kind| is_block_embed(&kind)) => {}
            _ if active_stamp(chunk.attrs.get(removed).cloned(), filter).is_none() => {
                return Some(true);
            }
            _ => {}
        }
    }
    before
        .first()
        .is_some_and(|chunk| chunk.start == 0)
        .then_some(false)
}

/// Resolves one story's tracked changes in place. `span` limits the walk to a story range
/// (`None` = the whole story, the by-id path); `filter` limits it to one revision id.
/// Returns the number of units physically removed inside `span`.
fn resolve_story(
    txn: &mut TransactionMut<'_>,
    story: &TextRef,
    mode: ResolveMode,
    span: Option<(u32, u32)>,
    filter: Option<&str>,
    resolved: &mut Vec<String>,
) -> u32 {
    let (span_start, span_end) = span.unwrap_or((0, u32::MAX));
    let (chunks, final_pilcrow) = if span.is_some() {
        (
            snapshot_range(story, txn, span_start, span_end),
            last_pilcrow(story, txn).map(|(index, _)| index),
        )
    } else {
        let chunks = snapshot(story, txn);
        let final_pilcrow = chunks.iter().rev().find_map(|chunk| match chunk.kind {
            ChunkKind::Pilcrow(_) => Some(chunk.start),
            _ => None,
        });
        (chunks, final_pilcrow)
    };
    let mut removed = 0;
    // Reverse walk so physical removals never shift the indices still to be visited.
    for (position, chunk) in chunks.iter().enumerate().rev() {
        let overlap_start = chunk.start.max(span_start);
        let overlap_end = chunk.end().min(span_end);
        if overlap_end <= overlap_start {
            continue;
        }
        match &chunk.kind {
            ChunkKind::Pilcrow(map) => {
                resolve_paragraph_property_changes(txn, map, mode, filter, resolved);
                // A suggested split/merge stamps BOTH the pilcrow unit's text attr and the
                // pPr marker; either signal (matching the filter) selects the mark.
                let ppr_ins = active_stamp(map_stamp(map, txn, PPR_INS), filter);
                let ppr_del = active_stamp(map_stamp(map, txn, PPR_DEL), filter);
                let attr_ins = active_stamp(chunk.attrs.get(INS).cloned(), filter);
                let attr_del = active_stamp(chunk.attrs.get(DEL).cloned(), filter);
                let ins_hit = ppr_ins.is_some() || attr_ins.is_some();
                let del_hit = ppr_del.is_some() || attr_del.is_some();
                let join = match mode {
                    ResolveMode::Accept => del_hit,
                    ResolveMode::Reject => ins_hit,
                };
                if join {
                    match mode {
                        ResolveMode::Accept => {
                            record(resolved, ppr_del.as_ref());
                            record(resolved, attr_del.as_ref());
                        }
                        ResolveMode::Reject => {
                            record(resolved, ppr_ins.as_ref());
                            record(resolved, attr_ins.as_ref());
                        }
                    }
                    let keep = opens_with_block(story, txn, chunk.start + 1)
                        && holds_content(&chunks[..position], txn, mode, filter).unwrap_or_else(
                            || {
                                // The paragraph starts before a range resolve's span.
                                span_start > 0
                                    && !snapshot_range(story, txn, span_start - 1, span_start)
                                        .first()
                                        .is_some_and(|unit| {
                                            matches!(unit.kind, ChunkKind::Pilcrow(_))
                                        })
                            },
                        );
                    if Some(chunk.start) == final_pilcrow || keep {
                        // The final paragraph mark can never be removed, nor one
                        // that keeps content out of a block's slot — clear instead.
                        let (ppr_key, attr_key) = match mode {
                            ResolveMode::Accept => (PPR_DEL, DEL),
                            ResolveMode::Reject => (PPR_INS, INS),
                        };
                        map.remove(txn, ppr_key);
                        clear_attr(txn, story, chunk.start, 1, attr_key);
                    } else {
                        story.remove_range(txn, chunk.start, 1);
                        removed += 1;
                    }
                } else {
                    match mode {
                        ResolveMode::Accept if ins_hit => {
                            record(resolved, ppr_ins.as_ref());
                            record(resolved, attr_ins.as_ref());
                            map.remove(txn, PPR_INS);
                            clear_attr(txn, story, chunk.start, 1, INS);
                        }
                        ResolveMode::Reject if del_hit => {
                            record(resolved, ppr_del.as_ref());
                            record(resolved, attr_del.as_ref());
                            map.remove(txn, PPR_DEL);
                            clear_attr(txn, story, chunk.start, 1, DEL);
                        }
                        _ => {}
                    }
                }
            }
            ChunkKind::Text(_) | ChunkKind::Embed(_) => {
                let ins = active_stamp(chunk.attrs.get(INS).cloned(), filter);
                let del = active_stamp(chunk.attrs.get(DEL).cloned(), filter);
                // A unit carrying BOTH stamps is removed in either mode: the
                // remove class wins over the keep class on the same content.
                let remove = match mode {
                    ResolveMode::Accept => del.is_some(),
                    ResolveMode::Reject => ins.is_some(),
                };
                if remove {
                    match mode {
                        ResolveMode::Accept => record(resolved, del.as_ref()),
                        ResolveMode::Reject => record(resolved, ins.as_ref()),
                    }
                    story.remove_range(txn, overlap_start, overlap_end - overlap_start);
                    removed += overlap_end - overlap_start;
                } else {
                    match mode {
                        ResolveMode::Accept if ins.is_some() => {
                            record(resolved, ins.as_ref());
                            clear_attr(txn, story, overlap_start, overlap_end - overlap_start, INS);
                        }
                        ResolveMode::Reject if del.is_some() => {
                            record(resolved, del.as_ref());
                            clear_attr(txn, story, overlap_start, overlap_end - overlap_start, DEL);
                        }
                        _ => {}
                    }
                }
            }
        }
    }
    removed
}

/// Every story, by id.
fn sorted_stories<T: ReadTxn>(txn: &T) -> Vec<(String, TextRef)> {
    let Some(stories) = txn.get_map(crate::STORIES) else {
        return Vec::new();
    };
    let mut ids: Vec<String> = stories.keys(txn).map(|key| key.to_string()).collect();
    ids.sort();
    ids.into_iter()
        .filter_map(|story_id| match stories.get(txn, &story_id) {
            Some(Out::YText(story)) => Some((story_id, story)),
            _ => None,
        })
        .collect()
}

impl EditingDoc {
    /// Accepts the targeted changes: pending insertions become plain content
    /// and pending deletions are carried out. See the module docs for the full
    /// matrix and the join rule.
    ///
    /// The receipt's `revision_ids` lists the ids resolved, deduplicated and
    /// in resolution order; a range target also echoes the surviving range.
    /// A [`ChangeTarget::Range`] errors when empty, a
    /// [`ChangeTarget::Revision`] when the id matches nothing.
    pub fn accept_change(&self, ctx: &EditCtx, target: &ChangeTarget) -> OpResult<Receipt> {
        self.resolve_change(ctx, target, ResolveMode::Accept)
    }

    /// Rejects the targeted changes — the inverse of
    /// [`EditingDoc::accept_change`]: pending insertions are rolled back and
    /// pending deletions restored to plain content. Same receipt and errors.
    pub fn reject_change(&self, ctx: &EditCtx, target: &ChangeTarget) -> OpResult<Receipt> {
        self.resolve_change(ctx, target, ResolveMode::Reject)
    }

    fn resolve_change(
        &self,
        ctx: &EditCtx,
        target: &ChangeTarget,
        mode: ResolveMode,
    ) -> OpResult<Receipt> {
        let mut txn = self.transact_for(ctx);
        let mut resolved: Vec<String> = Vec::new();
        match target {
            ChangeTarget::Range(range) => {
                let len = crate::format::range_len(range)?;
                if len == 0 {
                    return Err(OpError::EmptyRange);
                }
                let story = story_ref(&txn, &range.story)?;
                check_range(&story, &txn, range.start, len)?;
                resolve_table_row_revisions(
                    &mut txn,
                    &story,
                    &range.story,
                    mode == ResolveMode::Accept,
                    Some((range.start, range.end)),
                    None,
                    &mut resolved,
                )?;
                let removed = resolve_story(
                    &mut txn,
                    &story,
                    mode,
                    Some((range.start, range.end)),
                    None,
                    &mut resolved,
                );
                let loc_range =
                    loc_range_in_txn(&range.story, &story, &txn, range.start, range.end - removed)?;
                Ok(Receipt {
                    new_para_ids: Vec::new(),
                    revision_ids: resolved,
                    range: Some(loc_range),
                })
            }
            ChangeTarget::All => {
                for (story_id, story) in &sorted_stories(&txn) {
                    resolve_table_row_revisions(
                        &mut txn,
                        story,
                        story_id,
                        mode == ResolveMode::Accept,
                        None,
                        None,
                        &mut resolved,
                    )?;
                    resolve_story(&mut txn, story, mode, None, None, &mut resolved);
                    resolve_field_changes(&mut txn, story, mode == ResolveMode::Accept);
                }
                Ok(Receipt {
                    new_para_ids: Vec::new(),
                    revision_ids: resolved,
                    range: None,
                })
            }
            ChangeTarget::Revision(revision_id) => {
                for (story_id, story) in &sorted_stories(&txn) {
                    resolve_table_row_revisions(
                        &mut txn,
                        story,
                        story_id,
                        mode == ResolveMode::Accept,
                        None,
                        Some(revision_id.as_str()),
                        &mut resolved,
                    )?;
                    resolve_story(
                        &mut txn,
                        story,
                        mode,
                        None,
                        Some(revision_id.as_str()),
                        &mut resolved,
                    );
                }
                if resolved.is_empty() {
                    return Err(OpError::UnknownChange(revision_id.clone()));
                }
                Ok(Receipt {
                    new_para_ids: Vec::new(),
                    revision_ids: resolved,
                    range: None,
                })
            }
        }
    }
}

/// Resolves the tracked changes the fields in `story` keep as markup, in
/// field embeds and in the fields inline content controls hold.
fn resolve_field_changes(txn: &mut TransactionMut<'_>, story: &TextRef, accept: bool) {
    for chunk in snapshot(story, txn) {
        let ChunkKind::Embed(Some(map)) = &chunk.kind else {
            continue;
        };
        let Some(Out::Any(Any::String(kind))) = map.get(txn, KIND_KEY) else {
            continue;
        };
        match kind.as_ref() {
            "field" => {
                let projected = map.get(txn, "resultProjection").is_some();
                if let Some(Out::Any(Any::String(data))) = map.get(txn, "fieldData")
                    && let Some(resolved) = resolved_field(&data, projected, accept)
                {
                    for (key, value) in resolved {
                        map.insert(txn, key, value);
                    }
                }
            }
            "sdt" => {
                if let Some(Out::Any(content)) = map.get(txn, "content")
                    && let Some(content) = resolved_sdt_content(&content, accept)
                {
                    map.insert(txn, "content", content);
                }
            }
            _ => {}
        }
    }
}

/// An inline content control's content with its fields' kept changes resolved.
fn resolved_sdt_content(content: &Any, accept: bool) -> Option<Any> {
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
            let mut payload = payload.as_ref().clone();
            let updated = match entry.get("kind") {
                Some(Any::String(kind)) if kind.as_ref() == "field" => {
                    match payload.get("fieldData") {
                        Some(Any::String(data)) => {
                            let projected = payload.contains_key("resultProjection");
                            resolved_field(data, projected, accept).map(|resolved| {
                                payload.extend(
                                    resolved
                                        .into_iter()
                                        .map(|(key, value)| (key.to_owned(), value)),
                                )
                            })
                        }
                        _ => None,
                    }
                }
                Some(Any::String(kind)) if kind.as_ref() == "sdt" => payload
                    .get("content")
                    .and_then(|inner| resolved_sdt_content(inner, accept))
                    .map(|inner| {
                        payload.insert("content".to_owned(), inner);
                    }),
                _ => None,
            };
            if updated.is_none() {
                return item.clone();
            }
            changed = true;
            let mut entry = entry.as_ref().clone();
            entry.insert("payload".to_owned(), Any::Map(Arc::new(payload)));
            Any::Map(Arc::new(entry))
        })
        .collect();
    changed.then(|| Any::Array(Arc::from(items)))
}

/// A field embed's payload entries once the changes its data keeps as markup
/// resolve; `None` when it keeps none. A `projected` field shows only the
/// result its projected children leave (`field_to_units` in the seed).
fn resolved_field(data: &str, projected: bool, accept: bool) -> Option<Vec<(&'static str, Any)>> {
    use serde_json::Value;
    let mut field: Value = serde_json::from_str(data).ok()?;
    if !resolve_kept_changes(&mut field, accept) {
        return None;
    }
    let runs = if field["type"] == "simpleField" {
        field["content"].as_array().cloned().unwrap_or_default()
    } else if projected {
        let visible: Vec<Value> = field["structuredResult"]["inline"]
            .as_array()
            .into_iter()
            .flatten()
            .filter(|child| matches!(child["type"].as_str(), Some("run" | "rawXml" | "inlineSdt")))
            .cloned()
            .collect();
        crate::seed::shown_runs(&visible)
    } else {
        field["fieldResult"].as_array().cloned().unwrap_or_default()
    };
    let display: String = runs
        .iter()
        .flat_map(|run| run["content"].as_array().into_iter().flatten())
        .filter(|content| content["type"] == "text")
        .filter_map(|content| content["text"].as_str())
        .collect();
    let text = |key: &str| Any::String(Arc::from(field[key].as_str().unwrap_or_default()));
    Some(vec![
        ("instruction", text("instruction")),
        ("fieldType", text("fieldType")),
        ("hasCachedResult", Any::Bool(!display.is_empty())),
        ("displayText", Any::String(Arc::from(display))),
        ("fieldData", Any::String(Arc::from(field.to_string()))),
    ])
}

/// Resolves the raw tracked changes under `value`, refreshing what each field
/// they sit in shows; returns whether any resolved.
fn resolve_kept_changes(value: &mut serde_json::Value, accept: bool) -> bool {
    use serde_json::Value;
    match value {
        Value::Array(items) => {
            let mut changed = false;
            let mut resolved = Vec::with_capacity(items.len());
            for mut item in items.drain(..) {
                let raw = (item["type"] == "rawXml")
                    .then(|| item["xml"].as_str())
                    .flatten()
                    .and_then(|xml| docx_parse::paragraph::resolve_raw_changes(xml, accept));
                if let Some(fragments) = raw {
                    changed = true;
                    resolved.extend(fragments.into_iter().map(|xml| {
                        let shown: Vec<Value> = docx_parse::paragraph::shown_raw_runs(&xml)
                            .iter()
                            .filter_map(|run| serde_json::to_value(run).ok())
                            .collect();
                        serde_json::json!({ "type": "rawXml", "xml": xml, "shown": shown })
                    }));
                } else {
                    changed |= resolve_kept_changes(&mut item, accept);
                    resolved.push(item);
                }
            }
            *items = resolved;
            changed
        }
        Value::Object(entries) => {
            let changed: Vec<String> = entries
                .iter_mut()
                .filter_map(|(key, value)| resolve_kept_changes(value, accept).then(|| key.clone()))
                .collect();
            let field_type = entries
                .get("type")
                .and_then(Value::as_str)
                .unwrap_or_default();
            if matches!(field_type, "simpleField" | "complexField") {
                refresh_field(entries, &changed);
            }
            !changed.is_empty()
        }
        _ => false,
    }
}

/// Recomputes the runs a field shows from its structured code and result.
fn refresh_field(field: &mut serde_json::Map<String, serde_json::Value>, changed: &[String]) {
    use serde_json::Value;
    let shown = |field: &serde_json::Map<String, Value>, key: &str| {
        field
            .get(key)
            .and_then(|content| content["inline"].as_array())
            .map(|nodes| Value::Array(crate::seed::shown_runs(nodes)))
    };
    if changed.iter().any(|key| key == "structuredResult")
        && let Some(runs) = shown(field, "structuredResult")
    {
        let key = if field["type"] == "simpleField" {
            "content"
        } else {
            "fieldResult"
        };
        field.insert(key.to_owned(), runs);
    }
    if changed.iter().any(|key| key == "structuredCode")
        && let Some(runs) = shown(field, "structuredCode")
    {
        let instruction: String = runs
            .as_array()
            .into_iter()
            .flatten()
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
        field.insert("fieldCode".to_owned(), runs);
    }
}
