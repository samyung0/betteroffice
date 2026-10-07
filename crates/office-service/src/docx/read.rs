//! The session reads the TS made through the docx-edit WASM facade, as the
//! same JSON values (`crates/docx-edit/src/wasm.rs` builds them).

use std::collections::HashMap;

use docx_edit::{EditingDoc, SegmentContent};
use serde_json::{Value, json};
use yrs::types::text::YChange;
use yrs::{Map, Out, ReadTxn, Transact};

use crate::error::{Error, Result};
use crate::jsv::V;

fn attrs(values: &std::collections::BTreeMap<String, yrs::Any>) -> Result<Value> {
    serde_json::to_value(values).map_err(Error::engine)
}

/// `session.storyIds()`: every story id, sorted.
pub(crate) fn story_ids(doc: &EditingDoc) -> Vec<String> {
    let txn = doc.yrs_doc().transact();
    let Some(stories) = txn.get_map("stories") else {
        return Vec::new();
    };
    let mut ids: Vec<String> = stories.iter(&txn).map(|(id, _)| id.to_string()).collect();
    ids.sort();
    ids
}

/// `session.storySegments(story)`.
pub(crate) fn story_segments(doc: &EditingDoc, story: &str) -> Result<Vec<V>> {
    segment_values(doc.story_segments(story).map_err(Error::engine)?)
}

/// `story_segments` of every story by id, reading the bookmarks once.
pub(crate) fn all_story_segments(doc: &EditingDoc) -> Result<HashMap<String, Vec<V>>> {
    doc.all_story_segments()
        .into_iter()
        .map(|(story, segments)| Ok((story, segment_values(segments)?)))
        .collect()
}

fn segment_values(segments: Vec<docx_edit::StorySegment>) -> Result<Vec<V>> {
    segments
        .into_iter()
        .map(|segment| {
            let attributes = attrs(&segment.attributes)?;
            let value = match segment.content {
                SegmentContent::Text(text) => {
                    json!({ "kind": "text", "text": text, "attributes": attributes })
                }
                SegmentContent::Pilcrow(properties) => json!({
                    "kind": "pilcrow",
                    "paraId": properties.para_id,
                    "properties": attrs(&properties.values)?,
                    "attributes": attributes,
                }),
                SegmentContent::OtherEmbed { kind, payload } => json!({
                    "kind": "embed",
                    "embedKind": kind,
                    "payload": attrs(&payload)?,
                    "attributes": attributes,
                }),
            };
            Ok(V::from(&value))
        })
        .collect()
}

/// `session.storyObjectIds(story)`: the yrs id of each embedded map.
pub(crate) fn story_object_ids(doc: &EditingDoc, story: &str) -> Result<Vec<String>> {
    let txn = doc.yrs_doc().transact();
    let text = txn
        .get_map("stories")
        .and_then(|stories| stories.get(&txn, story))
        .and_then(|value| value.cast::<yrs::TextRef>().ok())
        .ok_or_else(|| Error::engine(docx_edit::EditError::StoryNotFound(story.to_owned())))?;
    Ok(yrs::types::text::Text::diff(&text, &txn, YChange::identity)
        .into_iter()
        .filter_map(|diff| match diff.insert {
            Out::YMap(map) => Some(format!("{:?}", map.as_ref().id())),
            _ => None,
        })
        .collect())
}

/// `session.paragraphs(story)`.
pub(crate) fn paragraphs(doc: &EditingDoc, story: &str) -> Result<Vec<V>> {
    doc.paragraphs(story)
        .map_err(Error::engine)?
        .into_iter()
        .map(|paragraph| {
            Ok(V::from(&json!({
                "paraId": paragraph.para_id,
                "text": paragraph.text,
                "properties": attrs(&paragraph.properties)?,
            })))
        })
        .collect()
}

/// `session.listComments()`.
pub(crate) fn list_comments(doc: &EditingDoc) -> Result<Vec<V>> {
    let comments = doc.list_comments().map_err(Error::engine)?;
    Ok(comments
        .into_iter()
        .map(|comment| {
            V::from(&json!({
                "id": comment.id, "author": comment.author, "date": comment.date,
                "done": comment.done, "parentId": comment.parent_id, "body": comment.body,
            }))
        })
        .collect())
}

/// `session.resolveComment(id)`.
pub(crate) fn resolve_comment(doc: &EditingDoc, id: &str) -> Result<Vec<V>> {
    Ok(doc
        .resolve_comment(id)
        .map_err(Error::engine)?
        .into_iter()
        .map(|anchor| {
            V::from(&json!({ "story": anchor.story, "start": anchor.start, "end": anchor.end }))
        })
        .collect())
}
