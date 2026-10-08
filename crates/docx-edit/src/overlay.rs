//! Override layout (DOCX spike, schema 2): the room holds only the source
//! blocks users touched. A story enters the room as its map entry plus one
//! deleted anchor per block boundary, written by the story's own writer; a
//! block's units (a chunk) enter between their two anchors, written by the
//! chunk's writer, the first time an edit needs them. Writers derive from the
//! source fingerprint, so every peer copying an entry writes the same items
//! and concurrent first edits converge on one copy.
//!
//! A session here materializes every chunk locally (the engine reads and
//! edits a whole document as today) and shares only the copies its updates
//! reference.

use std::collections::{BTreeSet, HashMap, HashSet};
use std::sync::{Arc, Mutex};

use sha2::{Digest, Sha256};
use yrs::block::UpdateParent;
use yrs::updates::decoder::Decode;
use yrs::updates::encoder::Encode;
use yrs::{
    Any, Assoc, ClientID, Doc, IdSet, IndexedSequence, Map, MapPrelim, OffsetKind, Options, Out,
    ReadTxn, StateVector, Text, TextPrelim, Transact, Update, UpdateBlock,
};

use crate::{COMMENTS, EditingDoc, RawOp, STORIES, anchor_value, bookmarks, deterministic};

/// Version of the copy function; it enters every writer and the meta seed.
pub const COPY_VERSION: &str = "docx-override/1";
/// The layout's own root: schema, fingerprint and copy version.
pub const META: &str = "docx:meta";
/// The deck schema of the override layout.
pub const SCHEMA: f64 = 2.0;
/// Copy writers live in [2^52, 2^53 − 1); sessions stay below.
pub const RESERVED_CLIENTS: u64 = 1 << 52;
/// Origin of the transaction that materializes copies; nothing tracks it.
pub const COPY_ORIGIN: &str = "docx:copy";
const ANCHOR: char = '\u{2060}';

/// The writer of `kind`/`key` for the source with `fingerprint`.
pub fn writer(fingerprint: &str, kind: &str, key: &str) -> u64 {
    let digest = Sha256::new()
        .chain_update(COPY_VERSION.as_bytes())
        .chain_update(b"\0")
        .chain_update(fingerprint.as_bytes())
        .chain_update(b"\0")
        .chain_update(kind.as_bytes())
        .chain_update(b"\0")
        .chain_update(key.as_bytes())
        .finalize();
    let mut bytes = [0_u8; 8];
    bytes.copy_from_slice(&digest[..8]);
    RESERVED_CLIENTS + u64::from_le_bytes(bytes) % (RESERVED_CLIENTS - 1)
}

/// SHA-256 hex of the source bytes.
pub fn fingerprint(base: &[u8]) -> String {
    Sha256::digest(base)
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

pub(crate) struct SourceBookmark {
    pub key: String,
    pub data: Any,
    /// Index relative to the chunk's start.
    pub index: u32,
    pub assoc: Assoc,
}

pub(crate) struct SourceChunk {
    /// Offset of the chunk's first unit in its story.
    pub start: u32,
    pub len: u32,
    /// The seed's insertions of this chunk, from index 0.
    pub ops: Vec<RawOp>,
    pub bookmarks: Vec<SourceBookmark>,
}

#[allow(dead_code)]
pub(crate) struct SourceStory {
    pub id: String,
    pub chunks: Vec<SourceChunk>,
    pub len: u32,
}

pub(crate) struct SourceComment {
    pub id: String,
    /// The coverage the seed anchors, story by story in seed order.
    pub ranges: Vec<(String, Vec<(u32, u32)>)>,
    /// The seed's `PatchComment` fields, if any.
    pub fields: Option<Vec<(String, Any)>>,
}

/// What the seed writes, cut into copy entries.
pub(crate) struct Source {
    pub stories: Vec<SourceStory>,
    pub comments: Vec<SourceComment>,
    pub referenced_fonts: Vec<String>,
}

#[allow(dead_code)]
#[derive(Clone, Copy, Debug)]
pub(crate) enum Role {
    Story(usize),
    Chunk(usize, usize),
    Comment(usize),
}

/// One story of the source, by entry.
pub struct StoryEntry {
    pub id: String,
    pub writer: u64,
    pub chunk_writers: Vec<u64>,
    pub chunk_starts: Vec<u32>,
    pub chunk_lens: Vec<u32>,
}

/// A session's view of the override layout: every writer's role, and which
/// copies the room already holds.
pub struct Overlay {
    pub fingerprint: String,
    pub stories: Vec<StoryEntry>,
    pub comment_writers: Vec<(String, u64)>,
    roles: HashMap<u64, Role>,
    shared: Mutex<HashSet<u64>>,
    /// Each story's copy as first written (entry and one deleted anchor run):
    /// a materialized session splits the run at every chunk.
    story_copies: Mutex<Vec<Vec<u8>>>,
}

impl Overlay {
    fn new(fingerprint: &str, source: &Source) -> Result<Self, String> {
        let mut roles = HashMap::new();
        let mut claim = |writer: u64, role: Role| -> Result<u64, String> {
            if roles.insert(writer, role).is_some() {
                return Err(format!("copy writer {writer} collides ({role:?})"));
            }
            Ok(writer)
        };
        let mut stories = Vec::with_capacity(source.stories.len());
        for (story_index, story) in source.stories.iter().enumerate() {
            let story_writer = claim(
                writer(fingerprint, "story", &story.id),
                Role::Story(story_index),
            )?;
            let mut chunk_writers = Vec::with_capacity(story.chunks.len());
            for chunk_index in 0..story.chunks.len() {
                chunk_writers.push(claim(
                    writer(fingerprint, "chunk", &format!("{}#{chunk_index}", story.id)),
                    Role::Chunk(story_index, chunk_index),
                )?);
            }
            stories.push(StoryEntry {
                id: story.id.clone(),
                writer: story_writer,
                chunk_writers,
                chunk_starts: story.chunks.iter().map(|chunk| chunk.start).collect(),
                chunk_lens: story.chunks.iter().map(|chunk| chunk.len).collect(),
            });
        }
        let mut comment_writers = Vec::with_capacity(source.comments.len());
        for (index, comment) in source.comments.iter().enumerate() {
            comment_writers.push((
                comment.id.clone(),
                claim(
                    writer(fingerprint, "comment", &comment.id),
                    Role::Comment(index),
                )?,
            ));
        }
        Ok(Self {
            fingerprint: fingerprint.to_owned(),
            stories,
            comment_writers,
            roles,
            shared: Mutex::new(HashSet::new()),
            story_copies: Mutex::new(Vec::new()),
        })
    }

    /// The story and chunk a chunk writer copies, in source order.
    pub fn chunk_of(&self, writer: u64) -> Option<(usize, usize)> {
        match self.roles.get(&writer) {
            Some(Role::Chunk(story, chunk)) => Some((*story, *chunk)),
            _ => None,
        }
    }

    /// Writers whose copies the room holds.
    pub fn shared(&self) -> HashSet<u64> {
        self.shared.lock().unwrap().clone()
    }

    /// Records the copies an update from the room carries.
    pub fn note_incoming(&self, update: &[u8]) -> Result<(), String> {
        let update = Update::decode_v1(update).map_err(|error| error.to_string())?;
        let mut shared = self.shared.lock().unwrap();
        for block in update.blocks() {
            let client = match block {
                UpdateBlock::Item(item) => item.id().client,
                UpdateBlock::Gc(range) | UpdateBlock::Skip(range) => range.client,
            };
            if self.roles.contains_key(&client.get()) {
                shared.insert(client.get());
            }
        }
        Ok(())
    }

    /// `update` (a local transaction's) preceded by the copies it needs that
    /// the room lacks: every reserved writer it names as an origin, parent or
    /// deletion, with the story of each chunk. Each copy is its writer's
    /// structs as `txn` holds them, with their deletions.
    pub fn augment<T: ReadTxn>(&self, txn: &T, update: &[u8]) -> Result<Vec<u8>, String> {
        let decoded = Update::decode_v1(update).map_err(|error| error.to_string())?;
        let mut named = BTreeSet::new();
        let mut note = |client: &ClientID| {
            if client.get() >= RESERVED_CLIENTS {
                named.insert(client.get());
            }
        };
        for block in decoded.blocks() {
            if let UpdateBlock::Item(item) = block {
                if let Some(id) = item.origin() {
                    note(&id.client);
                }
                if let Some(id) = item.right_origin() {
                    note(&id.client);
                }
                if let UpdateParent::Item(id) = item.update_parent() {
                    note(&id.client);
                }
            }
        }
        for (client, _) in decoded.delete_set().iter() {
            note(client);
        }
        let mut needed = BTreeSet::new();
        for writer in named {
            match self.roles.get(&writer) {
                Some(Role::Chunk(story, _)) => {
                    needed.insert(writer);
                    needed.insert(self.stories[*story].writer);
                }
                Some(_) => {
                    needed.insert(writer);
                }
                None => return Err(format!("update names unknown copy writer {writer}")),
            }
        }
        {
            let mut shared = self.shared.lock().unwrap();
            needed.retain(|writer| !shared.contains(writer));
            if needed.is_empty() {
                return Ok(update.to_vec());
            }
            shared.extend(needed.iter().copied());
        }
        // Stories go as first written; chunks and comments as this session
        // holds them, so a deleted copy travels without its content.
        let mut parts = Vec::new();
        {
            let story_copies = self.story_copies.lock().unwrap();
            needed.retain(|writer| match self.roles.get(writer) {
                Some(Role::Story(story)) if *story < story_copies.len() => {
                    parts.push(story_copies[*story].clone());
                    false
                }
                _ => true,
            });
        }
        let mut updates = parts
            .iter()
            .map(|bytes| Update::decode_v1(bytes).map_err(|error| error.to_string()))
            .collect::<Result<Vec<_>, _>>()?;
        if !needed.is_empty() {
            let local = txn.state_vector();
            let mut sv = StateVector::default();
            for (client, clock) in local.iter() {
                if !needed.contains(&client.get()) {
                    sv.set_max(*client, *clock);
                }
            }
            let copies = deterministic::encode_state_as_update_v1(txn, &sv);
            let mut copies = Update::decode_v1(&copies).map_err(|error| error.to_string())?;
            let keep = IdSet::from_iter(needed.iter().map(|writer| {
                let client = ClientID::new(*writer);
                (client, [0..local.get(&client)])
            }));
            copies.delete_set_mut().intersect_with(&keep);
            updates.push(copies);
        }
        updates.push(decoded);
        Ok(Update::merge_updates(updates).encode_v1())
    }
}

fn scratch(client: u64) -> Doc {
    let mut options = Options::with_client_id(ClientID::new(client));
    options.offset_kind = OffsetKind::Utf16;
    let doc = Doc::with_options(options);
    doc.get_or_insert_map(STORIES);
    doc.get_or_insert_map(COMMENTS);
    doc.get_or_insert_map(bookmarks::ROOT);
    doc.get_or_insert_map(META);
    doc
}

fn encode_since(doc: &Doc, since: &StateVector) -> Vec<u8> {
    deterministic::encode_state_as_update_v1(&doc.transact(), since)
}

/// The layout's shared seed: its meta root only, a function of the source
/// fingerprint.
pub fn meta_seed(fingerprint: &str) -> Vec<u8> {
    let doc = scratch(crate::seed::SEED_CLIENT_ID);
    {
        let mut txn = doc.transact_mut();
        let meta = txn.get_map(META).expect("declared");
        meta.insert(&mut txn, "schemaVersion", SCHEMA);
        meta.insert(&mut txn, "copyVersion", COPY_VERSION);
        meta.insert(&mut txn, "fingerprint", fingerprint);
    }
    encode_since(&doc, &StateVector::default())
}

/// A story's entry and anchors, live (for building chunk copies) or deleted
/// (as the room holds them).
fn story_update(story: &SourceStory, writer: u64, live: bool) -> Vec<u8> {
    let doc = scratch(writer);
    {
        let mut txn = doc.transact_mut();
        let stories = txn.get_map(STORIES).expect("declared");
        let text = stories.insert(&mut txn, story.id.as_str(), TextPrelim::new(""));
        let anchors = story.chunks.len() as u32 + 1;
        let run: String = std::iter::repeat_n(ANCHOR, anchors as usize).collect();
        text.insert(&mut txn, 0, &run);
        if !live {
            text.remove_range(&mut txn, 0, anchors);
        }
    }
    encode_since(&doc, &StateVector::default())
}

fn shifted(op: &RawOp, shift: u32) -> RawOp {
    let mut op = op.clone();
    match &mut op {
        RawOp::Insert { index, .. }
        | RawOp::InsertEmbed { index, .. }
        | RawOp::Delete { index, .. }
        | RawOp::Format { index, .. }
        | RawOp::SetEmbedAttr { index, .. }
        | RawOp::SetBookmark { index, .. } => *index += shift,
        RawOp::SetComment { .. } | RawOp::PatchComment { .. } | RawOp::RemoveComment { .. } => {}
    }
    op
}

/// Chunk `index` of `story` as its writer copies it: the seed's insertions
/// between anchors `index` and `index + 1`, then its bookmarks.
fn chunk_update(
    story: &SourceStory,
    live_story: &[u8],
    index: usize,
    writer: u64,
) -> Result<Vec<u8>, String> {
    let doc = scratch(writer);
    doc.transact_mut()
        .apply_update(Update::decode_v1(live_story).map_err(|error| error.to_string())?)
        .map_err(|error| error.to_string())?;
    let before = doc.transact().state_vector();
    let chunk = &story.chunks[index];
    let shift = index as u32 + 1;
    {
        let mut txn = doc.transact_mut();
        let ops = chunk.ops.iter().map(|op| shifted(op, shift)).collect();
        crate::raw::apply_raw_ops_to_story(&mut txn, &story.id, ops, true)
            .map_err(|error| error.to_string())?;
        if !chunk.bookmarks.is_empty() {
            let text = crate::story_ref(&txn, &story.id).map_err(|error| error.to_string())?;
            let root = txn.get_map(bookmarks::ROOT).expect("declared");
            for bookmark in &chunk.bookmarks {
                let sticky = text
                    .sticky_index(&txn, shift + bookmark.index, bookmark.assoc)
                    .ok_or("bookmark position does not resolve")?;
                let entry = root.insert(&mut txn, bookmark.key.as_str(), MapPrelim::default());
                entry.insert(&mut txn, "data", bookmark.data.clone());
                entry.insert(
                    &mut txn,
                    "anchors",
                    Any::Array(Arc::from([anchor_value(&story.id, &sticky, &sticky)])),
                );
            }
        }
    }
    Ok(encode_since(&doc, &before))
}

/// A source comment as its writer copies it, its anchors taken in `local`
/// (a document holding every chunk the coverage names).
fn comment_update<T: ReadTxn>(
    local: &T,
    comment: &SourceComment,
    writer: u64,
) -> Result<Vec<u8>, String> {
    let doc = scratch(writer);
    {
        let mut txn = doc.transact_mut();
        let comments = txn.get_map(COMMENTS).expect("declared");
        for (story_id, ranges) in &comment.ranges {
            if ranges.is_empty() {
                return Err("at least one anchored range is required".into());
            }
            let story = crate::story_ref(local, story_id).map_err(|error| error.to_string())?;
            let len = story.len(local);
            let mut anchors = Vec::with_capacity(ranges.len());
            for &(start, end) in ranges {
                if end <= start || end > len {
                    return Err(format!("invalid comment range {start}..{end}"));
                }
                let from = story
                    .sticky_index(local, start, Assoc::After)
                    .ok_or("start anchor could not be made")?;
                let to = story
                    .sticky_index(local, end, Assoc::Before)
                    .ok_or("end anchor could not be made")?;
                anchors.push(anchor_value(story_id, &from, &to));
            }
            let map = comments.insert(&mut txn, comment.id.as_str(), MapPrelim::default());
            map.insert(&mut txn, "author", "");
            map.insert(&mut txn, "date", "");
            map.insert(&mut txn, "parentId", Any::Null);
            map.insert(&mut txn, "done", false);
            map.insert(&mut txn, "body", Any::Null);
            map.insert(&mut txn, "anchors", Any::Array(Arc::from(anchors)));
        }
        if let Some(fields) = &comment.fields {
            let map = match comments.get(&txn, &comment.id) {
                Some(Out::YMap(map)) => map,
                _ => comments.insert(&mut txn, comment.id.as_str(), MapPrelim::default()),
            };
            if map.get(&txn, "anchors").is_none() {
                map.insert(&mut txn, "anchors", Any::Array(Arc::from([])));
            }
            for (key, value) in fields {
                map.insert(&mut txn, key.as_str(), value.clone());
            }
        }
    }
    Ok(encode_since(&doc, &StateVector::default()))
}

fn apply_merged(doc: &Doc, updates: Vec<Vec<u8>>) -> Result<(), String> {
    let decoded = updates
        .iter()
        .map(|bytes| Update::decode_v1(bytes).map_err(|error| error.to_string()))
        .collect::<Result<Vec<_>, _>>()?;
    doc.transact_mut_with(COPY_ORIGIN)
        .apply_update(Update::merge_updates(decoded))
        .map_err(|error| error.to_string())
}

/// Every copy of the source, as the room would hold them all: the meta seed,
/// the stories and their chunks, then the comments.
pub(crate) fn materialize(doc: &EditingDoc, overlay: &Overlay, source: &Source) -> Result<(), String> {
    let mut updates = vec![meta_seed(&overlay.fingerprint)];
    let mut story_copies = Vec::with_capacity(source.stories.len());
    for (story, entry) in source.stories.iter().zip(&overlay.stories) {
        let copy = story_update(story, entry.writer, false);
        story_copies.push(copy.clone());
        updates.push(copy);
        let live = story_update(story, entry.writer, true);
        for (index, writer) in entry.chunk_writers.iter().enumerate() {
            updates.push(chunk_update(story, &live, index, *writer)?);
        }
    }
    apply_merged(doc.yrs_doc(), updates)?;
    *overlay.story_copies.lock().unwrap() = story_copies;
    let comments = {
        let txn = doc.yrs_doc().transact();
        source
            .comments
            .iter()
            .zip(&overlay.comment_writers)
            .map(|(comment, (_, writer))| comment_update(&txn, comment, *writer))
            .collect::<Result<Vec<_>, _>>()?
    };
    apply_merged(doc.yrs_doc(), comments)
}

/// Seeds `document` with the override layout's local view of `envelope`:
/// every chunk and comment materialized under its writer, none shared yet.
/// Returns the fonts the source references.
pub fn open_chunked(
    document: &EditingDoc,
    envelope: docx_parse::S9WireEnvelope,
    fingerprint: &str,
) -> Result<Vec<String>, String> {
    document.set_package(Some(crate::seed::PackageContext::new(&envelope)));
    let source = crate::seed::chunk_source(crate::seed::lower(envelope)?)?;
    let overlay = Overlay::new(fingerprint, &source)?;
    materialize(document, &overlay, &source)?;
    *document.overlay.lock().unwrap() = Some(Arc::new(overlay));
    Ok(source.referenced_fonts)
}

/// Whether `state` holds the override layout's meta root.
pub fn is_chunked_state(state: &[u8]) -> bool {
    let doc = scratch(1);
    let Ok(update) = Update::decode_v1(state) else {
        return false;
    };
    if doc.transact_mut().apply_update(update).is_err() {
        return false;
    }
    let txn = doc.transact();
    txn.get_map(META).is_some_and(|meta| {
        matches!(meta.get(&txn, "schemaVersion"), Some(Out::Any(Any::Number(schema))) if schema == SCHEMA)
    })
}

impl EditingDoc {
    /// The override layout this session runs, if any.
    pub fn overlay(&self) -> Option<Arc<Overlay>> {
        self.overlay.lock().unwrap().clone()
    }

    /// A peer's update as an editor applies it (`EditSession::apply_peer_update`):
    /// then fields a deleted range now ends at re-read their shown text and
    /// duplicated paragraph ids are renamed, as system edits.
    pub fn apply_peer_update_v1(&self, bytes: &[u8]) -> crate::EditResult<()> {
        if let Some(overlay) = self.overlay() {
            overlay
                .note_incoming(bytes)
                .map_err(crate::EditError::InvalidUpdate)?;
        }
        self.applying_peer_update(|| {
            crate::ops::field_changes::refreshing_fields(self.yrs_doc(), || {
                self.apply_update_v1(bytes)
            })
        })
    }

    /// Applies an update from the room, recording the copies it carries.
    pub fn apply_shared_update(&self, bytes: &[u8]) -> crate::EditResult<()> {
        if let Some(overlay) = self.overlay() {
            overlay
                .note_incoming(bytes)
                .map_err(crate::EditError::InvalidUpdate)?;
        }
        self.apply_update_v1(bytes)
    }

    /// Calls `callback` with what each local transaction sends the room: its
    /// update, preceded in the override layout by the copies it needs. An
    /// applied remote update sends nothing, but for the formatting cleanup
    /// yrs ran inside it (deletions of redundant format items), which
    /// `forward_cleanups` sends as its own update: today the runtime sends
    /// only local transactions, so each peer keeps its own cleanup.
    pub fn observe_shared_updates(
        &self,
        forward_cleanups: bool,
        callback: impl Fn(Vec<u8>) + 'static,
    ) -> crate::EditResult<yrs::Subscription> {
        let overlay = self.overlay();
        self.yrs_doc()
            .observe_update_v1(move |txn, event| {
                let update = if txn.origin().is_some() {
                    event.update.clone()
                } else if forward_cleanups && !txn.cleanups().is_empty() {
                    let mut encoder = yrs::updates::encoder::EncoderV1::new();
                    yrs::encoding::write::Write::write_var(&mut encoder, 0_u32);
                    txn.cleanups().encode(&mut encoder);
                    yrs::updates::encoder::Encoder::to_vec(encoder)
                } else {
                    return;
                };
                let update = match &overlay {
                    Some(overlay) => overlay
                        .augment(txn, &update)
                        .expect("a local update names only this source's copies"),
                    None => update,
                };
                callback(update);
            })
            .map_err(|error| crate::EditError::InvalidUpdate(error.to_string()))
    }
}

fn json(value: &Any) -> serde_json::Value {
    use serde_json::Value;
    match value {
        Any::Null | Any::Undefined => Value::Null,
        Any::Bool(value) => Value::Bool(*value),
        Any::Number(value) => serde_json::Number::from_f64(*value).map_or(Value::Null, Value::Number),
        Any::BigInt(value) => Value::from(*value),
        Any::String(value) => Value::String(value.to_string()),
        Any::Buffer(value) => Value::from(value.to_vec()),
        Any::Array(values) => Value::Array(values.iter().map(json).collect()),
        Any::Map(values) => Value::Object(
            values
                .iter()
                .map(|(key, value)| (key.clone(), json(value)))
                .collect::<std::collections::BTreeMap<_, _>>()
                .into_iter()
                .collect(),
        ),
    }
}

/// Everything a reader of the document sees, as canonical JSON lines: each
/// story's segments, each comment's fields and resolved ranges, each
/// bookmark's position. Two replicas reading the same dump read the same
/// document.
pub fn read_dump(doc: &EditingDoc) -> String {
    use std::fmt::Write;
    let mut out = String::new();
    let (story_ids, comment_ids) = {
        let txn = doc.yrs_doc().transact();
        let mut stories: Vec<String> = txn
            .get_map(STORIES)
            .map(|map| map.keys(&txn).map(str::to_owned).collect())
            .unwrap_or_default();
        stories.sort();
        let mut comments: Vec<String> = txn
            .get_map(COMMENTS)
            .map(|map| map.keys(&txn).map(str::to_owned).collect())
            .unwrap_or_default();
        comments.sort();
        (stories, comments)
    };
    for story in &story_ids {
        let _ = writeln!(out, "story {story}");
        for segment in doc.story_segments(story).unwrap_or_default() {
            let attributes: serde_json::Map<String, serde_json::Value> = segment
                .attributes
                .iter()
                .map(|(key, value)| (key.clone(), json(value)))
                .collect();
            let content = match segment.content {
                crate::SegmentContent::Text(text) => serde_json::json!({ "text": text }),
                crate::SegmentContent::Pilcrow(properties) => serde_json::json!({
                    "pilcrow": properties.para_id,
                    "values": properties
                        .values
                        .iter()
                        .map(|(key, value)| (key.clone(), json(value)))
                        .collect::<serde_json::Map<_, _>>(),
                }),
                crate::SegmentContent::OtherEmbed { kind, payload } => serde_json::json!({
                    "embed": kind,
                    "payload": payload
                        .iter()
                        .map(|(key, value)| (key.clone(), json(value)))
                        .collect::<serde_json::Map<_, _>>(),
                }),
            };
            let _ = writeln!(out, "  {content} {}", serde_json::Value::Object(attributes));
        }
    }
    for id in &comment_ids {
        let fields = {
            let txn = doc.yrs_doc().transact();
            let comments = txn.get_map(COMMENTS).expect("declared");
            match comments.get(&txn, id) {
                Some(Out::YMap(map)) => {
                    let mut fields: Vec<(String, serde_json::Value)> = map
                        .iter(&txn)
                        .filter(|(key, _)| *key != "anchors")
                        .map(|(key, value)| {
                            let value = match value {
                                Out::Any(value) => json(&value),
                                other => serde_json::Value::String(format!("{other:?}")),
                            };
                            (key.to_owned(), value)
                        })
                        .collect();
                    fields.sort_by(|a, b| a.0.cmp(&b.0));
                    serde_json::Value::Object(fields.into_iter().collect())
                }
                other => serde_json::Value::String(format!("{other:?}")),
            }
        };
        let ranges = match doc.resolve_comment(id) {
            Ok(ranges) => ranges
                .into_iter()
                .map(|range| format!("{}:{}..{}", range.story, range.start, range.end))
                .collect::<Vec<_>>()
                .join(","),
            Err(error) => format!("unresolved: {error}"),
        };
        let _ = writeln!(out, "comment {id} {fields} [{ranges}]");
    }
    let txn = doc.yrs_doc().transact();
    let mut markers: Vec<_> = bookmarks::positions_by_story(&txn).into_iter().collect();
    markers.sort_by(|a, b| a.0.cmp(&b.0));
    for (story, positions) in markers {
        for (at, data) in positions {
            let _ = writeln!(out, "bookmark {story} {at} {}", json(&data));
        }
    }
    out
}

/// The source as a reader without a materialized session sees it: each
/// chunk's segments, read once from the chunk alone.
pub struct LogicalBase {
    overlay: Overlay,
    chunks: Vec<Vec<Vec<crate::StorySegment>>>,
}

fn segments_of_text<T: ReadTxn>(txn: &T, story: &yrs::TextRef) -> Vec<crate::StorySegment> {
    use yrs::types::text::YChange;
    story
        .diff(txn, YChange::identity)
        .into_iter()
        .map(|diff| crate::StorySegment {
            content: crate::segment_content(diff.insert, txn),
            attributes: crate::ordered_attrs(diff.attributes.as_deref()),
        })
        .collect()
}

/// Joins adjacent text segments with equal attributes, as `story_segments`
/// reads them.
fn merged(segments: Vec<crate::StorySegment>) -> Vec<crate::StorySegment> {
    let mut out: Vec<crate::StorySegment> = Vec::with_capacity(segments.len());
    for segment in segments {
        if let (crate::SegmentContent::Text(text), Some(previous)) = (&segment.content, out.last_mut())
            && let crate::SegmentContent::Text(before) = &mut previous.content
            && previous.attributes == segment.attributes
        {
            before.push_str(text);
            continue;
        }
        out.push(segment);
    }
    out
}

fn split_text(segment: &crate::StorySegment, at: u32) -> (crate::StorySegment, crate::StorySegment) {
    let crate::SegmentContent::Text(text) = &segment.content else {
        unreachable!("only text spans more than one unit");
    };
    let mut units = 0;
    let mut byte = text.len();
    for (index, ch) in text.char_indices() {
        if units >= at {
            byte = index;
            break;
        }
        units += ch.len_utf16() as u32;
    }
    let part = |text: &str| crate::StorySegment {
        content: crate::SegmentContent::Text(text.to_owned()),
        attributes: segment.attributes.clone(),
    };
    (part(&text[..byte]), part(&text[byte..]))
}

fn segment_len(segment: &crate::StorySegment) -> u32 {
    match &segment.content {
        crate::SegmentContent::Text(text) => text.encode_utf16().count() as u32,
        _ => 1,
    }
}

impl LogicalBase {
    pub fn new(base: &[u8]) -> Result<Self, String> {
        let envelope = crate::seed::parse_docx_for_edit(base)?;
        let source = crate::seed::chunk_source(crate::seed::lower(envelope)?)?;
        let overlay = Overlay::new(&fingerprint(base), &source)?;
        let mut chunks = Vec::with_capacity(source.stories.len());
        for (story, entry) in source.stories.iter().zip(&overlay.stories) {
            let finished = story_update(story, entry.writer, false);
            let live = story_update(story, entry.writer, true);
            let mut segments = Vec::with_capacity(story.chunks.len());
            for (index, writer) in entry.chunk_writers.iter().enumerate() {
                let doc = scratch(1);
                apply_merged(&doc, vec![finished.clone(), chunk_update(story, &live, index, *writer)?])?;
                let txn = doc.transact();
                let text = crate::story_ref(&txn, &story.id).map_err(|error| error.to_string())?;
                segments.push(segments_of_text(&txn, &text));
            }
            chunks.push(segments);
        }
        Ok(Self { overlay, chunks })
    }

    /// Story ids a room holds: the source's (unless the room deleted them)
    /// and the room's own.
    pub fn story_ids<T: ReadTxn>(&self, room: &T) -> Vec<String> {
        let sv = room.state_vector();
        let stories = room.get_map(STORIES);
        let mut ids: BTreeSet<String> = self
            .overlay
            .stories
            .iter()
            .filter(|story| {
                sv.get(&ClientID::new(story.writer)) == 0
                    || stories.as_ref().is_some_and(|map| map.get(room, &story.id).is_some())
            })
            .map(|story| story.id.clone())
            .collect();
        if let Some(map) = stories {
            ids.extend(map.keys(room).map(str::to_owned));
        }
        ids.into_iter().collect()
    }

    /// `story_id`'s segments (without projected bookmarks) from `room`.
    pub fn story_segments<T: ReadTxn>(&self, room: &T, story_id: &str) -> Option<Vec<crate::StorySegment>> {
        let sv = room.state_vector();
        let position = self.overlay.stories.iter().position(|story| story.id == story_id);
        let text = crate::story_ref(room, story_id).ok();
        let Some(position) = position else {
            return text.map(|text| merged(segments_of_text(room, &text)));
        };
        let entry = &self.overlay.stories[position];
        let copied = |writer: u64| sv.get(&ClientID::new(writer)) > 0;
        if !copied(entry.writer) {
            return Some(merged(self.chunks[position].concat()));
        }
        let text = text?;
        // Visible offsets of the anchors in the room's story, and the
        // formatting a span around an uncopied chunk gives its content: a
        // format edit across chunks writes markers only at its ends.
        let anchors = entry.chunk_writers.len() + 1;
        let mut boundaries = vec![0u32; anchors];
        let mut active: Vec<std::collections::BTreeMap<String, Any>> = vec![Default::default(); anchors];
        let mut current_attrs = std::collections::BTreeMap::new();
        let mut offset = 0u32;
        let mut item = <yrs::TextRef as AsRef<yrs::branch::Branch>>::as_ref(&text).start();
        while let Some(current) = item {
            let id = current.id();
            if id.client.get() == entry.writer {
                for clock in id.clock..id.clock + current.len() {
                    if (1..=anchors as u32).contains(&clock) {
                        boundaries[clock as usize - 1] = offset;
                        active[clock as usize - 1] = current_attrs.clone();
                    }
                }
            } else if !current.is_deleted() {
                if let yrs::block::ItemContent::Format(key, value) = current.content() {
                    if **value == Any::Null {
                        current_attrs.remove(key.as_ref());
                    } else {
                        current_attrs.insert(key.to_string(), (**value).clone());
                    }
                } else if current.is_countable() {
                    offset += current.len();
                }
            }
            item = current.right();
        }
        // The room's visible segments, cut at the anchors.
        let mut gaps: Vec<Vec<crate::StorySegment>> = vec![Vec::new(); anchors + 1];
        let mut at = 0u32;
        let mut gap = 0usize;
        for mut segment in segments_of_text(room, &text) {
            loop {
                while gap < anchors && boundaries[gap] <= at {
                    gap += 1;
                }
                let len = segment_len(&segment);
                let limit = if gap < anchors { boundaries[gap] } else { u32::MAX };
                if at + len <= limit {
                    at += len;
                    gaps[gap].push(segment);
                    break;
                }
                let (head, tail) = split_text(&segment, limit - at);
                at = limit;
                gaps[gap].push(head);
                segment = tail;
            }
        }
        // Gap k + 1 lies between anchors k and k + 1: chunk k.
        let mut out = std::mem::take(&mut gaps[0]);
        for (index, writer) in entry.chunk_writers.iter().enumerate() {
            if copied(*writer) {
                out.append(&mut gaps[index + 1]);
            } else {
                debug_assert!(gaps[index + 1].is_empty(), "content in an uncopied chunk's gap");
                out.extend(self.chunks[position][index].iter().map(|segment| {
                    let mut segment = segment.clone();
                    for (key, value) in &active[index] {
                        segment
                            .attributes
                            .entry(key.clone())
                            .or_insert_with(|| value.clone());
                    }
                    segment
                }));
            }
        }
        out.append(&mut gaps[anchors]);
        Some(merged(out))
    }
}

/// Stories' segments as canonical JSON lines, without the bookmarks a
/// paragraph mark projects (what [`LogicalBase`] reads).
pub fn segments_dump(stories: &[(String, Vec<crate::StorySegment>)]) -> String {
    use std::fmt::Write;
    let mut out = String::new();
    for (story, segments) in stories {
        let _ = writeln!(out, "story {story}");
        for segment in segments {
            let attributes: serde_json::Map<String, serde_json::Value> = segment
                .attributes
                .iter()
                .map(|(key, value)| (key.clone(), json(value)))
                .collect();
            let content = match &segment.content {
                crate::SegmentContent::Text(text) => serde_json::json!({ "text": text }),
                crate::SegmentContent::Pilcrow(properties) => serde_json::json!({
                    "pilcrow": properties.para_id,
                    "values": properties
                        .values
                        .iter()
                        .filter(|(key, _)| key.as_str() != "bookmarks")
                        .map(|(key, value)| (key.clone(), json(value)))
                        .collect::<serde_json::Map<_, _>>(),
                }),
                crate::SegmentContent::OtherEmbed { kind, payload } => serde_json::json!({
                    "embed": kind,
                    "payload": payload
                        .iter()
                        .map(|(key, value)| (key.clone(), json(value)))
                        .collect::<serde_json::Map<_, _>>(),
                }),
            };
            let _ = writeln!(out, "  {content} {}", serde_json::Value::Object(attributes));
        }
    }
    out
}

/// [`segments_dump`] of a session's every story.
pub fn session_segments_dump(doc: &EditingDoc) -> String {
    let mut ids: Vec<String> = doc.all_story_segments().into_keys().collect();
    ids.sort();
    let stories: Vec<_> = ids
        .into_iter()
        .map(|id| {
            let segments = doc.story_segments(&id).unwrap_or_default();
            (id, segments)
        })
        .collect();
    segments_dump(&stories)
}

/// [`segments_dump`] of a room read through [`LogicalBase`].
pub fn room_segments_dump(base: &LogicalBase, room: &Doc) -> String {
    let txn = room.transact();
    let stories: Vec<_> = base
        .story_ids(&txn)
        .into_iter()
        .filter_map(|id| base.story_segments(&txn, &id).map(|segments| (id, segments)))
        .collect();
    segments_dump(&stories)
}

/// The source lowered and cut into chunks without writing any Yrs: what a
/// session that reads uncopied chunks from the source pays at open (spike
/// measure). Returns the number of chunks.
pub fn lower_only(envelope: docx_parse::S9WireEnvelope) -> Result<usize, String> {
    let source = crate::seed::chunk_source(crate::seed::lower(envelope)?)?;
    Ok(source.stories.iter().map(|story| story.chunks.len()).sum())
}
