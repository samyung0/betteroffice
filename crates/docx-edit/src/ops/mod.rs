//! Mutating editing operations.

pub mod embed;
pub(crate) mod field_changes;
pub mod paragraph;
pub mod resolve;
pub mod table;
pub mod text;
pub mod toc;

use std::collections::BTreeMap;

use yrs::block::{ItemContent, ItemPtr};
use yrs::branch::Branch;
use yrs::types::Attrs;
use yrs::{Any, Map, MapRef, Out, ReadTxn, TextRef, TransactionMut};

use crate::segments::{SlotStart, apply_format};
use crate::{KIND_KEY, PARA_ID, PPR_CHANGE, is_pilcrow, map_string, out_len};

/// One formatting-run chunk of a story, snapshotted for index-stable reverse walks.
pub(crate) struct Chunk {
    pub start: u32,
    pub len: u32,
    pub kind: ChunkKind,
    pub attrs: BTreeMap<String, Any>,
}

pub(crate) enum ChunkKind {
    Text(String),
    Pilcrow(MapRef),
    Embed(Option<MapRef>),
}

impl Chunk {
    pub fn end(&self) -> u32 {
        self.start + self.len
    }

    pub fn attr_active(&self, key: &str) -> bool {
        matches!(self.attrs.get(key), Some(value) if *value != Any::Null)
    }

    /// The `author` of an `ins`/`del` revision value on this chunk, if any.
    pub fn revision_author(&self, key: &str) -> Option<String> {
        let Some(Any::Map(revision)) = self.attrs.get(key) else {
            return None;
        };
        match revision.get("author") {
            Some(Any::String(author)) => Some(author.to_string()),
            _ => None,
        }
    }
}

fn revision_id_for_author(value: &Any, author: &str) -> Option<String> {
    let Any::Map(revision) = value else {
        return None;
    };
    if !matches!(revision.get("author"), Some(Any::String(value)) if value.as_ref() == author) {
        return None;
    }
    match revision.get("id").or_else(|| revision.get("revisionId")) {
        Some(Any::String(id)) => Some(id.to_string()),
        Some(Any::Number(id)) if id.is_finite() => Some(id.to_string()),
        Some(Any::BigInt(id)) => Some(id.to_string()),
        _ => None,
    }
}

/// Reuse a same-author revision touching an edit boundary. This models Word's
/// continuous suggestion run: separately dispatched keystrokes and paragraph
/// breaks stay one revision while their stamped units remain adjacent.
pub(crate) fn adjacent_revision_id(
    chunks: &[Chunk],
    index: u32,
    key: &str,
    author: &str,
) -> Option<String> {
    chunks
        .iter()
        .rev()
        .filter(|chunk| chunk.end() == index)
        .chain(chunks.iter().filter(|chunk| chunk.start == index))
        .find_map(|chunk| {
            chunk
                .attrs
                .get(key)
                .and_then(|value| revision_id_for_author(value, author))
        })
}

/// Reuse a same-author paragraph-property revision at an edit boundary.
///
/// A list command commonly authors `pPrChange` on an empty paragraph before
/// the first character is typed. Treat that pending property change as the
/// start of the same continuous suggestion run so the list formatting, text,
/// and subsequent paragraph breaks resolve from one sidebar card.
pub(crate) fn adjacent_paragraph_change_revision_id<T: ReadTxn>(
    chunks: &[Chunk],
    index: u32,
    txn: &T,
    author: &str,
) -> Option<String> {
    chunks
        .iter()
        .rev()
        .filter(|chunk| chunk.end() == index)
        .chain(chunks.iter().filter(|chunk| chunk.start == index))
        .find_map(|chunk| {
            let ChunkKind::Pilcrow(map) = &chunk.kind else {
                return None;
            };
            let Some(Out::Any(Any::Array(changes))) = map.get(txn, PPR_CHANGE) else {
                return None;
            };
            changes.iter().rev().find_map(|change| {
                let Any::Map(change) = change else {
                    return None;
                };
                change
                    .get("info")
                    .and_then(|info| revision_id_for_author(info, author))
            })
        })
}

pub(crate) fn revision_id_in_range(
    chunks: &[Chunk],
    start: u32,
    end: u32,
    key: &str,
    author: &str,
) -> Option<String> {
    chunks
        .iter()
        .filter(|chunk| chunk.end() > start && chunk.start < end)
        .find_map(|chunk| {
            chunk
                .attrs
                .get(key)
                .and_then(|value| revision_id_for_author(value, author))
        })
}

fn chunk_at<T: ReadTxn>(insert: &Out, attrs: &Attrs, offset: u32, txn: &T) -> Chunk {
    let kind = match insert {
        Out::Any(Any::String(value)) => ChunkKind::Text(value.to_string()),
        Out::YMap(map) if is_pilcrow(map, txn) => ChunkKind::Pilcrow(map.clone()),
        Out::YMap(map) => ChunkKind::Embed(Some(map.clone())),
        _ => ChunkKind::Embed(None),
    };
    Chunk {
        start: offset,
        len: out_len(insert),
        kind,
        attrs: attrs
            .iter()
            .map(|(key, value)| (key.to_string(), value.clone()))
            .collect(),
    }
}

/// The chunks `TextRef::diff` reads a story as, from the first that ends
/// after `lo`, read item by item and only as far as asked: from `next`, at
/// story offset `at` with format attributes `attrs` open there (the story's
/// start, or a paragraph slot's). A run of text items with no format marker
/// or embed between them is one chunk.
pub(crate) struct Chunks<'a, T: ReadTxn> {
    txn: &'a T,
    next: Option<ItemPtr>,
    at: u32,
    attrs: Attrs,
    lo: u32,
    /// The text items of the run being read, and where it starts.
    run: Vec<ItemPtr>,
    run_start: u32,
    /// An embed's chunk read with the run before it, to give next.
    held: Option<Chunk>,
}

impl<'a, T: ReadTxn> Chunks<'a, T> {
    pub(crate) fn new(txn: &'a T, next: Option<ItemPtr>, at: u32, attrs: Attrs, lo: u32) -> Self {
        Self {
            txn,
            next,
            at,
            attrs,
            lo,
            run: Vec::new(),
            run_start: at,
            held: None,
        }
    }

    /// A story's chunks from `from`, a paragraph slot's start, else from
    /// the story's start.
    pub(crate) fn of(story: &TextRef, txn: &'a T, from: Option<&SlotStart>, lo: u32) -> Self {
        match from {
            Some((next, at, attrs)) => Self::new(txn, *next, *at, attrs.clone(), lo),
            None => Self::new(
                txn,
                AsRef::<Branch>::as_ref(story).start(),
                0,
                Attrs::new(),
                lo,
            ),
        }
    }

    /// The run read so far as a chunk, when it ends after `lo`.
    fn flush(&mut self) -> Option<Chunk> {
        if self.run.is_empty() {
            return None;
        }
        let chunk = (self.at > self.lo).then(|| {
            let mut text = String::new();
            for item in &self.run {
                if let ItemContent::String(part) = item.content() {
                    text.push_str(part.as_str());
                }
            }
            Chunk {
                start: self.run_start,
                len: self.at - self.run_start,
                kind: ChunkKind::Text(text),
                attrs: owned_attrs(&self.attrs),
            }
        });
        self.run.clear();
        chunk
    }
}

impl<T: ReadTxn> Iterator for Chunks<'_, T> {
    type Item = Chunk;

    fn next(&mut self) -> Option<Chunk> {
        if let Some(chunk) = self.held.take() {
            return Some(chunk);
        }
        while let Some(item) = self.next {
            self.next = item.right();
            if item.is_deleted() {
                continue;
            }
            match item.content() {
                ItemContent::Format(key, value) => {
                    let chunk = self.flush();
                    apply_format(&mut self.attrs, key, value);
                    if chunk.is_some() {
                        return chunk;
                    }
                }
                ItemContent::String(_) => {
                    if self.run.is_empty() {
                        self.run_start = self.at;
                    }
                    self.run.push(item);
                    self.at += item.len();
                }
                ItemContent::Type(_) | ItemContent::Embed(_) => {
                    let chunk = self.flush();
                    if let Some(value) = item.content().get_first() {
                        let len = out_len(&value);
                        if self.at + len > self.lo {
                            self.held = Some(chunk_at(&value, &self.attrs, self.at, self.txn));
                        }
                        self.at += len;
                    }
                    if chunk.is_some() {
                        return chunk;
                    }
                    if let Some(held) = self.held.take() {
                        return Some(held);
                    }
                }
                _ => {}
            }
        }
        self.flush()
    }
}

/// Format attributes as a chunk holds them.
fn owned_attrs(attrs: &Attrs) -> BTreeMap<String, Any> {
    attrs
        .iter()
        .map(|(key, value)| (key.to_string(), value.clone()))
        .collect()
}

/// The chunks of [`Chunks`] that overlap `[lo, hi)`.
pub(crate) fn chunks_from<T: ReadTxn>(
    txn: &T,
    next: Option<ItemPtr>,
    at: u32,
    attrs: Attrs,
    lo: u32,
    hi: u32,
) -> Vec<Chunk> {
    Chunks::new(txn, next, at, attrs, lo)
        .take_while(|chunk| chunk.start < hi)
        .collect()
}

pub(crate) fn snapshot<T: ReadTxn>(story: &TextRef, txn: &T) -> Vec<Chunk> {
    snapshot_range(story, txn, 0, u32::MAX)
}

/// `snapshot` over the chunks overlapping `[lo, hi)`.
pub(crate) fn snapshot_range<T: ReadTxn>(story: &TextRef, txn: &T, lo: u32, hi: u32) -> Vec<Chunk> {
    chunks_from(
        txn,
        AsRef::<Branch>::as_ref(story).start(),
        0,
        Attrs::new(),
        lo,
        hi,
    )
}

/// The story's last pilcrow as `(index, map)`.
pub(crate) fn last_pilcrow<T: ReadTxn>(story: &TextRef, txn: &T) -> Option<(u32, MapRef)> {
    let mut offset = 0;
    let mut last = None;
    let mut next = AsRef::<Branch>::as_ref(story).start();
    while let Some(item) = next {
        if !item.is_deleted()
            && let ItemContent::Type(_) = item.content()
            && let Some(Out::YMap(map)) = item.content().get_first()
            && is_pilcrow(&map, txn)
        {
            last = Some((offset, map));
        }
        offset += crate::segments::units(&item);
        next = item.right();
    }
    last
}

/// Captures every pilcrow property except the schema discriminator, plus the paraId.
pub(crate) fn capture_pilcrow<T: ReadTxn>(map: &MapRef, txn: &T) -> (String, Vec<(String, Any)>) {
    let para_id = map_string(map, txn, PARA_ID).unwrap_or_default();
    let props = map
        .iter(txn)
        .filter_map(|(key, value)| {
            if matches!(key, KIND_KEY | PARA_ID) {
                return None;
            }
            let Out::Any(value) = value else {
                return None;
            };
            Some((key.to_string(), value))
        })
        .collect();
    (para_id, props)
}

/// Replaces a pilcrow's ID and properties with the donor's.
pub(crate) fn adopt_pilcrow(
    txn: &mut TransactionMut<'_>,
    survivor: &MapRef,
    donor_para_id: &str,
    donor_props: &[(String, Any)],
) {
    let existing: Vec<String> = survivor
        .iter(txn)
        .map(|(key, _)| key.to_string())
        .filter(|key| key != KIND_KEY)
        .collect();
    for key in existing {
        survivor.remove(txn, &key);
    }
    survivor.insert(txn, PARA_ID, donor_para_id);
    for (key, value) in donor_props {
        survivor.insert(txn, key.clone(), value.clone());
    }
}

pub(crate) fn utf16_len(text: &str) -> u32 {
    text.encode_utf16().count() as u32
}
