// Segment/paragraph accessors used only by the wasm helpers read as dead code in
// native builds; the invalidation and paragraph paths below are live on both.
#![cfg_attr(not(feature = "wasm"), allow(dead_code))]

use std::collections::HashMap;
use std::sync::{Arc, Mutex};

use yrs::block::{Item, ItemContent, ItemPtr};
use yrs::branch::{Branch, BranchPtr};
use yrs::types::Attrs;
use yrs::{Any, ID, Out, ReadTxn, TransactionMut};

use crate::{KIND_KEY, PARA_ID, STORIES, is_pilcrow, map_string};

/// Whether the layout gives an embed its own block.
pub(crate) fn is_block_embed(kind: &str) -> bool {
    matches!(kind, "table" | "blockSdt" | "pageBreak" | "columnBreak")
}

/// The story units `TextRef::diff` reads a live item as: a string's UTF-16
/// length, one for an embed or nested type (a string embed's own length, as
/// `out_len` counts it); none for a format marker, a deleted item or content
/// the diff skips.
pub(crate) fn units(item: &Item) -> u32 {
    if item.is_deleted() {
        return 0;
    }
    match item.content() {
        ItemContent::String(_) => item.len(),
        ItemContent::Embed(Any::String(text)) => text.encode_utf16().count() as u32,
        ItemContent::Embed(_) | ItemContent::Type(_) => 1,
        _ => 0,
    }
}

/// Each live item of a story that holds units, from the story's start, with
/// its story offset: what `TextRef::diff` reads, item by item.
fn units_from<S: AsRef<Branch>>(story: &S) -> impl Iterator<Item = (u32, ItemPtr)> + use<S> {
    let mut next = story.as_ref().start();
    let mut at = 0;
    std::iter::from_fn(move || {
        loop {
            let item = next?;
            next = item.right();
            let units = units(&item);
            if units > 0 {
                let start = at;
                at += units;
                return Some((start, item));
            }
        }
    })
}

/// The map of a live embed item whose `_kind` is `kind`.
fn embed_of_kind<T: ReadTxn>(item: &Item, txn: &T, kind: &str) -> Option<yrs::MapRef> {
    if item.is_deleted() || !matches!(item.content(), ItemContent::Type(_)) {
        return None;
    }
    match item.content().get_first() {
        Some(Out::YMap(map)) if matches!(yrs::Map::get(&map, txn, KIND_KEY), Some(Out::Any(Any::String(found))) if &*found == kind) => {
            Some(map)
        }
        _ => None,
    }
}

/// The table embeds of a story with their story offsets, item by item.
pub(crate) fn tables<'a, S: AsRef<Branch>, T: ReadTxn>(
    story: &S,
    txn: &'a T,
) -> impl Iterator<Item = (u32, yrs::MapRef)> + use<'a, S, T> {
    units_from(story).filter_map(move |(at, item)| Some((at, embed_of_kind(&item, txn, "table")?)))
}

/// The format attributes open after a format marker, as `TextRef::diff`
/// carries them.
pub(crate) fn apply_format(attrs: &mut Attrs, key: &Arc<str>, value: &Any) {
    if *value == Any::Null {
        attrs.remove(key);
    } else {
        attrs.insert(key.clone(), value.clone());
    }
}

/// One paragraph's resolved geometry inside a story.
#[derive(Clone, Debug, PartialEq)]
pub(crate) struct ParaEntry {
    pub para_id: Arc<str>,
    /// Story index of the paragraph's first unit (after the previous pilcrow).
    pub start: u32,
    /// Story index of the paragraph's own pilcrow embed.
    pub pilcrow: u32,
    /// First content unit after the paragraph's leading block embeds.
    pub node_start: u32,
    /// The pilcrow's item, where a walk over the next slot starts.
    pub mark: ID,
    /// The format attributes open right after the pilcrow, which that walk
    /// starts with.
    pub after: Attrs,
}

#[derive(Clone, Debug, PartialEq)]
pub(crate) enum SegKind {
    /// A run of text between embeds (format markers do not split it).
    Text(Arc<str>),
    Pilcrow,
    /// `container`: a table or block content control, which Delete and
    /// Backspace never remove (the user selects it to delete it).
    Embed {
        container: bool,
        table: bool,
        /// A field whose code continues into the next paragraph: everything
        /// after it in its paragraph is that code.
        code_continues: bool,
        /// A field the render bridge draws as no text (no shown result, a
        /// comment reference): Backspace and Delete step over it.
        invisible: bool,
    },
}

#[derive(Clone, Debug, PartialEq)]
pub(crate) struct Seg {
    pub start: u32,
    pub kind: SegKind,
}

/// Segment geometry for one story, read once and then kept in step with each
/// commit that changes it ([`follow`]): the paragraphs a change sits in are
/// read again, from the mark before them, and the rest shift.
#[derive(Clone, Debug, PartialEq)]
pub(crate) struct SegmentIndex {
    len: u32,
    segs: Vec<Seg>,
    paras: Vec<ParaEntry>,
}

/// What a walk over items from a paragraph slot's start reads.
struct Scan {
    segs: Vec<Seg>,
    paras: Vec<ParaEntry>,
    /// The story offset after the last item read.
    at: u32,
}

/// Reads items from `next`, the first item of a paragraph slot at story
/// offset `at` with format attributes `attrs` open, through the pilcrow whose
/// item is `stop`, or to the story's end.
fn scan<T: ReadTxn>(
    txn: &T,
    mut next: Option<ItemPtr>,
    mut at: u32,
    mut attrs: Attrs,
    stop: Option<ID>,
) -> Scan {
    let (mut para_start, mut node_start) = (at, at);
    let mut segs = Vec::new();
    let mut paras = Vec::new();
    let mut text: Option<(u32, String)> = None;
    while let Some(item) = next {
        next = item.right();
        if item.is_deleted() {
            continue;
        }
        match item.content() {
            ItemContent::Format(key, value) => apply_format(&mut attrs, key, value),
            ItemContent::String(part) => {
                text.get_or_insert_with(|| (at, String::new()))
                    .1
                    .push_str(part.as_str());
                at += item.len();
            }
            ItemContent::Type(_) | ItemContent::Embed(_) => {
                if let Some((start, run)) = text.take() {
                    segs.push(Seg {
                        start,
                        kind: SegKind::Text(run.into()),
                    });
                }
                let units = units(&item);
                let kind = match item.content().get_first() {
                    Some(Out::YMap(map)) if is_pilcrow(&map, txn) => {
                        paras.push(ParaEntry {
                            para_id: map_string(&map, txn, PARA_ID).unwrap_or_default().into(),
                            start: para_start,
                            pilcrow: at,
                            node_start,
                            mark: *item.id(),
                            after: attrs.clone(),
                        });
                        para_start = at + 1;
                        node_start = at + 1;
                        SegKind::Pilcrow
                    }
                    value => {
                        let (kind, continuation, invisible) = match value {
                            Some(Out::YMap(map)) => {
                                let kind = map_string(&map, txn, KIND_KEY).unwrap_or_default();
                                // As the bridge draws it: PAGE and NUMPAGES show
                                // their own numbers, a numeric code nothing, any
                                // other field its shown result.
                                let invisible = kind == "field"
                                    && !matches!(
                                        map_string(&map, txn, "fieldType").as_deref(),
                                        Some("PAGE" | "NUMPAGES")
                                    )
                                    && (map_string(&map, txn, "instruction").is_some_and(|code| {
                                        crate::seed::numeric_field_instruction(&code)
                                    }) || map_string(&map, txn, "displayText")
                                        .is_none_or(|text| text.is_empty()));
                                let continuation = map_string(&map, txn, "fieldData")
                                    .filter(|data| data.contains("\"continuation\""))
                                    .and_then(|data| {
                                        serde_json::from_str::<serde_json::Value>(&data).ok()
                                    })
                                    .map(|field| field["continuation"].clone());
                                (kind, continuation, invisible)
                            }
                            _ => (String::new(), None, false),
                        };
                        let code_continues = continuation.as_ref().is_some_and(|continuation| {
                            continuation["separate"] == serde_json::Value::Bool(true)
                        });
                        if at == node_start && is_block_embed(&kind) {
                            node_start = at + 1;
                        }
                        SegKind::Embed {
                            container: matches!(kind.as_str(), "table" | "blockSdt"),
                            table: kind == "table",
                            code_continues,
                            invisible,
                        }
                    }
                };
                let pilcrow = kind == SegKind::Pilcrow;
                if units > 0 {
                    segs.push(Seg { start: at, kind });
                    at += units;
                }
                if pilcrow && Some(*item.id()) == stop {
                    break;
                }
            }
            _ => {}
        }
    }
    if let Some((start, run)) = text {
        segs.push(Seg {
            start,
            kind: SegKind::Text(run.into()),
        });
    }
    Scan { segs, paras, at }
}

/// `at` moved by `delta`.
fn moved(at: u32, delta: i64) -> u32 {
    (i64::from(at) + delta) as u32
}

impl SegmentIndex {
    pub(crate) fn build<T: ReadTxn>(story: &impl AsRef<Branch>, txn: &T) -> Self {
        let scan = scan(txn, story.as_ref().start(), 0, Attrs::new(), None);
        Self {
            len: scan.at,
            segs: scan.segs,
            paras: scan.paras,
        }
    }

    pub(crate) fn paras(&self) -> &[ParaEntry] {
        &self.paras
    }

    /// The first field in `[from, to)` whose code continues into the next paragraph.
    pub(crate) fn code_continues_in(&self, from: u32, to: u32) -> Option<u32> {
        self.segs[self.segs.partition_point(|seg| seg.start < from)..]
            .iter()
            .take_while(|seg| seg.start < to)
            .find(|seg| {
                matches!(
                    seg.kind,
                    SegKind::Embed {
                        code_continues: true,
                        ..
                    }
                )
            })
            .map(|seg| seg.start)
    }

    /// `(start, pilcrow)` span of `para_id`, matching the first paragraph with that id.
    pub(crate) fn para_span(&self, para_id: &str) -> Option<(u32, u32)> {
        let para = self.paras.iter().find(|para| &*para.para_id == para_id)?;
        Some((para.start, para.pilcrow))
    }

    /// How many tables come before story offset `at`.
    pub(crate) fn tables_before(&self, at: u32) -> usize {
        self.segs[..self.segs.partition_point(|seg| seg.start < at)]
            .iter()
            .filter(|seg| matches!(seg.kind, SegKind::Embed { table: true, .. }))
            .count()
    }

    /// The [`SlotMark`] of the slot holding story offset `at`; `None` past
    /// the story's last mark, which a walk from its start finds.
    pub(crate) fn slot_mark(&self, at: u32) -> SlotMark {
        let slot = self.slot_at(at);
        if slot == self.paras.len() {
            return None;
        }
        let previous = slot.checked_sub(1)?;
        let para = &self.paras[previous];
        Some((para.mark, para.pilcrow + 1))
    }

    /// The paragraph whose mark is `mark`, by its slot.
    pub(crate) fn slot_of(&self, mark: &ID) -> Option<usize> {
        self.paras.iter().position(|para| para.mark == *mark)
    }

    /// First paragraph whose pilcrow sits at or after `index` — the paragraph `index`
    /// resolves into.
    pub(crate) fn para_at(&self, index: u32) -> Option<&ParaEntry> {
        self.paras.get(self.slot_at(index))
    }

    /// The slot of [`Self::para_at`] (`paras().len()` past the last pilcrow).
    pub(crate) fn slot_at(&self, index: u32) -> usize {
        self.paras.partition_point(|para| para.pilcrow < index)
    }

    /// The segment covering `pos`, if any.
    pub(crate) fn segment_at(&self, pos: u32) -> Option<&Seg> {
        if pos >= self.len {
            return None;
        }
        self.segs
            .get(self.segs.partition_point(|seg| seg.start <= pos) - 1)
    }

    /// Where a walk over slot `slot` starts: its first item, its story offset
    /// and the format attributes open there. `None` when the mark before it
    /// is no longer a live item.
    pub(crate) fn slot_start<T: ReadTxn>(
        &self,
        story: &Branch,
        txn: &T,
        slot: usize,
    ) -> Option<SlotStart> {
        let Some(previous) = slot.checked_sub(1).map(|slot| &self.paras[slot]) else {
            return Some((story.start(), 0, Attrs::new()));
        };
        let mark = txn.store().get_item(&previous.mark)?;
        (!mark.is_deleted()).then(|| (mark.right(), previous.pilcrow + 1, previous.after.clone()))
    }

    /// Brings this index in step after `txn` changed `story` at `changed`
    /// (each item the change sits in at the story's level: itself, or the
    /// embed holding a nested type it is in). The paragraphs those items are
    /// in, and those whose marks went, are read again from the mark before
    /// them; the rest shift. `None` where it cannot follow, and the index is
    /// read again; it is left as it was.
    pub(crate) fn follow(
        &mut self,
        txn: &TransactionMut,
        story: &Branch,
        changed: &[ItemPtr],
    ) -> Option<()> {
        let inserted = txn.insert_set();
        let (mut first, mut last) = (usize::MAX, 0);
        for item in changed {
            // A mark this commit deleted: its paragraph joins the next.
            if matches!(item.content(), ItemContent::Type(_))
                && let Some(slot) = self.slot_of(item.id())
            {
                (first, last) = (first.min(slot), last.max(slot));
            }
            // The first old mark at or after the change closes its paragraph.
            let mut next = Some(*item);
            let slot = loop {
                let item = next?;
                if !item.is_deleted()
                    && !inserted.contains(item.id())
                    && let Some(Out::YMap(map)) = item.content().get_first()
                    && is_pilcrow(&map, txn)
                {
                    break self.slot_of(item.id())?;
                }
                next = item.right();
            };
            (first, last) = (first.min(slot), last.max(slot));
        }
        if first == usize::MAX {
            return Some(());
        }
        let (next, at, attrs) = self.slot_start(story, txn, first)?;
        let old = &self.paras[last];
        let scan = scan(txn, next, at, attrs, Some(old.mark));
        let new = scan.paras.last()?;
        if new.mark != old.mark || new.after != old.after {
            return None;
        }
        let (old_start, old_end) = (self.paras[first].start, old.pilcrow + 1);
        let delta = i64::from(scan.at) - i64::from(old_end);
        let len = moved(self.len, delta);
        if len != story.content_len() {
            return None;
        }
        self.len = len;
        let from = self.segs.partition_point(|seg| seg.start < old_start);
        let to = self.segs.partition_point(|seg| seg.start < old_end);
        let added = scan.segs.len();
        self.segs.splice(from..to, scan.segs);
        for seg in &mut self.segs[from + added..] {
            seg.start = moved(seg.start, delta);
        }
        let added = scan.paras.len();
        self.paras.splice(first..=last, scan.paras);
        for para in &mut self.paras[first + added..] {
            para.start = moved(para.start, delta);
            para.pilcrow = moved(para.pilcrow, delta);
            para.node_start = moved(para.node_start, delta);
        }
        Some(())
    }
}

/// Where a walk over a paragraph slot starts: its first item, its story
/// offset and the format attributes open there.
pub(crate) type SlotStart = (Option<ItemPtr>, u32, Attrs);

/// Where a walk over a paragraph slot starts, read after an edit at or
/// after the slot's start: the item of the mark before the slot and the
/// slot's story offset. `None` for the story's start (its first slot, or no
/// index in step); a walk whose mark went starts there too.
pub(crate) type SlotMark = Option<(ID, u32)>;

/// Segment indexes by story, each with the document epoch it is in step with.
pub(crate) type Indexes = HashMap<Box<str>, (u64, Arc<SegmentIndex>)>;

/// Changed items a commit may hold before the indexes are read again instead
/// of followed.
const FOLLOWED_ITEMS: u32 = 4_096;

/// The story an integrated item is in and the item at the story's level:
/// itself in a story's text, or the embed holding the nested type it is in;
/// `None` at the stories map's level (a story added or removed). Outside the
/// stories map: none.
pub(crate) fn story_level(
    mut item: ItemPtr,
    stories: BranchPtr,
) -> Option<(Arc<str>, Option<ItemPtr>)> {
    loop {
        let parent = item.parent_branch()?;
        if parent == stories {
            return Some((item.parent_sub()?.clone(), None));
        }
        let holder = parent.item()?;
        if holder.parent_branch() == Some(stories) {
            return Some((holder.parent_sub()?.clone(), Some(item)));
        }
        item = holder;
    }
}

/// After a commit that took the document from epoch `before` to the next:
/// each index in step with `before` follows the commit's changes to its
/// story; any other is dropped.
pub(crate) fn follow_commit(txn: &TransactionMut, indexes: &Mutex<Indexes>, before: u64) {
    let mut indexes = indexes.lock().unwrap();
    indexes.retain(|_, (epoch, _)| *epoch == before);
    if indexes.is_empty() {
        return;
    }
    let Some(stories) = txn.get_map(STORIES) else {
        indexes.clear();
        return;
    };
    let stories = BranchPtr::from(AsRef::<Branch>::as_ref(&stories));
    let sets = [txn.insert_set(), txn.delete_set()];
    let mut units = 0;
    for set in sets {
        for (_, ranges) in set.iter() {
            units += ranges
                .iter()
                .map(|range| range.end - range.start)
                .sum::<u32>();
        }
    }
    if units > FOLLOWED_ITEMS {
        indexes.clear();
        return;
    }
    let mut changed: HashMap<Arc<str>, Vec<Option<ItemPtr>>> = HashMap::new();
    for set in sets {
        for (client, ranges) in set.iter() {
            for range in ranges.iter() {
                for (_, item) in txn.store().blocks_in(client, range.clone()) {
                    if let Some((story, level)) = item.and_then(|item| story_level(item, stories))
                        && indexes.contains_key(&*story)
                    {
                        changed.entry(story).or_default().push(level);
                    }
                }
            }
        }
    }
    indexes.retain(|story, (epoch, index)| {
        *epoch = before.wrapping_add(1);
        let Some(levels) = changed.get(&**story) else {
            return true;
        };
        let Some(mut items) = levels.iter().copied().collect::<Option<Vec<_>>>() else {
            return false;
        };
        items.sort_by_key(|item| *item.id());
        items.dedup_by_key(|item| *item.id());
        let Some(branch) = stories
            .entry_item(story)
            .and_then(|item| match item.content() {
                ItemContent::Type(branch) if !item.is_deleted() => Some(BranchPtr::from(branch)),
                _ => None,
            })
        else {
            return false;
        };
        // In place unless a reader still holds this one.
        Arc::make_mut(index).follow(txn, &branch, &items).is_some()
    });
}

#[cfg(test)]
pub(crate) mod tests {
    use super::*;
    use crate::{EditCtx, EditingDoc, FormatPolicy, Position, RawOp, SegmentContent};
    use yrs::{Any, Map, Transact};

    /// The pre-index paragraph walk, kept as the reference oracle.
    fn reference_para_at(doc: &EditingDoc, story: &str, index: u32) -> Option<(String, u32, u32)> {
        let mut cursor = 0_u32;
        let mut para_start = 0_u32;
        let mut node_start = 0_u32;
        for segment in doc.story_segments(story).unwrap() {
            match segment.content {
                SegmentContent::Text(text) => cursor += text.encode_utf16().count() as u32,
                SegmentContent::Pilcrow(properties) => {
                    if index <= cursor {
                        return Some((
                            properties.para_id,
                            index.saturating_sub(para_start),
                            index.saturating_sub(node_start),
                        ));
                    }
                    cursor += 1;
                    para_start = cursor;
                    node_start = cursor;
                }
                SegmentContent::OtherEmbed { ref kind, .. } => {
                    if cursor == node_start && is_block_embed(kind) {
                        node_start = cursor + 1;
                    }
                    cursor += 1;
                }
            }
        }
        None
    }

    /// The pre-index unit classification (kind only; text widths stay in the caller).
    fn reference_seg_kind(doc: &EditingDoc, story: &str, pos: u32) -> Option<&'static str> {
        let mut cursor = 0_u32;
        for segment in doc.story_segments(story).unwrap() {
            let units = match &segment.content {
                SegmentContent::Text(text) => text.encode_utf16().count() as u32,
                _ => 1,
            };
            if pos >= cursor && pos < cursor + units {
                return Some(match segment.content {
                    SegmentContent::Text(_) => "text",
                    SegmentContent::Pilcrow(_) => "pilcrow",
                    SegmentContent::OtherEmbed { .. } => "embed",
                });
            }
            cursor += units;
        }
        None
    }

    /// The index `doc` serves for `story` equals one read afresh and agrees
    /// with the segment walks.
    pub(crate) fn assert_index_matches_segments(doc: &EditingDoc, story: &str) {
        let index = doc.segment_index(story).unwrap();
        let fresh = {
            let txn = doc.yrs_doc().transact();
            SegmentIndex::build(&crate::story_ref(&txn, story).unwrap(), &txn)
        };
        let lossy = lossy_starts(&fresh);
        assert_eq!(
            comparable(&index, &lossy),
            comparable(&fresh, &lossy),
            "the followed index of {story}"
        );
        let len = doc.story_len(story).unwrap();
        for pos in 0..len {
            let expected = reference_seg_kind(doc, story, pos);
            let seg = index.segment_at(pos).unwrap();
            let actual = match seg.kind {
                SegKind::Text(_) => "text",
                SegKind::Pilcrow => "pilcrow",
                SegKind::Embed { .. } => "embed",
            };
            assert_eq!(expected, Some(actual), "pos {pos}");
        }
        assert!(index.segment_at(len).is_none());
        let para_count = doc.paragraphs(story).unwrap().len() as u32;
        for index_pos in 0..=len {
            let expected = reference_para_at(doc, story, index_pos);
            let actual = index.para_at(index_pos).map(|para| {
                (
                    para.para_id.to_string(),
                    index_pos.saturating_sub(para.start),
                    index_pos.saturating_sub(para.node_start),
                )
            });
            assert_eq!(expected, actual, "index {index_pos}");
        }
        assert_eq!(index.paras.len() as u32, para_count);
        for paragraph in doc.paragraphs(story).unwrap() {
            let span = index.para_span(&paragraph.para_id).unwrap();
            let reference = {
                // The pre-index para-span walk, kept as the reference oracle.
                let mut offset = 0_u32;
                let mut para_start = 0_u32;
                let mut found = (u32::MAX, u32::MAX);
                for segment in doc.story_segments(story).unwrap() {
                    match segment.content {
                        SegmentContent::Text(text) => offset += text.encode_utf16().count() as u32,
                        SegmentContent::Pilcrow(properties) => {
                            if properties.para_id == paragraph.para_id {
                                found = (para_start, offset);
                                break;
                            }
                            offset += 1;
                            para_start = offset;
                        }
                        SegmentContent::OtherEmbed { .. } => offset += 1,
                    }
                }
                found
            };
            assert_eq!(span, reference, "para {}", paragraph.para_id);
        }
    }

    /// `A B [sdt] pilcrow(p1)` | `pilcrow(p2)` (empty paragraph) | `[table] [pageBreak] C pilcrow(p3)`
    fn seeded_doc() -> EditingDoc {
        let doc = EditingDoc::new(7);
        seed(&doc);
        doc
    }

    /// [`seeded_doc`]'s stories, in `doc`.
    fn seed(doc: &EditingDoc) {
        doc.create_story("body", "AB", "Normal", "left").unwrap();
        doc.apply_raw_ops(
            "body",
            vec![
                RawOp::InsertEmbed {
                    index: 2,
                    kind: "sdt".into(),
                    payload: vec![("embedId".into(), Any::from("control-1"))],
                    attrs: Default::default(),
                },
                RawOp::InsertEmbed {
                    index: 4,
                    kind: "pilcrow".into(),
                    payload: vec![("paraId".into(), Any::from("p-2"))],
                    attrs: Default::default(),
                },
                RawOp::InsertEmbed {
                    index: 5,
                    kind: "table".into(),
                    payload: vec![("embedId".into(), Any::from("t-1"))],
                    attrs: Default::default(),
                },
                RawOp::InsertEmbed {
                    index: 6,
                    kind: "pageBreak".into(),
                    payload: vec![("embedId".into(), Any::from("p-1"))],
                    attrs: Default::default(),
                },
                RawOp::Insert {
                    index: 7,
                    text: "C".into(),
                    attrs: Default::default(),
                },
                RawOp::InsertEmbed {
                    index: 8,
                    kind: "pilcrow".into(),
                    payload: vec![("paraId".into(), Any::from("p-3"))],
                    attrs: Default::default(),
                },
            ],
            &EditCtx::local(String::new(), String::new()),
        )
        .unwrap();
    }

    /// What [`random_edit`] starts from, in `doc`: [`seeded_doc`]'s body
    /// with four more paragraphs, and a header.
    pub(crate) fn random_doc(doc: &EditingDoc) {
        seed(doc);
        doc.create_story("hf:h1", "Header", "Header", "left")
            .unwrap();
        let ctx = EditCtx::local("Ada", "2026-10-22T00:00:00Z");
        for text in ["First paragraph", "Second one", "Third", "Fourth with more"] {
            doc.insert_text(&ctx, Position::new("body", 0), text, FormatPolicy::Plain)
                .unwrap();
            doc.split_paragraph(&ctx, Position::new("body", 0)).ok();
        }
    }

    #[test]
    fn segment_index_matches_segment_walks() {
        let doc = seeded_doc();
        assert_index_matches_segments(&doc, "body");
    }

    #[test]
    fn segment_index_rebuilds_after_local_and_remote_changes() {
        let doc = seeded_doc();
        let ctx = EditCtx::local(String::new(), String::new());
        let remote = EditingDoc::new(9);
        remote
            .apply_update_v1(&doc.encode_state_as_update_v1())
            .unwrap();
        doc.segment_index("body").unwrap();
        remote.segment_index("body").unwrap();
        doc.insert_text(&ctx, Position::new("body", 0), "Z", FormatPolicy::Plain)
            .unwrap();
        assert_index_matches_segments(&doc, "body");
        remote
            .apply_update_v1(
                &doc.encode_diff_v1(&remote.encode_state_vector_v1())
                    .unwrap(),
            )
            .unwrap();
        assert_index_matches_segments(&remote, "body");
        doc.delete_range(&ctx, crate::StoryRange::new("body", 0, 4))
            .unwrap();
        assert_index_matches_segments(&doc, "body");
    }

    /// xorshift64: the edits a seed draws.
    pub(crate) struct Draw(pub(crate) u64);

    impl Draw {
        pub(crate) fn below(&mut self, n: u32) -> u32 {
            self.0 ^= self.0 << 13;
            self.0 ^= self.0 >> 7;
            self.0 ^= self.0 << 17;
            (self.0 % u64::from(n.max(1))) as u32
        }
    }

    const PIECES: [&str; 5] = ["x", "two words", "漢字", "😀𝄞", ""];

    /// Each story of `doc` by id.
    fn story_ids(doc: &EditingDoc) -> Vec<String> {
        let txn = doc.yrs_doc().transact();
        let mut ids: Vec<String> = txn
            .get_map(STORIES)
            .unwrap()
            .keys(&txn)
            .map(str::to_owned)
            .collect();
        ids.sort();
        ids
    }

    /// One random step on `doc`: mostly its own edit ([`local_edit`]);
    /// otherwise a peer's edits of any kind on a state that diverged from
    /// this one, and maybe one of `doc`'s own meanwhile, exchanged both ways
    /// so they integrate concurrently.
    pub(crate) fn random_edit(
        doc: &EditingDoc,
        peer: &EditingDoc,
        undo: &crate::UndoSession,
        draw: &mut Draw,
    ) {
        if draw.below(8) != 0 {
            local_edit(doc, undo, draw);
            return;
        }
        let peer_undo = crate::UndoSession::new();
        for _ in 0..1 + draw.below(3) {
            local_edit(peer, &peer_undo, draw);
        }
        if draw.below(2) == 0 {
            local_edit(doc, undo, draw);
        }
        let to_peer = doc.encode_diff_v1(&peer.encode_state_vector_v1()).unwrap();
        doc.apply_update_v1(&peer.encode_diff_v1(&doc.encode_state_vector_v1()).unwrap())
            .unwrap();
        peer.apply_update_v1(&to_peer).unwrap();
    }

    /// One random edit of a random story of `doc`, as a person, a suggesting
    /// person, or an Undo or Redo makes it; errors are the edit's.
    fn local_edit(doc: &EditingDoc, undo: &crate::UndoSession, draw: &mut Draw) {
        let ids = story_ids(doc);
        let story = ids[draw.below(ids.len() as u32) as usize].clone();
        let len = doc.story_len(&story).unwrap();
        let at = draw.below(len + 1);
        let to = (at + draw.below(12)).min(len);
        let piece = PIECES[draw.below(PIECES.len() as u32) as usize];
        let ctx = if draw.below(4) == 0 {
            EditCtx::local("Ada", "2026-10-22T00:00:00Z").suggesting()
        } else {
            EditCtx::local("Ada", "2026-10-22T00:00:00Z")
        };
        undo.track(doc);
        let _ = match draw.below(14) {
            0..=2 => {
                // The receipt, read from the slot before the insert, is
                // where a read from the story's start puts the text.
                let landing = doc.inline_landing(&story, at, at).unwrap();
                doc.insert_text(
                    &ctx,
                    Position::new(&story, at),
                    piece,
                    FormatPolicy::Inherit,
                )
                .map(|receipt| {
                    if !piece.is_empty() {
                        let end = landing + piece.encode_utf16().count() as u32;
                        let range = crate::StoryRange::new(&story, landing, end);
                        assert_eq!(receipt.range, Some(doc.loc_range_of(&range).unwrap()));
                    }
                })
            }
            3..=4 => doc
                .replace_range(&ctx, crate::StoryRange::new(&story, at, to), piece)
                .map(drop),
            5 => doc
                .delete_range(&ctx, crate::StoryRange::new(&story, at, to))
                .map(|receipt| {
                    if !ctx.is_suggesting() {
                        let range = crate::StoryRange::new(&story, at, at);
                        assert_eq!(receipt.range, Some(doc.loc_range_of(&range).unwrap()));
                    }
                }),
            6 => doc
                .split_paragraph(&ctx, Position::new(&story, at))
                .map(drop),
            7 => {
                let paragraphs = doc.paragraphs(&story).unwrap();
                let para = &paragraphs[draw.below(paragraphs.len() as u32) as usize];
                let direction = if draw.below(2) == 0 {
                    crate::MergeDirection::Forward
                } else {
                    crate::MergeDirection::Backward
                };
                doc.merge_paragraphs(&ctx, &para.para_id, direction)
                    .map(drop)
            }
            8 => doc
                .format_range(
                    &ctx,
                    crate::StoryRange::new(&story, at, to),
                    &crate::InlineFormatDelta {
                        bold: crate::Patch::Set(draw.below(2) == 0),
                        font_family: if draw.below(3) == 0 {
                            crate::Patch::Set(crate::FontFamilyPatch {
                                ascii: "Georgia".into(),
                                h_ansi: None,
                            })
                        } else {
                            crate::Patch::Keep
                        },
                        ..Default::default()
                    },
                )
                .map(|receipt| {
                    let range = crate::StoryRange::new(&story, at, to);
                    assert_eq!(receipt.range, Some(doc.loc_range_of(&range).unwrap()));
                }),
            9 => doc
                .set_paragraph_attrs(
                    &ctx,
                    &crate::ParaSelector::Range(crate::StoryRange::new(&story, at, to)),
                    &crate::ParaAttrDelta {
                        alignment: crate::Patch::Set("center".into()),
                        ..Default::default()
                    },
                )
                .map(drop),
            10 => doc
                .insert_table(
                    &ctx,
                    Position::new(&story, at),
                    1 + draw.below(2),
                    1 + draw.below(2),
                    12,
                )
                .map(drop),
            11 => doc
                .delete_table(&ctx, &crate::TableLocator::new(&story, draw.below(2)))
                .map(drop),
            12 => {
                let paragraphs = doc.paragraphs(&story).unwrap();
                let para = &paragraphs[draw.below(paragraphs.len() as u32) as usize];
                if draw.below(2) == 0 {
                    doc.remove_paragraph(&ctx, &story, &para.para_id).map(drop)
                } else {
                    doc.insert_paragraph_before(&ctx, &story, &para.para_id)
                        .map(drop)
                }
            }
            _ => {
                if draw.below(2) == 0 {
                    undo.undo();
                } else {
                    undo.redo();
                }
                Ok(())
            }
        };
    }

    /// An index as comparable values. A text item an edit split inside a
    /// surrogate pair reads U+FFFD for each half once read again, where a
    /// followed index may still hold the pair: a text segment starting at
    /// one of `lossy` (where the fresh read holds U+FFFD) compares by its
    /// UTF-16 length.
    fn comparable(
        index: &SegmentIndex,
        lossy: &[u32],
    ) -> (u32, Vec<ParaEntry>, Vec<(u32, String)>) {
        let segs = index
            .segs
            .iter()
            .map(|seg| {
                let kind = match &seg.kind {
                    SegKind::Text(text) if lossy.contains(&seg.start) => {
                        format!("text of {} units", text.encode_utf16().count())
                    }
                    kind => format!("{kind:?}"),
                };
                (seg.start, kind)
            })
            .collect();
        (index.len, index.paras.clone(), segs)
    }

    /// Where `fresh`'s text segments holding U+FFFD start ([`comparable`]).
    fn lossy_starts(fresh: &SegmentIndex) -> Vec<u32> {
        fresh
            .segs
            .iter()
            .filter(|seg| matches!(&seg.kind, SegKind::Text(text) if text.contains('\u{FFFD}')))
            .map(|seg| seg.start)
            .collect()
    }

    /// Chunks as comparable values.
    fn chunk_values(
        chunks: &[crate::ops::Chunk],
    ) -> Vec<(u32, u32, String, std::collections::BTreeMap<String, Any>)> {
        chunks
            .iter()
            .map(|chunk| {
                let kind = match &chunk.kind {
                    crate::ops::ChunkKind::Text(text) => format!("text {text}"),
                    crate::ops::ChunkKind::Pilcrow(_) => "pilcrow".to_owned(),
                    crate::ops::ChunkKind::Embed(_) => "embed".to_owned(),
                };
                (chunk.start, chunk.len, kind, chunk.attrs.clone())
            })
            .collect()
    }

    /// Every index a document keeps follows random edits (local, suggested,
    /// Undo and Redo, a peer's of every kind integrated concurrently) to
    /// equal the index read afresh; the reads that start from it equal those
    /// from the story's start.
    #[test]
    fn segment_indexes_follow_random_edits() {
        let (mut followed, mut read) = (0, 0);
        for seed in 1..=24_u64 {
            let mut draw = Draw(seed.wrapping_mul(0x9E37_79B9_7F4A_7C15));
            let doc = EditingDoc::new(7);
            random_doc(&doc);
            let peer = EditingDoc::new(9);
            peer.apply_update_v1(&doc.encode_state_as_update_v1())
                .unwrap();
            let undo = crate::UndoSession::new();
            for ids in story_ids(&doc) {
                doc.segment_index(&ids).unwrap();
            }
            for step in 0..120 {
                let before: HashMap<String, Arc<SegmentIndex>> = doc
                    .segment_indexes
                    .lock()
                    .unwrap()
                    .iter()
                    .map(|(story, (_, index))| (story.to_string(), Arc::clone(index)))
                    .collect();
                random_edit(&doc, &peer, &undo, &mut draw);
                let epoch = doc.epoch.load(std::sync::atomic::Ordering::Relaxed);
                for story in story_ids(&doc) {
                    let in_step = doc
                        .segment_indexes
                        .lock()
                        .unwrap()
                        .get(story.as_str())
                        .is_some_and(|(at, _)| *at == epoch);
                    let index = doc.segment_index(&story).unwrap();
                    if before.contains_key(&story) {
                        if in_step {
                            followed += 1;
                        } else {
                            read += 1;
                        }
                    }
                    let fresh = {
                        let txn = doc.yrs_doc().transact();
                        SegmentIndex::build(&crate::story_ref(&txn, &story).unwrap(), &txn)
                    };
                    let lossy = lossy_starts(&fresh);
                    assert_eq!(
                        comparable(&index, &lossy),
                        comparable(&fresh, &lossy),
                        "seed {seed} step {step} story {story}"
                    );
                    let len = doc.story_len(&story).unwrap();
                    let lo = draw.below(len + 1);
                    let hi = lo + draw.below(20);
                    let txn = doc.yrs_doc().transact_mut();
                    let text = crate::story_ref(&txn, &story).unwrap();
                    assert_eq!(
                        chunk_values(&doc.snapshot_range_in(&story, &text, &txn, lo, hi)),
                        chunk_values(&crate::ops::snapshot_range(&text, &txn, lo, hi)),
                        "seed {seed} step {step} story {story} [{lo}, {hi})"
                    );
                }
            }
            for story in story_ids(&doc) {
                assert_index_matches_segments(&doc, &story);
            }
        }
        // Most commits are followed, not read again.
        assert!(
            followed > 4 * read,
            "followed {followed}, read again {read}"
        );
    }
}
