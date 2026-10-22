//! What an agent command's Undo re-creates ([`crate::CommandResult::touched`],
//! the Editor's design 2026-10-22): its target's live Yjs items in the
//! document before the command applies, nested types whole, and their
//! content's bytes; a PPTX `add_slide` and `set_shape_text` add the slide
//! and the story they wrote, read after they apply (their Redo re-creates
//! them). Capy prices an edit's Undo from it
//! (`sources::undo_cost::office`).

use std::ops::AddAssign;

use serde::{Deserialize, Serialize};
use yrs::block::{Item, ItemContent, ItemPtr};
use yrs::branch::Branch;
use yrs::{Any, Doc, ID, OffsetKind, ReadTxn, Transact};

/// The live items under a target and the bytes their content holds.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct Touched {
    /// Text, embedded values, format keys and values, map entries' values.
    pub content: u64,
    /// Text runs, format markers, embeds, map entries and nested types.
    pub items: u64,
}

impl AddAssign for Touched {
    fn add_assign(&mut self, other: Touched) {
        self.content += other.content;
        self.items += other.items;
    }
}

impl Touched {
    /// `item` when live, with everything a nested type under it holds.
    fn item(&mut self, item: &Item) {
        if item.is_deleted() {
            return;
        }
        self.items += 1;
        self.content += match item.content() {
            ItemContent::String(text) => text.len(OffsetKind::Bytes) as u64,
            ItemContent::Any(values) => values.iter().map(json_len).sum(),
            ItemContent::Embed(value) => json_len(value),
            ItemContent::Format(key, value) => key.len() as u64 + json_len(value),
            ItemContent::JSON(values) => values.iter().map(|value| value.len() as u64).sum(),
            ItemContent::Binary(bytes) => bytes.len() as u64,
            ItemContent::Type(branch) => {
                self.branch(branch);
                0
            }
            _ => 0,
        };
    }

    fn branch(&mut self, branch: &Branch) {
        let mut next = branch.start();
        while let Some(item) = next {
            self.item(&item);
            next = item.right();
        }
        for (_, item) in branch.map_items() {
            self.item(&item);
        }
    }
}

fn json_len(value: &Any) -> u64 {
    let mut json = String::new();
    value.to_json(&mut json);
    json.len() as u64
}

/// The live item holding story `story` of root map `root` (DOCX `stories`,
/// PPTX `pptx:stories`).
fn story_item(txn: &impl ReadTxn, root: &str, story: &str) -> Option<ItemPtr> {
    let stories = txn.get_map(root)?;
    let item = AsRef::<Branch>::as_ref(&stories).entry_item(story)?;
    (!item.is_deleted()).then_some(item)
}

/// Story `story` whole: its map entry and its text.
pub(crate) fn story(doc: &Doc, root: &str, story: &str) -> Touched {
    let mut touched = Touched::default();
    if let Some(item) = story_item(&doc.transact(), root, story) {
        touched.item(&item);
    }
    touched
}

/// Which items of a story range an Undo re-creates.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) enum Kept {
    /// Text, embeds and nested types (whole): what a delete takes out, the
    /// format markers staying where they are.
    Content,
    /// Those and the format markers too: what a format rewrites.
    Formatting,
}

/// Where a walk over a story starts: right after the item at a given
/// UTF-16 offset of the story (a paragraph mark before the range), or at
/// the story's start (`None`).
pub(crate) type Seek = Option<(ID, u32)>;

/// The items of story `story` in UTF-16 offsets `[start, end)` that `kept`
/// names (a paragraph is `[start, mark + 1)`, its mark included), walked
/// from `from`.
pub(crate) fn story_range(
    doc: &Doc,
    root: &str,
    story: &str,
    from: Seek,
    start: u32,
    end: u32,
    kept: Kept,
) -> Touched {
    let mut touched = Touched::default();
    walk(doc, root, story, from, |at, len, item| {
        if at > end {
            return false;
        }
        let inside = if len == 0 {
            kept == Kept::Formatting && at >= start
        } else {
            at < end && at + len > start
        };
        if inside {
            touched.item(item);
        }
        true
    });
    touched
}

/// The entries `keys` of the paragraph mark whose item is `mark` (a map
/// embedded in a story): the values a property write replaces.
pub(crate) fn mark_entries(doc: &Doc, mark: &ID, keys: &[&str]) -> Touched {
    let mut touched = Touched::default();
    if let Some(item) = doc.transact().store().get_item(mark)
        && !item.is_deleted()
        && let ItemContent::Type(mark) = item.content()
    {
        for key in keys {
            if let Some(entry) = mark.entry_item(key) {
                touched.item(&entry);
            }
        }
    }
    touched
}

/// Each item of story `story` in order from `from` with its UTF-16 offset
/// and live length (0 for a format marker or a deleted item), until `visit`
/// says stop.
fn walk(
    doc: &Doc,
    root: &str,
    story: &str,
    from: Seek,
    mut visit: impl FnMut(u32, u32, &Item) -> bool,
) {
    let txn = doc.transact();
    let (mut next, mut at) = match from {
        Some((mark, at)) => match txn.store().get_item(&mark) {
            Some(mark) => (mark.right(), at),
            None => return,
        },
        None => {
            let Some(holder) = story_item(&txn, root, story) else {
                return;
            };
            let ItemContent::Type(text) = holder.content() else {
                return;
            };
            (text.start(), 0)
        }
    };
    while let Some(item) = next {
        let len = if item.is_countable() && !item.is_deleted() {
            item.len()
        } else {
            0
        };
        if !visit(at, len, &item) {
            return;
        }
        at += len;
        next = item.right();
    }
}
