//! What an agent command's Undo re-creates ([`crate::CommandResult::touched`],
//! the Editor's design 2026-10-22): its target's live Yjs items in the
//! document before the command applies, nested types whole, and their
//! content's bytes. Capy prices an edit's Undo from it
//! (`sources::undo_cost::office`).

use std::ops::AddAssign;

use serde::{Deserialize, Serialize};
use yrs::block::{Item, ItemContent, ItemPtr};
use yrs::branch::Branch;
use yrs::{Any, Doc, OffsetKind, ReadTxn, Transact};

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

/// The items of story `story` from UTF-16 offset `start` through `end`
/// included (a paragraph, its mark at `end`), and the format markers there.
pub(crate) fn story_range(doc: &Doc, root: &str, story: &str, start: u32, end: u32) -> Touched {
    let txn = doc.transact();
    let mut touched = Touched::default();
    let Some(holder) = story_item(&txn, root, story) else {
        return touched;
    };
    let ItemContent::Type(text) = holder.content() else {
        return touched;
    };
    let mut at = 0;
    let mut next = text.start();
    while let Some(item) = next {
        if at > end {
            break;
        }
        let len = if item.is_countable() && !item.is_deleted() {
            item.len()
        } else {
            0
        };
        if at + len > start || (len == 0 && at >= start) {
            touched.item(&item);
        }
        at += len;
        next = item.right();
    }
    touched
}
