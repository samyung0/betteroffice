//! Memory safety of the patched yrs under DOCX edits, with garbage collection
//! on as sessions run it. Paragraph marks are maps embedded in story text, so
//! a peer's delete or merge of a paragraph whose attribute the user changed
//! and undid is the shape that left kept items under a freed branch.

use std::collections::HashSet;

use docx_edit::{
    DocUndoManager, EditCtx, EditingDoc, FormatPolicy, MergeDirection, Position, SeedParagraph,
    StoryRange, UndoCaptureMode,
};
use yrs::block::{Item, ItemContent};
use yrs::branch::Branch;
use yrs::{Any, Doc, ID, Out, ReadTxn, Transact};

/// Items the store holds that no root reaches: an item outlives its parent
/// only when garbage collection freed the parent's branch under it. Reads
/// each orphan only by id, never through its links.
fn orphans(doc: &Doc) -> Vec<ID> {
    fn walk(branch: &Branch, seen: &mut HashSet<ID>) {
        let mut next = branch.start();
        while let Some(item) = next {
            visit(&item, seen);
            next = item.right();
        }
        for (_, last) in branch.map_items() {
            let mut next = Some(last);
            while let Some(item) = next {
                visit(&item, seen);
                next = item.left();
            }
        }
    }
    fn visit(item: &Item, seen: &mut HashSet<ID>) {
        if seen.insert(*item.id())
            && let ItemContent::Type(branch) = item.content()
        {
            walk(branch, seen);
        }
    }
    let txn = doc.transact();
    let mut seen = HashSet::new();
    for (_, root) in txn.root_refs() {
        let branch: &Branch = match &root {
            Out::YMap(map) => map.as_ref(),
            Out::YArray(array) => array.as_ref(),
            Out::YText(text) => text.as_ref(),
            Out::UndefinedRef(branch) => branch,
            other => panic!("unexpected root {other:?}"),
        };
        walk(branch, &mut seen);
    }
    let store = txn.store();
    let mut orphans = Vec::new();
    for (client, end) in txn.state_vector().iter() {
        let mut clock = 0;
        while clock < *end {
            let (range, item) = store.get_block_range(&ID::new(*client, clock)).unwrap();
            if let Some(item) = item
                && !seen.contains(item.id())
            {
                orphans.push(*item.id());
            }
            clock = range.clock + range.len;
        }
    }
    orphans
}

struct Rng(u64);

impl Rng {
    fn next(&mut self) -> u64 {
        self.0 ^= self.0 << 13;
        self.0 ^= self.0 >> 7;
        self.0 ^= self.0 << 17;
        self.0
    }

    fn below(&mut self, n: usize) -> usize {
        if n == 0 {
            0
        } else {
            (self.next() % n as u64) as usize
        }
    }
}

struct Peer {
    doc: EditingDoc,
    undo: DocUndoManager,
}

fn ctx() -> EditCtx {
    EditCtx::local("test", "2026-01-01T00:00:00Z")
}

fn paragraph(text: &str) -> SeedParagraph {
    SeedParagraph {
        text: text.into(),
        p_style: "Normal".into(),
        alignment: "left".into(),
    }
}

fn sync(from: &Peer, to: &Peer) {
    let diff = from
        .doc
        .encode_diff_v1(&to.doc.encode_state_vector_v1())
        .unwrap();
    to.doc.apply_update_v1(&diff).unwrap();
}

/// One random edit, Undo or Redo, as a user would make it.
fn random_step(rng: &mut Rng, peer: &mut Peer, tag: &str) {
    let story = ["body", "fn1"][rng.below(2)];
    let len = peer.doc.story_len(story).unwrap_or(0);
    let paragraphs = peer.doc.paragraphs(story).unwrap_or_default();
    let doc = &peer.doc;
    // An edit the engine refuses (a stale position after a sync) is skipped.
    match rng.below(16) {
        0 | 1 if len > 0 => {
            let at = Position::new(story, rng.below(len as usize) as u32);
            let _ = doc.insert_text(&ctx(), at, tag, FormatPolicy::Inherit);
        }
        2 | 3 if len > 1 => {
            let start = rng.below(len as usize - 1) as u32;
            let end = (start + 1 + rng.below(6) as u32).min(len);
            let _ = doc.delete_range(&ctx(), StoryRange::new(story, start, end));
        }
        4..=6 if !paragraphs.is_empty() => {
            let id = &paragraphs[rng.below(paragraphs.len())].para_id;
            let key = ["alignment", "spacingAfter", "keepNext"][rng.below(3)];
            let _ = doc.set_paragraph_attr(id, key, Any::from(tag));
        }
        7 if len > 0 => {
            let at = Position::new(story, rng.below(len as usize) as u32);
            let _ = doc.split_paragraph(&ctx(), at);
        }
        8 if paragraphs.len() > 1 => {
            let id = &paragraphs[rng.below(paragraphs.len() - 1)].para_id;
            let _ = doc.merge_paragraphs(&ctx(), id, MergeDirection::Forward);
        }
        9 if doc.story_len("fn1").is_ok() && rng.below(2) == 0 => {
            let _ = doc.delete_story("fn1");
        }
        9 => {
            let _ = doc.seed_story("fn1", &[paragraph("note"), paragraph("more")]);
        }
        10..=12 => {
            peer.undo.undo();
        }
        13 | 14 => {
            peer.undo.redo();
        }
        _ => {}
    }
    peer.undo.add_undo_barrier();
}

/// Random edit, Undo and Redo schedules on two to four peers, delivered
/// pairwise in random orders. After every delivery no peer may hold an item
/// garbage collection cut off from its parent; at the end the peers
/// converge. On yrs without the GC fix, 19 of the first 100 schedules leave
/// orphans; without the redo parent check (Yjs #757), schedules 173 and 490
/// diverge.
#[test]
fn random_schedules_leave_no_item_without_its_parent() {
    let schedules: u64 = std::env::var("YRS_GC_SCHEDULES").map_or(500, |n| n.parse().unwrap());
    for schedule in 0..schedules {
        let mut rng = Rng(schedule.wrapping_mul(0x9E37_79B9_7F4A_7C15) | 3);
        let seed = EditingDoc::new(500);
        let body = ["alpha beta", "gamma", "delta epsilon", "zeta"].map(paragraph);
        seed.seed_story("body", &body).unwrap();
        seed.seed_story("fn1", &[paragraph("note"), paragraph("more")])
            .unwrap();
        let state = seed.encode_state_as_update_v1();
        let mut peers: Vec<Peer> = (0..2 + rng.below(3) as u64)
            .map(|index| {
                let doc = EditingDoc::new(600 + index);
                doc.apply_update_v1(&state).unwrap();
                let mut undo = doc.undo_manager();
                undo.set_capture_mode(UndoCaptureMode::Manual);
                Peer { doc, undo }
            })
            .collect();
        let count = peers.len();
        for round in 0..4 {
            for step in 0..3 {
                for (index, peer) in peers.iter_mut().enumerate() {
                    random_step(&mut rng, peer, &format!("[{index}.{round}.{step}]"));
                }
            }
            for _ in 0..count * 2 {
                let (from, to) = (rng.below(count), rng.below(count));
                if from != to {
                    sync(&peers[from], &peers[to]);
                    for (index, peer) in peers.iter().enumerate() {
                        let at = format!("schedule {schedule} round {round} sync {from}->{to}");
                        assert_eq!(orphans(peer.doc.yrs_doc()), [], "{at}: peer {index}");
                    }
                }
            }
        }
        for from in &peers {
            for to in &peers {
                if !std::ptr::eq(from, to) {
                    sync(from, to);
                }
            }
        }
        let read = |peer: &Peer| {
            (
                peer.doc.paragraphs("body").ok(),
                peer.doc.paragraphs("fn1").ok(),
            )
        };
        for peer in &peers[1..] {
            assert_eq!(read(peer), read(&peers[0]), "schedule {schedule}");
        }
    }
}
