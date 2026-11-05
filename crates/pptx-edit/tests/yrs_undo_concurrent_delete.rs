//! The patched yrs `UndoManager::adopt_concurrent_edits` (Epo 2026-11-04, Q549 A) on plain
//! yrs types: an Undo brings back deleted content with what a peer did inside it meanwhile.
use std::collections::HashSet;
use std::sync::Arc;

use yrs::undo::{Options, UndoManager};
use yrs::updates::decoder::Decode;
use yrs::{
    ClientID, Doc, GetString, Map, MapPrelim, MapRef, OffsetKind, Origin, Out, ReadTxn, Text,
    TextPrelim, TextRef, Transact, Update,
};

const SYSTEM: &str = "system";

fn doc(client: u64) -> Doc {
    let mut options = yrs::Options::with_client_id(ClientID::new(client));
    options.offset_kind = OffsetKind::Utf16;
    options.cleanup_formatting = false;
    Doc::with_options(options)
}

/// A manager as the engines build theirs: it tracks its client's transactions.
fn manager(doc: &Doc, client: u64, root: &MapRef, adopting: bool) -> UndoManager<()> {
    let mut manager = UndoManager::with_options(Options {
        capture_timeout_millis: 0,
        tracked_origins: HashSet::from([Origin::from(client)]),
        capture_transaction: None,
        timestamp: Arc::new(|| 0),
        init_undo_stack: Vec::new(),
        init_redo_stack: Vec::new(),
    });
    manager.expand_scope(doc, root);
    if adopting {
        manager.adopt_concurrent_edits();
    }
    manager
}

fn sync(from: &Doc, to: &Doc) {
    let update = from
        .transact()
        .encode_state_as_update_v1(&to.transact().state_vector());
    to.transact_mut()
        .apply_update(Update::decode_v1(&update).unwrap())
        .unwrap();
}

struct Pair {
    left: Doc,
    right: Doc,
    left_root: MapRef,
    right_root: MapRef,
    undo: UndoManager<()>,
}

/// `shape` holds a text `story` ("hello") and a value `x` (1), on both replicas.
fn pair(adopting: bool) -> Pair {
    let seed = doc(1);
    let root = seed.get_or_insert_map("shapes");
    {
        let mut txn = seed.transact_mut();
        let shape = root.insert(&mut txn, "shape", MapPrelim::default());
        shape.insert(&mut txn, "x", 1_i64);
        shape.insert(&mut txn, "story", TextPrelim::new("hello"));
    }
    let (left, right) = (doc(2), doc(3));
    let (left_root, right_root) = (
        left.get_or_insert_map("shapes"),
        right.get_or_insert_map("shapes"),
    );
    sync(&seed, &left);
    sync(&seed, &right);
    let undo = manager(&left, 2, &left_root, adopting);
    Pair {
        left,
        right,
        left_root,
        right_root,
        undo,
    }
}

fn shape<T: ReadTxn>(root: &MapRef, txn: &T) -> Option<MapRef> {
    root.get(txn, "shape").and_then(|value| value.cast().ok())
}

fn story(doc: &Doc, root: &MapRef) -> Option<String> {
    let txn = doc.transact();
    let text: TextRef = shape(root, &txn)?.get(&txn, "story")?.cast().ok()?;
    Some(text.get_string(&txn))
}

fn x(doc: &Doc, root: &MapRef) -> Option<i64> {
    let txn = doc.transact();
    match shape(root, &txn)?.get(&txn, "x")? {
        Out::Any(yrs::Any::Number(value)) => Some(value as i64),
        Out::Any(yrs::Any::BigInt(value)) => Some(value),
        _ => None,
    }
}

fn delete(pair: &Pair) {
    let mut txn = pair.left.transact_mut_with(2_u64);
    pair.left_root.remove(&mut txn, "shape");
}

fn type_right(pair: &Pair, at: u32, text: &str) {
    let mut txn = pair.right.transact_mut();
    let shape = shape(&pair.right_root, &txn).unwrap();
    let story: TextRef = shape.get(&txn, "story").unwrap().cast().unwrap();
    story.insert(&mut txn, at, text);
}

fn move_right(pair: &Pair, value: i64) {
    let mut txn = pair.right.transact_mut();
    let shape = shape(&pair.right_root, &txn).unwrap();
    shape.insert(&mut txn, "x", value);
}

fn undo(pair: &mut Pair) {
    assert!(pair.undo.undo_blocking());
}

fn assert_reads(pair: &Pair, text: &str, value: i64) {
    sync(&pair.left, &pair.right);
    sync(&pair.right, &pair.left);
    for (doc, root) in [
        (&pair.left, &pair.left_root),
        (&pair.right, &pair.right_root),
    ] {
        assert_eq!(story(doc, root).as_deref(), Some(text));
        assert_eq!(x(doc, root), Some(value));
    }
}

#[test]
fn without_adopting_the_peers_edits_are_lost() {
    let mut pair = pair(false);
    delete(&pair);
    type_right(&pair, 5, " PEER");
    move_right(&pair, 7);
    sync(&pair.right, &pair.left);
    undo(&mut pair);
    sync(&pair.left, &pair.right);
    assert_eq!(story(&pair.left, &pair.left_root).as_deref(), Some("hello"));
    assert_eq!(
        story(&pair.right, &pair.right_root).as_deref(),
        Some("hello")
    );
}

#[test]
fn an_undo_restores_the_peers_text_and_value() {
    let mut pair = pair(true);
    delete(&pair);
    type_right(&pair, 5, " PEER");
    move_right(&pair, 7);
    sync(&pair.right, &pair.left);
    undo(&mut pair);
    assert_reads(&pair, "hello PEER", 7);
}

#[test]
fn text_typed_mid_item_lands_in_place() {
    let mut pair = pair(true);
    delete(&pair);
    type_right(&pair, 2, "PEER");
    sync(&pair.right, &pair.left);
    undo(&mut pair);
    assert_reads(&pair, "hePEERllo", 1);
}

#[test]
fn updates_one_at_a_time_land_as_one() {
    let mut pair = pair(true);
    delete(&pair);
    for (index, letter) in ["P", "E", "E", "R"].iter().enumerate() {
        type_right(&pair, 5 + index as u32, letter);
        sync(&pair.right, &pair.left);
    }
    move_right(&pair, 7);
    sync(&pair.right, &pair.left);
    move_right(&pair, 8);
    sync(&pair.right, &pair.left);
    undo(&mut pair);
    assert_reads(&pair, "helloPEER", 8);
}

/// The peer's own later deletion of what it typed, and of seeded text, stands.
#[test]
fn what_the_peer_deleted_stays_deleted() {
    let mut pair = pair(true);
    delete(&pair);
    type_right(&pair, 5, " PEER");
    {
        let mut txn = pair.right.transact_mut();
        let shape = shape(&pair.right_root, &txn).unwrap();
        let story: TextRef = shape.get(&txn, "story").unwrap().cast().unwrap();
        story.remove_range(&mut txn, 8, 2); // "ER" of "hello PEER"
        story.remove_range(&mut txn, 0, 1); // "h"
    }
    sync(&pair.right, &pair.left);
    undo(&mut pair);
    assert_reads(&pair, "ello PE", 1);
}

/// The same deletion arriving in a later update.
#[test]
fn a_later_update_deleting_what_landed_takes_it_out() {
    let mut pair = pair(true);
    delete(&pair);
    type_right(&pair, 5, " PEER");
    sync(&pair.right, &pair.left);
    {
        let mut txn = pair.right.transact_mut();
        let shape = shape(&pair.right_root, &txn).unwrap();
        let story: TextRef = shape.get(&txn, "story").unwrap().cast().unwrap();
        story.remove_range(&mut txn, 5, 3);
    }
    sync(&pair.right, &pair.left);
    undo(&mut pair);
    assert_reads(&pair, "helloER", 1);
}

/// The peer heard the delete before its new value left it: its collection took that value, and
/// what it sends names the old one as deleted. The old value still comes back.
#[test]
fn a_value_its_writers_collection_took_leaves_the_old_one_in_the_step() {
    let mut pair = pair(true);
    delete(&pair);
    move_right(&pair, 7);
    sync(&pair.left, &pair.right);
    sync(&pair.right, &pair.left);
    undo(&mut pair);
    assert_reads(&pair, "hello", 1);
}

/// A step before the delete set a value over the one the peer replaced too. The peer's update
/// names the old value as deleted, which is a replacement, not a removal: undoing past the
/// restore brings it back, and the key keeps a value.
#[test]
fn undoing_past_the_restore_keeps_a_value_the_peer_replaced() {
    let seed = doc(1);
    let root = seed.get_or_insert_map("shapes");
    {
        let mut txn = seed.transact_mut();
        let shape = root.insert(&mut txn, "shape", MapPrelim::default());
        shape.insert(&mut txn, "x", 1_i64);
    }
    // The undoer's client id above the peer's: its value is the key's last.
    let (peer, undoer) = (doc(2), doc(3));
    let (peer_root, undoer_root) = (
        peer.get_or_insert_map("shapes"),
        undoer.get_or_insert_map("shapes"),
    );
    sync(&seed, &peer);
    sync(&seed, &undoer);
    let mut undo = manager(&undoer, 3, &undoer_root, true);
    {
        let mut txn = undoer.transact_mut_with(3_u64);
        let shape = shape(&undoer_root, &txn).unwrap();
        shape.insert(&mut txn, "x", 5_i64);
    }
    undoer_root.remove(&mut undoer.transact_mut_with(3_u64), "shape");
    {
        let mut txn = peer.transact_mut();
        let shape = shape(&peer_root, &txn).unwrap();
        shape.insert(&mut txn, "x", 7_i64);
    }
    sync(&peer, &undoer);
    assert!(undo.undo_blocking());
    assert_eq!(x(&undoer, &undoer_root), Some(5));
    assert!(undo.undo_blocking());
    assert_eq!(x(&undoer, &undoer_root), Some(1));
}

/// Both delete the whole shape, the peer after typing: its typing comes back, the delete
/// does not stop the Undo.
#[test]
fn the_peers_delete_of_the_whole_parent_does_not_stop_the_undo() {
    let mut pair = pair(true);
    delete(&pair);
    type_right(&pair, 5, " PEER");
    // Each edit goes out as it is made: the peer's own collection takes it once it deletes.
    sync(&pair.right, &pair.left);
    {
        let mut txn = pair.right.transact_mut();
        pair.right_root.remove(&mut txn, "shape");
    }
    sync(&pair.right, &pair.left);
    undo(&mut pair);
    assert_reads(&pair, "hello PEER", 1);
}

/// The peer's update arrives after the Undo: the host forwards it into the restored copy.
#[test]
fn an_update_after_the_undo_is_forwarded() {
    let mut pair = pair(true);
    delete(&pair);
    undo(&mut pair);
    type_right(&pair, 5, " LATE");
    move_right(&pair, 9);
    sync(&pair.right, &pair.left);
    assert_eq!(story(&pair.left, &pair.left_root).as_deref(), Some("hello"));
    assert!(pair.undo.forward_concurrent_edits(SYSTEM));
    assert!(!pair.undo.forward_concurrent_edits(SYSTEM));
    assert_reads(&pair, "hello LATE", 9);
    // Not a step: the next Undo finds nothing of it; the Redo removes the copy with its parent.
    assert!(!pair.undo.can_undo());
    assert!(pair.undo.redo_blocking());
    sync(&pair.left, &pair.right);
    assert_eq!(story(&pair.right, &pair.right_root), None);
    assert!(pair.undo.undo_blocking());
    assert_reads(&pair, "hello LATE", 9);
}

/// After Undo and Redo the content is deleted again; what lands then waits in that step.
#[test]
fn an_update_between_redo_and_the_next_undo_lands_in_that_step() {
    let mut pair = pair(true);
    delete(&pair);
    undo(&mut pair);
    assert!(pair.undo.redo_blocking());
    type_right(&pair, 5, " PEER");
    sync(&pair.right, &pair.left);
    assert!(!pair.undo.forward_concurrent_edits(SYSTEM));
    undo(&mut pair);
    assert_reads(&pair, "hello PEER", 1);
}

/// The local user added the shape and undid that; a peer's typing in it comes back on Redo.
#[test]
fn a_redo_of_an_undone_insert_keeps_the_peers_typing() {
    let (left, right) = (doc(2), doc(3));
    let left_root = left.get_or_insert_map("shapes");
    let right_root = right.get_or_insert_map("shapes");
    let mut undo = manager(&left, 2, &left_root, true);
    {
        let mut txn = left.transact_mut_with(2_u64);
        let shape = left_root.insert(&mut txn, "shape", MapPrelim::default());
        shape.insert(&mut txn, "story", TextPrelim::new("hello"));
    }
    sync(&left, &right);
    assert!(undo.undo_blocking());
    {
        let mut txn = right.transact_mut();
        let shape = shape(&right_root, &txn).unwrap();
        let story: TextRef = shape.get(&txn, "story").unwrap().cast().unwrap();
        story.insert(&mut txn, 5, " PEER");
    }
    sync(&right, &left);
    assert!(undo.redo_blocking());
    sync(&left, &right);
    assert_eq!(story(&left, &left_root).as_deref(), Some("hello PEER"));
    assert_eq!(story(&right, &right_root).as_deref(), Some("hello PEER"));
}

/// Another person's delete is no step here: nothing is kept for it.
#[test]
fn what_lands_in_a_peers_delete_is_collected() {
    let mut pair = pair(true);
    {
        let mut txn = pair.right.transact_mut();
        pair.right_root.remove(&mut txn, "shape");
    }
    let third = doc(4);
    let third_root = third.get_or_insert_map("shapes");
    sync(&pair.left, &third);
    sync(&pair.right, &pair.left);
    {
        let mut txn = third.transact_mut();
        let shape = shape(&third_root, &txn).unwrap();
        let story: TextRef = shape.get(&txn, "story").unwrap().cast().unwrap();
        story.insert(&mut txn, 5, "SECRET");
    }
    sync(&third, &pair.left);
    let state = pair
        .left
        .transact()
        .encode_state_as_update_v1(&yrs::StateVector::default());
    assert!(!state.windows(6).any(|window| window == b"SECRET"));
    assert!(!pair.undo.can_undo());
    let _ = &mut pair;
}

/// Clearing the stacks releases what was kept: the content goes with the next collection.
#[test]
fn clearing_history_releases_what_was_kept() {
    let mut pair = pair(true);
    delete(&pair);
    type_right(&pair, 5, " PEER");
    sync(&pair.right, &pair.left);
    pair.undo.clear_all();
    assert!(!pair.undo.can_undo());
    pair.left.transact_mut().gc(None);
    let state = pair
        .left
        .transact()
        .encode_state_as_update_v1(&yrs::StateVector::default());
    assert!(!state.windows(4).any(|window| window == b"PEER"));
}

/// What adopting costs a replica typing keeps coming to: one peer's keystrokes into a live
/// text, with and without adopting. Prints the two; the check is for an order of magnitude.
#[test]
fn adopting_costs_a_live_typing_update_nothing_measurable() {
    let run = |adopting: bool| {
        let pair = pair(adopting);
        let updates: Vec<Vec<u8>> = (0..20_000u32)
            .map(|index| {
                let before = pair.right.transact().state_vector();
                type_right(&pair, 5 + index, "x");
                pair.right.transact().encode_state_as_update_v1(&before)
            })
            .collect();
        let start = std::time::Instant::now();
        for update in updates {
            pair.left
                .transact_mut()
                .apply_update(Update::decode_v1(&update).unwrap())
                .unwrap();
        }
        start.elapsed()
    };
    let (off, on) = (run(false), run(true));
    let (off2, on2) = (run(false), run(true));
    println!("20,000 remote updates: off {off:?} {off2:?}, on {on:?} {on2:?}");
    assert!(on.min(on2) < off.max(off2) * 2);
}
