//! The patched yrs's Undo, as CI runs it: `third_party/yrs` sits outside the
//! workspace, so its own unit tests never run there.

use std::sync::Arc;

use yrs::types::ToJson;
use yrs::undo::{Options, UndoManager};
use yrs::updates::decoder::Decode;
use yrs::{Any, ClientID, Doc, ID, Map, MapPrelim, MapRef, ReadTxn, Transact, Update};

/// Every peer takes what it lacks from every other, untracked.
fn exchange(docs: &[&Doc]) {
    for from in docs {
        for to in docs {
            if !std::ptr::eq(from, to) {
                let known = to.transact().state_vector();
                let diff = from.transact().encode_diff_v1(&known);
                let update = Update::decode_v1(&diff).unwrap();
                to.transact_mut_with("remote").apply_update(update).unwrap();
            }
        }
    }
}

/// An Undo manager over `map` with one step per transaction.
fn manager(doc: &Doc, map: &MapRef) -> UndoManager<()> {
    let mut manager = UndoManager::with_options(Options {
        capture_timeout_millis: 0,
        ..Options::default()
    });
    manager.expand_scope(doc, map);
    manager
}

fn json(doc: &Doc, map: &MapRef) -> Any {
    map.to_json(&doc.transact())
}

fn object(entries: &[(&str, f64)]) -> Any {
    Any::Map(Arc::new(
        entries
            .iter()
            .map(|(key, value)| (key.to_string(), Any::from(*value)))
            .collect(),
    ))
}

fn nested<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> MapRef {
    map.get(txn, key).unwrap().cast::<MapRef>().unwrap()
}

/// `UndoManager::try_process`: an Undo step passes an inserted item whose
/// redone copy garbage collection took with a removed parent, and still
/// undoes the rest of the step, as Yjs's `followRedone` reaches a GC struct.
#[test]
fn undo_passes_an_insertion_whose_redone_copy_was_collected() {
    let (d1, d2) = (Doc::with_client_id(1), Doc::with_client_id(2));
    let (m1, m2) = (d1.get_or_insert_map("m"), d2.get_or_insert_map("m"));
    m1.insert(&mut d1.transact_mut(), "p", MapPrelim::default());
    exchange(&[&d1, &d2]);
    let mut undo = manager(&d2, &m2);
    {
        // Step 1: "x" in p and "y" beside it.
        let mut txn = d2.transact_mut();
        nested(&m2, &txn, "p").insert(&mut txn, "x", 1);
        m2.insert(&mut txn, "y", 1);
    }
    m2.remove(&mut d2.transact_mut(), "p"); // step 2
    undo.undo_blocking(); // restores p as a redone copy, with "x" in it
    {
        let mut txn = d2.transact_mut();
        nested(&m2, &txn, "p").insert(&mut txn, "x", 2); // step 3
    }
    undo.undo_blocking();
    m2.insert(&mut d2.transact_mut(), "z", 1); // clears the redo stack: un-keeps p's copy

    exchange(&[&d1, &d2]);
    m1.remove(&mut d1.transact_mut(), "p");
    exchange(&[&d1, &d2]); // collects p's copy and the "x" copies in it

    undo.undo_blocking(); // takes back "z"
    undo.undo_blocking(); // takes back step 1: "y" goes, though "x" is gone
    exchange(&[&d1, &d2]);
    assert_eq!(json(&d2, &m2), object(&[]));
    assert_eq!(json(&d1, &m1), object(&[]));
}

/// `Item::gc`: clearing the redo stack un-keeps map "p" though "number",
/// which the undo stack still holds, stays kept. Collecting "p" after a
/// remote removal takes "number" with it (upstream y-crdt #667, PR #682).
#[test]
fn gc_collects_kept_children_of_a_collected_type() {
    let (d1, d2) = (Doc::with_client_id(1), Doc::with_client_id(2));
    let (s1, s2) = (
        d1.get_or_insert_map("styles"),
        d2.get_or_insert_map("styles"),
    );
    s1.insert(
        &mut d1.transact_mut(),
        "p",
        MapPrelim::from([("kind", "number")]),
    );
    exchange(&[&d1, &d2]);
    let number = ID::new(ClientID::new(1), 1);
    let mut undo = manager(&d2, &s2);
    {
        let mut txn = d2.transact_mut();
        nested(&s2, &txn, "p").insert(&mut txn, "kind", "bullet");
    }
    undo.undo_blocking();
    s2.insert(&mut d2.transact_mut(), "q", MapPrelim::default()); // clears the redo stack

    s1.remove(&mut d1.transact_mut(), "p");
    exchange(&[&d1, &d2]);
    assert!(d2.transact().store().get_item(&number).is_none());
    while undo.undo_blocking() {}
    exchange(&[&d1, &d2]);
    assert_eq!(json(&d1, &s1), json(&d2, &s2));
}

/// `Item::redo` (Epo 2026-10-08 narrow Undo rule): d2's set of "x" wins over
/// d1's concurrent one. Undo on d2 restores the value its set replaced (Yjs
/// leaves the key without one); a later live value from the other peer is
/// kept.
#[test]
fn undo_of_a_concurrent_map_set_restores_the_replaced_value() {
    let (d1, d2) = (Doc::with_client_id(1), Doc::with_client_id(2));
    let (m1, m2) = (d1.get_or_insert_map("m"), d2.get_or_insert_map("m"));
    m1.insert(&mut d1.transact_mut(), "x", 0);
    exchange(&[&d1, &d2]);
    let mut undo = manager(&d2, &m2);
    m1.insert(&mut d1.transact_mut(), "x", 1);
    m2.insert(&mut d2.transact_mut(), "x", 2);
    exchange(&[&d1, &d2]);
    assert_eq!(json(&d2, &m2), object(&[("x", 2.0)]));

    undo.undo_blocking();
    exchange(&[&d1, &d2]);
    assert_eq!(json(&d1, &m1), object(&[("x", 0.0)]));
    assert_eq!(json(&d2, &m2), object(&[("x", 0.0)]));

    undo.redo_blocking();
    exchange(&[&d1, &d2]);
    m1.insert(&mut d1.transact_mut(), "x", 3);
    exchange(&[&d1, &d2]);
    undo.undo_blocking();
    exchange(&[&d1, &d2]);
    assert_eq!(json(&d1, &m1), object(&[("x", 3.0)]));
    assert_eq!(json(&d2, &m2), object(&[("x", 3.0)]));
}

/// `UndoManager::clear_undo` and `clear_redo` clear only their own stack.
#[test]
fn clearing_one_stack_keeps_the_other() {
    let doc = Doc::with_client_id(1);
    let map = doc.get_or_insert_map("m");
    let mut undo = manager(&doc, &map);
    map.insert(&mut doc.transact_mut(), "a", 1);
    map.insert(&mut doc.transact_mut(), "b", 1);
    undo.undo_blocking();
    assert!(undo.can_undo() && undo.can_redo());

    undo.clear_undo();
    assert!(!undo.can_undo() && undo.can_redo());
    undo.redo_blocking();
    assert_eq!(json(&doc, &map), object(&[("a", 1.0), ("b", 1.0)]));

    map.insert(&mut doc.transact_mut(), "c", 1);
    undo.undo_blocking();
    undo.clear_redo();
    assert!(undo.can_undo() && !undo.can_redo());
    undo.undo_blocking();
    assert_eq!(json(&doc, &map), object(&[("a", 1.0)]));
}

/// `Item::redo` keeps the restored value's parents, as Yjs's `keepItem`
/// does: after clearing the redo stack un-kept map "p", Undo restores "k"
/// into it and keeps "p" again, so a peer's removal of "p" does not collect
/// it or the restored value (Yjs 13.6.31 the same:
/// `yrs-gc/probes/yjs_redo_keeps_parent.cjs`).
#[test]
fn redo_keeps_the_restored_values_parent() {
    let (d1, d2) = (Doc::with_client_id(1), Doc::with_client_id(2));
    let (m1, m2) = (d1.get_or_insert_map("m"), d2.get_or_insert_map("m"));
    m1.insert(
        &mut d1.transact_mut(),
        "p",
        MapPrelim::from([("k", 0), ("j", 0)]),
    );
    exchange(&[&d1, &d2]);
    let mut undo = manager(&d2, &m2);
    {
        let mut txn = d2.transact_mut();
        nested(&m2, &txn, "p").remove(&mut txn, "k"); // step 1: a pure deletion
    }
    {
        let mut txn = d2.transact_mut();
        nested(&m2, &txn, "p").insert(&mut txn, "j", 1); // step 2
    }
    undo.undo_blocking(); // takes back step 2
    undo.clear_redo(); // un-keeps "p"
    undo.undo_blocking(); // takes back step 1: restores "k" as a copy
    let restored = {
        let txn = d2.transact();
        let p = nested(&m2, &txn, "p");
        let (_, item) = p
            .as_ref()
            .map_items()
            .find(|(key, _)| &***key == "k")
            .unwrap();
        *item.id()
    };

    m1.remove(&mut d1.transact_mut(), "p");
    exchange(&[&d1, &d2]);
    assert!(d2.transact().store().get_item(&restored).is_some());
}
