//! Patched yrs (third_party/yrs, BetterOffice): structs that arrive before a
//! struct they depend on wait in a pending update and are applied as soon as
//! the missing struct comes, whatever order the updates arrive in. yrs places
//! an early struct behind a skip, so its clock counts past the hole; the
//! pending update waits on the held state instead (`BlockStore::get_state`),
//! and lists every client its stacked structs wait on. yrs's own tests do not
//! run in CI, so these live here.

use std::collections::HashSet;
use std::sync::{Arc, Mutex};

use yrs::block::ClientID;
use yrs::types::ToJson;
use yrs::undo::Options;
use yrs::updates::decoder::Decode;
use yrs::{
    Any, Array, ArrayPrelim, Doc, GetString, Map, MapPrelim, Origin, Out, ReadTxn, Text, Transact,
    UndoManager, Update,
};

fn apply(doc: &Doc, update: &[u8]) {
    doc.transact_mut()
        .apply_update(Update::decode_v1(update).unwrap())
        .unwrap();
}

fn pending(doc: &Doc) -> bool {
    let txn = doc.transact();
    txn.store().pending_update().is_some() || txn.store().pending_ds().is_some()
}

fn text(doc: &Doc, name: &str) -> String {
    let text = doc.get_or_insert_text(name);
    text.get_string(&doc.transact())
}

/// Each transaction of `edit` on `doc`, as an update.
fn recorded(doc: &Doc, edits: &[&dyn Fn(&Doc)]) -> Vec<Vec<u8>> {
    let log = Arc::new(Mutex::new(Vec::new()));
    let sink = log.clone();
    let subscription = doc
        .observe_update_v1(move |_, event| sink.lock().unwrap().push(event.update.clone()))
        .unwrap();
    for edit in edits {
        edit(doc);
    }
    drop(subscription);
    log.lock().unwrap().clone()
}

#[test]
fn a_struct_waiting_on_a_skipped_unit_applies_when_the_unit_arrives() {
    // U2 depends on U1; U3 does not and is placed behind a skip over U1 and
    // U2, which moves the client's clock past the hole U1 then fills.
    let author = Doc::with_client_id(1);
    let a = author.get_or_insert_text("a");
    let b = author.get_or_insert_text("b");
    let updates = recorded(
        &author,
        &[
            &|doc| a.insert(&mut doc.transact_mut(), 0, "A"),
            &|doc| a.insert(&mut doc.transact_mut(), 1, "B"),
            &|doc| b.insert(&mut doc.transact_mut(), 0, "C"),
        ],
    );
    for order in [
        [0, 1, 2],
        [0, 2, 1],
        [1, 0, 2],
        [1, 2, 0],
        [2, 0, 1],
        [2, 1, 0],
    ] {
        let doc = Doc::with_client_id(2);
        for i in order {
            apply(&doc, &updates[i]);
        }
        assert_eq!(text(&doc, "a"), "AB", "{order:?}");
        assert_eq!(text(&doc, "b"), "C", "{order:?}");
        assert!(!pending(&doc), "{order:?}");
    }
}

#[test]
fn structs_stacked_on_a_struct_that_waits_apply_when_its_unit_arrives() {
    // One update brings B's "b" (after A's "c"), B's "d", A's map entry and
    // A's "e" (after "d"), but not "c". "b" waits on "c"; A's entry goes in
    // behind a skip over "c"; "e" then waits on "d", which waits on "b". Only
    // "c" arriving lets them in, so the pending update must wait on A too.
    let a = Doc::with_client_id(1);
    let b = Doc::with_client_id(2);
    let (ta, tb) = (a.get_or_insert_text("t"), b.get_or_insert_text("t"));
    let first = recorded(&a, &[&|doc| ta.insert(&mut doc.transact_mut(), 0, "a")]);
    let c = recorded(&a, &[&|doc| ta.insert(&mut doc.transact_mut(), 1, "c")]);
    apply(&b, &first[0]);
    apply(&b, &c[0]);
    let from_b = recorded(
        &b,
        &[&|doc| tb.insert(&mut doc.transact_mut(), 2, "b"), &|doc| {
            tb.insert(&mut doc.transact_mut(), 3, "d")
        }],
    );
    for update in &from_b {
        apply(&a, update);
    }
    let m = a.get_or_insert_map("m");
    let late = recorded(
        &a,
        &[
            &|doc| {
                m.insert(&mut doc.transact_mut(), "k", 1.0);
            },
            &|doc| ta.insert(&mut doc.transact_mut(), 4, "e"),
        ],
    );
    let bundle = Update::merge_updates(
        from_b
            .iter()
            .chain(&late)
            .map(|update| Update::decode_v1(update).unwrap())
            .collect::<Vec<_>>(),
    );
    let doc = Doc::with_client_id(3);
    apply(&doc, &first[0]);
    doc.transact_mut().apply_update(bundle).unwrap();
    apply(&doc, &c[0]);
    assert_eq!(text(&doc, "t"), "acbde");
    let entry = doc.get_or_insert_map("m").get(&doc.transact(), "k");
    assert!(matches!(entry, Some(Out::Any(Any::Number(n))) if n == 1.0));
    assert!(!pending(&doc));
    assert_eq!(doc.transact().state_vector(), a.transact().state_vector());
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

/// Every update 2-4 peers produce in one random session of map, array and
/// text edits, Undo, Redo and syncs, in the order they were made.
fn session(seed: u64) -> Vec<Vec<u8>> {
    let mut rng = Rng(seed.wrapping_mul(0x9E37_79B9_7F4A_7C15) | 1);
    let peers = 2 + rng.below(3);
    let log = Arc::new(Mutex::new(Vec::new()));
    let mut docs = Vec::new();
    let mut managers = Vec::new();
    let mut subscriptions = Vec::new();
    for i in 0..peers {
        // Garbage collection off: on capy-ci 6779b9ea a type collected under
        // an item Undo keeps can make a peer encode freed memory.
        let doc = Doc::with_options(yrs::Options {
            skip_gc: true,
            ..yrs::Options::with_client_id(ClientID::new(1_000 + i as u64))
        });
        let mut manager = UndoManager::with_options(Options {
            capture_timeout_millis: 0,
            tracked_origins: HashSet::from([Origin::from(1_000 + i as u64)]),
            ..Options::default()
        });
        manager.expand_scope(&doc, &doc.get_or_insert_map("m"));
        manager.expand_scope(&doc, &doc.get_or_insert_array("a"));
        manager.expand_scope(&doc, &doc.get_or_insert_text("t"));
        let sink = log.clone();
        subscriptions.push(
            doc.observe_update_v1(move |_, event| sink.lock().unwrap().push(event.update.clone()))
                .unwrap(),
        );
        docs.push(doc);
        managers.push(manager);
    }
    for step in 0..80 {
        let i = rng.below(peers);
        let doc = &docs[i];
        let origin = 1_000 + i as u64;
        match rng.below(10) {
            0 => {
                let key = format!("k{}", rng.below(4));
                let value = MapPrelim::from([("v", Any::from(step as f64))]);
                doc.get_or_insert_map("m")
                    .insert(&mut doc.transact_mut_with(origin), key, value);
            }
            1 => {
                let key = format!("k{}", rng.below(4));
                doc.get_or_insert_map("m")
                    .remove(&mut doc.transact_mut_with(origin), &key);
            }
            2 => {
                let array = doc.get_or_insert_array("a");
                let mut txn = doc.transact_mut_with(origin);
                let at = rng.below(array.len(&txn) as usize + 1) as u32;
                array.insert(&mut txn, at, ArrayPrelim::from([Any::from(1.0)]));
            }
            3 => {
                let array = doc.get_or_insert_array("a");
                let mut txn = doc.transact_mut_with(origin);
                let len = array.len(&txn);
                if len > 0 {
                    array.remove(&mut txn, rng.below(len as usize) as u32);
                }
            }
            4 | 5 => {
                let text = doc.get_or_insert_text("t");
                let mut txn = doc.transact_mut_with(origin);
                let at = rng.below(text.len(&txn) as usize + 1) as u32;
                text.insert(&mut txn, at, "ab");
            }
            6 => {
                let text = doc.get_or_insert_text("t");
                let mut txn = doc.transact_mut_with(origin);
                let len = text.len(&txn);
                if len > 1 {
                    text.remove_range(&mut txn, rng.below(len as usize - 1) as u32, 1);
                }
            }
            7 => {
                managers[i].undo_blocking();
            }
            8 => {
                managers[i].redo_blocking();
            }
            _ => {
                let j = rng.below(peers);
                if i != j {
                    let state = docs[j].transact().state_vector();
                    let diff = doc.transact().encode_state_as_update_v1(&state);
                    docs[j]
                        .transact_mut_with("remote")
                        .apply_update(Update::decode_v1(&diff).unwrap())
                        .unwrap();
                }
            }
        }
    }
    drop(subscriptions);
    log.lock().unwrap().clone()
}

fn canonical(value: &Any) -> String {
    match value {
        Any::Map(map) => {
            let mut entries: Vec<String> = map
                .iter()
                .map(|(key, value)| format!("{key}:{}", canonical(value)))
                .collect();
            entries.sort();
            format!("{{{}}}", entries.join(","))
        }
        Any::Array(items) => format!(
            "[{}]",
            items.iter().map(canonical).collect::<Vec<_>>().join(",")
        ),
        other => format!("{other:?}"),
    }
}

fn content(doc: &Doc) -> String {
    let (m, a, t) = (
        doc.get_or_insert_map("m"),
        doc.get_or_insert_array("a"),
        doc.get_or_insert_text("t"),
    );
    let txn = doc.transact();
    format!(
        "{}|{}|{}",
        canonical(&m.to_json(&txn)),
        canonical(&a.to_json(&txn)),
        t.get_string(&txn)
    )
}

#[test]
fn peer_updates_delivered_in_any_order_end_complete_with_nothing_pending() {
    for seed in 0..8 {
        let updates = session(seed);
        let reference = Doc::new();
        for update in &updates {
            apply(&reference, update);
        }
        assert!(!pending(&reference));
        let expected = content(&reference);
        let state = reference.transact().state_vector();
        for round in 0..4 {
            let mut rng = Rng(seed * 31 + round + 1);
            let mut order: Vec<usize> = (0..updates.len()).collect();
            for i in (1..order.len()).rev() {
                order.swap(i, rng.below(i + 1));
            }
            let doc = Doc::new();
            for (k, &i) in order.iter().enumerate() {
                apply(&doc, &updates[i]);
                // Some deliveries repeat an earlier update, as a resync does.
                if round % 2 == 1 && rng.below(5) == 0 {
                    apply(&doc, &updates[order[rng.below(k + 1)]]);
                }
            }
            assert!(!pending(&doc), "seed {seed} round {round}");
            assert_eq!(
                doc.transact().state_vector(),
                state,
                "seed {seed} round {round}"
            );
            assert_eq!(content(&doc), expected, "seed {seed} round {round}");
        }
    }
}
