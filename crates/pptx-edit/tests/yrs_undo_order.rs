//! Patched yrs (third_party/yrs, BetterOffice): an Undo or Redo restores the
//! items of its step in the order the step's deletions list them, as Yjs does,
//! not in hash order. The order gives the restored copies their clocks, and a
//! copy restored right after its left neighbour's merges with it into one
//! block, which a later Undo of the neighbour's own insertion deleted whole.
//! yrs's own tests do not run in CI, so these live here.

use std::collections::HashSet;

use pptx_edit::{DeckSession, EditCtx, ShapeDraft, ShapeRect, TextStyle};
use yrs::block::ClientID;
use yrs::undo::Options;
use yrs::updates::decoder::Decode;
use yrs::{
    Any, Array, ArrayPrelim, ArrayRef, Doc, GetString, Map, MapPrelim, MapRef, Origin, Out,
    ReadTxn, StateVector, Text, TextPrelim, TextRef, Transact, UndoManager, Update,
};

const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

fn manager(doc: &Doc, scope: &MapRef, origin: u64) -> UndoManager {
    let mut manager = UndoManager::with_options(Options {
        capture_timeout_millis: 0,
        tracked_origins: HashSet::from([Origin::from(origin)]),
        ..Options::default()
    });
    manager.expand_scope(doc, scope);
    manager
}

/// The numbers in the list at "list" of `root`.
fn numbers(root: &MapRef, txn: &impl ReadTxn) -> Vec<f64> {
    let list = root.get(txn, "list").unwrap().cast::<ArrayRef>().unwrap();
    list.iter(txn)
        .map(|value| match value {
            Out::Any(Any::Number(number)) => number,
            other => panic!("{other:?}"),
        })
        .collect()
}

#[test]
fn undo_after_a_list_is_restored_removes_only_its_own_insertion() {
    // Each step inserts at the front, so a later item sits left of an earlier
    // one. Replacing the list deletes them all; Undo restores them as copies.
    // Restored in clock order, no copy follows its left neighbour's; restored
    // in hash order, some usually do, and merge with it.
    for _ in 0..16 {
        let doc = Doc::with_client_id(1);
        let root = doc.get_or_insert_map("m");
        let mut undo = manager(&doc, &root, 1);
        root.insert(
            &mut doc.transact_mut_with(1u64),
            "list",
            ArrayPrelim::default(),
        );
        for value in 0..10 {
            let mut txn = doc.transact_mut_with(1u64);
            let list = root.get(&txn, "list").unwrap().cast::<ArrayRef>().unwrap();
            list.insert(&mut txn, 0, value as f64);
        }
        root.insert(&mut doc.transact_mut_with(1u64), "list", 0.0);
        assert!(undo.undo_blocking());
        let mut expected: Vec<f64> = (0..10).rev().map(|value| value as f64).collect();
        for _ in 0..10 {
            assert_eq!(numbers(&root, &doc.transact()), expected);
            assert!(undo.undo_blocking());
            expected.remove(0);
        }
        assert!(numbers(&root, &doc.transact()).is_empty());
    }
}

#[test]
fn undo_after_restoring_a_text_box_removes_only_its_own_typing() {
    // The user types twelve words, each at the start of a new text box,
    // removes the box and presses Undo: the box comes back, and each further
    // Undo must take back one word only. In hash order about one run in five
    // lost more, so the scenario runs 40 times.
    let context = EditCtx::local("local");
    let words: Vec<String> = (0..12).map(|i| format!("w{i} ")).collect();
    let seed = DeckSession::open(FIXTURE, 101)
        .unwrap()
        .encode_state_as_update_v1();
    for _ in 0..40 {
        let session = DeckSession::open_from_update_with_source(&seed, FIXTURE, 101).unwrap();
        let slide = session.snapshot().unwrap().slides[0].id.clone();
        let draft = ShapeDraft {
            name: "Note".to_owned(),
            rect: ShapeRect {
                x: 100_000,
                y: 100_000,
                width: 2_000_000,
                height: 500_000,
            },
            text: String::new(),
            style: TextStyle::default(),
        };
        let shape = session
            .add_text_box(&context, &slide, &draft)
            .unwrap()
            .shape_id;
        let story = session
            .snapshot()
            .unwrap()
            .slides
            .iter()
            .flat_map(|slide| &slide.shapes)
            .find(|candidate| candidate.id == shape)
            .unwrap()
            .text_stories[0]
            .id
            .clone();
        let text = || session.story(&story).unwrap().plain_text();
        for word in &words {
            session.add_undo_barrier();
            session
                .insert_text(&context, &story, 0, word, &TextStyle::default())
                .unwrap();
        }
        session.add_undo_barrier();
        session.remove_shape(&context, &slide, &shape).unwrap();
        session.add_undo_barrier();
        assert!(session.undo());
        for typed in (0..words.len()).rev() {
            let expected: String = words[..=typed].iter().rev().map(String::as_str).collect();
            assert_eq!(text(), expected);
            assert!(session.undo());
        }
        assert_eq!(text(), "");
    }
}

#[test]
fn undo_after_restoring_a_peers_text_keeps_the_peers_text() {
    // The user writes "X" and then "Hello" in a text; a peer adds " world";
    // the user deletes the text and presses Undo twice. The restored copies
    // of "Hello" and " world" merge (the user's client restores both), and
    // the second Undo must take back "Hello" only (Yjs 13.6.31 deletes both).
    let user = Doc::with_client_id(101);
    let peer = Doc::with_client_id(202);
    let root = user.get_or_insert_map("stories");
    let mut undo = manager(&user, &root, 101);
    {
        let mut txn = user.transact_mut_with(101u64);
        let text = root.insert(&mut txn, "s", TextPrelim::new(""));
        text.insert(&mut txn, 0, "X");
    }
    let story = |doc: &Doc| -> TextRef {
        let root = doc.get_or_insert_map("stories");
        let txn = doc.transact();
        root.get(&txn, "s").unwrap().cast::<TextRef>().unwrap()
    };
    let typed = story(&user);
    typed.insert(&mut user.transact_mut_with(101u64), 1, "Hello");
    let update = user
        .transact()
        .encode_state_as_update_v1(&peer.transact().state_vector());
    peer.transact_mut()
        .apply_update(Update::decode_v1(&update).unwrap())
        .unwrap();
    let theirs = story(&peer);
    theirs.insert(&mut peer.transact_mut(), 6, " world");
    let update = peer
        .transact()
        .encode_state_as_update_v1(&user.transact().state_vector());
    user.transact_mut_with("remote")
        .apply_update(Update::decode_v1(&update).unwrap())
        .unwrap();
    root.remove(&mut user.transact_mut_with(101u64), "s");
    assert!(undo.undo_blocking());
    let restored = story(&user);
    assert_eq!(restored.get_string(&user.transact()), "XHello world");
    assert!(undo.undo_blocking());
    assert_eq!(restored.get_string(&user.transact()), "X world");
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

/// Every map and array of `doc` under its root map "m", with their paths, in
/// a fixed order (keys sorted).
fn containers(doc: &Doc) -> Vec<Out> {
    fn walk(value: Out, txn: &impl ReadTxn, depth: usize, out: &mut Vec<Out>) {
        match &value {
            Out::YMap(map) => {
                out.push(value.clone());
                if depth < 3 {
                    let mut entries: Vec<(String, Out)> = map
                        .iter(txn)
                        .map(|(key, value)| (key.to_owned(), value))
                        .collect();
                    entries.sort_by(|a, b| a.0.cmp(&b.0));
                    for (_, value) in entries {
                        walk(value, txn, depth + 1, out);
                    }
                }
            }
            Out::YArray(array) => {
                out.push(value.clone());
                if depth < 3 {
                    for value in array.iter(txn) {
                        walk(value, txn, depth + 1, out);
                    }
                }
            }
            _ => {}
        }
    }
    let mut out = Vec::new();
    let root = doc.get_or_insert_map("m");
    walk(Out::YMap(root), &doc.transact(), 0, &mut out);
    out
}

/// One random schedule of nested map and array edits, Undo, Redo and syncs on
/// 2-4 peers; every peer's whole state at the end. Garbage collection is off:
/// on capy-ci 6779b9ea a type collected under an item Undo keeps leaves that
/// item pointing into freed memory, so encoding it is not repeatable (fixed
/// on capy/yrs-gc-fix).
fn schedule(seed: u64) -> Vec<Vec<u8>> {
    let mut rng = Rng(seed.wrapping_mul(0x9E37_79B9_7F4A_7C15) | 1);
    let peers = 2 + rng.below(3);
    let docs: Vec<Doc> = (0..peers)
        .map(|i| {
            Doc::with_options(yrs::Options {
                skip_gc: true,
                ..yrs::Options::with_client_id(ClientID::new(100 + i as u64))
            })
        })
        .collect();
    let mut managers: Vec<UndoManager> = docs
        .iter()
        .enumerate()
        .map(|(i, doc)| manager(doc, &doc.get_or_insert_map("m"), 100 + i as u64))
        .collect();
    for _ in 0..80 {
        let i = rng.below(peers);
        let doc = &docs[i];
        let origin = 100 + i as u64;
        match rng.below(20) {
            0..=8 => {
                let all = containers(doc);
                let target = all[rng.below(all.len())].clone();
                let n = rng.below(1_000) as f64;
                let kind = rng.below(3);
                let mut txn = doc.transact_mut_with(origin);
                match target {
                    Out::YMap(map) => {
                        let key = format!("k{}", rng.below(4));
                        match kind {
                            0 => {
                                map.insert(&mut txn, key, n);
                            }
                            1 => {
                                map.insert(&mut txn, key, MapPrelim::from([("v", Any::from(n))]));
                            }
                            _ => {
                                map.insert(&mut txn, key, ArrayPrelim::from([Any::from(n)]));
                            }
                        }
                    }
                    Out::YArray(array) => {
                        let len = array.len(&txn);
                        if len > 0 && kind == 0 {
                            array.remove(&mut txn, rng.below(len as usize) as u32);
                        } else {
                            let at = rng.below(len as usize + 1) as u32;
                            array.insert(&mut txn, at, ArrayPrelim::from([Any::from(n)]));
                        }
                    }
                    _ => unreachable!(),
                }
            }
            9..=12 => {
                managers[i].undo_blocking();
            }
            13..=14 => {
                managers[i].redo_blocking();
            }
            15 => managers[i].clear_all(),
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
    docs.iter()
        .map(|doc| {
            doc.transact()
                .encode_state_as_update_v1(&StateVector::default())
        })
        .collect()
}

#[test]
fn random_undo_schedules_build_the_same_documents_every_time() {
    // Undo and Redo give restored copies their clocks in a fixed order, so the
    // same edits make the same documents.
    for seed in 0..60 {
        assert!(schedule(seed) == schedule(seed), "seed {seed}");
    }
}
