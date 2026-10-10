//! A yrs transaction answers "what did it insert" from its insert set instead
//! of building whole state vectors, whose cost grows with every client the
//! document ever saw. Its update and its events must stay what the state
//! vectors gave, and applying a peer's keystroke must cost the same at any
//! client count (counted in bytes allocated).

use std::alloc::{GlobalAlloc, Layout, System};
use std::cell::{Cell, RefCell};
use std::rc::Rc;
use std::sync::Arc;

use docx_edit::{EditCtx, EditingDoc, FormatPolicy, Position};
use yrs::types::{Attrs, EntryChange};
use yrs::updates::decoder::Decode;
use yrs::{
    Any, ClientID, DeepObservable, Doc, GetString, Map, MapPrelim, MapRef, Options, ReadTxn,
    StateVector, Text, Transact, Update,
};

fn doc(client: u64) -> Doc {
    Doc::with_options(Options::with_client_id(ClientID::new(client)))
}

/// An update's structs, one line each, sorted: its content as a set.
fn structs(update: &[u8]) -> Vec<String> {
    let update = Update::decode_v1(update).unwrap();
    let mut blocks: Vec<String> = update.blocks().map(|block| format!("{block:?}")).collect();
    blocks.sort();
    blocks
}

struct Rng(u64);

impl Rng {
    fn below(&mut self, n: u32) -> u32 {
        self.0 ^= self.0 << 13;
        self.0 ^= self.0 >> 7;
        self.0 ^= self.0 << 17;
        (self.0 % u64::from(n.max(1))) as u32
    }
}

#[test]
fn each_update_holds_the_structs_a_state_vector_diff_gives() {
    // A document many clients wrote into.
    let main = doc(7);
    let text = main.get_or_insert_text("story");
    for client in 0..400 {
        let writer = doc(1_000 + client);
        let writer_text = writer.get_or_insert_text("story");
        writer
            .transact_mut()
            .apply_update(
                Update::decode_v1(
                    &main
                        .transact()
                        .encode_state_as_update_v1(&StateVector::default()),
                )
                .unwrap(),
            )
            .unwrap();
        {
            let mut txn = writer.transact_mut();
            let length = writer_text.len(&txn);
            writer_text.insert(&mut txn, length / 2, "w");
        }
        let update = writer
            .transact()
            .encode_state_as_update_v1(&main.transact().state_vector());
        main.transact_mut()
            .apply_update(Update::decode_v1(&update).unwrap())
            .unwrap();
    }
    let sent = Rc::new(RefCell::new(Vec::new()));
    let sink = Rc::clone(&sent);
    let _subscription = main
        .observe_update_v1(move |_txn, event| sink.borrow_mut().push(event.update.clone()))
        .unwrap();
    let mut rng = Rng(0x9E37_79B9_7F4A_7C15);
    for step in 0..300 {
        let before = main.transact().state_vector();
        let count = sent.borrow().len();
        {
            let mut txn = main.transact_mut();
            let length = text.len(&txn);
            match rng.below(4) {
                0 => text.insert(&mut txn, rng.below(length + 1), "abc"),
                1 if length > 3 => text.remove_range(&mut txn, rng.below(length - 3), 2),
                2 if length > 3 => {
                    let at = rng.below(length - 3);
                    text.format(
                        &mut txn,
                        at,
                        3,
                        Attrs::from([(Arc::from("bold"), Any::Bool(step % 2 == 0))]),
                    );
                }
                _ => text.insert(&mut txn, 0, "z"),
            }
        }
        let expected = structs(&main.transact().encode_state_as_update_v1(&before));
        match &sent.borrow()[count..] {
            [update] => assert_eq!(structs(update), expected, "step {step}"),
            [] => assert!(expected.is_empty(), "step {step} changed and sent nothing"),
            more => panic!("step {step} sent {} updates", more.len()),
        }
    }
}

/// A struct that arrives before one its client wrote earlier is integrated
/// behind a skip, where the state vector stops. Its transaction's update,
/// read from the insert set, holds that struct; diffing the state vectors, as
/// upstream yrs does, gave an update without blocks. Pinned: the one place
/// the insert set and a state-vector diff differ.
#[test]
fn a_struct_integrated_behind_a_skip_is_in_its_update() {
    let five = ClientID::new(5);
    let writer = doc(5);
    let text = writer.get_or_insert_text("story");
    text.insert(&mut writer.transact_mut(), 0, "abc");
    let first = writer
        .transact()
        .encode_state_as_update_v1(&StateVector::default());
    // 5#3 "X" stays behind; 5#4 "Y" follows "c".
    text.insert(&mut writer.transact_mut(), 0, "X");
    text.insert(&mut writer.transact_mut(), 4, "Y");
    let mut has_x = StateVector::default();
    has_x.set_max(five, 4);
    let only_y = writer.transact().encode_state_as_update_v1(&has_x);
    let main = doc(7);
    let main_text = main.get_or_insert_text("story");
    main.transact_mut()
        .apply_update(Update::decode_v1(&first).unwrap())
        .unwrap();
    let sent = Rc::new(RefCell::new(Vec::new()));
    let sink = Rc::clone(&sent);
    let _subscription = main
        .observe_update_v1(move |_txn, event| sink.borrow_mut().push(event.update.clone()))
        .unwrap();
    main.transact_mut()
        .apply_update(Update::decode_v1(&only_y).unwrap())
        .unwrap();
    assert_eq!(main_text.get_string(&main.transact()), "abcY");
    assert_eq!(main.transact().state_vector().get(&five), 3);
    let sent = sent.borrow();
    assert_eq!(sent.len(), 1);
    let blocks = structs(&sent[0]);
    assert_eq!(blocks.len(), 1, "{blocks:?}");
    assert!(
        blocks[0].contains("<5#4>") && blocks[0].contains("'Y'"),
        "{blocks:?}"
    );
}

#[test]
fn events_tell_new_types_and_entries_from_older_ones() {
    let doc = doc(3);
    let root = doc.get_or_insert_map("root");
    let events = Arc::new(std::sync::Mutex::new(Vec::new()));
    let sink = Arc::clone(&events);
    let _subscription = root.observe_deep(move |txn, deep| {
        for event in deep.iter() {
            if let yrs::types::Event::Map(event) = event {
                let mut keys: Vec<String> = event
                    .keys(txn)
                    .iter()
                    .map(|(key, change)| {
                        let kind = match change {
                            EntryChange::Inserted(_) => "inserted",
                            EntryChange::Updated(_, _) => "updated",
                            EntryChange::Removed(_) => "removed",
                        };
                        format!("{key}:{kind}")
                    })
                    .collect();
                keys.sort();
                sink.lock().unwrap().push(keys.join(","));
            }
        }
    });
    // A nested map made and filled in one transaction: one event, on the root.
    {
        let mut txn = doc.transact_mut();
        let child: MapRef = root.insert(&mut txn, "child", MapPrelim::default());
        child.insert(&mut txn, "inner", 1);
        root.insert(&mut txn, "a", 1);
    }
    assert_eq!(*events.lock().unwrap(), ["a:inserted,child:inserted"]);
    events.lock().unwrap().clear();
    // The older map changes: an event on it; an older entry is updated, an
    // entry set twice in one transaction is inserted.
    {
        let mut txn = doc.transact_mut();
        let child = root.get(&txn, "child").unwrap().cast::<MapRef>().unwrap();
        child.insert(&mut txn, "inner", 2);
        root.insert(&mut txn, "a", 2);
        root.insert(&mut txn, "b", 1);
        root.insert(&mut txn, "b", 2);
    }
    let mut seen = events.lock().unwrap().clone();
    seen.sort();
    assert_eq!(seen, ["a:updated,b:inserted", "inner:updated"]);
}

/// Counts the bytes each thread allocates, so the tests running beside one
/// another in this binary do not count each other's.
struct Counting;

thread_local! {
    static ALLOCATED: Cell<usize> = const { Cell::new(0) };
}

unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let _ = ALLOCATED.try_with(|bytes| bytes.set(bytes.get() + layout.size()));
        unsafe { System.alloc(layout) }
    }

    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        unsafe { System.dealloc(ptr, layout) }
    }

    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        let grown = new_size.saturating_sub(layout.size());
        let _ = ALLOCATED.try_with(|bytes| bytes.set(bytes.get() + grown));
        unsafe { System.realloc(ptr, layout, new_size) }
    }
}

#[global_allocator]
static ALLOCATOR: Counting = Counting;

/// Bytes this thread allocated while `run` ran.
fn allocated(run: impl FnOnce()) -> usize {
    let before = ALLOCATED.with(Cell::get);
    run();
    ALLOCATED.with(Cell::get) - before
}

/// Bytes allocated applying a peer's keystroke to an editor document that
/// `writers` other clients typed into, with an update observer as the WASM
/// session has. The peer wrote before, so the keystroke adds no client.
fn peer_keystroke_bytes(writers: u64) -> usize {
    let ctx = EditCtx::local("Ada", "2026-10-09T00:00:00Z");
    let base = EditingDoc::new(1);
    base.create_story("body", "Hello.", "Normal", "left")
        .unwrap();
    let state = base.encode_state_as_update_v1();
    let replica = |client| {
        let doc = EditingDoc::new(client);
        doc.apply_update_v1(&state).unwrap();
        doc
    };
    let main = replica(7);
    for client in 0..writers {
        let writer = replica(2_000 + client);
        let before = writer.encode_state_vector_v1();
        writer
            .insert_text(&ctx, Position::new("body", 0), "w", FormatPolicy::Inherit)
            .unwrap();
        main.apply_update_v1(&writer.encode_diff_v1(&before).unwrap())
            .unwrap();
    }
    // The peer types after "Hello.", so integrating its keystroke scans no
    // concurrent items.
    let peer = replica(1_000);
    peer.apply_update_v1(&main.encode_state_as_update_v1())
        .unwrap();
    let end = writers as u32 + 6;
    peer.insert_text(&ctx, Position::new("body", end), "a", FormatPolicy::Inherit)
        .unwrap();
    main.apply_update_v1(&peer.encode_diff_v1(&main.encode_state_vector_v1()).unwrap())
        .unwrap();
    let before = peer.encode_state_vector_v1();
    peer.insert_text(
        &ctx,
        Position::new("body", end + 1),
        "b",
        FormatPolicy::Inherit,
    )
    .unwrap();
    let keystroke = peer.encode_diff_v1(&before).unwrap();
    let _subscription = main.yrs_doc().observe_update_v1(|_txn, _event| {}).unwrap();
    allocated(|| main.apply_update_v1(&keystroke).unwrap())
}

#[test]
fn a_peer_keystroke_allocates_the_same_at_any_client_count() {
    // A state vector allocates per client the document saw: with 1,000
    // writers, each one built would add tens of KiB to every keystroke.
    let few = peer_keystroke_bytes(10);
    let many = peer_keystroke_bytes(1_000);
    assert!(
        many <= few + 1_024,
        "a keystroke allocated {many} B with 1,000 writers, {few} B with 10"
    );
}
