//! A yrs transaction answers "what did it insert" from its insert set instead
//! of building whole state vectors, whose cost grows with every client the
//! document ever saw. Its update and its events must stay what the state
//! vectors gave.

use std::cell::RefCell;
use std::rc::Rc;
use std::sync::Arc;

use yrs::types::{Attrs, EntryChange};
use yrs::updates::decoder::Decode;
use yrs::{
    Any, ClientID, DeepObservable, Doc, Map, MapPrelim, MapRef, Options, ReadTxn, StateVector,
    Text, Transact, Update,
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
            .apply_update(Update::decode_v1(&main.transact().encode_state_as_update_v1(&StateVector::default())).unwrap())
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
