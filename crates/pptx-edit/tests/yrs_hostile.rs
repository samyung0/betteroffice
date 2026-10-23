//! Security regressions for the patched yrs: updates a peer can craft.
//! `third_party/yrs` sits outside the cargo workspace, so its own unit tests
//! never run in CI; these do.

use yrs::block::ClientID;
use yrs::encoding::write::Write;
use yrs::types::text::YChange;
use yrs::updates::decoder::Decode;
use yrs::{Any, Doc, GetString, ID, Out, ReadTxn, Text, Transact, Update, UpdateBlock};

/// An update of one client with one item whose content is a single `Any`
/// value starting with `value`, under the root text "t".
fn any_item(value: &[u8]) -> Vec<u8> {
    let mut bytes = Vec::new();
    bytes.write_var(1u32); // clients
    bytes.write_var(1u32); // structs
    bytes.write_var(77u64); // client
    bytes.write_var(0u32); // clock
    bytes.write_u8(8); // item with Any content, no origins
    bytes.write_var(1u32); // parent is a root, by name
    bytes.write_string("t");
    bytes.write_var(1u32); // one value
    bytes.extend_from_slice(value);
    bytes.write_var(0u32); // empty delete set
    bytes
}

/// `depth` arrays nested in one another around a null.
fn nested_arrays(depth: usize) -> Vec<u8> {
    let mut value = Vec::new();
    for _ in 0..depth {
        value.write_u8(117);
        value.write_var(1u32);
    }
    value.write_u8(126);
    value
}

/// Security regression (notes REVIEW1 M3): an update listing a client with
/// no structs panicked in `BlockSet::exclude` when the document knew that
/// client, killing whatever task applied it.
#[test]
fn a_client_listed_with_no_structs_applies_as_nothing() {
    let doc = Doc::with_client_id(1);
    let text = doc.get_or_insert_text("t");
    text.insert(&mut doc.transact_mut(), 0, "abc");
    let before = doc.transact().state_vector();
    let update = Update::decode_v1(&[0x01, 0x00, 0x01, 0x01, 0x00, 0x18]).unwrap();
    doc.transact_mut().apply_update(update).unwrap();
    assert_eq!(text.get_string(&doc.transact()), "abc");
    assert_eq!(doc.transact().state_vector(), before);
}

/// Security regression (notes REVIEW1 M2): a 20 KB update holding a value
/// nested 10,000 deep overflowed the decoding thread's stack, an abort of the
/// whole process. Decoding stops past 128 levels with an error, also on a
/// thread with a tokio worker's 2 MiB stack.
#[test]
fn a_value_nested_past_128_levels_is_refused_at_decode() {
    assert!(Update::decode_v1(&any_item(&nested_arrays(128))).is_ok());
    assert!(Update::decode_v1(&any_item(&nested_arrays(129))).is_err());
    let deep = any_item(&nested_arrays(10_000));
    let refused = std::thread::Builder::new()
        .stack_size(2 << 20)
        .spawn(move || Update::decode_v1(&deep).is_err())
        .unwrap()
        .join()
        .unwrap();
    assert!(refused);
}

/// Security regression (yrs find-index, 2026-10-22): a lookup of a clock the
/// store does not hold, such as an id an update names, finds nothing. Here
/// the client holds one unit.
#[test]
fn a_lookup_past_a_single_unit_finds_nothing() {
    let doc = Doc::with_client_id(1);
    let text = doc.get_or_insert_text("t");
    text.insert(&mut doc.transact_mut(), 0, "a");
    let client = ClientID::new(1);
    let txn = doc.transact();
    let store = txn.store();
    assert!(store.get_item(&ID::new(client, 0)).is_some());
    for clock in [1, 5, 1_000] {
        assert!(store.get_item(&ID::new(client, clock)).is_none());
        assert!(store.get_block_range(&ID::new(client, clock)).is_none());
        assert_eq!(store.blocks_in(&client, clock..clock + 1).count(), 0);
    }
}

/// Security regression (yrs find-index, 2026-10-22): as above, for a client
/// holding several blocks, while every held clock still finds the block
/// holding it.
#[test]
fn a_lookup_past_several_blocks_finds_nothing() {
    let doc = Doc::with_client_id(1);
    let text = doc.get_or_insert_text("t");
    // Typed at the start, so no block merges with the one before: clocks 0,
    // 1-3 and 4.
    for typed in ["a", "xyz", "b"] {
        text.insert(&mut doc.transact_mut(), 0, typed);
    }
    let client = ClientID::new(1);
    let txn = doc.transact();
    let store = txn.store();
    for (clock, start, len) in [(0, 0, 1), (1, 1, 3), (2, 1, 3), (3, 1, 3), (4, 4, 1)] {
        let (range, _) = store.get_block_range(&ID::new(client, clock)).unwrap();
        assert_eq!((range.clock, range.len), (start, len), "clock {clock}");
    }
    for clock in [5, 9, 1_000, u32::MAX] {
        assert!(store.get_item(&ID::new(client, clock)).is_none());
        assert!(store.get_block_range(&ID::new(client, clock)).is_none());
        assert_eq!(
            store
                .blocks_in(&client, clock..clock.saturating_add(1))
                .count(),
            0
        );
    }
}

// Stored states (gc on) and updates from replicas that never saw the
// deletion, written by browser Yjs 13.6.31, which gives the expected texts.
// The root text "content" holds paragraphs as embedded texts.

/// "L", then a deleted paragraph that held three units.
const L_GONE3: &[u8] = &[
    1, 3, 1, 0, 4, 1, 7, 99, 111, 110, 116, 101, 110, 116, 1, 76, 129, 1, 0, 1, 0, 3, 1, 1, 1, 1, 4,
];
/// "L", a deleted paragraph of three units, "R".
const L_GONE3_R: &[u8] = &[
    1, 4, 1, 0, 4, 1, 7, 99, 111, 110, 116, 101, 110, 116, 1, 76, 129, 1, 0, 1, 0, 3, 132, 1, 1, 1,
    82, 1, 1, 1, 1, 4,
];
/// "L", then a deleted empty paragraph.
const L_GONE_EMPTY: &[u8] = &[
    1, 2, 1, 0, 4, 1, 7, 99, 111, 110, 116, 101, 110, 116, 1, 76, 129, 1, 0, 1, 1, 1, 1, 1, 1,
];
/// A deleted paragraph holding an empty embedded text.
const GONE_NESTED: &[u8] = &[
    1, 2, 1, 0, 1, 1, 7, 99, 111, 110, 116, 101, 110, 116, 1, 0, 1, 1, 1, 1, 0, 2,
];
/// A deleted paragraph with one unit and an attribute.
const GONE_ATTRIBUTE: &[u8] = &[
    1, 2, 1, 0, 1, 1, 7, 99, 111, 110, 116, 101, 110, 116, 1, 0, 2, 1, 1, 1, 0, 3,
];
/// "abc" with "b" deleted (a deleted item, not collected).
const A_DELETED_C: &[u8] = &[
    1, 3, 1, 0, 4, 1, 7, 99, 111, 110, 116, 101, 110, 116, 1, 97, 129, 1, 0, 1, 132, 1, 1, 1, 99,
    1, 1, 1, 1, 1,
];
/// A live paragraph "hel world" after deletions, then a deleted paragraph
/// whose children are collected (clocks 15..24).
const LIVE_THEN_GONE: &[u8] = &[
    1, 8, 1, 0, 7, 1, 7, 99, 111, 110, 116, 101, 110, 116, 6, 40, 0, 1, 0, 4, 116, 121, 112, 101,
    1, 119, 1, 112, 40, 0, 1, 0, 2, 105, 100, 1, 119, 2, 115, 49, 4, 0, 1, 0, 3, 104, 101, 108,
    129, 1, 5, 2, 132, 1, 7, 6, 32, 119, 111, 114, 108, 100, 129, 1, 0, 1, 0, 9, 1, 1, 2, 6, 2, 14,
    10,
];

/// One "X" of client 10 with origin `origin` and right origin `right`.
const fn x_between(origin: [u8; 2], right: [u8; 2]) -> [u8; 12] {
    [
        1, 1, 10, 0, 196, origin[0], origin[1], right[0], right[1], 1, 88, 0,
    ]
}

struct Case {
    name: &'static str,
    state: &'static [u8],
    updates: &'static [&'static [u8]],
    text: &'static str,
    collected: bool,
}

const CASES: &[Case] = &[
    Case {
        name: "origin and right origin collected",
        state: L_GONE3_R,
        updates: &[&x_between([1, 2], [1, 3])],
        text: "LR",
        collected: true,
    },
    Case {
        name: "origin collected, no right origin",
        state: L_GONE3,
        updates: &[&[1, 1, 10, 0, 132, 1, 4, 1, 88, 0]],
        text: "L",
        collected: true,
    },
    Case {
        name: "right origin collected, no origin",
        state: L_GONE3,
        updates: &[&[1, 1, 10, 0, 68, 1, 2, 1, 88, 0]],
        text: "L",
        collected: true,
    },
    Case {
        name: "parent is a collected type",
        state: L_GONE_EMPTY,
        updates: &[&[1, 1, 10, 0, 4, 0, 1, 1, 1, 88, 0]],
        text: "L",
        collected: true,
    },
    Case {
        name: "parent is a collected unit",
        state: GONE_NESTED,
        updates: &[&[1, 1, 10, 0, 4, 0, 1, 1, 1, 88, 0]],
        text: "",
        collected: true,
    },
    Case {
        name: "after an item that was collected on arrival",
        state: L_GONE3,
        updates: &[
            &x_between([1, 2], [1, 3]),
            &[1, 1, 11, 0, 196, 10, 0, 1, 3, 1, 89, 0],
        ],
        text: "L",
        collected: true,
    },
    Case {
        name: "the same two items in one update",
        state: L_GONE3,
        updates: &[&[
            2, 1, 11, 0, 196, 10, 0, 1, 3, 1, 89, 1, 10, 0, 196, 1, 2, 1, 3, 1, 88, 0,
        ]],
        text: "L",
        collected: true,
    },
    Case {
        name: "map entry whose origin is a collected value",
        state: GONE_ATTRIBUTE,
        updates: &[&[
            1, 1, 10, 0, 168, 1, 2, 1, 119, 5, 114, 105, 103, 104, 116, 1, 1, 1, 2, 1,
        ]],
        text: "",
        collected: true,
    },
    Case {
        name: "origin collected, right origin live",
        state: LIVE_THEN_GONE,
        updates: &[&x_between([1, 20], [1, 8])],
        text: "[hel world]",
        collected: true,
    },
    Case {
        name: "origin live, right origin collected",
        state: LIVE_THEN_GONE,
        updates: &[&x_between([1, 8], [1, 20])],
        text: "[hel world]",
        collected: true,
    },
    Case {
        name: "control: origin deleted but held as an item",
        state: A_DELETED_C,
        updates: &[&x_between([1, 1], [1, 2])],
        text: "aXc",
        collected: false,
    },
];

/// The root's strings, with each embedded paragraph's text in brackets.
fn content_text(doc: &Doc) -> String {
    let root = doc.get_or_insert_text("content");
    let txn = doc.transact();
    root.diff(&txn, YChange::identity)
        .into_iter()
        .map(|chunk| match chunk.insert {
            Out::Any(Any::String(text)) => text.to_string(),
            Out::YXmlText(paragraph) => format!("[{}]", paragraph.get_string(&txn)),
            other => panic!("unexpected chunk {other:?}"),
        })
        .collect()
}

/// Security regression (yrs GC neighbours, 2026-10-23): an item next to a
/// collected unit is collected, as in Yjs: each update's items integrate as
/// collected units over their whole length, the client's clock advances, and
/// the document's text equals Yjs's. Fails if the patch in `src/update.rs`
/// (missing_dependency) is reverted.
#[test]
fn an_item_next_to_a_collected_unit_is_collected_as_in_yjs() {
    for case in CASES {
        let doc = Doc::with_client_id(999);
        doc.transact_mut()
            .apply_update(Update::decode_v1(case.state).unwrap())
            .unwrap();
        for bytes in case.updates {
            let update = Update::decode_v1(bytes).unwrap();
            let items: Vec<(ID, u32)> = update
                .blocks()
                .filter_map(|block| match block {
                    UpdateBlock::Item(item) => Some((*item.id(), item.len())),
                    _ => None,
                })
                .collect();
            doc.transact_mut().apply_update(update).unwrap();
            let txn = doc.transact();
            let store = txn.store();
            assert!(store.pending_update().is_none(), "{}: pending", case.name);
            for (id, len) in items {
                assert_eq!(
                    txn.state_vector().get(&id.client),
                    id.clock + len,
                    "{}: clock of {id}",
                    case.name
                );
                let (range, item) = store.get_block_range(&id).unwrap();
                assert!(range.clock <= id.clock && id.clock + len <= range.clock + range.len);
                assert_eq!(
                    item.is_none(),
                    case.collected,
                    "{}: {id} collected",
                    case.name
                );
            }
        }
        assert_eq!(content_text(&doc), case.text, "{}", case.name);
    }
}
