//! Security regressions for the patched yrs's integration of an update:
//! the struct picker, the work budget and the block list's push.
//! `third_party/yrs` sits outside the cargo workspace, so its own unit tests
//! never run in CI; these do.

use yrs::block::ClientID;
use yrs::encoding::write::Write;
use yrs::error::UpdateError;
use yrs::updates::decoder::Decode;
use yrs::{Doc, ID, ReadTxn, Text, Transact, Update};

/// Security regression (notes REVIEW4 M9): 20,000 forged clients with one
/// struct each, each struct's origin the one of the client before, made the
/// struct picker recurse once per drained client and overflow a 2 MiB stack
/// (a tokio worker's), aborting the process.
#[test]
fn a_chain_of_clients_integrates_on_a_task_stack() {
    let doc = Doc::with_client_id(1);
    let text = doc.get_or_insert_text("t");
    text.insert(&mut doc.transact_mut(), 0, "h");
    let clients = 20_000u64;
    let mut bytes: Vec<u8> = Vec::new();
    bytes.write_var(clients);
    for k in 0..clients {
        bytes.write_var(1u32);
        bytes.write_var(1_000 + k);
        bytes.write_var(0u32);
        bytes.write_u8(0x80 | 4); // an origin, a string
        bytes.write_var(if k == 0 { 1 } else { 1_000 + k - 1 });
        bytes.write_var(0u32);
        bytes.write_string("x");
    }
    bytes.write_var(0u32);
    let len = std::thread::Builder::new()
        .stack_size(2 << 20)
        .spawn(move || {
            let mut txn = doc.transact_mut();
            txn.apply_update(Update::decode_v1(&bytes).unwrap())
                .unwrap();
            text.len(&txn)
        })
        .unwrap()
        .join()
        .unwrap();
    assert_eq!(len, 1 + clients as u32);
}

/// A text of `n` characters each typed at its start, all deleted, and an
/// update of `n` one-character items with only a right origin, each a
/// tombstone from the last back: yrs scans from the start of the text to
/// each right origin, n(n-1)/2 items in all, and splits nothing.
fn fanned(n: u32) -> (Doc, Vec<u8>) {
    let doc = Doc::with_client_id(1);
    let text = doc.get_or_insert_text("t");
    for _ in 0..n {
        text.insert(&mut doc.transact_mut(), 0, "t");
    }
    text.remove_range(&mut doc.transact_mut(), 0, n);
    let mut bytes: Vec<u8> = Vec::new();
    bytes.write_var(1u32);
    bytes.write_var(n);
    bytes.write_var(1_000u64);
    bytes.write_var(0u32);
    for clock in 0..n {
        bytes.write_u8(0x40 | 4); // a right origin, a string
        bytes.write_var(1u64);
        bytes.write_var(clock);
        bytes.write_string("x");
    }
    bytes.write_var(0u32);
    (doc, bytes)
}

/// Security regression (notes REVIEW4 M10-M12, Epo 2026-10-08): the work an
/// update costs yrs is counted by yrs itself, so Capy can refuse one past a
/// budget on a trial copy. A conflict scan counts a step per item it passes,
/// a split one step plus one per 128 blocks it moves; past the budget apply
/// stops; the budget ends with its transaction.
#[test]
fn integration_work_is_counted_and_stops_past_the_budget() {
    let (doc, bytes) = fanned(100);
    let mut txn = doc.transact_mut();
    txn.limit_work(u64::MAX);
    txn.apply_update(Update::decode_v1(&bytes).unwrap())
        .unwrap();
    txn.commit();
    assert_eq!(txn.work(), 100 * 99 / 2);
    drop(txn);

    let (doc, bytes) = fanned(3_000); // 4,498,500 steps in full
    let mut txn = doc.transact_mut();
    txn.limit_work(10_000);
    let applied = txn.apply_update(Update::decode_v1(&bytes).unwrap());
    assert!(
        matches!(applied, Err(UpdateError::WorkBudgetExceeded(10_001))),
        "{applied:?}"
    );
    txn.commit();
    assert_eq!(txn.work(), 10_001);
    drop(txn);

    // Between "a" and "bc": the "abc" block splits, and the 300 blocks after
    // it in client 5's list move.
    let doc = Doc::with_client_id(5);
    let text = doc.get_or_insert_text("t");
    text.insert(&mut doc.transact_mut(), 0, "abc");
    for _ in 0..300 {
        text.insert(&mut doc.transact_mut(), 0, "p");
    }
    let mut bytes: Vec<u8> = Vec::new();
    bytes.write_var(1u32);
    bytes.write_var(1u32);
    bytes.write_var(77u64);
    bytes.write_var(0u32);
    bytes.write_u8(0x80 | 0x40 | 4);
    bytes.write_var(5u64);
    bytes.write_var(0u32);
    bytes.write_var(5u64);
    bytes.write_var(1u32);
    bytes.write_string("x");
    bytes.write_var(0u32);
    let mut txn = doc.transact_mut();
    txn.limit_work(u64::MAX);
    txn.apply_update(Update::decode_v1(&bytes).unwrap())
        .unwrap();
    txn.commit();
    assert_eq!(txn.work(), 1 + 300 / 128);
    drop(txn);

    let (doc, bytes) = fanned(100);
    {
        let mut txn = doc.transact_mut();
        txn.limit_work(10);
    }
    let mut txn = doc.transact_mut();
    txn.apply_update(Update::decode_v1(&bytes).unwrap())
        .unwrap();
    assert_eq!(txn.get_text("t").unwrap().len(&txn), 100);
}

/// Security regression (notes round 8): an update listing client 77 twice,
/// each section a struct at clock 0, made yrs place the second over the
/// first in the client's block list, freeing the first while the text still
/// linked it (a use after free). The block list keeps the first. (Decoding
/// refuses the shape once capy/yrs-gc-fix lands; Capy refuses it at admit.)
#[test]
fn a_struct_placed_twice_keeps_the_first() {
    let doc = Doc::with_client_id(1);
    let text = doc.get_or_insert_text("t");
    text.insert(&mut doc.transact_mut(), 0, "hello");
    let mut bytes: Vec<u8> = Vec::new();
    bytes.write_var(2u32);
    for origin in [1u32, 3] {
        bytes.write_var(1u32);
        bytes.write_var(77u64);
        bytes.write_var(0u32);
        bytes.write_u8(0x80 | 4);
        bytes.write_var(1u64);
        bytes.write_var(origin);
        bytes.write_string("x");
    }
    bytes.write_var(0u32);
    let Ok(update) = Update::decode_v1(&bytes) else {
        return;
    };
    let mut txn = doc.transact_mut();
    txn.apply_update(update).unwrap();
    let item = txn
        .store()
        .get_item(&ID::new(ClientID::new(77), 0))
        .expect("the struct");
    assert_eq!(item.origin().copied(), Some(ID::new(ClientID::new(1), 1)));
}
