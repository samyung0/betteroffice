//! The yrs patch that compacts a client's block list once per commit: a commit
//! deleting in many places merges the blocks around each deleted range as
//! before, then removes every merged block in one pass. Removing them range by
//! range moved the list's tail once per range, quadratic in the ranges (a DOCX
//! seed deleted a placeholder in each of thousands of stories).

use yrs::updates::decoder::Decode;
use yrs::{Doc, GetString, ReadTxn, Text, Transact, Update};

const LEN: u32 = 40_000;

#[test]
fn a_commit_deleting_in_many_places_merges_in_one_pass() {
    let doc = Doc::with_client_id(1);
    let text = doc.get_or_insert_text("t");
    let source: String = (0..LEN)
        .map(|at| char::from(b'a' + (at % 26) as u8))
        .collect();
    text.insert(&mut doc.transact_mut(), 0, &source);
    // Two characters deleted one by one at each place (4k + 1 and 4k + 2,
    // from the start, so each split is of the last block and moves none):
    // the two deleted blocks merge at the commit.
    let places = (LEN / 4) as usize;
    let work = {
        let mut txn = doc.transact_mut();
        txn.limit_work(u64::MAX);
        for place in 0..LEN / 4 {
            let at = 2 * place + 1;
            text.remove_range(&mut txn, at, 1);
            text.remove_range(&mut txn, at, 1);
        }
        txn.commit();
        txn.work()
    };
    // A few steps per place; removing merged blocks range by range moved the
    // list's tail each time (816,256 steps here before the patch).
    assert!(work < 10 * places as u64, "{work} steps");

    let kept: String = source
        .chars()
        .enumerate()
        .filter(|(at, _)| at % 4 == 0 || at % 4 == 3)
        .map(|(_, c)| c)
        .collect();
    let txn = doc.transact();
    assert_eq!(text.get_string(&txn), kept);
    // A kept block before each place, its merged deleted block, the last kept.
    let state = txn.encode_state_as_update_v1(&Default::default());
    let update = Update::decode_v1(&state).unwrap();
    assert_eq!(update.blocks().count(), 2 * places + 1);

    let copy = Doc::with_client_id(2);
    copy.transact_mut().apply_update(update).unwrap();
    let copied = copy.get_or_insert_text("t");
    assert_eq!(copied.get_string(&copy.transact()), kept);
}
