//! The yrs patch `StickyIndex::at_many` places many sticky indexes in one walk
//! of a text and must answer exactly what `StickyIndex::at` answers for each,
//! whatever the request order, over split, deleted, formatted and embedded
//! items.

use std::sync::Arc;

use yrs::branch::{Branch, BranchPtr};
use yrs::types::Attrs;
use yrs::{Any, Assoc, Doc, GetString, MapPrelim, StickyIndex, Text, Transact};

#[test]
fn at_many_answers_as_at_for_every_index() {
    let doc = Doc::with_client_id(1);
    let text = doc.get_or_insert_text("story");
    // Each edit its own transaction, so its items stay apart.
    text.insert(&mut doc.transact_mut(), 0, "hello world");
    text.insert(&mut doc.transact_mut(), 5, ", dear");
    text.insert_embed(&mut doc.transact_mut(), 3, MapPrelim::default());
    let bold = Attrs::from([(Arc::from("bold"), Any::Bool(true))]);
    text.format(&mut doc.transact_mut(), 2, 6, bold);
    text.remove_range(&mut doc.transact_mut(), 8, 3);
    text.insert(&mut doc.transact_mut(), 0, "😀");
    text.push(&mut doc.transact_mut(), " end");
    let txn = doc.transact();
    let len = text.len(&txn);
    assert!(text.get_string(&txn).contains("end"));
    let branch = BranchPtr::from(AsRef::<Branch>::as_ref(&text));
    // Every index past the end too, both ways, backwards with repeats.
    let mut requests = Vec::new();
    for index in (0..=len + 1).rev() {
        requests.push((index, Assoc::After));
        requests.push((index, Assoc::Before));
        requests.push((index / 2, Assoc::After));
    }
    let batch = StickyIndex::at_many(&txn, branch, &requests);
    assert_eq!(batch.len(), requests.len());
    for ((index, assoc), placed) in requests.iter().zip(batch) {
        assert_eq!(
            placed,
            StickyIndex::at(&txn, branch, *index, *assoc),
            "{index} {assoc:?}"
        );
    }
    assert!(StickyIndex::at_many(&txn, branch, &[]).is_empty());
}
