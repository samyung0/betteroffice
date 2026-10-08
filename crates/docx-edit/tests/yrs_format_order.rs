//! yrs writes one edit's format items in key order: the attributes come as a
//! hash map, which iterates in a different order on every replica, so the
//! same edit wrote differently ordered items, and later concurrent edits
//! next to them integrated differently.

use std::sync::Arc;

use yrs::block::ItemContent;
use yrs::types::Attrs;
use yrs::{Any, ClientID, Doc, Options, Text, TextRef, Transact};

const KEYS: [&str; 10] = [
    "bold",
    "italic",
    "underline",
    "strike",
    "color",
    "fontSize",
    "fontFamily",
    "highlight",
    "language",
    "spacing",
];

/// The keys of the live format items in `text`, in document order.
fn format_keys(doc: &Doc, text: &TextRef) -> Vec<Arc<str>> {
    let _txn = doc.transact();
    let mut keys = Vec::new();
    let mut item = <TextRef as AsRef<yrs::branch::Branch>>::as_ref(text).start();
    while let Some(at) = item {
        if let (false, ItemContent::Format(key, _)) = (at.is_deleted(), at.content()) {
            keys.push(key.clone());
        }
        item = at.right();
    }
    keys
}

fn sorted() -> Vec<Arc<str>> {
    let mut keys: Vec<Arc<str>> = KEYS.iter().map(|key| Arc::from(*key)).collect();
    keys.sort();
    keys
}

#[test]
fn a_format_writes_its_items_in_key_order() {
    for client in 1..=8 {
        let doc = Doc::with_options(Options::with_client_id(ClientID::new(client)));
        let text = doc.get_or_insert_text("story");
        let attrs: Attrs = KEYS
            .iter()
            .map(|key| (Arc::from(*key), Any::Bool(true)))
            .collect();
        {
            let mut txn = doc.transact_mut();
            text.insert(&mut txn, 0, "abcdefgh");
            text.format(&mut txn, 2, 3, attrs);
        }
        // The start markers, then the end markers that restore the old values.
        let expected: Vec<Arc<str>> = sorted().into_iter().chain(sorted()).collect();
        assert_eq!(format_keys(&doc, &text), expected, "client {client}");
    }
}

#[test]
fn an_insert_with_attributes_writes_its_items_in_key_order() {
    for client in 1..=8 {
        let doc = Doc::with_options(Options::with_client_id(ClientID::new(client)));
        let text = doc.get_or_insert_text("story");
        let attrs: Attrs = KEYS
            .iter()
            .map(|key| (Arc::from(*key), Any::Bool(true)))
            .collect();
        {
            let mut txn = doc.transact_mut();
            text.insert(&mut txn, 0, "abcdefgh");
            text.insert_with_attributes(&mut txn, 4, "XY", attrs);
        }
        let expected: Vec<Arc<str>> = sorted().into_iter().chain(sorted()).collect();
        assert_eq!(format_keys(&doc, &text), expected, "client {client}");
    }
}
