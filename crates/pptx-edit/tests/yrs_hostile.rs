//! Security regressions for the patched yrs: updates a peer can craft.
//! `third_party/yrs` sits outside the cargo workspace, so its own unit tests
//! never run in CI; these do.

use yrs::encoding::write::Write;
use yrs::updates::decoder::Decode;
use yrs::{Doc, GetString, ReadTxn, Text, Transact, Update};

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
