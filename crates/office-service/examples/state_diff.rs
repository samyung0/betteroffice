//! `state_diff <a.bin> <b.bin>`: the delete-set difference of two DOCX room
//! states and what each differing item holds (spike diagnostics).

use yrs::updates::decoder::Decode;
use yrs::{ReadTxn, Transact};

fn load(path: &str) -> yrs::Doc {
    let mut options = yrs::Options::with_client_id(yrs::ClientID::new(5));
    options.offset_kind = yrs::OffsetKind::Utf16;
    options.skip_gc = true;
    let doc = yrs::Doc::with_options(options);
    let bytes = std::fs::read(path).unwrap();
    doc.transact_mut()
        .apply_update(yrs::Update::decode_v1(&bytes).unwrap())
        .unwrap();
    doc
}

fn deleted(doc: &yrs::Doc) -> yrs::IdSet {
    let bytes = doc
        .transact()
        .encode_state_as_update_v1(&yrs::StateVector::default());
    yrs::Update::decode_v1(&bytes).unwrap().delete_set().clone()
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let (a, b) = (load(&args[1]), load(&args[2]));
    let reference = args.get(3).map(|path| load(path));
    let (da, db) = (deleted(&a), deleted(&b));
    for (label, only, doc) in [
        ("only in a", da.diff(&db), &a),
        ("only in b", db.diff(&da), &b),
    ] {
        println!("{label}:");
        let txn = doc.transact();
        for (client, ranges) in only.iter() {
            for range in ranges.iter() {
                let reference_txn = reference.as_ref().map(|doc| doc.transact());
                let item = reference_txn
                    .as_ref()
                    .and_then(|txn| txn.store().get_item(&yrs::ID::new(*client, range.start)))
                    .or_else(|| txn.store().get_item(&yrs::ID::new(*client, range.start)));
                let content = item
                    .map(|item| format!("{:?}", item.content()))
                    .unwrap_or_else(|| "?".into());
                println!(
                    "  {client:?} {:?} {}",
                    range,
                    &content[..content.len().min(160)]
                );
            }
        }
    }
}
