//! Remote keystrokes on an engine session, both layouts: a peer types 200
//! characters into the middle paragraph; a second session applies each
//! message as an editor does (`apply_peer_update_v1`), and again as a plain
//! update (`apply_update_v1`). Median per message.
//! `remote_keys <file.docx>...`

use std::cell::RefCell;
use std::rc::Rc;
use std::time::Instant;

use docx_edit::{EditCtx, EditingDoc, FormatPolicy, Position, SegmentContent, overlay};

fn open(base: &[u8], chunked: bool, client: u64) -> EditingDoc {
    let doc = EditingDoc::new(client);
    if chunked {
        let envelope = docx_edit::parse_docx_for_edit(base).unwrap();
        overlay::open_chunked(&doc, envelope, &overlay::fingerprint(base)).unwrap();
    } else {
        docx_edit::seed_from_docx(&doc, base).unwrap();
    }
    doc
}

fn middle_mark(doc: &EditingDoc) -> u32 {
    let (mut offset, mut marks) = (0, Vec::new());
    for segment in doc.story_segments("body").unwrap() {
        match segment.content {
            SegmentContent::Text(text) => offset += text.encode_utf16().count() as u32,
            SegmentContent::Pilcrow(_) => {
                marks.push(offset);
                offset += 1;
            }
            SegmentContent::OtherEmbed { .. } => offset += 1,
        }
    }
    marks[marks.len() / 2]
}

fn median(mut values: Vec<f64>) -> f64 {
    values.sort_by(|a, b| a.partial_cmp(b).unwrap());
    values[values.len() / 2]
}

fn main() {
    for path in std::env::args().skip(1) {
        let base = std::fs::read(&path).unwrap();
        for chunked in [false, true] {
            let typist = open(&base, chunked, 701);
            let sent = Rc::new(RefCell::new(Vec::new()));
            let sink = Rc::clone(&sent);
            let _subscription = typist
                .observe_shared_updates(true, move |update| sink.borrow_mut().push(update))
                .unwrap();
            let at = middle_mark(&typist);
            let ctx = EditCtx::local("Bench", "2026-10-08T00:00:00Z");
            let mut local_times = Vec::new();
            for index in 0..200 {
                let started = Instant::now();
                typist
                    .insert_text(&ctx, Position::new("body", at + index), "x", FormatPolicy::Inherit)
                    .unwrap();
                local_times.push(started.elapsed().as_secs_f64() * 1000.0);
            }
            // The cost yrs pays per transaction for every client the
            // document holds (its before and after state vectors).
            let clients = {
                use yrs::{ReadTxn, Transact};
                typist.yrs_doc().transact().state_vector().len()
            };
            let empty_ms = {
                use yrs::Transact;
                let started = Instant::now();
                for _ in 0..100 {
                    drop(typist.yrs_doc().transact_mut());
                }
                started.elapsed().as_secs_f64() * 1000.0 / 100.0
            };
            // Items in the body story (live, deleted), and blocks overall.
            let (live, deleted) = {
                use yrs::{Map, ReadTxn, Transact};
                let txn = typist.yrs_doc().transact();
                let body = txn
                    .get_map("stories")
                    .and_then(|stories| stories.get(&txn, "body"))
                    .and_then(|value| value.cast::<yrs::TextRef>().ok())
                    .unwrap();
                let (mut live, mut deleted) = (0, 0);
                let mut item = <yrs::TextRef as AsRef<yrs::branch::Branch>>::as_ref(&body).start();
                while let Some(current) = item {
                    if current.is_deleted() {
                        deleted += 1;
                    } else {
                        live += 1;
                    }
                    item = current.right();
                }
                (live, deleted)
            };
            let messages = sent.borrow().clone();
            let mut peer_times = Vec::new();
            let mut plain_times = Vec::new();
            for (peer, plain) in [(true, false), (false, true)] {
                let receiver = open(&base, chunked, 702);
                for message in &messages {
                    let started = Instant::now();
                    if peer {
                        receiver.apply_peer_update_v1(message).unwrap();
                    } else {
                        receiver.apply_update_v1(message).unwrap();
                    }
                    let elapsed = started.elapsed().as_secs_f64() * 1000.0;
                    if plain {
                        plain_times.push(elapsed);
                    } else {
                        peer_times.push(elapsed);
                    }
                }
            }
            println!(
                "{}",
                serde_json::json!({
                    "file": std::path::Path::new(&path).file_stem().unwrap().to_string_lossy(),
                    "layout": if chunked { "chunked" } else { "seeded" },
                    "messages": messages.len(),
                    "peerApplyMedianMs": median(peer_times[1..].to_vec()),
                    "plainApplyMedianMs": median(plain_times[1..].to_vec()),
                    "firstPeerApplyMs": peer_times[0],
                    "localKeyMedianMs": median(local_times[1..].to_vec()),
                    "clients": clients,
                    "emptyTransactionMs": empty_ms,
                    "bodyLiveItems": live,
                    "bodyDeletedItems": deleted,
                })
            );
        }
    }
}
