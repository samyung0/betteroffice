//! Where a remote keystroke's apply spends its time on an override-layout
//! session (spike): decode, integrate, commit, per message, median.
//! `apply_probe <file.docx>`

use std::cell::RefCell;
use std::rc::Rc;
use std::time::Instant;

use docx_edit::{EditCtx, EditingDoc, FormatPolicy, Position, SegmentContent, overlay};
use yrs::updates::decoder::Decode;
use yrs::{ReadTxn, Transact, Update};

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

fn median(mut values: Vec<f64>) -> f64 {
    values.sort_by(|a, b| a.partial_cmp(b).unwrap());
    values[values.len() / 2]
}

fn main() {
    let path = std::env::args().nth(1).unwrap();
    let base = std::fs::read(&path).unwrap();
    for chunked in [false, true] {
        let typist = open(&base, chunked, 701);
        let sent = Rc::new(RefCell::new(Vec::new()));
        let sink = Rc::clone(&sent);
        let _subscription = typist
            .observe_shared_updates(true, move |update| sink.borrow_mut().push(update))
            .unwrap();
        let mut marks = Vec::new();
        let mut offset = 0;
        for segment in typist.story_segments("body").unwrap() {
            match segment.content {
                SegmentContent::Text(text) => offset += text.encode_utf16().count() as u32,
                SegmentContent::Pilcrow(_) => {
                    marks.push(offset);
                    offset += 1;
                }
                SegmentContent::OtherEmbed { .. } => offset += 1,
            }
        }
        let at = marks[marks.len() / 2];
        let ctx = EditCtx::local("Bench", "2026-10-08T00:00:00Z");
        for index in 0..100 {
            typist
                .insert_text(
                    &ctx,
                    Position::new("body", at + index),
                    "x",
                    FormatPolicy::Inherit,
                )
                .unwrap();
        }
        let messages = sent.borrow().clone();
        let receiver = open(&base, chunked, 702);
        let (mut decode, mut integrate, mut commit) = (Vec::new(), Vec::new(), Vec::new());
        for message in &messages {
            let started = Instant::now();
            let update = Update::decode_v1(message).unwrap();
            let decoded = Instant::now();
            let mut txn = receiver.yrs_doc().transact_mut();
            txn.apply_update(update).unwrap();
            let integrated = Instant::now();
            drop(txn);
            let committed = Instant::now();
            decode.push((decoded - started).as_secs_f64() * 1000.0);
            integrate.push((integrated - decoded).as_secs_f64() * 1000.0);
            commit.push((committed - integrated).as_secs_f64() * 1000.0);
        }
        // Reading: the body's segments and every story's segments, best of 5.
        let best = |f: &dyn Fn()| {
            (0..5)
                .map(|_| {
                    let started = Instant::now();
                    f();
                    started.elapsed().as_secs_f64() * 1000.0
                })
                .fold(f64::MAX, f64::min)
        };
        let body_read_ms = best(&|| {
            receiver.story_segments("body").unwrap();
        });
        let story_ids: Vec<String> = {
            use yrs::Map;
            let txn = receiver.yrs_doc().transact();
            let stories = txn.get_map("stories").unwrap();
            stories.keys(&txn).map(str::to_owned).collect()
        };
        let all_read_ms = best(&|| {
            for story in &story_ids {
                receiver.story_segments(story).unwrap();
            }
        });
        let pending = {
            let txn = receiver.yrs_doc().transact();
            (
                txn.store().pending_update().is_some(),
                txn.store().pending_ds().is_some(),
            )
        };
        println!(
            "{}",
            serde_json::json!({
                "file": std::path::Path::new(&path).file_stem().unwrap().to_string_lossy(),
                "layout": if chunked { "chunked" } else { "seeded" },
                "messages": messages.len(),
                "decodeMedianMs": median(decode[1..].to_vec()),
                "integrateMedianMs": median(integrate[1..].to_vec()),
                "commitMedianMs": median(commit[1..].to_vec()),
                "firstIntegrateMs": integrate[0],
                "firstCommitMs": commit[0],
                "bodyReadMs": body_read_ms,
                "allStoriesReadMs": all_read_ms,
                "stories": story_ids.len(),
                "pendingUpdate": pending.0,
                "pendingDeleteSet": pending.1,
            })
        );
    }
}
