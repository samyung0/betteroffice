//! DOCX override layout (spike) against today's seeded layout, one build,
//! layouts interleaved per run: seed and room bytes, stored change, room load
//! and heap as the server holds the room, engine open, a remote keystroke,
//! save effects and export.
//! `chunked_bench <out.jsonl> <runs> <file.docx>...`

use std::alloc::{GlobalAlloc, Layout as AllocLayout, System};
use std::cell::RefCell;
use std::rc::Rc;
use std::sync::atomic::{AtomicIsize, Ordering};
use std::time::Instant;

use docx_edit::overlay;
use docx_edit::{EditCtx, EditingDoc, FormatPolicy, Position, SegmentContent, StoryRange};
use office_service::{Checkpoint, Determinism, Format};
use yrs::updates::decoder::Decode;
use yrs::{Doc, OffsetKind, Options, ReadTxn, StateVector, Transact, Update};

struct Counting;
static LIVE: AtomicIsize = AtomicIsize::new(0);
static PEAK: AtomicIsize = AtomicIsize::new(0);
unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: AllocLayout) -> *mut u8 {
        let live =
            LIVE.fetch_add(layout.size() as isize, Ordering::Relaxed) + layout.size() as isize;
        PEAK.fetch_max(live, Ordering::Relaxed);
        unsafe { System.alloc(layout) }
    }
    unsafe fn dealloc(&self, ptr: *mut u8, layout: AllocLayout) {
        LIVE.fetch_sub(layout.size() as isize, Ordering::Relaxed);
        unsafe { System.dealloc(ptr, layout) }
    }
    unsafe fn realloc(&self, ptr: *mut u8, layout: AllocLayout, new_size: usize) -> *mut u8 {
        let delta = new_size as isize - layout.size() as isize;
        let live = LIVE.fetch_add(delta, Ordering::Relaxed) + delta;
        PEAK.fetch_max(live, Ordering::Relaxed);
        unsafe { System.realloc(ptr, layout, new_size) }
    }
}
#[global_allocator]
static ALLOC: Counting = Counting;

/// Bytes allocated by `f` that it keeps, and its peak above the start.
fn heap<T>(f: impl FnOnce() -> T) -> (T, isize, isize) {
    let before = LIVE.load(Ordering::Relaxed);
    PEAK.store(before, Ordering::Relaxed);
    let value = f();
    (
        value,
        LIVE.load(Ordering::Relaxed) - before,
        PEAK.load(Ordering::Relaxed) - before,
    )
}

fn ms<T>(f: impl FnOnce() -> T) -> (T, f64) {
    let started = Instant::now();
    let value = f();
    (value, started.elapsed().as_secs_f64() * 1000.0)
}

fn best<T>(n: usize, mut f: impl FnMut() -> T) -> f64 {
    (0..n)
        .map(|_| ms(|| std::hint::black_box(f())).1)
        .fold(f64::MAX, f64::min)
}

fn room_doc() -> Doc {
    Doc::with_options(Options {
        offset_kind: OffsetKind::Utf16,
        ..Default::default()
    })
}

fn load(layers: &[&[u8]]) -> Doc {
    let doc = room_doc();
    {
        let mut txn = doc.transact_mut();
        for layer in layers {
            txn.apply_update(Update::decode_v1(layer).unwrap()).unwrap();
        }
    }
    doc
}

fn state(doc: &Doc) -> Vec<u8> {
    doc.transact()
        .encode_state_as_update_v1(&StateVector::default())
}

struct Peer {
    doc: EditingDoc,
    sent: Rc<RefCell<Vec<Vec<u8>>>>,
    _subscription: yrs::Subscription,
}

fn open(base: &[u8], chunked: bool, client: u64) -> Peer {
    let doc = EditingDoc::new(client);
    if chunked {
        let envelope = docx_edit::parse_docx_for_edit(base).unwrap();
        overlay::open_chunked(&doc, envelope, &overlay::fingerprint(base)).unwrap();
    } else {
        docx_edit::seed_from_docx(&doc, base).unwrap();
    }
    let sent = Rc::new(RefCell::new(Vec::new()));
    let sink = Rc::clone(&sent);
    let subscription = doc
        .observe_shared_updates(true, move |update| sink.borrow_mut().push(update))
        .unwrap();
    Peer {
        doc,
        sent,
        _subscription: subscription,
    }
}

fn ctx() -> EditCtx {
    EditCtx::local("Bench", "2026-10-08T00:00:00Z")
}

/// Body paragraphs as (first unit, mark index), with text.
fn paragraphs(doc: &EditingDoc) -> (u32, Vec<(u32, u32)>) {
    let (mut offset, mut start, mut out) = (0, 0, Vec::new());
    for segment in doc.story_segments("body").unwrap() {
        match segment.content {
            SegmentContent::Text(text) => offset += text.encode_utf16().count() as u32,
            SegmentContent::Pilcrow(_) => {
                if offset > start {
                    out.push((start, offset));
                }
                offset += 1;
                start = offset;
            }
            SegmentContent::OtherEmbed { .. } => offset += 1,
        }
    }
    (offset, out)
}

/// Runs `scenario` on `peer`; returns the messages it sent.
fn run(peer: &Peer, scenario: &str) -> Vec<Vec<u8>> {
    let doc = &peer.doc;
    let (len, paragraphs) = paragraphs(doc);
    let middle = paragraphs[paragraphs.len() / 2];
    match scenario {
        "one-edit" => {
            doc.insert_text(
                &ctx(),
                Position::new("body", middle.1),
                " edited",
                FormatPolicy::Inherit,
            )
            .unwrap();
        }
        "ten-edits" => {
            for index in 0..10 {
                let (_, paragraphs) = self::paragraphs(doc);
                let (_, mark) = paragraphs[(index * 2 + 1) * paragraphs.len() / 20];
                doc.insert_text(
                    &ctx(),
                    Position::new("body", mark),
                    " edited",
                    FormatPolicy::Inherit,
                )
                .unwrap();
            }
        }
        "typing-200" => {
            for index in 0..200 {
                doc.insert_text(
                    &ctx(),
                    Position::new("body", middle.1 + index),
                    "x",
                    FormatPolicy::Inherit,
                )
                .unwrap();
            }
        }
        "select-all" => {
            doc.delete_range(&ctx(), StoryRange::new("body", 0, len - 1))
                .unwrap();
        }
        other => panic!("unknown scenario {other}"),
    }
    std::mem::take(&mut *peer.sent.borrow_mut())
}

fn main() {
    use std::io::Write;
    let args: Vec<String> = std::env::args().collect();
    let mut out = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(&args[1])
        .unwrap();
    let runs: usize = args[2].parse().unwrap();
    let scenarios: Vec<String> = std::env::var("BENCH_SCENARIOS")
        .unwrap_or_else(|_| "one-edit,ten-edits,typing-200,select-all".into())
        .split(',')
        .map(str::to_owned)
        .collect();
    let heavy = std::env::var("BENCH_ENGINE").map_or(true, |value| value != "0");
    for path in &args[3..] {
        let base = std::fs::read(path).unwrap();
        let name = std::path::Path::new(path)
            .file_stem()
            .unwrap()
            .to_string_lossy()
            .into_owned();
        let sha = office_service::sha256_hex(&base);
        let determinism = Determinism {
            seed: "0000000000000000000000000000000000000000000000000000000000000000",
            now: "2026-10-08T00:00:00.000Z",
        };
        let seeded_seed = office_service::seed(Format::Docx, &base).unwrap();
        // The override layout's local materialization, for the WASM parity check.
        if let Ok(dir) = std::env::var("BENCH_DUMP_DIR") {
            let doc = EditingDoc::new(9);
            let envelope = docx_edit::parse_docx_for_edit(&base).unwrap();
            overlay::open_chunked(&doc, envelope, &overlay::fingerprint(&base)).unwrap();
            std::fs::write(
                format!("{dir}/{name}.materialized.bin"),
                doc.encode_state_as_update_v1(),
            )
            .unwrap();
        }
        let meta_seed = overlay::meta_seed(&overlay::fingerprint(&base));
        for scenario in &scenarios {
            for run_index in 0..runs {
                let order: [bool; 2] = if run_index % 2 == 0 {
                    [false, true]
                } else {
                    [true, false]
                };
                for chunked in order {
                    let seed = if chunked { &meta_seed } else { &seeded_seed };
                    let (peer, open_ms) = ms(|| open(&base, chunked, 701));
                    let entries = peer.doc.overlay().map(|overlay| {
                        (
                            overlay.stories.len(),
                            overlay
                                .stories
                                .iter()
                                .map(|story| story.chunk_writers.len())
                                .sum::<usize>(),
                            overlay.comment_writers.len(),
                        )
                    });
                    let messages = run(&peer, scenario);
                    let sent: usize = messages.iter().map(Vec::len).sum();
                    let room = {
                        let mut layers: Vec<&[u8]> = vec![seed];
                        layers.extend(messages.iter().map(Vec::as_slice));
                        state(&load(&layers))
                    };
                    if let Ok(dir) = std::env::var("BENCH_DUMP_DIR")
                        && run_index == 0
                        && !messages.is_empty()
                    {
                        let stem = format!(
                            "{dir}/{name}.{}.{scenario}",
                            if chunked { "chunked" } else { "seeded" }
                        );
                        let before: Vec<&[u8]> = std::iter::once(seed.as_slice())
                            .chain(messages[..messages.len() - 1].iter().map(Vec::as_slice))
                            .collect();
                        std::fs::write(format!("{stem}.state.bin"), &room).unwrap();
                        std::fs::write(format!("{stem}.before.bin"), state(&load(&before)))
                            .unwrap();
                        std::fs::write(format!("{stem}.last.bin"), messages.last().unwrap())
                            .unwrap();
                    }
                    let seed_sv = load(&[seed]).transact().state_vector();
                    let change = load(&[&room])
                        .transact()
                        .encode_state_as_update_v1(&seed_sv);
                    let state_vector = load(&[&room]).transact().state_vector().len();
                    let (_, room_heap, _) = heap(|| load(&[seed, &change]));
                    let room_load_ms = best(3, || load(&[seed, &change]));
                    let encode_ms = {
                        let doc = load(&[&room]);
                        best(3, || state(&doc))
                    };
                    let mut line = serde_json::json!({
                        "file": name,
                        "scenario": scenario,
                        "layout": if chunked { "chunked" } else { "seeded" },
                        "run": run_index,
                        "baseBytes": base.len(),
                        "seedBytes": seed.len(),
                        "roomBytes": room.len(),
                        "storedChangeBytes": change.len(),
                        "sentBytes": sent,
                        "messages": messages.len(),
                        "lastMessageBytes": messages.last().map_or(0, Vec::len),
                        "stateVectorClients": state_vector,
                        "roomLoadMs": room_load_ms,
                        "roomHeapBytes": room_heap,
                        "wholeEncodeMs": encode_ms,
                        "peerOpenMs": open_ms,
                        "sourceStories": entries.map(|entries| entries.0),
                        "sourceChunks": entries.map(|entries| entries.1),
                        "sourceComments": entries.map(|entries| entries.2),
                    });
                    drop(peer);
                    if heavy {
                        // A session on the room as office-service and the
                        // runtime open it: today the room's whole state, the
                        // override layout its copies over a local materialization.
                        let ((engine, _), engine_heap, engine_peak) = heap(|| {
                            ms(|| {
                                let doc = EditingDoc::new(9);
                                let envelope = docx_edit::parse_docx_for_edit(&base).unwrap();
                                if chunked {
                                    overlay::open_chunked(
                                        &doc,
                                        envelope,
                                        &overlay::fingerprint(&base),
                                    )
                                    .unwrap();
                                    doc.apply_shared_update(&room).unwrap();
                                } else {
                                    drop(envelope);
                                    doc.load_state_v1(&room).unwrap();
                                }
                                doc
                            })
                        });
                        let engine_open_ms = best(2, || {
                            let doc = EditingDoc::new(9);
                            let envelope = docx_edit::parse_docx_for_edit(&base).unwrap();
                            if chunked {
                                overlay::open_chunked(&doc, envelope, &overlay::fingerprint(&base))
                                    .unwrap();
                                doc.apply_shared_update(&room).unwrap();
                            } else {
                                drop(envelope);
                                doc.load_state_v1(&room).unwrap();
                            }
                            doc
                        });
                        // A remote keystroke on that session.
                        let keystroke = if scenario == "typing-200" {
                            messages.last().cloned()
                        } else {
                            None
                        };
                        let remote_ms = keystroke.map(|keystroke| {
                            let before: Vec<&[u8]> = std::iter::once(seed.as_slice())
                                .chain(messages[..messages.len() - 1].iter().map(Vec::as_slice))
                                .collect();
                            let earlier = state(&load(&before));
                            let doc = EditingDoc::new(9);
                            if chunked {
                                let envelope = docx_edit::parse_docx_for_edit(&base).unwrap();
                                overlay::open_chunked(&doc, envelope, &overlay::fingerprint(&base))
                                    .unwrap();
                                doc.apply_shared_update(&earlier).unwrap();
                            } else {
                                doc.load_state_v1(&earlier).unwrap();
                            }
                            ms(|| doc.apply_peer_update_v1(&keystroke).unwrap()).1
                        });
                        drop(engine);
                        let indexed =
                            office_service::baseline(&base, checkpoint(&sha, seed)).unwrap();
                        let (_, effects_ms) = ms(|| {
                            office_service::save_effects(&base, &indexed, checkpoint(&sha, &room))
                                .unwrap()
                        });
                        let (_, export_ms) = ms(|| {
                            office_service::export(&base, checkpoint(&sha, &room), determinism)
                                .unwrap()
                        });
                        let object = line.as_object_mut().unwrap();
                        object.insert(
                            "engineOpenFirstMs".into(),
                            engine_open_first_ms(&base, chunked).into(),
                        );
                        object.insert(
                            "parseMs".into(),
                            best(2, || docx_edit::parse_docx_for_edit(&base).unwrap()).into(),
                        );
                        object.insert(
                            "parseAndLowerMs".into(),
                            best(2, || {
                                overlay::lower_only(docx_edit::parse_docx_for_edit(&base).unwrap())
                                    .unwrap()
                            })
                            .into(),
                        );
                        object.insert("engineOpenMs".into(), engine_open_ms.into());
                        object.insert("engineHeapBytes".into(), engine_heap.into());
                        object.insert("enginePeakBytes".into(), engine_peak.into());
                        object.insert("remoteKeystrokeMs".into(), remote_ms.into());
                        object.insert("saveEffectsMs".into(), effects_ms.into());
                        object.insert("exportMs".into(), export_ms.into());
                    }
                    writeln!(out, "{line}").unwrap();
                    eprintln!("{line}");
                }
            }
        }
    }
}

fn checkpoint<'a>(sha: &'a str, state: &'a [u8]) -> Checkpoint<'a> {
    Checkpoint {
        format: Format::Docx,
        schema_version: 1,
        base_sha256: sha,
        state,
    }
}

/// Parse plus the layout's own seeding: today's whole seed, or the override
/// layout's materialization, without any room state.
fn engine_open_first_ms(base: &[u8], chunked: bool) -> f64 {
    ms(|| {
        let doc = EditingDoc::new(9);
        if chunked {
            let envelope = docx_edit::parse_docx_for_edit(base).unwrap();
            overlay::open_chunked(&doc, envelope, &overlay::fingerprint(base)).unwrap();
        } else {
            docx_edit::seed_from_docx(&doc, base).unwrap();
        }
        doc
    })
    .1
}
