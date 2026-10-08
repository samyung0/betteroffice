//! Override seeds (schema 6) against today's seeded deck (schema 5), PPTX, on
//! the same build, interleaved: state bytes, the server room's load, per-save
//! encodes, heap per loaded room, the engine's open and save effects.
//! `override_bench <out.jsonl> <runs> <file.pptx>...`

use std::alloc::{GlobalAlloc, Layout as AllocLayout, System};
use std::sync::atomic::{AtomicIsize, Ordering};
use std::time::Instant;

use office_service::{Checkpoint, Determinism, Format, sha256_hex};
use pptx_edit::{DeckSession, EditCtx, Layout, ShapeSnapshot, StorySnapshot, TextStyle};
use yrs::updates::decoder::Decode;
use yrs::updates::encoder::Encode;
use yrs::{Doc, OffsetKind, Options, ReadTxn, StateVector, Transact, Update};

struct Counting;
static LIVE: AtomicIsize = AtomicIsize::new(0);
static PEAK: AtomicIsize = AtomicIsize::new(0);
unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: AllocLayout) -> *mut u8 {
        let now =
            LIVE.fetch_add(layout.size() as isize, Ordering::Relaxed) + layout.size() as isize;
        PEAK.fetch_max(now, Ordering::Relaxed);
        unsafe { System.alloc(layout) }
    }
    unsafe fn dealloc(&self, ptr: *mut u8, layout: AllocLayout) {
        LIVE.fetch_sub(layout.size() as isize, Ordering::Relaxed);
        unsafe { System.dealloc(ptr, layout) }
    }
    unsafe fn realloc(&self, ptr: *mut u8, layout: AllocLayout, new_size: usize) -> *mut u8 {
        let delta = new_size as isize - layout.size() as isize;
        let now = LIVE.fetch_add(delta, Ordering::Relaxed) + delta;
        PEAK.fetch_max(now, Ordering::Relaxed);
        unsafe { System.realloc(ptr, layout, new_size) }
    }
}
#[global_allocator]
static ALLOC: Counting = Counting;

fn live() -> isize {
    LIVE.load(Ordering::Relaxed)
}

/// Peak heap above the current level while `f` runs, and what stays after it.
fn measure<T>(f: impl FnOnce() -> T) -> (f64, isize, isize, T) {
    let before = live();
    PEAK.store(before, Ordering::Relaxed);
    let started = Instant::now();
    let value = f();
    let ms = started.elapsed().as_secs_f64() * 1000.0;
    (
        ms,
        PEAK.load(Ordering::Relaxed) - before,
        live() - before,
        value,
    )
}

fn best_ms<T>(n: usize, mut f: impl FnMut() -> T) -> f64 {
    let mut best = f64::MAX;
    for _ in 0..n {
        let started = Instant::now();
        std::hint::black_box(f());
        best = best.min(started.elapsed().as_secs_f64() * 1000.0);
    }
    best
}

/// The Rust server's room document: UTF-16 offsets, GC on, a random client.
fn room_doc() -> Doc {
    Doc::with_options(Options {
        offset_kind: OffsetKind::Utf16,
        ..Default::default()
    })
}

/// `replay`: seed plus stored change in one transaction.
fn load_room(layers: &[&[u8]]) -> Doc {
    let doc = room_doc();
    {
        let mut txn = doc.transact_mut();
        for layer in layers {
            txn.apply_update(Update::decode_v1(layer).unwrap()).unwrap();
        }
    }
    doc
}

fn encode(doc: &Doc) -> Vec<u8> {
    doc.transact()
        .encode_state_as_update_v1(&StateVector::default())
}

fn all_shapes(shapes: &[ShapeSnapshot]) -> Vec<&ShapeSnapshot> {
    let mut out = Vec::new();
    for shape in shapes {
        out.push(shape);
        out.extend(all_shapes(&shape.children));
    }
    out
}

/// Text stories per slide, in deck order.
fn stories_by_slide(session: &DeckSession) -> Vec<Vec<StorySnapshot>> {
    session
        .snapshot()
        .unwrap()
        .slides
        .iter()
        .map(|slide| {
            all_shapes(&slide.shapes)
                .into_iter()
                .flat_map(|shape| shape.text_stories.clone())
                .filter(|story| story.length > 1)
                .collect()
        })
        .collect()
}

#[derive(Clone, Copy)]
enum Scenario {
    /// One word typed into a story on the middle slide.
    OneEdit,
    /// One word in each of ten stories on ten slides.
    TenEdits,
    /// 200 keystrokes into one story.
    Typing,
    /// One shape moved.
    MoveShape,
}

impl Scenario {
    fn name(self) -> &'static str {
        match self {
            Scenario::OneEdit => "one-edit",
            Scenario::TenEdits => "ten-edits",
            Scenario::Typing => "typing-200",
            Scenario::MoveShape => "move-shape",
        }
    }
}

/// Runs the scenario; returns the state before its last keystroke (the
/// durable state of a delta save) and after.
fn edit(session: &DeckSession, scenario: Scenario) -> (Vec<u8>, Vec<u8>) {
    let ctx = EditCtx::local("bench");
    let style = TextStyle::default();
    let slides = stories_by_slide(session);
    let with_text: Vec<usize> = (0..slides.len())
        .filter(|i| !slides[*i].is_empty())
        .collect();
    let middle = with_text[with_text.len() / 2];
    let mut before_last = session.encode_state_as_update_v1();
    match scenario {
        Scenario::OneEdit => {
            let story = &slides[middle][0];
            session
                .insert_text(&ctx, &story.id, 1, "edited ", &style)
                .unwrap();
        }
        Scenario::TenEdits => {
            for slide in with_text
                .iter()
                .step_by((with_text.len() / 10).max(1))
                .take(10)
            {
                before_last = session.encode_state_as_update_v1();
                let story = &slides[*slide][0];
                session
                    .insert_text(&ctx, &story.id, 1, "edited ", &style)
                    .unwrap();
            }
        }
        Scenario::Typing => {
            let story = &slides[middle][0];
            for key in 0..200_u32 {
                before_last = session.encode_state_as_update_v1();
                session
                    .insert_text(&ctx, &story.id, 1 + key, "x", &style)
                    .unwrap();
            }
        }
        Scenario::MoveShape => {
            let deck = session.snapshot().unwrap();
            let slide = &deck.slides[middle];
            let shape = &slide.shapes[0];
            session
                .move_shape(&ctx, &slide.id, &shape.id, shape.x + 12_700, shape.y)
                .unwrap();
        }
    }
    (before_last, session.encode_state_as_update_v1())
}

fn state_vector(update: &[u8]) -> StateVector {
    let doc = room_doc();
    doc.transact_mut()
        .apply_update(Update::decode_v1(update).unwrap())
        .unwrap();
    doc.transact().state_vector()
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let out_path = &args[1];
    let runs: usize = args[2].parse().unwrap();
    let mut out = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(out_path)
        .unwrap();
    use std::io::Write;
    let scenarios = [
        Scenario::OneEdit,
        Scenario::TenEdits,
        Scenario::Typing,
        Scenario::MoveShape,
    ];
    for path in &args[3..] {
        let base = std::fs::read(path).unwrap();
        let name = std::path::Path::new(path)
            .file_stem()
            .unwrap()
            .to_string_lossy()
            .into_owned();
        let base_sha = sha256_hex(&base);
        // The indexed baseline Capy caches per base: seed(base)'s entries.
        let seeded_seed = office_service::seed(Format::Pptx, &base).unwrap();
        let indexed = office_service::baseline(
            &base,
            Checkpoint {
                format: Format::Pptx,
                schema_version: 1,
                base_sha256: &base_sha,
                state: &seeded_seed,
            },
        )
        .unwrap();
        for run in 0..runs {
            let layouts = if run % 2 == 0 {
                [Layout::Seeded, Layout::Overrides]
            } else {
                [Layout::Overrides, Layout::Seeded]
            };
            for layout in layouts {
                let label = match layout {
                    Layout::Seeded => "seeded",
                    Layout::Overrides => "overrides",
                };
                // Seed: what a NULL row stands for, built from the file.
                let (seed_ms, _, _, seed_session) =
                    measure(|| DeckSession::open_with_layout(&base, 42, layout).unwrap());
                let seed = seed_session.encode_state_as_update_v1();
                drop(seed_session);
                let seed_sv = state_vector(&seed);
                for scenario in scenarios {
                    let session =
                        DeckSession::open_from_update_with_source(&seed, &base, 43).unwrap();
                    let (before_last, state) = edit(&session, scenario);
                    drop(session);
                    if let (Ok(dir), 0) = (std::env::var("DUMP_DIR"), run) {
                        let stem = format!("{dir}/{name}.{label}.{}", scenario.name());
                        std::fs::write(format!("{stem}.state.bin"), &state).unwrap();
                        // The last keystroke alone: what a co-editor receives.
                        let last = {
                            let doc = load_room(&[&state]);
                            doc.transact()
                                .encode_state_as_update_v1(&state_vector(&before_last))
                        };
                        std::fs::write(format!("{stem}.last.bin"), &last).unwrap();
                        std::fs::write(format!("{stem}.before.bin"), &before_last).unwrap();
                    }
                    let change = {
                        let doc = load_room(&[&state]);
                        doc.transact().encode_state_as_update_v1(&seed_sv)
                    };
                    // Server room: load seed + change, heap held, encodes per save.
                    let (load_ms, load_peak, room_heap, room) =
                        measure(|| load_room(&[&seed, &change]));
                    let load_best = best_ms(5, || load_room(&[&seed, &change]));
                    let whole_ms = best_ms(5, || encode(&room));
                    let durable_sv = state_vector(&before_last);
                    let delta = room.transact().encode_state_as_update_v1(&durable_sv);
                    let delta_ms =
                        best_ms(5, || room.transact().encode_state_as_update_v1(&durable_sv));
                    let sv_bytes = room.transact().state_vector().encode_v1().len();
                    let merge_ms = best_ms(3, || {
                        let doc = load_room(&[&seed]);
                        doc.transact_mut()
                            .apply_update(Update::decode_v1(&encode(&room)).unwrap())
                            .unwrap();
                        encode(&doc)
                    });
                    drop(room);
                    // Engine: open from base + whole state, save effects, export.
                    let checkpoint = Checkpoint {
                        format: Format::Pptx,
                        schema_version: 1,
                        base_sha256: &base_sha,
                        state: &state,
                    };
                    let (open_ms, open_peak, open_heap, engine) = measure(|| {
                        DeckSession::open_from_update_with_source(&state, &base, 44).unwrap()
                    });
                    let engine_doc_heap = {
                        let before = live();
                        let doc = load_room(&[&state]);
                        let heap = live() - before;
                        drop(doc);
                        heap
                    };
                    drop(engine);
                    let (effects_ms, effects_peak, _, effects) = measure(|| {
                        office_service::save_effects(&base, &indexed, checkpoint).unwrap()
                    });
                    let effects_best = best_ms(2, || {
                        office_service::save_effects(&base, &indexed, checkpoint).unwrap()
                    });
                    let (export_ms, export_peak, _, exported) = measure(|| {
                        office_service::export(
                            &base,
                            checkpoint,
                            Determinism {
                                seed: &"0".repeat(64),
                                now: "2026-10-08T00:00:00.000Z",
                            },
                        )
                        .unwrap()
                    });
                    let line = serde_json::json!({
                        "file": name,
                        "baseBytes": base.len(),
                        "layout": label,
                        "run": run,
                        "scenario": scenario.name(),
                        "seedBytes": seed.len(),
                        "seedOpenMs": seed_ms,
                        "roomStateBytes": state.len(),
                        "storedChangeBytes": change.len(),
                        "deltaRowBytes": delta.len(),
                        "stateVectorBytes": sv_bytes,
                        "roomLoadMs": load_best,
                        "roomLoadFirstMs": load_ms,
                        "roomLoadPeakBytes": load_peak,
                        "roomHeapBytes": room_heap,
                        "wholeEncodeMs": whole_ms,
                        "deltaEncodeMs": delta_ms,
                        "mergeRebuildMs": merge_ms,
                        "engineOpenMs": open_ms,
                        "engineOpenPeakBytes": open_peak,
                        "engineOpenHeldBytes": open_heap,
                        "engineDocHeapBytes": engine_doc_heap,
                        "saveEffectsMs": effects_best,
                        "saveEffectsFirstMs": effects_ms,
                        "saveEffectsPeakBytes": effects_peak,
                        "effects": effects.len(),
                        "effectIds": effects.iter().map(|e| format!("{}:{:?}", e.id, e.operation)).collect::<Vec<_>>(),
                        "exportMs": export_ms,
                        "exportPeakBytes": export_peak,
                        "exportBytes": exported.len(),
                        "exportSha": sha256_hex(&exported),
                    });
                    writeln!(out, "{line}").unwrap();
                    eprintln!(
                        "{name} {label} run {run} {}: state {} B, change {} B, load {:.2} ms, heap {} KiB, effects {:.1} ms",
                        scenario.name(),
                        state.len(),
                        change.len(),
                        load_best,
                        room_heap / 1024,
                        effects_best
                    );
                }
            }
        }
    }
}
