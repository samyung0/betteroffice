//! Today's DOCX room on the native engine, for the override-seeds projection:
//! seed bytes, the server room's load and heap with one agent edit, whole
//! encode per save, the engine's save effects, and the block counts the
//! override layout's copies would be cut from.
//! `docx_today <out.jsonl> <runs> <file.docx>...`

use std::alloc::{GlobalAlloc, Layout as AllocLayout, System};
use std::sync::atomic::{AtomicIsize, Ordering};
use std::time::Instant;

use office_service::{Checkpoint, Command, Format, sha256_hex};
use yrs::updates::decoder::Decode;
use yrs::{Doc, OffsetKind, Options, ReadTxn, StateVector, Transact, Update};

struct Counting;
static LIVE: AtomicIsize = AtomicIsize::new(0);
unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: AllocLayout) -> *mut u8 {
        LIVE.fetch_add(layout.size() as isize, Ordering::Relaxed);
        unsafe { System.alloc(layout) }
    }
    unsafe fn dealloc(&self, ptr: *mut u8, layout: AllocLayout) {
        LIVE.fetch_sub(layout.size() as isize, Ordering::Relaxed);
        unsafe { System.dealloc(ptr, layout) }
    }
    unsafe fn realloc(&self, ptr: *mut u8, layout: AllocLayout, new_size: usize) -> *mut u8 {
        LIVE.fetch_add(
            new_size as isize - layout.size() as isize,
            Ordering::Relaxed,
        );
        unsafe { System.realloc(ptr, layout, new_size) }
    }
}
#[global_allocator]
static ALLOC: Counting = Counting;

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

fn best_ms<T>(n: usize, mut f: impl FnMut() -> T) -> f64 {
    let mut best = f64::MAX;
    for _ in 0..n {
        let started = Instant::now();
        std::hint::black_box(f());
        best = best.min(started.elapsed().as_secs_f64() * 1000.0);
    }
    best
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
    for path in &args[3..] {
        let base = std::fs::read(path).unwrap();
        let name = std::path::Path::new(path)
            .file_stem()
            .unwrap()
            .to_string_lossy()
            .into_owned();
        let sha = sha256_hex(&base);
        let seed = office_service::seed(Format::Docx, &base).unwrap();
        let checkpoint = |state| Checkpoint {
            format: Format::Docx,
            schema_version: 1,
            base_sha256: &sha,
            state,
        };
        let entries = office_service::inspect(&base, checkpoint(&seed)).unwrap();
        let target = &entries[entries.len() / 2];
        let edited = office_service::apply_commands(
            &base,
            checkpoint(&seed),
            &[Command::ReplaceText {
                target_id: target.id.clone(),
                expected_text: target.value.clone(),
                text: format!("{} edited", target.value),
            }],
        )
        .unwrap()
        .state;
        let seed_sv = load(&[&seed]).transact().state_vector();
        let change = load(&[&edited])
            .transact()
            .encode_state_as_update_v1(&seed_sv);
        let indexed = office_service::baseline(&base, checkpoint(&seed)).unwrap();
        for run in 0..runs {
            let before = LIVE.load(Ordering::Relaxed);
            let room = load(&[&seed, &change]);
            let heap = LIVE.load(Ordering::Relaxed) - before;
            let load_ms = best_ms(3, || load(&[&seed, &change]));
            let encode_ms = best_ms(3, || {
                room.transact()
                    .encode_state_as_update_v1(&StateVector::default())
            });
            drop(room);
            let effects_ms = best_ms(2, || {
                office_service::save_effects(&base, &indexed, checkpoint(&edited)).unwrap()
            });
            let line = serde_json::json!({
                "file": name,
                "run": run,
                "baseBytes": base.len(),
                "seedBytes": seed.len(),
                "storedChangeBytes": change.len(),
                "editableEntries": entries.len(),
                "roomLoadMs": load_ms,
                "roomHeapBytes": heap,
                "wholeEncodeMs": encode_ms,
                "saveEffectsMs": effects_ms,
            });
            writeln!(out, "{line}").unwrap();
            eprintln!("{line}");
        }
    }
}
