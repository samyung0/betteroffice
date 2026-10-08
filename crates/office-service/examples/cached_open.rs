//! Engine opens of a room, natively: today (parse, then the room's whole
//! state), the override layout materializing locally (parse, lower, copy,
//! then the room), and the override layout from a materialization cached per
//! base (parse, load it, then the room). Interleaved, best of each.
//! `cached_open <out.jsonl> <runs> <states dir> <file.docx>...` (rooms from
//! `chunked_bench`'s dumps: `<stem>.<layout>.one-edit.state.bin`).

use std::time::Instant;

use docx_edit::{EditingDoc, overlay};

fn ms<T>(f: impl FnOnce() -> T) -> (T, f64) {
    let started = Instant::now();
    let value = f();
    (value, started.elapsed().as_secs_f64() * 1000.0)
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
    let states = &args[3];
    for path in &args[4..] {
        let base = std::fs::read(path).unwrap();
        let name = std::path::Path::new(path)
            .file_stem()
            .unwrap()
            .to_string_lossy()
            .into_owned();
        let seeded_room =
            std::fs::read(format!("{states}/{name}.seeded.one-edit.state.bin")).unwrap();
        let chunked_room =
            std::fs::read(format!("{states}/{name}.chunked.one-edit.state.bin")).unwrap();
        let materialization = {
            let doc = EditingDoc::new(9);
            let envelope = docx_edit::parse_docx_for_edit(&base).unwrap();
            overlay::open_chunked(&doc, envelope, &overlay::fingerprint(&base)).unwrap();
            doc.encode_state_as_update_v1()
        };
        for run in 0..runs {
            let mut modes = ["today", "materialize", "cached"];
            if run % 2 == 1 {
                modes.reverse();
            }
            for mode in modes {
                let (_, open_ms) = ms(|| {
                    let doc = EditingDoc::new(9);
                    let envelope = docx_edit::parse_docx_for_edit(&base).unwrap();
                    match mode {
                        "today" => {
                            drop(envelope);
                            doc.load_state_v1(&seeded_room).unwrap();
                        }
                        "materialize" => {
                            overlay::open_chunked(&doc, envelope, &overlay::fingerprint(&base))
                                .unwrap();
                            doc.apply_shared_update(&chunked_room).unwrap();
                        }
                        _ => {
                            drop(envelope);
                            doc.load_state_v1(&materialization).unwrap();
                            doc.apply_update_v1(&chunked_room).unwrap();
                        }
                    }
                    doc
                });
                // Save effects and export, as office-service runs them per
                // save and per publication.
                let sha = office_service::sha256_hex(&base);
                let (room, seed) = match mode {
                    "today" => (
                        &seeded_room,
                        office_service::seed(office_service::Format::Docx, &base).unwrap(),
                    ),
                    _ => (
                        &chunked_room,
                        overlay::meta_seed(&overlay::fingerprint(&base)),
                    ),
                };
                // SAFETY of the switch: set before the calls of this mode only.
                unsafe {
                    std::env::set_var(
                        "OFFICE_DOCX_MATERIALIZATION_CACHE",
                        if mode == "cached" { "1" } else { "0" },
                    );
                }
                let checkpoint = |state| office_service::Checkpoint {
                    format: office_service::Format::Docx,
                    schema_version: 1,
                    base_sha256: &sha,
                    state,
                };
                let indexed = office_service::baseline(&base, checkpoint(&seed)).unwrap();
                let (_, effects_ms) =
                    ms(|| office_service::save_effects(&base, &indexed, checkpoint(room)).unwrap());
                let determinism = office_service::Determinism {
                    seed: "0000000000000000000000000000000000000000000000000000000000000000",
                    now: "2026-10-08T00:00:00.000Z",
                };
                let (_, export_ms) =
                    ms(|| office_service::export(&base, checkpoint(room), determinism).unwrap());
                let line = serde_json::json!({
                    "file": name,
                    "mode": mode,
                    "run": run,
                    "openMs": open_ms,
                    "saveEffectsMs": effects_ms,
                    "exportMs": export_ms,
                    "materializationBytes": materialization.len(),
                    "seededRoomBytes": seeded_room.len(),
                });
                writeln!(out, "{line}").unwrap();
                eprintln!("{line}");
            }
        }
    }
}
