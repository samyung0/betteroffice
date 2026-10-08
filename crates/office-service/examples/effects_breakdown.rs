//! Where a PPTX save's engine time goes, per layout: parse the source, build
//! the source index, open the session from the room state, snapshot, the
//! effects baseline, one slide's scope (what an effects pass over the touched
//! slides only would read). `effects_breakdown <runs> <state dir> <file.pptx>...`

use std::time::Instant;

use office_service::{Checkpoint, Format, sha256_hex};
use pptx_edit::DeckSession;

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
    let args: Vec<String> = std::env::args().collect();
    let runs: usize = args[1].parse().unwrap();
    let dir = &args[2];
    for path in &args[3..] {
        let base = std::fs::read(path).unwrap();
        let name = std::path::Path::new(path)
            .file_stem()
            .unwrap()
            .to_string_lossy()
            .into_owned();
        let sha = sha256_hex(&base);
        let parse_ms = best_ms(runs, || pptx_parse::parse_pptx(&base).unwrap());
        for layout in ["seeded", "overrides"] {
            let state = std::fs::read(format!("{dir}/{name}.{layout}.one-edit.state.bin")).unwrap();
            let open_ms = best_ms(runs, || {
                DeckSession::open_from_update_with_source(&state, &base, 77).unwrap()
            });
            let session = DeckSession::open_from_update_with_source(&state, &base, 77).unwrap();
            let snapshot_ms = best_ms(runs, || session.snapshot().unwrap());
            let slides = session.slide_ids().unwrap().len();
            let scope_ms = best_ms(runs, || session.slide_scope(slides / 2).unwrap());
            let baseline_ms = best_ms(runs, || {
                office_service::baseline(
                    &base,
                    Checkpoint {
                        format: Format::Pptx,
                        schema_version: 1,
                        base_sha256: &sha,
                        state: &state,
                    },
                )
                .unwrap()
            });
            println!(
                "{}",
                serde_json::json!({
                    "file": name,
                    "layout": layout,
                    "slides": slides,
                    "parseMs": parse_ms,
                    "openMs": open_ms,
                    "snapshotMs": snapshot_ms,
                    "oneSlideScopeMs": scope_ms,
                    "baselineMs": baseline_ms,
                })
            );
        }
    }
}
