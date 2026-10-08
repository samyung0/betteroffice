//! One remote keystroke applied by a PPTX editor session, natively, per layout:
//! `remote_apply <state dir> <name> <file.pptx>`.
use pptx_edit::DeckSession;
use std::time::Instant;
fn main() {
    let args: Vec<String> = std::env::args().collect();
    let (dir, name, base) = (&args[1], &args[2], std::fs::read(&args[3]).unwrap());
    for layout in ["seeded", "overrides"] {
        let before = std::fs::read(format!("{dir}/{name}.{layout}.one-edit.before.bin")).unwrap();
        let last = std::fs::read(format!("{dir}/{name}.{layout}.one-edit.last.bin")).unwrap();
        for _ in 0..3 {
            let peer = DeckSession::open_from_update_with_source(&before, &base, 23456).unwrap();
            let t = Instant::now();
            peer.apply_update_v1(&last).unwrap();
            let apply = t.elapsed().as_secs_f64() * 1000.0;
            let t = Instant::now();
            let _ = peer.encode_state_as_update_v1();
            let encode = t.elapsed().as_secs_f64() * 1000.0;
            let t = Instant::now();
            let _ = peer.snapshot().unwrap();
            let snapshot = t.elapsed().as_secs_f64() * 1000.0;
            println!(
                "{name} {layout}: apply_update_v1 {apply:.1} ms (state encode {encode:.1}, snapshot {snapshot:.1})"
            );
        }
    }
}
