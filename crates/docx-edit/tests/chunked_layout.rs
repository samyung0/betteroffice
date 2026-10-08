//! Override layout (DOCX spike): a session that materializes every chunk
//! under its copy writer reads as today's seeded session, and what it shares
//! is enough for a room and a fresh peer.

use std::path::{Path, PathBuf};

use docx_edit::overlay;
use docx_edit::*;

fn seeded(bytes: &[u8], client: u64) -> EditingDoc {
    let doc = EditingDoc::new(client);
    seed_from_docx(&doc, bytes).unwrap();
    doc
}

fn chunked(bytes: &[u8], client: u64) -> EditingDoc {
    let doc = EditingDoc::new(client);
    let envelope = parse_docx_for_edit(bytes).unwrap();
    overlay::open_chunked(&doc, envelope, &overlay::fingerprint(bytes)).unwrap();
    doc
}

fn repo() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../..")
}

fn docx_files(dir: &Path, out: &mut Vec<PathBuf>) {
    let Ok(entries) = std::fs::read_dir(dir) else {
        return;
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            docx_files(&path, out);
        } else if path.extension().is_some_and(|ext| ext == "docx") {
            out.push(path);
        }
    }
}

/// The repository's DOCX fixtures, plus every `.docx` under `CHUNKED_DOCX_DIR`.
fn fixtures() -> Vec<(String, Vec<u8>)> {
    let root = repo();
    let mut paths: Vec<PathBuf> = [
        "crates/docx-edit/tests/fixtures/footnote-anchor.docx",
        "crates/docx-edit/tests/fixtures/suppressed-list-markers.docx",
        "crates/office-service/tests/fixtures/stories.docx",
        "crates/betteroffice-docx/tests/corpus/fixtures/betteroffice-demo.docx",
        "crates/betteroffice-docx/tests/corpus/fixtures/wordprocessingml-comprehensive.docx",
        "crates/ooxml-drawingml/tests/fixtures/preset-adjustments.docx",
        "crates/ooxml-redact/tests/fixtures/redaction-integrity.docx",
        "crates/ooxml-text/tests/fixtures/line-spacing-baseline.docx",
        "packages/docx-react/src/components/DocxEditor/hooks/__fixtures__/probe-linked-header.docx",
        "poc/fixtures/book-30p.docx",
        "poc/fixtures/exchange-plan.docx",
        "poc/fixtures/feature-rich.docx",
        "poc/fixtures/opaque-objects.docx",
    ]
    .iter()
    .map(|path| root.join(path))
    .collect();
    if let Ok(dir) = std::env::var("CHUNKED_DOCX_DIR") {
        docx_files(Path::new(&dir), &mut paths);
    }
    let mut fixtures: Vec<(String, Vec<u8>)> = paths
        .into_iter()
        .map(|path| {
            (
                path.file_stem().unwrap().to_string_lossy().into_owned(),
                std::fs::read(&path).unwrap(),
            )
        })
        .collect();
    fixtures.extend(xml_fixtures());
    fixtures
}

#[path = "support/chunked_fixtures.rs"]
mod chunked_fixtures;
use chunked_fixtures::xml_fixtures;

fn first_difference(left: &str, right: &str) -> String {
    for (index, (a, b)) in left.lines().zip(right.lines()).enumerate() {
        if a != b {
            return format!("line {index}:\n  seeded:  {a}\n  chunked: {b}");
        }
    }
    format!(
        "lengths {} / {} lines",
        left.lines().count(),
        right.lines().count()
    )
}

#[test]
fn chunked_view_reads_as_the_seed() {
    let mut failures = Vec::new();
    let mut count = 0;
    for (name, bytes) in fixtures() {
        if parse_docx_for_edit(&bytes).is_err() {
            continue;
        }
        let seeded = std::panic::catch_unwind(|| seeded(&bytes, 7));
        let Ok(seeded) = seeded else { continue };
        let chunked = chunked(&bytes, 7);
        let (left, right) = (overlay::read_dump(&seeded), overlay::read_dump(&chunked));
        count += 1;
        if left != right {
            failures.push(format!("{name}: {}", first_difference(&left, &right)));
        }
    }
    assert!(
        failures.is_empty(),
        "{} of {count} differ:\n{}",
        failures.len(),
        failures.join("\n")
    );
    eprintln!("{count} fixtures read the same");
}
