//! Pins the bytes a DOCX seed and save produce. The memory round's export fixes
//! (working copies dropped, forms dropped, batched bookmark positions, story
//! segments) kept these bytes; a change here is a change to what Word gets.
//! The fixtures cover headers and footers, footnotes, fields and a TOC, tables,
//! comments, and bookmarks in the body, a header and a footnote.

use office_service::{Checkpoint, Determinism, Format, export, seed, sha256_hex};

fn saved(base: &[u8]) -> String {
    let state = seed(Format::Docx, base).unwrap();
    let base_sha256 = sha256_hex(base);
    let saved = export(
        base,
        Checkpoint {
            format: Format::Docx,
            schema_version: 1,
            base_sha256: &base_sha256,
            state: &state,
        },
        Determinism {
            seed: &sha256_hex(b"capy-s0"),
            now: "2026-10-20T00:00:00.000Z",
        },
    )
    .unwrap();
    sha256_hex(&saved)
}

#[test]
fn a_seeded_document_saves_the_same_bytes() {
    let cases: [(&str, &[u8], &str); 6] = [
        (
            "wordprocessingml-comprehensive",
            include_bytes!("../../../poc/fixtures/wordprocessingml-comprehensive.docx"),
            "e14a1c63f619b775d41b8fba08d54027d81c7000d641feb909031346804d3a21",
        ),
        (
            "exchange-plan",
            include_bytes!("../../../poc/fixtures/exchange-plan.docx"),
            "650cf18ca39464d9570b4477c735395856164dae0cd7a136a9084ee22e8d1d78",
        ),
        (
            "feature-rich",
            include_bytes!("../../../poc/fixtures/feature-rich.docx"),
            "bf8fee4c2f9f37c31c41e18781e7eae942f838760e3a147278a3a5e76e608d9a",
        ),
        (
            "stories",
            include_bytes!("fixtures/stories.docx"),
            "2e6f561efa51dbdc911c7d827aecbdb4568994078690796728c9cf8ac276bbc1",
        ),
        (
            "footnote-anchor",
            include_bytes!("../../docx-edit/tests/fixtures/footnote-anchor.docx"),
            "f57623eef9ab30f945a7fefd80e8591f78d7d7df533795d6e84fa8cc79e6ade0",
        ),
        (
            "bookmarks-across-stories",
            include_bytes!("fixtures/bookmarks-across-stories.docx"),
            "f3b79a27cd4a3299cfd3838b534f44a509e209420d7f3abcf7ebe2d0c488b197",
        ),
    ];
    for (name, base, golden) in cases {
        assert_eq!(saved(base), golden, "{name}");
    }
}
