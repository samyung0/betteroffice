//! Pins the bytes a DOCX seed and save produce. The memory round's export fixes
//! (working copies dropped, forms dropped, batched bookmark positions, story
//! segments) kept these bytes; a change here is a change to what Word gets.
//! Saving runs with their source formatting and cells with only the borders
//! and margins an edit changed (Epo 2026-10-21) changed the three fixtures
//! with styles; the others have none to resolve.
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
            "738686f6b8d95ac9f1bdf3a566e92709fbf05ec49660132c88981f289e9d29c7",
        ),
        (
            "exchange-plan",
            include_bytes!("../../../poc/fixtures/exchange-plan.docx"),
            "f14019ebb6d2259efcfa086e3cb0a04a42ca9c7626a04aeaa18bf3d8540f57b4",
        ),
        (
            "feature-rich",
            include_bytes!("../../../poc/fixtures/feature-rich.docx"),
            "3ddeff1960b564bd34a6d28618d205c17dfde500d286a34a8fcf1d9d2a57107f",
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
