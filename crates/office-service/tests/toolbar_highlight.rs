//! A highlight the editor's toolbar sets (docx-edit `format_range`) is saved:
//! the toolbar wrote it as `{color}` while the save read a plain colour name,
//! so it was dropped on export.

use docx_edit::{EditCtx, EditingDoc, InlineFormatDelta, Patch, StoryRange};
use office_service::{Checkpoint, Determinism, Format, export, seed, sha256_hex};

#[test]
fn a_toolbar_highlight_is_saved() {
    let base = include_bytes!("fixtures/stories.docx");
    let doc = EditingDoc::new(4101);
    doc.load_state_v1(&seed(Format::Docx, base).unwrap())
        .unwrap();
    doc.format_range(
        &EditCtx::local("", ""),
        StoryRange::new("body", 0, 5),
        &InlineFormatDelta {
            highlight: Patch::Set("#FFFF00".into()),
            ..InlineFormatDelta::default()
        },
    )
    .unwrap();
    let state = doc.encode_state_as_update_v1();
    let sha = sha256_hex(base);
    let saved = export(
        base,
        Checkpoint {
            format: Format::Docx,
            schema_version: 1,
            base_sha256: &sha,
            state: &state,
        },
        Determinism {
            seed: &"0".repeat(64),
            now: "2026-10-21T00:00:00.000Z",
        },
    )
    .unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let document = parts
        .iter()
        .find(|(path, _)| path == "word/document.xml")
        .map(|(_, xml)| String::from_utf8_lossy(xml).into_owned())
        .unwrap();
    // The run holding the highlighted text carries it.
    let text = document.find(">Intro</w:t>").expect("the highlighted run");
    let run = &document[document[..text].rfind("<w:r>").unwrap()..text];
    assert!(run.contains(r#"<w:highlight w:val="yellow"/>"#), "{run}");
}
