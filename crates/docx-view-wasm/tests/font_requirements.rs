//! The faces a document's layout asks for: an East Asian face only for text
//! that uses it.

use docx_view_wasm::DocxViewDocument;
use serde_json::Value;

fn required_families(path: &str) -> Vec<String> {
    let bytes = std::fs::read(format!("{}/{path}", env!("CARGO_MANIFEST_DIR"))).unwrap();
    let document = DocxViewDocument::open(&bytes).unwrap();
    let requirements: Vec<Value> = serde_json::from_str(
        &document
            .font_requirements_json(&document.layout_request_json())
            .unwrap(),
    )
    .unwrap();
    requirements
        .iter()
        .map(|requirement| requirement["family"].as_str().unwrap().to_owned())
        .collect()
}

#[test]
fn an_english_document_on_an_east_asian_theme_asks_for_no_east_asian_face() {
    let families = required_families("tests/fixtures/english-east-asian-theme.docx");
    assert!(
        families.iter().any(|family| family == "Calibri"),
        "{families:?}"
    );
    assert!(
        !families
            .iter()
            .any(|family| ooxml_text::word_fonts::is_east_asian_family(family)),
        "{families:?}"
    );
}

#[test]
fn a_japanese_document_asks_for_its_east_asian_faces() {
    let families = required_families("../../poc/fixtures/exchange-plan.docx");
    assert!(
        families
            .iter()
            .any(|family| ooxml_text::word_fonts::is_east_asian_family(family)),
        "{families:?}"
    );
}
