#![cfg(feature = "base-latin")]
//! Upstream #2093's engine check; its paged-export case needs upstream's
//! structured export, which Capy's engine does not have.

#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use docx_edit::{EngineSession, seed_from_docx};
use ooxml_text::FontId;
use serde_json::{Value, json};

#[test]
fn backstop_measurements_remeasure_when_the_first_real_font_arrives() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let bytes = fixture::unrevised_docx();
    let engine = EngineSession::new(9665);
    seed_from_docx(engine.doc(), &bytes).unwrap();
    let backstop = FontId::base_latin(false, false).to_u32();
    let mut request = fixture::region_request(&engine, &bytes, backstop);
    let before: Value = serde_json::from_str(
        &engine
            .layout_document_with_regions_json(&request.to_string())
            .unwrap(),
    )
    .unwrap();
    let real = docx_layout::register_measure_font_bytes(fixture::OTHER_FONT).unwrap();
    assert_eq!(real, 0);
    for chain in request["measurement"]["fontChains"]
        .as_object_mut()
        .unwrap()
        .values_mut()
    {
        *chain = json!([real]);
    }
    let warm: Value = serde_json::from_str(
        &engine
            .layout_document_with_regions_json(&request.to_string())
            .unwrap(),
    )
    .unwrap();
    assert_ne!(before["measured"], warm["measured"]);
    let cold: Value = docx_layout::with_private_measure_fonts(|| {
        assert_eq!(
            docx_layout::register_measure_font_bytes(fixture::OTHER_FONT).unwrap(),
            0
        );
        let cold = EngineSession::new(9665);
        seed_from_docx(cold.doc(), &bytes).unwrap();
        serde_json::from_str(
            &cold
                .layout_document_with_regions_json(&request.to_string())
                .unwrap(),
        )
        .unwrap()
    });
    assert_eq!(warm, cold);
}
