use std::collections::BTreeMap;

use pptx_edit::{DeckSession, EditCtx, ShapeSnapshot, ShapeStroke};

const SOURCE: &[u8] = include_bytes!("fixtures/deck-schema-v2-connectors.pptx");
const SLIDE: &str = "ppt/slides/slide1.xml";

fn context() -> EditCtx {
    EditCtx::local("connectors")
}

fn parts(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    ooxml_opc::unzip_parts(bytes).unwrap().into_iter().collect()
}

fn slide_xml(bytes: &[u8]) -> String {
    String::from_utf8(parts(bytes).remove(SLIDE).unwrap()).unwrap()
}

fn element<'a>(xml: &'a str, open: &str, close: &str) -> &'a str {
    let start = xml.find(open).unwrap_or_else(|| panic!("missing {open}"));
    let end = start + xml[start..].find(close).unwrap() + close.len();
    &xml[start..end]
}

fn connector(xml: &str) -> &str {
    element(xml, "<p:cxnSp>", "</p:cxnSp>")
}

fn shape(xml: &str, id: u32) -> &str {
    element(
        xml,
        &format!("<p:sp><p:nvSpPr><p:cNvPr id=\"{id}\""),
        "</p:sp>",
    )
}

fn shape_named(session: &DeckSession, name: &str) -> (String, ShapeSnapshot) {
    let slide = session.snapshot().unwrap().slides.remove(0);
    let shape = slide.shapes.into_iter().find(|shape| shape.name == name);
    (
        slide.id,
        shape.unwrap_or_else(|| panic!("no shape named {name}")),
    )
}

#[test]
fn a_fresh_deck_models_the_connector_and_saves_byte_for_byte() {
    let session = DeckSession::open(SOURCE, 20).unwrap();
    let names: Vec<String> = session.snapshot().unwrap().slides[0]
        .shapes
        .iter()
        .map(|shape| shape.name.clone())
        .collect();
    assert_eq!(names, ["Before", "Straight Arrow Connector 2", "After"]);
    assert_eq!(parts(&session.save().unwrap()), parts(SOURCE));
}

#[test]
fn an_edit_after_a_connector_reaches_the_shape_after_it() {
    let session = DeckSession::open(SOURCE, 21).unwrap();
    let (slide_id, after) = shape_named(&session, "After");
    session
        .move_shape(&context(), &slide_id, &after.id, 952_500, 1_047_750)
        .unwrap();

    let source = slide_xml(SOURCE);
    let saved = slide_xml(&session.save().unwrap());
    assert!(shape(&saved, 4).contains(r#"<a:off x="952500" y="1047750"/>"#));
    assert_eq!(connector(&saved), connector(&source));
    assert_eq!(shape(&saved, 2), shape(&source, 2));
}

#[test]
fn a_connector_edit_keeps_its_joins() {
    let session = DeckSession::open(SOURCE, 22).unwrap();
    let (slide_id, connector_shape) = shape_named(&session, "Straight Arrow Connector 2");
    session
        .set_shape_stroke(
            &context(),
            &slide_id,
            &connector_shape.id,
            &ShapeStroke {
                color: Some("#0000FF".to_owned()),
                width_pt: Some(3.0),
            },
        )
        .unwrap();

    let saved = slide_xml(&session.save().unwrap());
    let written = connector(&saved);
    assert!(written.starts_with(
        r#"<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="3" name="Straight Arrow Connector 2"/><p:cNvCxnSpPr><a:stCxn id="2" idx="1"/><a:endCxn id="4" idx="2"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr>"#
    ));
    assert!(written.contains(r#"<a:srgbClr val="0000FF"/>"#));
    assert_eq!(shape(&saved, 4), shape(&slide_xml(SOURCE), 4));
}
