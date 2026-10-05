//! `w:ptab` lays out as ECMA-376 §17.3.3.23 describes: the text after it is
//! aligned left, centred or right at the margin or indent position, with the
//! tab's own leader, and a position the line has already passed is found on
//! the next line.

use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
/// Letter page, 1in margins: the text area runs from 96 to 720 px.
const LEFT: f64 = 96.0;
const RIGHT: f64 = 720.0;
const CENTER: f64 = (LEFT + RIGHT) / 2.0;

const RPR: &str = r#"<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr>"#;

fn r(text: &str) -> String {
    format!(r#"<w:r>{RPR}<w:t xml:space="preserve">{text}</w:t></w:r>"#)
}

fn ptab(alignment: &str, relative_to: &str, leader: &str) -> String {
    format!(
        r#"<w:r>{RPR}<w:ptab w:relativeTo="{relative_to}" w:alignment="{alignment}" w:leader="{leader}"/></w:r>"#
    )
}

fn p(inner: &str, ppr: &str) -> String {
    format!(r#"<w:p><w:pPr><w:spacing w:before="0" w:after="120"/>{ppr}</w:pPr>{inner}</w:p>"#)
}

fn document(header: &str, body: &str) -> Vec<u8> {
    const W: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;
    let sect = r#"<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/><w:cols w:space="720"/></w:sectPr>"#;
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/_rels/document.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml", format!(r#"<w:document {W}><w:body>{body}{sect}</w:body></w:document>"#)),
        ("word/header1.xml", format!(r#"<w:hdr {W}>{header}</w:hdr>"#)),
    ];
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name.to_owned(), value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

/// The first page's text and decoration primitives.
fn page(header: &str, body: &str) -> Vec<Value> {
    let bytes = document(header, body);
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(&bytes, Default::default())
        .unwrap()
        .document
        .package;
    let engine = EngineSession::new(71733);
    seed_from_docx(engine.doc(), &bytes).unwrap();
    let layout = engine
        .layout_document_with_regions_json(
            &json!({
                "bodyStory": "body", "renderEnv": {},
                "regions": {"sections": [{"properties": package.document.final_section_properties}], "settings": package.settings},
                "measurement": {"fontChains": {"arial|0|0": [font]}, "defaults": {"fontFamily": "Arial", "fontSize": 12}}
            })
            .to_string(),
        )
        .unwrap();
    let display: Value =
        serde_json::from_str(&engine.build_display_list_json(&layout).unwrap()).unwrap();
    let page = &display["pages"][0];
    [&page["header"]["primitives"], &page["primitives"]]
        .into_iter()
        .filter_map(Value::as_array)
        .flatten()
        .cloned()
        .collect()
}

/// Where `label` is drawn: (left, right, baseline). Text paints a primitive per
/// cluster, so the label is found in each line's glyphs read left to right.
fn span(prims: &[Value], label: &str) -> (f64, f64, f64) {
    let mut glyphs: Vec<(f64, f64, f64, &str)> = prims
        .iter()
        .filter(|p| p["kind"] == "text" && p["leaderGlyphs"].is_null())
        .map(|p| {
            let x = p["x"].as_f64().unwrap();
            (
                p["baselineY"].as_f64().unwrap(),
                x,
                x + p["width"].as_f64().unwrap(),
                p["text"].as_str().unwrap(),
            )
        })
        .collect();
    glyphs.sort_by(|a, b| a.0.total_cmp(&b.0).then(a.1.total_cmp(&b.1)));
    for line in glyphs.chunk_by(|a, b| a.0 == b.0) {
        let text: String = line.iter().map(|g| g.3).collect();
        if let Some(at) = text.find(label) {
            let mut offset = 0;
            let (mut first, mut last) = (None, None);
            for glyph in line {
                if offset >= at && offset < at + label.len() {
                    first.get_or_insert(glyph.1);
                    last = Some(glyph.2);
                }
                offset += glyph.3.len();
            }
            return (first.unwrap(), last.unwrap(), line[0].0);
        }
    }
    panic!("{label:?} is not drawn");
}

fn x(prims: &[Value], label: &str) -> f64 {
    span(prims, label).0
}

fn end(prims: &[Value], label: &str) -> f64 {
    span(prims, label).1
}

fn mid(prims: &[Value], label: &str) -> f64 {
    let (left, right, _) = span(prims, label);
    (left + right) / 2.0
}

fn near(actual: f64, expected: f64, what: &str) {
    assert!(
        (actual - expected).abs() < 1.0,
        "{what}: {actual} is not {expected}"
    );
}

#[test]
fn margin_tabs_centre_and_right_align_a_header() {
    let header = p(
        &format!(
            "{}{}{}{}{}",
            r("Left text"),
            ptab("center", "margin", "none"),
            r("Centered"),
            ptab("right", "margin", "dot"),
            r("Page 7")
        ),
        "",
    );
    let prims = page(&header, &p(&r("Body"), ""));
    near(x(&prims, "Left text"), LEFT, "left text");
    near(mid(&prims, "Centered"), CENTER, "centred text");
    near(end(&prims, "Page 7"), RIGHT, "right text");
    // the dot leader fills the gap before the right-aligned text
    let leader = prims
        .iter()
        .find(|p| p["kind"] == "decoration" && p["dotted"] == true)
        .expect("dot leader");
    let leader_x = leader["x"].as_f64().unwrap();
    near(leader_x, end(&prims, "Centered"), "leader start");
    near(
        leader_x + leader["w"].as_f64().unwrap(),
        x(&prims, "Page 7"),
        "leader end",
    );
    // the centred tab has no leader of its own
    assert_eq!(
        prims.iter().filter(|p| p["kind"] == "decoration").count(),
        1
    );
}

#[test]
fn indent_tabs_align_between_the_paragraph_indents() {
    // 1in left and 0.5in right indent: the indents run from 192 to 672 px.
    let indents = r#"<w:ind w:left="1440" w:right="720"/>"#;
    let body = p(
        &format!(
            "{}{}{}{}{}",
            r("A"),
            ptab("center", "indent", "none"),
            r("Mid"),
            ptab("right", "indent", "none"),
            r("End")
        ),
        indents,
    ) + &p(
        &format!(
            "{}{}{}",
            r("B"),
            ptab("center", "margin", "none"),
            r("Centre")
        ),
        indents,
    );
    let prims = page("", &body);
    near(mid(&prims, "Mid"), (192.0 + 672.0) / 2.0, "indent centre");
    near(end(&prims, "End"), 672.0, "indent right");
    near(
        mid(&prims, "Centre"),
        CENTER,
        "margin centre in an indented paragraph",
    );
}

#[test]
fn a_passed_position_is_found_on_the_next_line() {
    let sentence = "This sentence runs well past the middle of the line";
    let body = p(
        &format!(
            "{}{}{}",
            r(sentence),
            ptab("center", "margin", "none"),
            r("X")
        ),
        "",
    ) + &p(
        &format!("{}{}", ptab("right", "margin", "none"), r("Only right")),
        "",
    );
    let prims = page("", &body);
    assert!(
        span(&prims, "X").2 > span(&prims, sentence).2,
        "X moves to the next line"
    );
    near(mid(&prims, "X"), CENTER, "centred on the next line");
    near(
        end(&prims, "Only right"),
        RIGHT,
        "right tab at the line start",
    );
}
