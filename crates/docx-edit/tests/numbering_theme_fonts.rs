//! Upstream #1921 for list markers: a numbering level's theme font wins over
//! its authored name per slot, as in Word, while the seeded model keeps the
//! authored name, and a marker font set on the paragraph keeps winning.

use docx_edit::bridge::{RenderEnv, yrs_doc_to_layout_blocks};
use docx_edit::{EditingDoc, seed_from_docx};
use docx_layout::types::LayoutBlock;
use yrs::Any;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn document(fonts: &str, major: &str) -> Vec<u8> {
    let numbering = format!(
        r#"<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:rPr>{fonts}</w:rPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#
    );
    let theme = format!(
        r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="Synthetic"><a:majorFont><a:latin typeface="{major}"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>"#
    );
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/_rels/document.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="numbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="theme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml", format!(r#"<w:document xmlns:w="{W}"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Title</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#)),
        ("word/numbering.xml", numbering),
        ("word/theme/theme1.xml", theme),
    ];
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name.to_owned(), value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

/// The render env the host sends: the parsed package's theme.
fn env(bytes: &[u8]) -> RenderEnv {
    let theme = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package
        .theme;
    serde_json::from_value(serde_json::json!({ "fontTheme": theme })).unwrap()
}

/// `(paraId, marker font)` of the body's list paragraph.
fn marker_font(doc: &EditingDoc, env: &RenderEnv) -> (String, Option<String>) {
    let blocks = yrs_doc_to_layout_blocks(doc, "body", env).unwrap();
    let LayoutBlock::Paragraph(paragraph) = &blocks[0] else {
        panic!("missing paragraph");
    };
    let attrs = paragraph.attrs.as_ref().unwrap();
    assert_eq!(attrs.list_marker.as_deref(), Some("1."));
    let para_id = doc.paragraphs("body").unwrap()[0].para_id.clone();
    (para_id, attrs.list_marker_font_family.clone())
}

#[test]
fn numbering_theme_fonts_win_per_slot_and_keep_the_model() {
    for (fonts, expected) in [
        (
            r#"<w:rFonts w:ascii="Times New Roman" w:hAnsi="Arial" w:hAnsiTheme="minorHAnsi"/>"#,
            "Times New Roman",
        ),
        (
            r#"<w:rFonts w:ascii="Times New Roman" w:asciiTheme="minorHAnsi" w:hAnsi="Arial"/>"#,
            "Calibri",
        ),
        (
            r#"<w:rFonts w:ascii="" w:hAnsi="Arial" w:hAnsiTheme="unknown"/>"#,
            "Arial",
        ),
    ] {
        let bytes = document(fonts, "Calibri Light");
        let doc = EditingDoc::new(75304);
        seed_from_docx(&doc, &bytes).unwrap();
        let original = doc.encode_state_as_update_v1();
        assert_eq!(
            marker_font(&doc, &env(&bytes)).1.as_deref(),
            Some(expected),
            "{fonts}"
        );
        assert_eq!(doc.encode_state_as_update_v1(), original);
    }
}

#[test]
fn an_edited_marker_font_wins_over_the_numbering_theme_font() {
    let bytes = document(
        r#"<w:rFonts w:ascii="Cambria" w:hAnsi="Cambria" w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>"#,
        "Calibri",
    );
    let doc = EditingDoc::new(75305);
    seed_from_docx(&doc, &bytes).unwrap();
    let env = env(&bytes);
    let (para_id, inherited) = marker_font(&doc, &env);
    assert_eq!(inherited.as_deref(), Some("Calibri"));
    assert_eq!(
        marker_font(&doc, &RenderEnv::default()).1.as_deref(),
        Some("Cambria")
    );
    doc.set_paragraph_attr(&para_id, "listMarkerFontFamily", Any::from("Courier New"))
        .unwrap();
    assert_eq!(marker_font(&doc, &env).1.as_deref(), Some("Courier New"));
}
