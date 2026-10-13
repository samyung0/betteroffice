//! A section's page borders (`w:pgBorders`) reach the viewer's display list.

use docx_view_wasm::DocxViewDocument;
use serde_json::{Value, json};

const W: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#;
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Two sections: the first with a page border measured from the text, on
/// every page; the second with a border measured from the page edge, on its
/// first page only, drawn behind the text.
fn bordered() -> Vec<u8> {
    let paragraph = |text: &str| format!(r#"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"#);
    let first_section = r#"<w:p><w:pPr><w:sectPr><w:pgBorders><w:top w:val="single" w:sz="8" w:space="24" w:color="FF0000"/><w:left w:val="single" w:sz="8" w:space="4" w:color="FF0000"/><w:bottom w:val="single" w:sz="8" w:space="24" w:color="FF0000"/><w:right w:val="single" w:sz="8" w:space="4" w:color="FF0000"/></w:pgBorders><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:pPr></w:p>"#;
    let body = format!(
        r#"{}{first_section}{}<w:sectPr><w:pgBorders w:offsetFrom="page" w:display="firstPage" w:zOrder="back"><w:top w:val="double" w:sz="4" w:space="24" w:color="0000FF"/><w:bottom w:val="double" w:sz="4" w:space="24" w:color="0000FF"/></w:pgBorders><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>"#,
        paragraph("First section."),
        paragraph("Second section."),
    );
    let parts = [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_owned(),
        ),
        (
            "_rels/.rels",
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>"#
            ),
        ),
        (
            "word/document.xml",
            format!(r#"<w:document {W}><w:body>{body}</w:body></w:document>"#),
        ),
    ];
    let parts = parts
        .into_iter()
        .map(|(path, xml)| (path.to_owned(), xml.into_bytes()))
        .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn section_page_borders_reach_the_viewer_display_list() {
    let mut document = DocxViewDocument::open(&bordered()).unwrap();
    let mut request: Value = serde_json::from_str(&document.layout_request_json()).unwrap();
    request["measurement"] = json!({
        "fontChains": {},
        "defaults": { "fontSize": 11, "fontFamily": "Calibri" },
        "authoritativeShaping": false,
    });
    assert_eq!(document.layout(&request.to_string()).unwrap(), 2);
    let list: Value = serde_json::from_str(&document.display_pages_json(0, 2).unwrap()).unwrap();
    let borders = |page: usize| list["pages"][page]["pageBorders"].clone();

    // From the text: 1 inch margins less the border spaces (24 pt and 4 pt).
    let first = borders(0);
    assert_eq!(first.as_array().map(Vec::len), Some(1), "{first}");
    let first = &first[0];
    assert_eq!(first["x"], json!(90.667));
    assert_eq!(first["y"], json!(64));
    assert_eq!(first["zOrder"], "front");
    for side in ["top", "right", "bottom", "left"] {
        assert_eq!(first[side]["color"], "#FF0000", "{side}");
        assert_eq!(first[side]["style"], "solid", "{side}");
    }

    // From the page edge, top and bottom only, behind the text.
    let second = borders(1);
    assert_eq!(second.as_array().map(Vec::len), Some(1), "{second}");
    let second = &second[0];
    assert_eq!(second["y"], json!(32));
    assert_eq!(second["zOrder"], "back");
    assert_eq!(second["top"]["style"], "double");
    assert_eq!(second["left"], Value::Null);
}
