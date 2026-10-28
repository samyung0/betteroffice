//! A reading layout (`renderEnv.reading` plus a one-column section) reflows
//! the text into the column the host asks for.

use docx_view_wasm::DocxViewDocument;
use serde_json::{Value, json};

const W: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#;
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// A 12pt body, a 35pt title and a page break, an indented paragraph and a
/// 575px table.
fn document() -> Vec<u8> {
    let cell = |text: &str| {
        format!(
            r#"<w:tc><w:tcPr><w:tcW w:w="1725" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>"#
        )
    };
    let body = format!(
        r#"<w:p><w:r><w:rPr><w:sz w:val="70"/></w:rPr><w:t>Title</w:t></w:r></w:p><w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:pPr><w:ind w:firstLine="720"/></w:pPr><w:r><w:t>Body</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="8625" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>{}</w:tblGrid><w:tr>{}</w:tr></w:tbl><w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1802" w:bottom="1440" w:left="1802"/><w:cols w:num="2"/></w:sectPr>"#,
        r#"<w:gridCol w:w="1725"/>"#.repeat(5),
        ["One", "Two", "Three", "Four", "Five"].map(cell).concat(),
    );
    let parts = [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"#.to_owned(),
        ),
        (
            "_rels/.rels",
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>"#
            ),
        ),
        (
            "word/_rels/document.xml.rels",
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{R}/styles" Target="styles.xml"/></Relationships>"#
            ),
        ),
        (
            "word/styles.xml",
            format!(
                r#"<w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>"#
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

/// What `openDocumentViewer`'s reading option sends (packages/docx/src/viewer.ts).
fn reading_request(mut request: Value, width: f64, padding: f64) -> Value {
    for section in request["regions"]["sections"].as_array_mut().unwrap() {
        let properties = section["properties"].as_object_mut().unwrap();
        for key in ["columns", "headerReferences", "footerReferences"] {
            properties.remove(key);
        }
        for (key, value) in [
            ("pageWidth", width * 15.0),
            ("pageHeight", 2400.0 * 15.0),
            ("marginTop", 0.0),
            ("marginBottom", 0.0),
            ("marginLeft", padding * 15.0),
            ("marginRight", padding * 15.0),
            ("columnCount", 1.0),
        ] {
            properties.insert(key.to_owned(), json!(value));
        }
    }
    request["options"]["continuousFlow"] = json!(true);
    request["renderEnv"]["reading"] = json!({
        "bodySize": request["bodyFontSize"],
        "textSize": 12,
        "contentWidth": width - 2.0 * padding,
    });
    request["measurement"] = json!({
        "fontChains": {},
        "defaults": { "fontSize": 12, "fontFamily": "Calibri" },
        "authoritativeShaping": false,
    });
    request
}

fn font_px(run: &Value) -> f64 {
    let font = run["font"].as_str().unwrap();
    font.split_whitespace()
        .find_map(|part| part.strip_suffix("px")?.parse().ok())
        .unwrap()
}

#[test]
fn a_reading_layout_reflows_into_the_column() {
    let mut document = DocxViewDocument::open(&document()).unwrap();
    let request: Value = serde_json::from_str(&document.layout_request_json()).unwrap();
    assert_eq!(request["bodyFontSize"], json!(12.0));
    let request = reading_request(request, 360.0, 16.0);
    let pages = document.layout(&request.to_string()).unwrap();
    let list: Value =
        serde_json::from_str(&document.display_pages_json(0, pages).unwrap()).unwrap();
    assert_eq!(list["pages"][0]["width"], json!(360));
    // The one page ends where its text does, well short of 2400px.
    assert_eq!(pages, 1);
    assert!(list["pages"][0]["height"].as_f64().unwrap() < 600.0);
    // The page break starts no page: everything flows on the first.
    let runs = list["pages"][0]["primitives"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|primitive| primitive["kind"] == "text")
        .collect::<Vec<_>>();
    let run = |text: &str| *runs.iter().find(|run| run["text"] == text).unwrap();

    // 35pt eases to 12pt × √(35/12); the body reads at 16px.
    assert!((font_px(run("Title")) - 16.0 * (35.0_f64 / 12.0).sqrt()).abs() < 0.1);
    assert_eq!(font_px(run("Body")), 16.0);
    // The 48px first-line indent is halved, inside the 16px padding.
    assert_eq!(run("Body")["x"].as_f64(), Some(40.0));
    // The table's 575px grid fits the 328px column, in one column of text.
    for cell in ["One", "Two", "Three", "Four", "Five"] {
        let run = run(cell);
        let right = run["x"].as_f64().unwrap() + run["width"].as_f64().unwrap();
        assert!(right <= 344.5, "{cell} ends at {right}");
    }
}
