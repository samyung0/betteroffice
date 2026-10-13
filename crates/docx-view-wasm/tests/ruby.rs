//! A phonetic guide (`w:ruby`) shows its base text in its place.

use docx_view_wasm::DocxViewDocument;
use serde_json::{Value, json};

const W: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#;
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

#[test]
fn a_phonetic_guide_shows_its_base_text() {
    let body = r#"<w:p><w:r><w:t xml:space="preserve">読み: </w:t></w:r><w:r><w:ruby><w:rubyPr><w:hps w:val="10"/><w:hpsRaise w:val="18"/><w:hpsBaseText w:val="21"/><w:lid w:val="ja-JP"/></w:rubyPr><w:rt><w:r><w:t>かんじ</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>漢字</w:t></w:r></w:rubyBase></w:ruby></w:r><w:r><w:t xml:space="preserve"> です</w:t></w:r></w:p>"#;
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
            format!(
                r#"<w:document {W}><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>"#
            ),
        ),
    ]
    .into_iter()
    .map(|(path, xml)| (path.to_owned(), xml.into_bytes()))
    .collect::<Vec<_>>();
    let mut document = DocxViewDocument::open(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap();
    let mut request: Value = serde_json::from_str(&document.layout_request_json()).unwrap();
    request["measurement"] = json!({
        "fontChains": {},
        "defaults": { "fontSize": 11, "fontFamily": "Calibri" },
        "authoritativeShaping": false,
    });
    document.layout(&request.to_string()).unwrap();
    let list: Value = serde_json::from_str(&document.display_pages_json(0, 1).unwrap()).unwrap();
    let text: String = list["pages"][0]["primitives"]
        .as_array()
        .unwrap()
        .iter()
        .filter_map(|primitive| primitive["text"].as_str())
        .collect();
    assert_eq!(text, "読み: 漢字 です");
}
