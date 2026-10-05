//! The selection reports a file's highlight as the toolbar's highlight picker
//! reads it: a `w:highlight` by its Word name, a clear shading fill as a hex.

use docx_edit::{EngineSession, StoryRange, seed_from_docx};

fn document(body: &str) -> Vec<u8> {
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml", format!(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}</w:body></w:document>"#)),
    ];
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name.to_owned(), value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

#[test]
fn a_seeded_highlight_is_reported_by_name_or_hex() {
    let body = r#"<w:p><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t xml:space="preserve">one </w:t></w:r><w:r><w:rPr><w:shd w:val="clear" w:color="auto" w:fill="F4CCCC"/></w:rPr><w:t xml:space="preserve">two </w:t></w:r><w:r><w:t>six</w:t></w:r></w:p>"#;
    let engine = EngineSession::new(71734);
    seed_from_docx(engine.doc(), &document(body)).unwrap();
    let highlight = |start, end| {
        engine
            .doc()
            .selection_context(&StoryRange::new("body", start, end))
            .unwrap()
            .highlight
    };
    assert_eq!(highlight(0, 3).as_deref(), Some("yellow"));
    assert_eq!(highlight(4, 7).as_deref(), Some("#F4CCCC"));
    assert_eq!(highlight(0, 7), None);
    assert_eq!(highlight(8, 11), None);
}
