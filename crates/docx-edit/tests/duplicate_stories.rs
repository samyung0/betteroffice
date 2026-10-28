//! A package naming one story twice (two footnotes with the same id) fails
//! the seed, as it did when each story was created on its own, instead of
//! seeding both notes' paragraphs into one story.

use docx_edit::{EditingDoc, seed_from_docx};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const OFFICE: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml";

#[test]
fn a_story_named_twice_fails_the_seed() {
    let note = |text: &str| {
        format!(r#"<w:footnote w:id="1"><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:footnote>"#)
    };
    let package = ooxml_opc::rezip_parts(&[
        (
            "[Content_Types].xml".to_owned(),
            format!(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{OFFICE}.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="{OFFICE}.footnotes+xml"/></Types>"#
            )
            .into_bytes(),
        ),
        (
            "_rels/.rels".to_owned(),
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>"#
            )
            .into_bytes(),
        ),
        (
            "word/_rels/document.xml.rels".to_owned(),
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{R}/footnotes" Target="footnotes.xml"/></Relationships>"#
            )
            .into_bytes(),
        ),
        (
            "word/document.xml".to_owned(),
            format!(
                r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>Body</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p><w:sectPr/></w:body></w:document>"#
            )
            .into_bytes(),
        ),
        (
            "word/footnotes.xml".to_owned(),
            format!(r#"<w:footnotes xmlns:w="{W}">{}{}</w:footnotes>"#, note("one"), note("two"))
                .into_bytes(),
        ),
    ])
    .unwrap();
    let error = seed_from_docx(&EditingDoc::new(7), &package).unwrap_err();
    assert!(error.contains("\"fn:1\" already exists"), "{error}");
}
