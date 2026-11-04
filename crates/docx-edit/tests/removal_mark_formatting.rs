//! Removing the end of a formatted paragraph leaves the formatting of the paragraph marks and
//! the table after it as they were (upstream #1433; Capy's engine already keeps it).

use std::collections::BTreeMap;

use docx_edit::*;
use yrs::Any;

fn ctx() -> EditCtx {
    EditCtx::local("Ada", "2026-10-28T00:00:00Z")
}

fn seeded(body: &str) -> EditingDoc {
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml", format!(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>{body}<w:sectPr/></w:body></w:document>"#)),
    ];
    let bytes = ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, xml)| (name.to_owned(), xml.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap();
    let doc = EditingDoc::new(7001);
    seed_from_docx(&doc, &bytes).unwrap();
    doc
}

fn formatted(text: &str) -> String {
    format!(
        r#"<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>"#
    )
}

/// The formatting of every unit that is not text: paragraph marks and block embeds.
fn embed_formatting(doc: &EditingDoc) -> Vec<BTreeMap<String, Any>> {
    doc.story_segments("body")
        .unwrap()
        .into_iter()
        .filter(|segment| !matches!(segment.content, SegmentContent::Text(_)))
        .map(|segment| segment.attributes)
        .collect()
}

#[test]
fn replacing_the_end_of_a_formatted_paragraph_keeps_the_marks_and_table_after_it() {
    let table = r#"<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p w14:paraId="00000003"><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#;
    let doc = seeded(&format!(
        r#"<w:p w14:paraId="00000001">{}</w:p><w:p w14:paraId="00000002"/>{table}<w:p w14:paraId="00000004">{}</w:p>"#,
        formatted("Dose of [compound]"),
        formatted("Tail")
    ));
    let before = embed_formatting(&doc);
    doc.replace_range(&ctx(), StoryRange::new("body", 8, 18), "Agent X")
        .unwrap();
    assert_eq!(embed_formatting(&doc), before);
}

#[test]
fn deleting_the_end_of_a_formatted_paragraph_keeps_the_marks_after_it() {
    let doc = seeded(&format!(
        r#"<w:p w14:paraId="00000001">{}</w:p><w:p w14:paraId="00000002"/><w:p w14:paraId="00000003">{}</w:p>"#,
        formatted("One"),
        formatted("Two")
    ));
    let before = embed_formatting(&doc);
    doc.delete_range(&ctx(), StoryRange::new("body", 2, 3))
        .unwrap();
    assert_eq!(embed_formatting(&doc), before);
}
