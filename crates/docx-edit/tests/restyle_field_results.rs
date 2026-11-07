//! Upstream #1932 (restyles update the live formatting of field results), on
//! Capy's engine: a field result reads its size and weight the way the text
//! beside it does, before and after a paragraph style change. Capy's restyle
//! leaves run formatting to the host (Q658), so both keep the seeded style's.

use std::collections::BTreeMap;

use docx_edit::bridge::{RenderEnv, yrs_doc_to_layout_blocks};
use docx_edit::{EditCtx, EditingDoc, ParaSelector, seed_from_docx};
use serde_json::Value;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn document(body: &str) -> Vec<u8> {
    let styles = format!(
        r#"<w:styles xmlns:w="{W}"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Big"><w:name w:val="Big"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Small"><w:name w:val="Small"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style></w:styles>"#
    );
    let parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"#.to_owned(),
        ),
        (
            "_rels/.rels".to_owned(),
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned(),
        ),
        (
            "word/_rels/document.xml.rels".to_owned(),
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#.to_owned(),
        ),
        (
            "word/document.xml".to_owned(),
            format!(r#"<w:document xmlns:w="{W}"><w:body>{body}<w:sectPr/></w:body></w:document>"#),
        ),
        ("word/styles.xml".to_owned(), styles),
    ];
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name, value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

/// `(text sizes, field sizes, text bold, field bold)` of the body's runs.
fn sizes(doc: &EditingDoc) -> (Vec<Value>, Vec<Value>, Vec<Value>, Vec<Value>) {
    let blocks =
        serde_json::to_value(yrs_doc_to_layout_blocks(doc, "body", &RenderEnv::default()).unwrap())
            .unwrap();
    let runs = blocks[0]["runs"].as_array().unwrap();
    let pick = |kind: &str, key: &str| {
        runs.iter()
            .filter(|run| run["kind"] == kind)
            .map(|run| run[key].clone())
            .collect::<Vec<_>>()
    };
    (
        pick("text", "fontSize"),
        pick("field", "fontSize"),
        pick("text", "bold"),
        pick("field", "bold"),
    )
}

#[test]
fn field_results_follow_a_restyle_as_the_text_beside_them() {
    for field in [
        r#"<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>"#,
        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
    ] {
        let bytes = document(&format!(
            r#"<w:p><w:pPr><w:pStyle w:val="Big"/></w:pPr><w:r><w:t xml:space="preserve">Page </w:t></w:r>{field}</w:p>"#
        ));
        let doc = EditingDoc::new(7);
        seed_from_docx(&doc, &bytes).unwrap();
        let (text, fields, text_bold, field_bold) = sizes(&doc);
        assert!(!fields.is_empty(), "{field}");
        assert!(
            fields.iter().all(|size| Some(size) == text.first()),
            "{field}: {text:?} {fields:?}"
        );
        assert!(
            field_bold
                .iter()
                .all(|bold| Some(bold) == text_bold.first()),
            "{field}"
        );

        let para = doc.paragraphs("body").unwrap()[0].para_id.clone();
        let previous = BTreeMap::from([("Big".to_owned(), BTreeMap::new())]);
        doc.apply_paragraph_style(
            &EditCtx::local("", "2026-11-07T00:00:00Z"),
            &ParaSelector::One(para),
            "Small",
            &BTreeMap::new(),
            &previous,
        )
        .unwrap();
        let (text, fields, text_bold, field_bold) = sizes(&doc);
        assert!(
            fields.iter().all(|size| Some(size) == text.first()),
            "{field}: {text:?} {fields:?}"
        );
        assert!(
            field_bold
                .iter()
                .all(|bold| Some(bold) == text_bold.first()),
            "{field}"
        );
    }
}
