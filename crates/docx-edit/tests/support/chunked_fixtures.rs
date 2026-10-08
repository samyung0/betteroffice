//! Small DOCX documents for the override-layout checks (shared by the
//! docx-edit and office-service tests through `#[path]`).

const W: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#;

fn docx(body: &str, extra: &[(&str, String)]) -> Vec<u8> {
    let mut parts = vec![(
        "word/document.xml".to_owned(),
        format!("<w:document {W}><w:body>{body}</w:body></w:document>").into_bytes(),
    )];
    for (name, xml) in extra {
        parts.push(((*name).to_owned(), xml.clone().into_bytes()));
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// Small documents for the hard cases: a field across paragraphs, comments
/// across paragraphs, tracked changes, bookmarks, breaks, sections, lists,
/// a table.
pub fn xml_fixtures() -> Vec<(String, Vec<u8>)> {
    let comments = format!(
        r#"<w:comments {W}><w:comment w:id="1" w:author="Ada" w:date="2026-01-01T00:00:00Z"><w:p><w:r><w:t>Note one</w:t></w:r></w:p></w:comment><w:comment w:id="2" w:author="Bo" w:date="2026-01-02T00:00:00Z"><w:p><w:r><w:t>Note two</w:t></w:r></w:p></w:comment></w:comments>"#
    );
    let numbering = format!(
        r#"<w:numbering {W}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#
    );
    vec![
        (
            "xml-mixed".to_owned(),
            docx(
                concat!(
                    r#"<w:p w14:paraId="10000001"><w:r><w:t xml:space="preserve">Alpha </w:t></w:r><w:commentRangeStart w:id="1"/><w:r><w:rPr><w:b/></w:rPr><w:t>bold start</w:t></w:r></w:p>"#,
                    r#"<w:p w14:paraId="10000002"><w:bookmarkStart w:id="5" w:name="mark"/><w:r><w:t xml:space="preserve">Beta with a </w:t></w:r><w:ins w:id="7" w:author="Ada" w:date="2026-01-01T00:00:00Z"><w:r><w:t>tracked insert</w:t></w:r></w:ins><w:del w:id="8" w:author="Bo" w:date="2026-01-01T00:00:00Z"><w:r><w:delText> gone</w:delText></w:r></w:del><w:bookmarkEnd w:id="5"/></w:p>"#,
                    r#"<w:p w14:paraId="10000003"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>first item</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>"#,
                    r#"<w:p w14:paraId="10000004"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>second item</w:t></w:r></w:p>"#,
                    r#"<w:p w14:paraId="10000005"><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \o "1-3" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>toc line one</w:t></w:r></w:p>"#,
                    r#"<w:p w14:paraId="10000006"><w:r><w:t>toc line two</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
                    r#"<w:p w14:paraId="10000007"><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:br w:type="page"/><w:t>after break</w:t></w:r></w:p>"#,
                    r#"<w:p w14:paraId="10000008"><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr><w:commentRangeStart w:id="2"/><w:r><w:t>section one ends</w:t></w:r></w:p>"#,
                    r#"<w:tbl><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p w14:paraId="10000009"><w:r><w:t>cell a</w:t></w:r></w:p></w:tc><w:tc><w:p w14:paraId="1000000A"><w:r><w:t>cell b</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
                    r#"<w:p w14:paraId="1000000B"><w:r><w:t>after table</w:t></w:r><w:commentRangeEnd w:id="2"/><w:r><w:commentReference w:id="2"/></w:r></w:p>"#,
                    r#"<w:p w14:paraId="1000000C"><w:r><w:br w:type="page"/></w:r><w:r><w:t>leading break</w:t></w:r></w:p>"#,
                    r#"<w:p w14:paraId="1000000D"><w:r><w:t>last</w:t></w:r></w:p>"#,
                    r#"<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>"#
                ),
                &[
                    ("word/comments.xml", comments.clone()),
                    ("word/numbering.xml", numbering.clone()),
                ],
            ),
        ),
        (
            "xml-table-last".to_owned(),
            docx(
                r#"<w:p><w:r><w:t>before</w:t></w:r></w:p><w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>only cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
                &[],
            ),
        ),
        (
            "xml-collapsed-bookmark".to_owned(),
            docx(
                r#"<w:p w14:paraId="20000001"><w:r><w:t>one</w:t></w:r></w:p><w:p w14:paraId="20000002"><w:bookmarkStart w:id="1" w:name="here"/><w:bookmarkEnd w:id="1"/><w:r><w:t>two</w:t></w:r></w:p><w:p w14:paraId="20000002"><w:r><w:t>duplicate paraId</w:t></w:r></w:p>"#,
                &[],
            ),
        ),
    ]
}

