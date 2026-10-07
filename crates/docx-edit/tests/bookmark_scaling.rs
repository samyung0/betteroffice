//! Reading a story's bookmarks must stay linear in its length.

use std::time::{Duration, Instant};

use docx_edit::{EditingDoc, seed_from_docx};

const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;

const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;

/// `count` paragraphs, each holding one bookmark around its text.
fn bookmarked(count: usize) -> EditingDoc {
    let body: String = (0..count)
        .map(|index| {
            format!(
                r#"<w:p><w:bookmarkStart w:id="{index}" w:name="b{index}"/><w:r><w:t>Paragraph {index} of the handbook.</w:t></w:r><w:bookmarkEnd w:id="{index}"/></w:p>"#
            )
        })
        .collect();
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}</w:body></w:document>"#
    );
    let bytes = ooxml_opc::rezip_parts(&[
        ("[Content_Types].xml".to_owned(), CONTENT_TYPES.into()),
        ("_rels/.rels".to_owned(), ROOT_RELS.into()),
        ("word/document.xml".to_owned(), document.into_bytes()),
    ])
    .expect("synthetic package zips");
    let doc = EditingDoc::new(7);
    seed_from_docx(&doc, &bytes).expect("seed succeeds");
    doc
}

fn read_time(doc: &EditingDoc) -> Duration {
    (0..3)
        .map(|_| {
            let start = Instant::now();
            doc.story_segments("body").expect("body reads");
            doc.all_story_segments();
            start.elapsed()
        })
        .min()
        .unwrap()
}

#[test]
fn bookmark_reads_stay_linear_in_story_length() {
    let mut per_paragraph = Vec::new();
    for count in [300usize, 2400] {
        let doc = bookmarked(count);
        let markers = doc
            .story_segments("body")
            .unwrap()
            .iter()
            .filter(|segment| format!("{:?}", segment.content).contains("bookmark"))
            .count();
        assert!(markers > 0, "the story's paragraphs carry their bookmarks");
        let millis = read_time(&doc).as_secs_f64() * 1000.0 / count as f64;
        println!("{count:>5} bookmarked paragraphs   {millis:>7.4} ms/paragraph");
        per_paragraph.push(millis);
    }
    let ratio = per_paragraph[1] / per_paragraph[0];
    assert!(
        ratio < 3.0,
        "per-paragraph bookmark read cost grew {ratio:.1}x from 300 to 2400 paragraphs; anchors resolve per walk again"
    );
}

/// Lowering a story resolves its comment anchors in one walk as well: one
/// comment over every paragraph keeps the per-paragraph cost flat.
#[test]
fn comment_ranges_lower_linearly_in_story_length() {
    let mut per_paragraph = Vec::new();
    for count in [300usize, 2400] {
        let doc = bookmarked(count);
        let mut at = 0_u32;
        let mut comments = Vec::new();
        for (index, paragraph) in doc.paragraphs("body").unwrap().iter().enumerate() {
            let len = paragraph.text.encode_utf16().count() as u32;
            comments.push(docx_edit::RawOp::SetComment {
                id: index.to_string(),
                ranges: vec![(at, at + len)],
                author: "Ada".into(),
                date: "2026-10-07T00:00:00Z".into(),
                body: yrs::Any::Null,
            });
            at += len + 1;
        }
        doc.apply_raw_ops("body", comments, &docx_edit::EditCtx::local("Ada", ""))
            .unwrap();
        let env = docx_edit::bridge::RenderEnv::default();
        let elapsed = (0..3)
            .map(|_| {
                let start = Instant::now();
                let blocks = docx_edit::bridge::yrs_doc_to_layout_blocks(&doc, "body", &env)
                    .expect("body lowers");
                assert!(format!("{:?}", blocks[count / 2]).contains("comment_ids: Some"));
                start.elapsed()
            })
            .min()
            .unwrap();
        let millis = elapsed.as_secs_f64() * 1000.0 / count as f64;
        println!("{count:>5} commented paragraphs   {millis:>7.4} ms/paragraph");
        per_paragraph.push(millis);
    }
    let ratio = per_paragraph[1] / per_paragraph[0];
    assert!(
        ratio < 3.0,
        "per-paragraph lowering cost grew {ratio:.1}x from 300 to 2400 commented paragraphs; comment anchors resolve per walk again"
    );
}
