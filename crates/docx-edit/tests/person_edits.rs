//! A person's edits on a 10,000-paragraph document as the editor makes them:
//! each edit through the paragraph-keyed caret, then the resident layout.
//! Prints the median edit and layout time of typing, a paste, a selection
//! delete and Backspace.
//!
//! Run: `cargo test -p betteroffice-docx-edit --release --test person_edits -- --ignored --nocapture`

use std::time::{Duration, Instant};

use docx_edit::{EditCtx, EngineSession, FormatPolicy, Loc, Position, StoryRange, seed_from_docx};

const PARAGRAPHS: usize = 10_000;
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"#;

const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;

const DOC_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#;

const STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"#;

/// Paragraphs of a plain run and a coloured one, so the story holds format
/// markers as real documents do.
fn build_docx() -> Vec<u8> {
    let mut body = String::new();
    for index in 0..PARAGRAPHS {
        body.push_str(&format!(
            r#"<w:p><w:r><w:t xml:space="preserve">Paragraph {index}: the quick brown fox jumps over </w:t></w:r><w:r><w:rPr><w:color w:val="C00000"/></w:rPr><w:t>the lazy dog.</w:t></w:r></w:p>"#
        ));
    }
    body.push_str(r#"<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>"#);
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}</w:body></w:document>"#
    );
    ooxml_opc::rezip_parts(&[
        ("[Content_Types].xml".to_owned(), CONTENT_TYPES.into()),
        ("_rels/.rels".to_owned(), ROOT_RELS.into()),
        ("word/_rels/document.xml.rels".to_owned(), DOC_RELS.into()),
        ("word/styles.xml".to_owned(), STYLES.into()),
        ("word/document.xml".to_owned(), document.into_bytes()),
    ])
    .expect("synthetic package zips")
}

fn millis(spent: Duration) -> f64 {
    spent.as_secs_f64() * 1000.0
}

fn median(mut times: Vec<f64>) -> f64 {
    times.sort_by(f64::total_cmp);
    times[times.len() / 2]
}

#[test]
#[ignore = "bench; run explicitly with --ignored --nocapture"]
fn person_edits_on_ten_thousand_paragraphs() {
    let engine = EngineSession::new(7);
    let doc = engine.doc();
    seed_from_docx(doc, &build_docx()).unwrap();
    let ids: Vec<String> = doc
        .paragraphs("body")
        .unwrap()
        .into_iter()
        .map(|para| para.para_id.to_string())
        .collect();
    assert_eq!(ids.len(), PARAGRAPHS);

    docx_layout::clear_measure_fonts();
    let font_id = docx_layout::register_measure_font(FONT).unwrap();
    let request = serde_json::json!({
        "bodyStory": "body",
        "regions": {"sections": [{"sectionId": "main", "properties": {}}]},
        "measurement": {
            "fontChains": {"calibri|0|0": [font_id]},
            "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
            "authoritativeShaping": true
        },
        "renderEnv": {}
    })
    .to_string();
    let extras = serde_json::json!({"fontChains": {"calibri|0|0": [font_id]}}).to_string();
    let started = Instant::now();
    engine.layout_document_with_regions_json(&request).unwrap();
    engine.build_display_list_frame(&extras, 0).unwrap();
    let cold = millis(started.elapsed());

    let ctx = EditCtx::local("", "");
    // Each sample is (edit, layout) in milliseconds. The editor finds its
    // paragraph-keyed caret in the segment index; `at` reads the story
    // instead, so it stays out of the time.
    let at = |para: usize, offset: u32| {
        doc.locate(&Loc::new("body", ids[para].as_str(), offset))
            .unwrap()
    };
    let sample = |edit: &mut dyn FnMut()| {
        let started = Instant::now();
        edit();
        let edited = millis(started.elapsed());
        let started = Instant::now();
        let epoch = engine.stats().frame_epoch;
        engine.apply_and_layout("body", epoch).unwrap();
        (edited, millis(started.elapsed()))
    };

    let typing: Vec<(f64, f64)> = (0..100)
        .map(|key| {
            let caret = at(5_000, 20 + key);
            sample(&mut || {
                doc.insert_text(&ctx, caret.clone(), "x", FormatPolicy::Inherit)
                    .unwrap();
            })
        })
        .collect();

    // What locating a caret from the story's start costs (`loc_at`), as an
    // edit's receipt did for its range.
    let walks: Vec<f64> = (0..20)
        .map(|_| {
            let caret = at(5_000, 20);
            let started = Instant::now();
            doc.loc_at(&caret).unwrap();
            millis(started.elapsed())
        })
        .collect();

    // A paste of 20 lines: each line's text, then a split after it.
    let pastes: Vec<(f64, f64)> = (0..3)
        .map(|paste| {
            let start = at(2_000 + paste * 1_000, 10);
            sample(&mut || {
                let mut caret = start.clone();
                for line in 0..20 {
                    let text = format!("Pasted line {line}, as the clipboard held it.");
                    doc.insert_text(&ctx, caret.clone(), &text, FormatPolicy::Inherit)
                        .unwrap();
                    caret = Position::new("body", caret.index + text.encode_utf16().count() as u32);
                    if line < 19 {
                        doc.split_paragraph(&ctx, caret.clone()).unwrap();
                        caret = Position::new("body", caret.index + 1);
                    }
                }
            })
        })
        .collect();

    // A selection over ten paragraph marks, deleted.
    let deletes: Vec<(f64, f64)> = (0..3)
        .map(|delete| {
            let para = 7_000 + delete * 100;
            let (start, end) = (at(para, 10), at(para + 10, 10));
            sample(&mut || {
                doc.delete_range(&ctx, StoryRange::new("body", start.index, end.index))
                    .unwrap();
            })
        })
        .collect();

    let backspaces: Vec<(f64, f64)> = (0..50)
        .map(|key| {
            let caret = at(8_000, 60 - key);
            sample(&mut || {
                doc.delete_range(&ctx, StoryRange::new("body", caret.index - 1, caret.index))
                    .unwrap();
            })
        })
        .collect();

    let stats = engine.stats();
    println!("{PARAGRAPHS} paragraphs, cold layout {cold:.0} ms; medians, edit + layout:");
    for (name, samples) in [
        ("typing, a key (100)", typing),
        ("a paste of 20 lines (3)", pastes),
        ("a delete over 10 marks (3)", deletes),
        ("Backspace, a key (50)", backspaces),
    ] {
        let (edit, layout): (Vec<f64>, Vec<f64>) = samples.into_iter().unzip();
        println!(
            "  {name}: {:.2} ms + {:.2} ms",
            median(edit),
            median(layout)
        );
    }
    println!("  loc_at at paragraph 5,000: {:.2} ms", median(walks));
    println!(
        "  fast path hits {}, last fallback {:?}",
        stats.region_fast_path_hits, stats.region_fast_path_fallback,
    );
}
