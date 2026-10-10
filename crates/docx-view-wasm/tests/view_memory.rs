//! Counts what the viewer holds while it opens a dense-text ladder, with a
//! counting allocator standing in for the wasm one.

use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicUsize, Ordering::Relaxed};
use std::sync::{Mutex, PoisonError};

use docx_view_wasm::DocxViewDocument;
use serde_json::{Value, json};

struct Counting;

static LIVE: AtomicUsize = AtomicUsize::new(0);
static PEAK: AtomicUsize = AtomicUsize::new(0);
/// Bytes allocated so far, freed or not.
static ALLOCATED: AtomicUsize = AtomicUsize::new(0);
/// One test counts at a time.
static COUNTING: Mutex<()> = Mutex::new(());

fn grew(bytes: usize) {
    ALLOCATED.fetch_add(bytes, Relaxed);
    let live = LIVE.fetch_add(bytes, Relaxed) + bytes;
    PEAK.fetch_max(live, Relaxed);
}

// SAFETY: every method forwards to `System` with its own arguments.
unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let ptr = unsafe { System.alloc(layout) };
        if !ptr.is_null() {
            grew(layout.size());
        }
        ptr
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        let ptr = unsafe { System.alloc_zeroed(layout) };
        if !ptr.is_null() {
            grew(layout.size());
        }
        ptr
    }

    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        let next = unsafe { System.realloc(ptr, layout, new_size) };
        if !next.is_null() {
            LIVE.fetch_sub(layout.size(), Relaxed);
            grew(new_size);
        }
        next
    }

    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        LIVE.fetch_sub(layout.size(), Relaxed);
        unsafe { System.dealloc(ptr, layout) }
    }
}

#[global_allocator]
static ALLOCATOR: Counting = Counting;

/// `run`'s result and the peak live bytes above those live when it started.
fn peak_during<T>(run: impl FnOnce() -> T) -> (T, usize) {
    let start = LIVE.load(Relaxed);
    PEAK.store(start, Relaxed);
    let value = run();
    (value, PEAK.load(Relaxed).saturating_sub(start))
}

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const SENTENCE: &str = "The exchange committee reviews every application against the published criteria before the spring deadline.";
const W: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#;
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// `paragraphs` paragraphs of dense text; with `note`, the first carries a
/// footnote.
fn ladder(paragraphs: usize, note: bool) -> Vec<u8> {
    let text = [SENTENCE; 5].join(" ");
    let paragraph = format!(r#"<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>"#);
    let reference = if note {
        r#"<w:r><w:footnoteReference w:id="1"/></w:r>"#
    } else {
        ""
    };
    let body = format!(
        r#"<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r>{reference}</w:p>{}"#,
        paragraph.repeat(paragraphs - 1)
    );
    let mut parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>"#.to_owned(),
        ),
        (
            "_rels/.rels".to_owned(),
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>"#
            ),
        ),
        (
            "word/document.xml".to_owned(),
            format!(
                r#"<w:document {W}><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>"#
            ),
        ),
    ];
    if note {
        parts.push((
            "word/_rels/document.xml.rels".to_owned(),
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="{R}/footnotes" Target="footnotes.xml"/></Relationships>"#
            ),
        ));
        parts.push((
            "word/footnotes.xml".to_owned(),
            format!(
                r#"<w:footnotes {W}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p></w:footnote></w:footnotes>"#
            ),
        ));
    }
    let parts = parts
        .into_iter()
        .map(|(path, xml)| (path, xml.into_bytes()))
        .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// `bytes` opened, its request measuring every font with `font`.
fn open(bytes: &[u8], font: u32) -> (DocxViewDocument, String) {
    let document = DocxViewDocument::open(bytes).unwrap();
    let mut request: Value = serde_json::from_str(&document.layout_request_json()).unwrap();
    let requirements: Vec<Value> = serde_json::from_str(
        &document
            .font_requirements_json(&request.to_string())
            .unwrap(),
    )
    .unwrap();
    let chains = requirements
        .iter()
        .map(|requirement| {
            (
                requirement["key"].as_str().unwrap().to_owned(),
                json!([font]),
            )
        })
        .collect::<serde_json::Map<_, _>>();
    request["measurement"] = json!({
        "fontChains": chains,
        "defaults": { "fontSize": 11, "fontFamily": "Calibri" },
        "authoritativeShaping": true,
    });
    (document, request.to_string())
}

/// Body lines on the first page, and whether it has a note area.
fn first_page(document: &mut DocxViewDocument) -> (usize, bool) {
    let page: Value = serde_json::from_str(&document.display_pages_json(0, 1).unwrap()).unwrap();
    let page = &page["pages"][0];
    let lines = page["primitives"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|primitive| {
            primitive["blockKey"]
                .as_str()
                .is_some_and(|key| key.starts_with("body"))
        })
        .filter_map(|primitive| primitive["glyphs"][0]["y"].as_f64())
        .map(f64::to_bits)
        .collect::<std::collections::BTreeSet<_>>()
        .len();
    (
        lines,
        page["noteAreas"]
            .as_array()
            .is_some_and(|areas| !areas.is_empty()),
    )
}

/// Without notes the region pass keeps one measured arena (its note
/// fixpoint replays a copy only when a note can reserve space); with one, the
/// note still takes its space from the first page.
#[test]
fn the_region_pass_copies_the_arena_only_for_notes() {
    let _counting = COUNTING.lock().unwrap_or_else(PoisonError::into_inner);
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();

    let (mut plain, request) = open(&ladder(120, false), font);
    let before = LIVE.load(Relaxed);
    let (_, peak) = peak_during(|| plain.layout(&request).unwrap());
    let retained = LIVE.load(Relaxed) - before;
    eprintln!("region pass: peak {peak} bytes, retained {retained} bytes");
    assert!(peak * 4 <= retained * 5, "peak {peak}, retained {retained}");

    let (mut noted, request) = open(&ladder(120, true), font);
    noted.layout(&request).unwrap();
    let (plain_lines, _) = first_page(&mut plain);
    let (noted_lines, note_shown) = first_page(&mut noted);
    assert!(note_shown);
    assert!(noted_lines < plain_lines, "{noted_lines} < {plain_lines}");
}

/// Pages per `displayPagesJson` call, as `packages/docx/src/viewer.ts` asks.
const PAGE_BATCH: u32 = 8;

struct Display {
    pages: u32,
    batch_peak: usize,
    primitives: usize,
    text_primitives: usize,
    glyph_runs: usize,
    glyph_run_chars: usize,
}

/// `bytes` laid out and built a batch at a time, counting what each batch
/// allocated and what it holds.
fn display(bytes: &[u8], font: u32) -> Display {
    let (mut document, request) = open(bytes, font);
    let pages = document.layout(&request).unwrap();
    let laid_out = LIVE.load(Relaxed);
    let mut display = Display {
        pages,
        batch_peak: 0,
        primitives: 0,
        text_primitives: 0,
        glyph_runs: 0,
        glyph_run_chars: 0,
    };
    for start in (0..pages).step_by(PAGE_BATCH as usize) {
        let (json, peak) = peak_during(|| {
            document
                .display_pages_json(start, (start + PAGE_BATCH).min(pages))
                .unwrap()
        });
        display.batch_peak = display.batch_peak.max(peak);
        let batch: Value = serde_json::from_str(&json).unwrap();
        for primitive in batch["pages"]
            .as_array()
            .unwrap()
            .iter()
            .flat_map(|page| page["primitives"].as_array().unwrap())
        {
            display.primitives += 1;
            match primitive["kind"].as_str().unwrap() {
                "text" => display.text_primitives += 1,
                "glyphRun" => {
                    display.glyph_runs += 1;
                    display.glyph_run_chars += primitive["text"].as_str().unwrap().chars().count();
                }
                _ => {}
            }
        }
    }
    // A batch leaves nothing behind in the document.
    assert!(LIVE.load(Relaxed) <= laid_out + 64 * 1024);
    display
}

/// Measured text is joined glyph runs, not a primitive per character, and a
/// batch costs the same whatever the document's length: the display list is
/// never built in one piece.
#[test]
fn the_display_list_is_one_batch_of_joined_glyph_runs() {
    let _counting = COUNTING.lock().unwrap_or_else(PoisonError::into_inner);
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let small = display(&ladder(80, false), font);
    let large = display(&ladder(240, false), font);
    eprintln!(
        "pages {} / {}; batch peak {} / {} bytes; {} primitives ({} text, {} glyph runs of {} chars)",
        small.pages,
        large.pages,
        small.batch_peak,
        large.batch_peak,
        large.primitives,
        large.text_primitives,
        large.glyph_runs,
        large.glyph_run_chars,
    );
    assert!(small.pages >= PAGE_BATCH && large.pages >= 3 * small.pages - 1);
    assert_eq!(large.text_primitives, 0);
    assert!(large.glyph_run_chars >= 20 * large.glyph_runs);
    assert!(large.primitives <= 100 * large.pages as usize);
    assert!(large.batch_peak * 4 <= small.batch_peak * 5);
}

/// `sections` one-page sections sharing a header with PAGE and NUMPAGES
/// fields, so the display extras carry a measured header per section.
fn sectioned(sections: usize) -> Vec<u8> {
    let text = [SENTENCE; 5].join(" ");
    let field = |code: &str| {
        format!(
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> {code} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
        )
    };
    let properties = r#"<w:headerReference w:type="default" r:id="rIdH"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>"#;
    let paragraph = format!(r#"<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>"#);
    let section_break = format!(r#"<w:p><w:pPr><w:sectPr>{properties}</w:sectPr></w:pPr></w:p>"#);
    let body = vec![paragraph.clone(); sections].join(&section_break);
    let parts = [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>"#.to_owned(),
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
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH" Type="{R}/header" Target="header1.xml"/></Relationships>"#
            ),
        ),
        (
            "word/header1.xml",
            format!(
                r#"<w:hdr {W}><w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r>{}<w:r><w:t xml:space="preserve"> of </w:t></w:r>{}</w:p>{paragraph}</w:hdr>"#,
                field("PAGE"),
                field("NUMPAGES")
            ),
        ),
        (
            "word/document.xml",
            format!(
                r#"<w:document {W} xmlns:r="{R}"><w:body>{body}<w:sectPr>{properties}</w:sectPr></w:body></w:document>"#
            ),
        ),
    ]
    .map(|(path, xml)| (path.to_owned(), xml.into_bytes()));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// Pages laid out and bytes allocated a page while building them all.
fn allocated_per_page(bytes: &[u8], font: u32) -> (u32, usize) {
    let (mut document, request) = open(bytes, font);
    let pages = document.layout(&request).unwrap();
    let start = ALLOCATED.load(Relaxed);
    for first in (0..pages).step_by(PAGE_BATCH as usize) {
        document
            .display_pages_json(first, (first + PAGE_BATCH).min(pages))
            .unwrap();
    }
    (pages, (ALLOCATED.load(Relaxed) - start) / pages as usize)
}

/// The display extras, which carry a measured header per section, are read
/// once per document: a page costs the same with 8 sections as with 40.
#[test]
fn the_display_extras_are_read_once_per_document() {
    let _counting = COUNTING.lock().unwrap_or_else(PoisonError::into_inner);
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let (few_pages, few) = allocated_per_page(&sectioned(8), font);
    let (many_pages, many) = allocated_per_page(&sectioned(40), font);
    eprintln!("pages {few_pages} / {many_pages}; allocated a page {few} / {many} bytes");
    assert_eq!((few_pages, many_pages), (8, 40));
    assert!(many * 4 <= few * 5, "{many} > 1.25 x {few}");
}
