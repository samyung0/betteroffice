//! Counts what the viewer holds while it opens a dense-text ladder, with a
//! counting allocator standing in for the wasm one.

use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::Mutex;
use std::sync::atomic::{AtomicUsize, Ordering::Relaxed};

use docx_view_wasm::DocxViewDocument;
use serde_json::{Value, json};

struct Counting;

static LIVE: AtomicUsize = AtomicUsize::new(0);
static PEAK: AtomicUsize = AtomicUsize::new(0);
/// One test counts at a time.
static COUNTING: Mutex<()> = Mutex::new(());

fn grew(bytes: usize) {
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
fn first_page(document: &DocxViewDocument) -> (usize, bool) {
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
    let _counting = COUNTING.lock().unwrap();
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
    let (plain_lines, _) = first_page(&plain);
    let (noted_lines, note_shown) = first_page(&noted);
    assert!(note_shown);
    assert!(noted_lines < plain_lines, "{noted_lines} < {plain_lines}");
}
