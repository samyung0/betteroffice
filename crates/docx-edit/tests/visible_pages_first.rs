//! An editor's first frame builds the pages in view and leaves the rest
//! unbuilt; later frames build them a few at a time (Q815 C). Once every page
//! is built, the list is a whole build's, edits made in between included.

use docx_edit::{EditCtx, EditingDoc, EngineSession, FormatPolicy, Position};
use docx_layout::display_list::DisplayList;
use serde_json::json;

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const PARAGRAPHS: usize = 150;

struct Fixture {
    state: Vec<u8>,
    request: String,
    extras: String,
    starts: Vec<u32>,
}

impl Fixture {
    /// One-line paragraphs over many small pages under a one-paragraph header.
    fn new() -> Self {
        docx_layout::clear_measure_fonts();
        let font = docx_layout::register_measure_font(FONT).unwrap();
        let doc = EditingDoc::new(1);
        let ctx = EditCtx::local("", "");
        doc.create_story("hf:rIdH", "Running head", "Normal", "left")
            .unwrap();
        doc.create_story("body", "", "Normal", "left").unwrap();
        let mut starts = Vec::new();
        let mut cursor = 0;
        for index in 0..PARAGRAPHS {
            starts.push(cursor);
            let text = format!("Paragraph {index} of the body");
            doc.insert_text(
                &ctx,
                Position::new("body", cursor),
                &text,
                FormatPolicy::Inherit,
            )
            .unwrap();
            cursor += text.len() as u32;
            if index + 1 < PARAGRAPHS {
                doc.split_paragraph(&ctx, Position::new("body", cursor))
                    .unwrap();
                cursor += 1;
            }
        }
        let request = json!({
            "bodyStory": "body",
            "regions": {"sections": [{
                "sectionId": "main",
                "pageSize": {"w": 400, "h": 300},
                "margins": {"top": 40, "right": 30, "bottom": 30, "left": 30, "header": 10, "footer": 10},
                "headerFooterRefs": {"headerDefault": "rIdH"}
            }]},
            "measurement": {
                "fontChains": {"calibri|0|0": [font]},
                "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
                "authoritativeShaping": true
            },
            "renderEnv": {}
        })
        .to_string();
        let extras = json!({"fontChains": {"calibri|0|0": [font]}}).to_string();
        Self {
            state: doc.encode_state_as_update_v1(),
            request,
            extras,
            starts,
        }
    }

    /// A worker replica laid out as a bootstrap does, holding a first frame of
    /// `first` pages.
    fn worker(&self, first: usize) -> EngineSession {
        let worker = EngineSession::new(9);
        worker.doc().load_state_v1(&self.state).unwrap();
        worker
            .layout_document_with_regions_retained_json(&self.request)
            .unwrap();
        worker
            .build_display_list_frame_first(&self.extras, 0, first)
            .unwrap();
        worker
    }

    /// A fresh engine's whole build of `worker`'s document.
    fn whole(&self, worker: &EngineSession) -> DisplayList {
        let fresh = EngineSession::new(10);
        fresh
            .doc()
            .load_state_v1(&worker.doc().encode_state_as_update_v1())
            .unwrap();
        fresh
            .relayout_frame(&self.request, &self.extras, 0)
            .unwrap();
        fresh.with_display_list(Clone::clone).unwrap()
    }
}

fn list(engine: &EngineSession) -> DisplayList {
    engine.with_display_list(Clone::clone).unwrap()
}

fn unbuilt(engine: &EngineSession) -> Vec<usize> {
    list(engine)
        .pages
        .iter()
        .enumerate()
        .filter_map(|(index, page)| page.unbuilt.then_some(index))
        .collect()
}

/// Builds the unbuilt pages `batch` at a time, from the last page back.
fn build_rest(engine: &EngineSession, batch: usize) {
    loop {
        let mut pages = unbuilt(engine);
        if pages.is_empty() {
            return;
        }
        pages.reverse();
        pages.truncate(batch);
        let epoch = engine.stats().frame_epoch;
        engine
            .build_unbuilt_display_pages_frame(&pages, epoch)
            .unwrap();
    }
}

/// `(full, upserts)` from a FrameDelta header and table.
fn frame_ops(frame: &[u8]) -> (bool, usize) {
    let u32_at = |offset: usize| u32::from_le_bytes(frame[offset..offset + 4].try_into().unwrap());
    let count = u32_at(52) as usize;
    let upserts = (0..count).filter(|op| frame[80 + op * 48] == 1).count();
    (u32_at(12) & 1 == 1, upserts)
}

fn type_at(worker: &EngineSession, index: u32, text: &str) {
    let peer = EditingDoc::new(2);
    peer.load_state_v1(&worker.doc().encode_state_as_update_v1())
        .unwrap();
    peer.insert_text(
        &EditCtx::local("", ""),
        Position::new("body", index),
        text,
        FormatPolicy::Inherit,
    )
    .unwrap();
    let update = peer
        .encode_diff_v1(&worker.doc().encode_state_vector_v1())
        .unwrap();
    worker.doc().apply_update_v1(&update).unwrap();
}

#[test]
fn the_first_frame_builds_only_the_first_pages() {
    let fixture = Fixture::new();
    let worker = fixture.worker(2);
    let pages = list(&worker).pages;
    assert!(pages.len() > 8, "the fixture spans pages: {}", pages.len());
    assert!(
        pages[..2]
            .iter()
            .all(|page| !page.unbuilt && !page.primitives.is_empty())
    );
    assert!(
        pages[2..]
            .iter()
            .all(|page| page.unbuilt && page.primitives.is_empty())
    );
    let whole = fixture.whole(&worker);
    for (page, built) in pages.iter().zip(&whole.pages) {
        assert_eq!(
            (page.width.clone(), page.height.clone()),
            (built.width.clone(), built.height.clone())
        );
    }
    assert_eq!(pages[..2], whole.pages[..2]);
    assert_eq!(
        worker.layout_page_heights().len(),
        pages.len(),
        "a height for every page"
    );
}

#[test]
fn pages_built_later_make_the_whole_build() {
    let fixture = Fixture::new();
    let worker = fixture.worker(2);
    let pages = list(&worker).pages.len();
    let epoch = worker.stats().frame_epoch;
    // Out of order, with a page already built and one past the end.
    let frame = worker
        .build_unbuilt_display_pages_frame(&[6, 5, 1, 7, pages + 3], epoch)
        .unwrap();
    assert_eq!(
        frame_ops(&frame),
        (false, 3),
        "only the three pages built go"
    );
    assert!(!list(&worker).pages[5].unbuilt);
    build_rest(&worker, 3);
    assert_eq!(list(&worker), fixture.whole(&worker));
}

#[test]
fn an_edit_while_pages_are_unbuilt_keeps_them_unbuilt_and_builds_them_from_the_new_layout() {
    let fixture = Fixture::new();
    let worker = fixture.worker(2);
    let before = unbuilt(&worker);
    // An edit on a built page that moves every later position.
    type_at(&worker, fixture.starts[1] + 2, "inserted text ");
    let epoch = worker.stats().frame_epoch;
    worker
        .relayout_frame(&fixture.request, &fixture.extras, epoch)
        .unwrap();
    let after = unbuilt(&worker);
    assert!(!after.is_empty() && after.iter().all(|page| before.contains(page)));
    // The unbuilt pages' spans moved with the text: a fresh first frame of the
    // edited document places them the same way.
    let fresh = Fixture {
        state: worker.doc().encode_state_as_update_v1(),
        ..fixture
    };
    let fresh_list = list(&fresh.worker(2));
    let edited = list(&worker);
    for &page in &after {
        assert_eq!(edited.pages[page], fresh_list.pages[page], "page {page}");
    }
    // An edit on an unbuilt page (its paragraph moved by the 14 characters
    // typed above), then the rest.
    type_at(&worker, fresh.starts[PARAGRAPHS - 3] + 14 + 2, "more ");
    let epoch = worker.stats().frame_epoch;
    worker
        .relayout_frame(&fresh.request, &fresh.extras, epoch)
        .unwrap();
    build_rest(&worker, 4);
    assert_eq!(list(&worker), fresh.whole(&worker));
}
