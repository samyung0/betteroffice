//! A resident worker replica lays out a peer's updates the way it lays out a
//! local edit (`EngineSession::relayout_frame`): only changed paragraphs are
//! re-lowered and re-measured, only changed pages are rebuilt and shipped, and
//! a burst of updates costs one layout and one frame. Work is counted, never
//! timed; every frame state is compared with a fresh engine's full layout.

use docx_edit::{EditCtx, EditingDoc, EngineSession, FormatPolicy, Position};
use serde_json::json;

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const PARAGRAPHS: usize = 150;

struct Fixture {
    state: Vec<u8>,
    request: String,
    extras: String,
    /// Story index where each body paragraph starts.
    starts: Vec<u32>,
}

impl Fixture {
    /// One-line paragraphs over several pages under a one-paragraph header.
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

    fn replica(&self, client_id: u64) -> EditingDoc {
        let doc = EditingDoc::new(client_id);
        doc.load_state_v1(&self.state).unwrap();
        doc
    }

    /// A worker replica holding the first frame.
    fn worker(&self) -> (EngineSession, Vec<u8>) {
        let worker = EngineSession::new(9);
        worker.doc().load_state_v1(&self.state).unwrap();
        let first = worker
            .relayout_frame(&self.request, &self.extras, 0)
            .unwrap();
        (worker, first)
    }

    fn relayout(&self, worker: &EngineSession) -> Vec<u8> {
        let epoch = worker.stats().frame_epoch;
        worker
            .relayout_frame(&self.request, &self.extras, epoch)
            .unwrap()
    }

    /// The worker's retained display list equals a fresh engine's full pass.
    fn assert_same_as_fresh(&self, worker: &EngineSession) {
        let fresh = EngineSession::new(10);
        fresh
            .doc()
            .load_state_v1(&worker.doc().encode_state_as_update_v1())
            .unwrap();
        fresh
            .relayout_frame(&self.request, &self.extras, 0)
            .unwrap();
        assert!(fresh.with_display_list(|_| ()).is_some());
        assert_eq!(
            worker.with_display_list(Clone::clone),
            fresh.with_display_list(Clone::clone)
        );
    }
}

fn apply(worker: &EngineSession, peer: &EditingDoc) {
    let update = peer
        .encode_diff_v1(&worker.doc().encode_state_vector_v1())
        .unwrap();
    worker.doc().apply_update_v1(&update).unwrap();
}

fn type_at(peer: &EditingDoc, index: u32, text: &str) {
    peer.insert_text(
        &EditCtx::local("", ""),
        Position::new("body", index),
        text,
        FormatPolicy::Inherit,
    )
    .unwrap();
}

/// `(full, upserts, other operations)` from a FrameDelta header and table.
fn frame_ops(frame: &[u8]) -> (bool, usize, usize) {
    let u32_at = |offset: usize| u32::from_le_bytes(frame[offset..offset + 4].try_into().unwrap());
    let count = u32_at(52) as usize;
    let upserts = (0..count).filter(|op| frame[80 + op * 48] == 1).count();
    (u32_at(12) & 1 == 1, upserts, count - upserts)
}

#[test]
fn a_peer_keystroke_relays_out_one_paragraph_and_ships_one_page() {
    let fixture = Fixture::new();
    let (worker, first) = fixture.worker();
    let pages = worker.stats().retained_pages;
    assert!(pages >= 6, "the fixture spans pages: {pages}");
    let peer = fixture.replica(2);
    type_at(&peer, fixture.starts[PARAGRAPHS / 2] + 4, "x");
    apply(&worker, &peer);

    let before = worker.stats();
    let frame = fixture.relayout(&worker);
    let after = worker.stats();
    assert_eq!(
        after.paragraph_relowerings - before.paragraph_relowerings,
        1
    );
    assert_eq!(
        after.lower_cache_misses, before.lower_cache_misses,
        "no story is lowered in full, the header included"
    );
    assert_eq!(
        after.resident_measure_calls - before.resident_measure_calls,
        1
    );
    assert_eq!(
        after.region_fast_path_hits - before.region_fast_path_hits,
        1
    );
    assert_eq!(
        after.incremental_display_builds - before.incremental_display_builds,
        1
    );
    assert_eq!(
        after.rebuilt_display_pages - before.rebuilt_display_pages,
        1
    );
    let (full, upserts, _) = frame_ops(&frame);
    assert!(!full);
    assert_eq!(upserts, 1, "only the edited page is re-sent");
    assert!(
        frame.len() * pages < first.len() * 2,
        "{} B for one page of a {pages}-page, {} B frame",
        frame.len(),
        first.len()
    );
    fixture.assert_same_as_fresh(&worker);
}

#[test]
fn a_burst_of_peer_updates_costs_one_layout_and_one_frame() {
    let fixture = Fixture::new();
    let (worker, _) = fixture.worker();
    let peer = fixture.replica(2);
    let paragraphs = [10, PARAGRAPHS / 2, PARAGRAPHS - 10];
    let mut updates = 0;
    for paragraph in paragraphs {
        for _ in 0..4 {
            type_at(&peer, fixture.starts[paragraph] + 2, "y");
            apply(&worker, &peer);
            updates += 1;
        }
    }
    assert_eq!(updates, 12);

    let before = worker.stats();
    let frame = fixture.relayout(&worker);
    let after = worker.stats();
    assert_eq!(
        after.region_fast_path_hits - before.region_fast_path_hits,
        1
    );
    assert_eq!(after.pagination_calls - before.pagination_calls, 1);
    assert_eq!(
        after.incremental_pagination_calls - before.incremental_pagination_calls,
        1
    );
    // More than one update since the last frame: the body is lowered again
    // in full, once; the header is not.
    assert_eq!(after.lower_cache_misses - before.lower_cache_misses, 1);
    assert_eq!(after.display_builds - before.display_builds, 1);
    assert_eq!(after.frame_epoch - before.frame_epoch, 1);
    assert_eq!(
        after.resident_measure_calls - before.resident_measure_calls,
        paragraphs.len() as u64
    );
    let (full, upserts, _) = frame_ops(&frame);
    assert!(!full);
    assert!(upserts <= paragraphs.len());
    fixture.assert_same_as_fresh(&worker);
}

#[test]
fn a_peer_edit_outside_the_body_takes_the_full_region_pass() {
    let fixture = Fixture::new();
    let (worker, _) = fixture.worker();
    let peer = fixture.replica(2);
    peer.insert_text(
        &EditCtx::local("", ""),
        Position::new("hf:rIdH", 0),
        "Changed ",
        FormatPolicy::Inherit,
    )
    .unwrap();
    type_at(&peer, fixture.starts[3] + 1, "z");
    apply(&worker, &peer);

    let before = worker.stats();
    let frame = fixture.relayout(&worker);
    let after = worker.stats();
    // The body-only fast path would keep the old header measurement.
    fixture.assert_same_as_fresh(&worker);
    assert_eq!(after.region_fast_path_hits, before.region_fast_path_hits);
    assert_eq!(after.region_fast_path_fallback, Some("outside body"));
    assert!(!frame_ops(&frame).0);

    // The full pass consumed the header edit: the next body keystroke is fast.
    type_at(&peer, fixture.starts[5] + 1, "z");
    apply(&worker, &peer);
    let before = worker.stats();
    fixture.relayout(&worker);
    assert_eq!(
        worker.stats().region_fast_path_hits - before.region_fast_path_hits,
        1
    );
    fixture.assert_same_as_fresh(&worker);
}

#[test]
fn a_changed_layout_request_takes_the_full_region_pass() {
    let mut fixture = Fixture::new();
    let (worker, _) = fixture.worker();
    let peer = fixture.replica(2);
    type_at(&peer, fixture.starts[3] + 1, "w");
    apply(&worker, &peer);
    fixture.request = fixture.request.replace("\"left\":30", "\"left\":60");
    let before = worker.stats();
    fixture.relayout(&worker);
    assert_eq!(
        worker.stats().region_fast_path_hits,
        before.region_fast_path_hits
    );
    fixture.assert_same_as_fresh(&worker);
}
