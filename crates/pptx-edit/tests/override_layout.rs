//! The override layout (schema 6) against the seeded one (schema 5): the same
//! deck after the same edits, one copy per source entry under concurrent
//! first edits, and convergence of two and three peers on random schedules
//! run in lockstep in both layouts.

use std::collections::BTreeSet;

use pptx_edit::{
    DeckSession, DeckSnapshot, EditCtx, Layout, ShapeDraft, ShapeRect, ShapeSnapshot,
    StorySnapshot, TextStyle, TextStylePatch, UndoCaptureMode,
};

const DEMO: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

fn fixtures() -> Vec<(String, Vec<u8>)> {
    let mut files: Vec<(String, Vec<u8>)> = vec![("demo".into(), DEMO.to_vec())];
    for name in [
        "hidden-shapes",
        "blip-shadow",
        "chart-text-overflow",
        "deck-schema-v2-connectors",
        "deck-schema-v2-nested-connectors",
        "metafile-tracking",
        "modern-comments",
        "run-spacing-shadow",
    ] {
        let path = format!("{}/tests/fixtures/{name}.pptx", env!("CARGO_MANIFEST_DIR"));
        files.push((name.into(), std::fs::read(path).unwrap()));
    }
    if let Ok(dir) = std::env::var("PPTX_DIR") {
        let mut extra: Vec<_> = std::fs::read_dir(dir)
            .unwrap()
            .filter_map(|entry| entry.ok().map(|entry| entry.path()))
            .filter(|path| path.extension().is_some_and(|ext| ext == "pptx"))
            .collect();
        extra.sort();
        for path in extra {
            let name = path.file_stem().unwrap().to_string_lossy().into_owned();
            files.push((name, std::fs::read(path).unwrap()));
        }
    }
    files
}

fn ctx() -> EditCtx {
    EditCtx::local("test")
}

fn manual(session: DeckSession) -> DeckSession {
    session.set_undo_capture_mode(UndoCaptureMode::Manual);
    session
}

fn open(bytes: &[u8], client: u64, layout: Layout) -> DeckSession {
    manual(DeckSession::open_with_layout(bytes, client, layout).unwrap())
}

fn all_shapes(shapes: &[ShapeSnapshot]) -> Vec<&ShapeSnapshot> {
    let mut out = Vec::new();
    for shape in shapes {
        out.push(shape);
        out.extend(all_shapes(&shape.children));
    }
    out
}

fn stories(deck: &DeckSnapshot) -> Vec<&StorySnapshot> {
    deck.slides
        .iter()
        .flat_map(|slide| all_shapes(&slide.shapes))
        .flat_map(|shape| &shape.text_stories)
        .collect()
}

/// Reserved writer ids (copies of source entries) a state holds.
fn copies(session: &DeckSession) -> usize {
    use yrs::updates::decoder::Decode;
    let sv = yrs::StateVector::decode_v1(&session.encode_state_vector_v1()).unwrap();
    sv.iter()
        .filter(|(client, _)| {
            let client = client.get();
            ((1_u64 << 52)..(1_u64 << 53) - 1).contains(&client)
        })
        .count()
}

fn sync(from: &DeckSession, to: &DeckSession) {
    let diff = from.encode_diff_v1(&to.encode_state_vector_v1()).unwrap();
    to.apply_update_v1(&diff).unwrap();
}

fn sync_all(peers: &[DeckSession]) {
    for _ in 0..2 {
        for from in peers {
            for to in peers {
                if !std::ptr::eq(from, to) {
                    sync(from, to);
                }
            }
        }
    }
}

#[test]
fn override_open_projects_the_seeded_deck() {
    for (name, bytes) in fixtures() {
        let seeded = DeckSession::open_with_layout(&bytes, 7, Layout::Seeded).unwrap();
        let overrides = DeckSession::open_with_layout(&bytes, 7, Layout::Overrides).unwrap();
        assert_eq!(overrides.layout(), Layout::Overrides);
        assert_eq!(
            seeded.snapshot().unwrap(),
            overrides.snapshot().unwrap(),
            "{name}"
        );
        assert_eq!(copies(&overrides), 0, "{name}");
        let state = overrides.encode_state_as_update_v1();
        assert!(state.len() < 400, "{name}: {} bytes", state.len());
        let reopened = DeckSession::open_from_update_with_source(&state, &bytes, 8).unwrap();
        assert_eq!(reopened.layout(), Layout::Overrides);
        assert_eq!(reopened.snapshot().unwrap(), seeded.snapshot().unwrap());
        for index in 0..seeded.snapshot().unwrap().slides.len() {
            assert_eq!(
                seeded.slide_scope(index).unwrap(),
                overrides.slide_scope(index).unwrap()
            );
        }
        assert_eq!(seeded.save().unwrap(), overrides.save().unwrap(), "{name}");
    }
}

/// Every source story copied on its first edit reads as the seed wrote it:
/// one edit in each story gives the same deck and the same saved file.
#[test]
fn every_story_copy_matches_the_seed() {
    for (name, bytes) in fixtures() {
        let seeded = open(&bytes, 9, Layout::Seeded);
        let overrides = open(&bytes, 9, Layout::Overrides);
        let ids: Vec<String> = stories(&seeded.snapshot().unwrap())
            .iter()
            .map(|story| story.id.clone())
            .collect();
        for id in &ids {
            for session in [&seeded, &overrides] {
                session
                    .insert_text(&ctx(), id, 0, "+", &TextStyle::default())
                    .unwrap();
            }
            assert_eq!(seeded.story(id).unwrap(), overrides.story(id).unwrap());
        }
        assert_eq!(
            seeded.snapshot().unwrap(),
            overrides.snapshot().unwrap(),
            "{name}"
        );
        assert_eq!(copies(&overrides), ids.len(), "{name}");
        assert_eq!(seeded.save().unwrap(), overrides.save().unwrap(), "{name}");
    }
}

/// Deck, slide and shape edits on source entries: same deck, same file, and
/// the state reopens to it.
#[test]
fn deck_and_shape_edits_match_the_seeded_layout() {
    for (name, bytes) in fixtures() {
        let seeded = open(&bytes, 11, Layout::Seeded);
        let overrides = open(&bytes, 11, Layout::Overrides);
        let deck = seeded.snapshot().unwrap();
        let first = &deck.slides[0];
        let edits = |session: &DeckSession| {
            let shapes = all_shapes(&first.shapes);
            if let Some(shape) = first.shapes.first() {
                session
                    .move_shape(&ctx(), &first.id, &shape.id, 1000, 2000)
                    .unwrap();
                let _ = session.set_shape_fill(&ctx(), &first.id, &shape.id, Some("#336699"));
                session
                    .bring_to_front(&ctx(), &first.id, &shape.id)
                    .unwrap();
            }
            if let Some(last) = first.shapes.last() {
                session.remove_shape(&ctx(), &first.id, &last.id).unwrap();
            }
            session
                .add_text_box(
                    &ctx(),
                    &first.id,
                    &ShapeDraft {
                        name: "Added".into(),
                        rect: ShapeRect {
                            x: 10,
                            y: 10,
                            width: 1_000_000,
                            height: 400_000,
                        },
                        text: "new box".into(),
                        style: TextStyle::default(),
                    },
                )
                .unwrap();
            session.set_slide_notes(&ctx(), &first.id, "notes").unwrap();
            let slide = session.insert_slide(&ctx(), 1, None).unwrap();
            session.move_slide(&ctx(), &slide.slide_id, 0).unwrap();
            if deck.slides.len() > 2 {
                session.delete_slide(&ctx(), &deck.slides[2].id).unwrap();
            }
            shapes.len()
        };
        edits(&seeded);
        edits(&overrides);
        assert_eq!(
            seeded.snapshot().unwrap(),
            overrides.snapshot().unwrap(),
            "{name}"
        );
        assert_eq!(seeded.save().unwrap(), overrides.save().unwrap(), "{name}");
        let reopened = DeckSession::open_from_update_with_source(
            &overrides.encode_state_as_update_v1(),
            &bytes,
            12,
        )
        .unwrap();
        assert_eq!(reopened.snapshot().unwrap(), seeded.snapshot().unwrap());
        while seeded.undo() {
            assert!(overrides.undo(), "{name}");
            assert_eq!(seeded.snapshot().unwrap(), overrides.snapshot().unwrap());
        }
        assert!(!overrides.undo());
        assert_eq!(overrides.snapshot().unwrap(), deck, "{name}");
    }
}

/// Three peers make their first edit of one source story and one source
/// shape at once: one copy each, every peer's text kept, the source text once.
#[test]
fn concurrent_first_edits_converge_on_one_copy() {
    let peers: Vec<DeckSession> = (0..3)
        .map(|peer| open(DEMO, 501 + peer, Layout::Overrides))
        .collect();
    let deck = peers[0].snapshot().unwrap();
    let (slide, shape, story) = deck
        .slides
        .iter()
        .find_map(|slide| {
            let shape = slide.shapes.iter().find(|shape| {
                shape
                    .text_stories
                    .first()
                    .is_some_and(|story| story.plain_text().len() > 4)
            })?;
            Some((slide, shape, shape.text_stories[0].clone()))
        })
        .unwrap();
    let base_text = story.plain_text();
    let end = story.length - 1;
    for (peer, session) in peers.iter().enumerate() {
        let at = [0, 2, end][peer];
        session
            .insert_text(
                &ctx(),
                &story.id,
                at,
                &format!("<p{peer}>"),
                &TextStyle::default(),
            )
            .unwrap();
    }
    peers[1]
        .move_shape(&ctx(), &slide.id, &shape.id, 5_000, 6_000)
        .unwrap();
    peers[2]
        .resize_shape(&ctx(), &slide.id, &shape.id, 2_000_000, 900_000)
        .unwrap();
    sync_all(&peers);
    let first = peers[0].snapshot().unwrap();
    for session in &peers[1..] {
        assert_eq!(session.snapshot().unwrap(), first);
    }
    let merged = peers[0].story(&story.id).unwrap().plain_text();
    for peer in 0..3 {
        assert_eq!(merged.matches(&format!("<p{peer}>")).count(), 1, "{merged}");
    }
    let stripped = merged
        .replace("<p0>", "")
        .replace("<p1>", "")
        .replace("<p2>", "");
    assert_eq!(stripped, base_text);
    let moved = all_shapes(
        &first
            .slides
            .iter()
            .find(|s| s.id == slide.id)
            .unwrap()
            .shapes,
    )
    .into_iter()
    .find(|s| s.id == shape.id)
    .unwrap()
    .clone();
    assert_eq!((moved.x, moved.y), (5_000, 6_000));
    assert_eq!((moved.width, moved.height), (2_000_000, 900_000));
    // One story and one shape were copied, whoever copied them.
    for session in &peers {
        assert_eq!(copies(session), 2);
    }
}

/// Undo after a first edit reverts the edit, keeps the copy, and keeps a
/// peer's edit in the same copy.
#[test]
fn undo_reverts_only_the_users_edit() {
    let left = open(DEMO, 611, Layout::Overrides);
    let right = open(DEMO, 612, Layout::Overrides);
    let deck = left.snapshot().unwrap();
    let story = stories(&deck)
        .into_iter()
        .find(|story| {
            let first: String = story.paragraphs[0]
                .runs
                .iter()
                .map(|run| run.text.as_str())
                .collect();
            first.is_ascii() && first.len() > 4
        })
        .unwrap()
        .clone();
    let base = story.plain_text();
    left.insert_text(&ctx(), &story.id, 0, "{L}", &TextStyle::default())
        .unwrap();
    left.add_undo_barrier();
    right
        .insert_text(&ctx(), &story.id, 2, "{R}", &TextStyle::default())
        .unwrap();
    let peers = [left, right];
    sync_all(&peers);
    assert!(peers[0].undo());
    sync_all(&peers);
    let text = peers[1].story(&story.id).unwrap().plain_text();
    assert_eq!(text, format!("{}{{R}}{}", &base[..2], &base[2..]));
    assert_eq!(peers[0].snapshot().unwrap(), peers[1].snapshot().unwrap());
    assert!(peers[0].redo());
    sync_all(&peers);
    let text = peers[1].story(&story.id).unwrap().plain_text();
    assert_eq!(text, format!("{{L}}{}{{R}}{}", &base[..2], &base[2..]));
}

/// A deleted source slide stays deleted though a peer edited it meanwhile;
/// undoing the delete brings it back as the deleter saw it, in both layouts.
#[test]
fn delete_of_a_source_slide_against_a_concurrent_edit() {
    for layout in [Layout::Seeded, Layout::Overrides] {
        let left = open(DEMO, 621, layout);
        let right = open(DEMO, 622, layout);
        let deck = left.snapshot().unwrap();
        let slide = deck
            .slides
            .iter()
            .find(|slide| {
                slide
                    .shapes
                    .iter()
                    .any(|shape| !shape.text_stories.is_empty())
            })
            .unwrap();
        let shape = slide
            .shapes
            .iter()
            .find(|shape| !shape.text_stories.is_empty())
            .unwrap();
        left.delete_slide(&ctx(), &slide.id).unwrap();
        right
            .move_shape(&ctx(), &slide.id, &shape.id, 77, 88)
            .unwrap();
        let peers = [left, right];
        sync_all(&peers);
        for session in &peers {
            let now = session.snapshot().unwrap();
            assert!(now.slides.iter().all(|s| s.id != slide.id), "{layout:?}");
        }
        assert!(peers[0].undo());
        // Without garbage collection the peer's move is kept as live items
        // in the deleted record, and Yjs will not redo the record's x and y
        // over them: the restored record lacks them and the merge is refused
        // (both layouts). With collection the moved items are gone and the
        // record comes back whole.
        let diff = peers[0]
            .encode_diff_v1(&peers[1].encode_state_vector_v1())
            .unwrap();
        if let Err(error) = peers[1].apply_update_v1(&diff) {
            assert!(
                std::env::var_os("PPTX_EDIT_SKIP_GC").is_some(),
                "{layout:?}: {error}"
            );
            eprintln!("{layout:?} without GC: {error}");
            continue;
        }
        sync_all(&peers);
        let back = peers[1].snapshot().unwrap();
        let restored = back.slides.iter().find(|s| s.id == slide.id).unwrap();
        let moved = restored.shapes.iter().find(|s| s.id == shape.id).unwrap();
        // The delete removed the shape's record with the slide (the override
        // layout copies the slide's entries first, so the peer's copy of the
        // same record went too); the undo restores the deleter's view.
        assert_eq!((moved.x, moved.y), (shape.x, shape.y), "{layout:?}");
    }
}

/// A caret anchored in a story nobody copied yet follows text a peer
/// inserts before it: the anchor names the items the copy will have.
#[test]
fn anchor_in_an_uncopied_story_follows_a_peers_insert() {
    let left = open(DEMO, 631, Layout::Overrides);
    let right = open(DEMO, 632, Layout::Overrides);
    let deck = left.snapshot().unwrap();
    let story = stories(&deck)
        .into_iter()
        .find(|story| story.plain_text().len() > 6)
        .unwrap()
        .clone();
    let anchor = left.anchor_caret(&story.id, 4).unwrap();
    assert_eq!(left.resolve_caret_anchor(&anchor), Some(4));
    right
        .insert_text(&ctx(), &story.id, 1, "abc", &TextStyle::default())
        .unwrap();
    sync(&right, &left);
    assert_eq!(left.resolve_caret_anchor(&anchor), Some(7));
}

/// The first path where two snapshots differ, for readable failures.
fn difference(left: &DeckSnapshot, right: &DeckSnapshot) -> String {
    fn walk(path: &str, a: &serde_json::Value, b: &serde_json::Value, out: &mut Vec<String>) {
        use serde_json::Value;
        if out.len() >= 12 {
            return;
        }
        match (a, b) {
            (Value::Object(x), Value::Object(y)) => {
                let keys: BTreeSet<&String> = x.keys().chain(y.keys()).collect();
                for key in keys {
                    let (l, r) = (
                        x.get(key).unwrap_or(&Value::Null),
                        y.get(key).unwrap_or(&Value::Null),
                    );
                    walk(&format!("{path}.{key}"), l, r, out);
                }
            }
            (Value::Array(x), Value::Array(y)) => {
                for index in 0..x.len().max(y.len()) {
                    let (l, r) = (
                        x.get(index).unwrap_or(&Value::Null),
                        y.get(index).unwrap_or(&Value::Null),
                    );
                    walk(&format!("{path}[{index}]"), l, r, out);
                }
            }
            _ if a == b => {}
            _ => out.push(format!(
                "{path}: {} != {}",
                a.to_string().chars().take(80).collect::<String>(),
                b.to_string().chars().take(80).collect::<String>()
            )),
        }
    }
    let (a, b) = (
        serde_json::to_value(left).unwrap(),
        serde_json::to_value(right).unwrap(),
    );
    let mut out = Vec::new();
    walk("", &a, &b, &mut out);
    out.join(
        "
",
    )
}

// ---------------------------------------------------------------------------
// Random schedules in lockstep: every op runs on the seeded and the override
// world with the same choices; the two must read the same after every op,
// every partial sync and the final full sync, and each world must converge.

struct Rng(u64);
impl Rng {
    fn next(&mut self) -> u64 {
        self.0 ^= self.0 << 13;
        self.0 ^= self.0 >> 7;
        self.0 ^= self.0 << 17;
        self.0
    }
    fn below(&mut self, n: usize) -> usize {
        if n == 0 {
            0
        } else {
            (self.next() % n as u64) as usize
        }
    }
    fn chance(&mut self, percent: u64) -> bool {
        self.next() % 100 < percent
    }
}

/// Paragraph text spans `[start, end)` of a story, in story offsets.
fn paragraphs(story: &StorySnapshot) -> Vec<(u32, u32)> {
    let mut spans = Vec::new();
    let mut at = 0;
    for paragraph in &story.paragraphs {
        let len: u32 = paragraph
            .runs
            .iter()
            .map(|run| run.text.encode_utf16().count() as u32)
            .sum();
        spans.push((at, at + len));
        at += len + 1;
    }
    spans
}

#[derive(Debug)]
enum Op {
    Insert(String, u32, String),
    Delete(String, u32, u32),
    Replace(String, u32, u32, String),
    Format(String, u32, u32, bool),
    Split(String, u32),
    Align(String, u32, u32),
    Move(String, String, i64, i64),
    Resize(String, String, i64, i64),
    Fill(String, String),
    Front(String, String),
    Remove(String, String),
    AddBox(String, String),
    Notes(String, String),
    InsertSlide(u32),
    DeleteSlide(String),
    MoveSlide(String, u32),
    Undo,
    Redo,
}

/// An op chosen from what the peer sees. `focus` sends early ops to the
/// first text story and its shape so peers collide on first edits.
fn choose(rng: &mut Rng, deck: &DeckSnapshot, peer: usize, step: usize, focus: bool) -> Op {
    let story_list: Vec<&StorySnapshot> = stories(deck);
    let marker = format!("<{peer}.{step}>");
    let slide_count = deck.slides.len();
    let pick_slide = |rng: &mut Rng| &deck.slides[rng.below(slide_count)];
    let shape_on = |rng: &mut Rng| {
        for _ in 0..8 {
            let slide = pick_slide(rng);
            if !slide.shapes.is_empty() {
                let shape = &slide.shapes[rng.below(slide.shapes.len())];
                return Some((slide.id.clone(), shape.id.clone()));
            }
        }
        None
    };
    let kind = if focus { rng.below(6) } else { rng.below(18) };
    let story = if focus || story_list.is_empty() {
        story_list.first().copied()
    } else {
        Some(story_list[rng.below(story_list.len())])
    };
    let text_op = |rng: &mut Rng, story: &StorySnapshot, kind: usize| {
        let spans = paragraphs(story);
        let (start, end) = spans[rng.below(spans.len())];
        let a = start + rng.below((end - start + 1) as usize) as u32;
        let b = a + rng.below((end - a + 1) as usize) as u32;
        match kind {
            0 | 1 => Op::Insert(story.id.clone(), a, marker.clone()),
            2 => Op::Delete(story.id.clone(), a, b),
            3 => {
                let z = rng.below(story.length as usize) as u32;
                Op::Replace(story.id.clone(), a.min(z), a.max(z), marker.clone())
            }
            4 => Op::Format(
                story.id.clone(),
                a,
                b.max(a + 1).min(end.max(a + 1)),
                rng.chance(50),
            ),
            _ => Op::Split(story.id.clone(), a),
        }
    };
    match (kind, story) {
        (0..=4, Some(story)) => text_op(rng, story, kind),
        (5, Some(story)) if focus => {
            let shape = deck.slides.iter().find_map(|slide| {
                all_shapes(&slide.shapes)
                    .into_iter()
                    .find(|shape| shape.text_stories.iter().any(|s| s.id == story.id))
                    .map(|shape| (slide.id.clone(), shape.id.clone()))
            });
            match shape {
                Some((slide, shape)) => {
                    Op::Move(slide, shape, rng.below(9_000) as i64, peer as i64)
                }
                None => text_op(rng, story, 0),
            }
        }
        (5, Some(story)) => text_op(rng, story, 5),
        (6, Some(story)) => {
            let spans = paragraphs(story);
            let (start, end) = spans[rng.below(spans.len())];
            Op::Align(story.id.clone(), start, end)
        }
        (7, _) => shape_on(rng).map_or(Op::Undo, |(s, h)| {
            Op::Move(
                s,
                h,
                rng.below(1_000_000) as i64,
                rng.below(1_000_000) as i64,
            )
        }),
        (8, _) => shape_on(rng).map_or(Op::Undo, |(s, h)| {
            Op::Resize(
                s,
                h,
                1 + rng.below(3_000_000) as i64,
                1 + rng.below(3_000_000) as i64,
            )
        }),
        (9, _) => shape_on(rng).map_or(Op::Undo, |(s, h)| Op::Fill(s, h)),
        (10, _) => shape_on(rng).map_or(Op::Undo, |(s, h)| Op::Front(s, h)),
        (11, _) if rng.chance(40) => shape_on(rng).map_or(Op::Undo, |(s, h)| Op::Remove(s, h)),
        (12, _) => Op::AddBox(pick_slide(rng).id.clone(), marker),
        (13, _) => Op::Notes(pick_slide(rng).id.clone(), marker),
        (14, _) => Op::InsertSlide(rng.below(slide_count + 1) as u32),
        (15, _) if slide_count > 2 && rng.chance(30) => Op::DeleteSlide(pick_slide(rng).id.clone()),
        (15, _) => Op::MoveSlide(pick_slide(rng).id.clone(), rng.below(slide_count) as u32),
        (16, _) => Op::Undo,
        (17, _) => Op::Redo,
        _ => Op::Undo,
    }
}

/// Undo or Redo as the user sees it: pressed again while the deck reads the
/// same. A step can change only content nobody sees: formatting already-bold
/// text in either layout, or, in the override layout, a story the user copied
/// on a slide a peer deleted meanwhile (the seeded layout's delete removed
/// that story, so Yjs passes over the step there).
fn visible_step(session: &DeckSession, step: fn(&DeckSession) -> bool) -> bool {
    let before = session.snapshot().unwrap();
    while step(session) {
        if session.snapshot().unwrap() != before {
            return true;
        }
    }
    false
}

fn run(session: &DeckSession, op: &Op) -> Result<(), String> {
    let style = TextStyle::default();
    let result = match op {
        Op::Insert(story, at, text) => session
            .insert_text(&ctx(), story, *at, text, &style)
            .map(drop),
        Op::Delete(story, a, b) => session.delete_text(&ctx(), story, *a, *b).map(drop),
        Op::Replace(story, a, b, text) => session
            .replace_text(&ctx(), story, *a, *b, text, &style)
            .map(drop),
        Op::Format(story, a, b, bold) => session
            .format_text(
                &ctx(),
                story,
                *a,
                *b,
                &TextStylePatch {
                    bold: Some(*bold),
                    ..Default::default()
                },
            )
            .map(drop),
        Op::Split(story, at) => session.insert_paragraph_break(&ctx(), story, *at).map(drop),
        Op::Align(story, a, b) => session
            .set_paragraph_alignment(&ctx(), story, *a, *b, Some("ctr"))
            .map(drop),
        Op::Move(slide, shape, x, y) => session.move_shape(&ctx(), slide, shape, *x, *y).map(drop),
        Op::Resize(slide, shape, w, h) => {
            session.resize_shape(&ctx(), slide, shape, *w, *h).map(drop)
        }
        Op::Fill(slide, shape) => session
            .set_shape_fill(&ctx(), slide, shape, Some("#AA3300"))
            .map(drop),
        Op::Front(slide, shape) => session.bring_to_front(&ctx(), slide, shape).map(drop),
        Op::Remove(slide, shape) => session.remove_shape(&ctx(), slide, shape).map(drop),
        Op::AddBox(slide, text) => session
            .add_text_box(
                &ctx(),
                slide,
                &ShapeDraft {
                    name: text.clone(),
                    rect: ShapeRect {
                        x: 1,
                        y: 2,
                        width: 900_000,
                        height: 300_000,
                    },
                    text: text.clone(),
                    style: TextStyle::default(),
                },
            )
            .map(drop),
        Op::Notes(slide, text) => session.set_slide_notes(&ctx(), slide, text),
        Op::InsertSlide(at) => session.insert_slide(&ctx(), *at, None).map(drop),
        Op::DeleteSlide(slide) => session.delete_slide(&ctx(), slide).map(drop),
        Op::MoveSlide(slide, to) => session.move_slide(&ctx(), slide, *to).map(drop),
        Op::Undo => {
            visible_step(session, DeckSession::undo);
            Ok(())
        }
        Op::Redo => {
            visible_step(session, DeckSession::redo);
            Ok(())
        }
    };
    session.add_undo_barrier();
    result.map_err(|error| error.to_string())
}

fn lockstep(bytes: &[u8], seed: u64, peers: usize, rounds: usize, ops: usize) -> (usize, usize) {
    let mut rng = Rng(seed.wrapping_mul(0x9E37_79B9_7F4A_7C15) | 1);
    let layouts: Vec<Layout> = match std::env::var("OVERRIDE_ONLY").as_deref() {
        Ok("seeded") => vec![Layout::Seeded],
        Ok("override") => vec![Layout::Overrides],
        _ => vec![Layout::Seeded, Layout::Overrides],
    };
    let both = layouts.len() == 2;
    let worlds: Vec<Vec<DeckSession>> = layouts
        .iter()
        .map(|layout| {
            (0..peers)
                .map(|peer| open(bytes, 700 + peer as u64, *layout))
                .collect()
        })
        .collect();
    let (mut applied, mut refused) = (0, 0);
    for round in 0..rounds {
        for step in 0..ops {
            for peer in 0..peers {
                let deck = worlds[0][peer].snapshot().unwrap();
                let op = choose(
                    &mut rng,
                    &deck,
                    peer,
                    round * ops + step,
                    round == 0 && step < 2,
                );
                let seeded = run(&worlds[0][peer], &op);
                let overrides = if both {
                    run(&worlds[1][peer], &op)
                } else {
                    seeded.clone()
                };
                if std::env::var("OVERRIDE_TRACE").ok() == Some(seed.to_string()) {
                    let json = serde_json::to_string(&worlds[0][peer].snapshot().unwrap()).unwrap();
                    eprintln!(
                        "r{round} s{step} p{peer} {op:?} -> {seeded:?} #{}",
                        json.len()
                    );
                }
                assert_eq!(
                    seeded, overrides,
                    "seed {seed} round {round} peer {peer} {op:?}"
                );
                if seeded.is_ok() {
                    applied += 1
                } else {
                    refused += 1
                }
                if !both {
                    continue;
                }
                let (left, right) = (
                    worlds[0][peer].snapshot().unwrap(),
                    worlds[1][peer].snapshot().unwrap(),
                );
                assert!(
                    left == right,
                    "seed {seed} round {round} peer {peer} after {op:?}: {}",
                    difference(&left, &right)
                );
            }
        }
        // A partial sync: some pairs, in a random order, the same in both worlds.
        let pairs: Vec<(usize, usize)> = (0..peers * 2)
            .map(|_| (rng.below(peers), rng.below(peers)))
            .filter(|(from, to)| from != to)
            .collect();
        for (from, to) in pairs {
            if std::env::var("OVERRIDE_TRACE").ok() == Some(seed.to_string()) {
                eprintln!("r{round} sync {from}->{to}");
            }
            for (layout, world) in worlds.iter().enumerate() {
                let diff = world[from]
                    .encode_diff_v1(&world[to].encode_state_vector_v1())
                    .unwrap();
                if let Err(error) = world[to].apply_update_v1(&diff) {
                    panic!("seed {seed} round {round} layout {layout} sync {from}->{to}: {error}");
                }
            }
            if std::env::var("OVERRIDE_TRACE").ok() == Some(seed.to_string()) {
                let json = serde_json::to_string(&worlds[0][to].snapshot().unwrap()).unwrap();
                eprintln!("r{round} synced {from}->{to} #{}", json.len());
            }
            if !both {
                continue;
            }
            let (left, right) = (
                worlds[0][to].snapshot().unwrap(),
                worlds[1][to].snapshot().unwrap(),
            );
            assert!(
                left == right,
                "seed {seed} round {round} sync {from}->{to}: {}",
                difference(&left, &right)
            );
        }
    }
    for world in &worlds {
        sync_all(world);
        let first = world[0].snapshot().unwrap();
        for session in &world[1..] {
            assert_eq!(
                session.snapshot().unwrap(),
                first,
                "seed {seed}: peers diverge"
            );
        }
    }
    if !both {
        return (applied, refused);
    }
    assert_eq!(
        worlds[0][0].snapshot().unwrap(),
        worlds[1][0].snapshot().unwrap(),
        "seed {seed}"
    );
    assert_eq!(
        worlds[0][0].save().unwrap(),
        worlds[1][0].save().unwrap(),
        "seed {seed}"
    );
    // The override state reopens to the same deck.
    let reopened = DeckSession::open_from_update_with_source(
        &worlds[1][0].encode_state_as_update_v1(),
        bytes,
        799,
    )
    .unwrap();
    assert_eq!(
        reopened.snapshot().unwrap(),
        worlds[0][0].snapshot().unwrap()
    );
    (applied, refused)
}

#[test]
fn random_schedules_match_the_seeded_layout_and_converge() {
    let seeds: u64 = std::env::var("OVERRIDE_SEEDS")
        .ok()
        .and_then(|v| v.parse().ok())
        .unwrap_or(24);
    let mut totals = (0, 0);
    let mut files = BTreeSet::new();
    let only: Option<u64> = std::env::var("OVERRIDE_SEED")
        .ok()
        .and_then(|v| v.parse().ok());
    for (name, bytes) in fixtures().into_iter().filter(|(name, _)| {
        matches!(
            name.as_str(),
            "demo" | "deck-schema-v2-nested-connectors" | "modern-comments"
        ) || std::env::var("PPTX_DIR").is_ok()
    }) {
        for seed in 0..seeds {
            for peers in [2, 3] {
                if only.is_some_and(|only| only != seed * 31 + peers as u64) {
                    continue;
                }
                let (applied, refused) = lockstep(&bytes, seed * 31 + peers as u64, peers, 4, 3);
                totals.0 += applied;
                totals.1 += refused;
            }
        }
        files.insert(name);
    }
    eprintln!(
        "lockstep: {files:?}, {} ops applied, {} refused alike",
        totals.0, totals.1
    );
    assert!(totals.0 > totals.1);
}
