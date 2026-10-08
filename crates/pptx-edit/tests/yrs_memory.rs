//! Memory safety of the patched yrs under real deck edits, with garbage
//! collection on as sessions and the collaboration server run it.

use std::collections::HashSet;

use pptx_edit::{DeckSession, EditCtx, ShapeDraft, ShapeRect, TextStyle, UndoCaptureMode};
use yrs::block::ItemContent;
use yrs::branch::Branch;
use yrs::updates::decoder::Decode;
use yrs::updates::encoder::Encode;
use yrs::{Doc, ID, Out, ReadTxn, Transact, Update};

const DEMO: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

fn ctx() -> EditCtx {
    EditCtx::local("test")
}

fn sync(from: &DeckSession, to: &DeckSession) {
    let diff = from.encode_diff_v1(&to.encode_state_vector_v1()).unwrap();
    to.apply_update_v1(&diff).unwrap();
}

/// Items the store holds that no root reaches. An item outlives its parent
/// only when garbage collection freed the parent's branch under it, so the
/// item's parent and neighbour links point into freed memory. Walks from the
/// roots and reads each orphan only by id, never through its links.
fn orphans(doc: &Doc) -> Vec<ID> {
    fn walk(branch: &Branch, seen: &mut HashSet<ID>) {
        let mut next = branch.start();
        while let Some(item) = next {
            visit(&item, seen);
            next = item.right();
        }
        for (_, last) in branch.map_items() {
            let mut next = Some(last);
            while let Some(item) = next {
                visit(&item, seen);
                next = item.left();
            }
        }
    }
    fn visit(item: &yrs::block::Item, seen: &mut HashSet<ID>) {
        seen.insert(*item.id());
        if let ItemContent::Type(branch) = item.content() {
            walk(branch, seen);
        }
    }
    let txn = doc.transact();
    let mut seen = HashSet::new();
    for (_, root) in txn.root_refs() {
        let branch: &Branch = match &root {
            Out::YMap(map) => map.as_ref(),
            Out::YArray(array) => array.as_ref(),
            Out::YText(text) => text.as_ref(),
            Out::UndefinedRef(branch) => branch,
            other => panic!("unexpected root {other:?}"),
        };
        walk(branch, &mut seen);
    }
    let store = txn.store();
    let mut orphans = Vec::new();
    for (client, end) in txn.state_vector().iter() {
        let mut clock = 0;
        while clock < *end {
            let (range, item) = store.get_block_range(&ID::new(*client, clock)).unwrap();
            if let Some(item) = item
                && !seen.contains(item.id())
            {
                orphans.push(*item.id());
            }
            clock = range.clock + range.len;
        }
    }
    orphans
}

/// The investigation's seed-561 schedule, shrunk: a user resizes a shape and
/// presses Undo, and a new edit clears the redo stack. Clearing un-keeps the
/// shape's map though Undo still keeps its old size values. A peer then
/// deletes the slide, and garbage collection must take the kept values with
/// the shape's map instead of leaving them pointing into its freed branch.
#[test]
fn a_deleted_slide_takes_the_values_undo_kept_on_it() {
    let left = DeckSession::open(DEMO, 901).unwrap();
    let right = DeckSession::open(DEMO, 902).unwrap();
    right.set_undo_capture_mode(UndoCaptureMode::Manual);
    let deck = right.snapshot().unwrap();
    let slide = &deck.slides[1];
    let (resized, moved) = (&slide.shapes[0].id, &slide.shapes[1].id);

    right
        .resize_shape(&ctx(), &slide.id, resized, 1_000_000, 900_000)
        .unwrap();
    right.add_undo_barrier();
    assert!(right.undo());
    right.move_shape(&ctx(), &slide.id, moved, 5, 7).unwrap();
    right.add_undo_barrier();
    left.delete_slide(&ctx(), &slide.id).unwrap();
    sync(&left, &right);
    assert_eq!(orphans(right.yrs_doc()), []);

    let stored = right.encode_state_as_update_v1();
    let reopened = DeckSession::open_from_update_with_source(&stored, DEMO, 903).unwrap();
    while right.undo() {}
    sync(&right, &left);
    sync(&left, &right);
    assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
    assert_eq!(
        reopened.snapshot().unwrap().slides.len(),
        deck.slides.len() - 1
    );
}

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
}

/// One random edit, Undo or Redo on what `session` shows, as a user would
/// make it; an edit the engine refuses is skipped.
fn random_step(rng: &mut Rng, session: &DeckSession, tag: &str) {
    let deck = session.snapshot().unwrap();
    let Some(slide) = deck.slides.get(rng.below(deck.slides.len())) else {
        session.insert_slide(&ctx(), 0, None).unwrap();
        session.add_undo_barrier();
        return;
    };
    let shape = (!slide.shapes.is_empty()).then(|| &slide.shapes[rng.below(slide.shapes.len())].id);
    let story = slide
        .shapes
        .iter()
        .flat_map(|shape| &shape.text_stories)
        .nth(rng.below(4));
    let n = rng.below(1_000_000) as i64;
    let _ = match (rng.below(17), shape, story) {
        (0, Some(shape), _) => session
            .move_shape(&ctx(), &slide.id, shape, n, n / 2)
            .map(drop),
        (1 | 2, Some(shape), _) => session
            .resize_shape(&ctx(), &slide.id, shape, n + 1, n / 2 + 1)
            .map(drop),
        (3, Some(shape), _) => session
            .set_shape_fill(&ctx(), &slide.id, shape, Some("#AA3300"))
            .map(drop),
        (4, Some(shape), _) => session.remove_shape(&ctx(), &slide.id, shape).map(drop),
        (5, _, Some(story)) => {
            let at = rng.below(story.length as usize + 1) as u32;
            session
                .insert_text(&ctx(), &story.id, at, tag, &TextStyle::default())
                .map(drop)
        }
        (6, _, Some(story)) if story.length > 1 => {
            let at = rng.below(story.length as usize - 1) as u32;
            session.delete_text(&ctx(), &story.id, at, at + 1).map(drop)
        }
        (7, ..) => session.set_slide_notes(&ctx(), &slide.id, tag),
        (8, ..) => {
            let draft = ShapeDraft {
                name: tag.to_owned(),
                rect: ShapeRect {
                    x: 1,
                    y: 2,
                    width: 900_000,
                    height: 300_000,
                },
                text: tag.to_owned(),
                style: TextStyle::default(),
            };
            session.add_text_box(&ctx(), &slide.id, &draft).map(drop)
        }
        (9, ..) => session
            .insert_slide(&ctx(), rng.below(deck.slides.len() + 1) as u32, None)
            .map(drop),
        (10 | 11, ..) if deck.slides.len() > 2 => session.delete_slide(&ctx(), &slide.id).map(drop),
        (12, ..) => session
            .move_slide(&ctx(), &slide.id, rng.below(deck.slides.len()) as u32)
            .map(drop),
        (13..=15, ..) => {
            session.undo();
            Ok(())
        }
        _ => {
            session.redo();
            Ok(())
        }
    };
    session.add_undo_barrier();
}

/// Random edit, Undo and Redo schedules on three peers, delivered pairwise in
/// random orders, and to a plain room document as each is made. After every
/// delivery no session or room may hold an item garbage collection cut off
/// from its parent; at the end the peers converge and the room reopens.
#[test]
fn random_schedules_leave_no_item_without_its_parent() {
    let schedules: u64 = std::env::var("YRS_GC_SCHEDULES").map_or(16, |n| n.parse().unwrap());
    for schedule in 0..schedules {
        let mut rng = Rng(schedule.wrapping_mul(0x9E37_79B9_7F4A_7C15) | 1);
        let peers: Vec<DeckSession> = (0..3)
            .map(|peer| {
                let session = DeckSession::open(DEMO, 700 + peer).unwrap();
                session.set_undo_capture_mode(UndoCaptureMode::Manual);
                session
            })
            .collect();
        let room = Doc::with_options(yrs::Options {
            offset_kind: yrs::OffsetKind::Utf16,
            ..Default::default()
        });
        let check = |at: &str| {
            for (index, peer) in peers.iter().enumerate() {
                assert_eq!(
                    orphans(peer.yrs_doc()),
                    [],
                    "schedule {schedule} {at}: peer {index}"
                );
            }
            assert_eq!(orphans(&room), [], "schedule {schedule} {at}: room");
        };
        for round in 0..4 {
            for step in 0..3 {
                for (index, peer) in peers.iter().enumerate() {
                    random_step(&mut rng, peer, &format!("<{index}.{round}.{step}>"));
                    let known = room.transact().state_vector();
                    let diff = peer.yrs_doc().transact().encode_diff_v1(&known);
                    room.transact_mut()
                        .apply_update(Update::decode_v1(&diff).unwrap())
                        .unwrap();
                }
            }
            for _ in 0..peers.len() * 2 {
                let (from, to) = (rng.below(peers.len()), rng.below(peers.len()));
                if from != to {
                    sync(&peers[from], &peers[to]);
                    check(&format!("round {round} sync {from}->{to}"));
                }
            }
        }
        for from in &peers {
            for to in &peers {
                if !std::ptr::eq(from, to) {
                    sync(from, to);
                }
            }
        }
        check("final");
        let deck = peers[0].snapshot().unwrap();
        for peer in &peers[1..] {
            assert_eq!(peer.snapshot().unwrap(), deck, "schedule {schedule}");
        }
        let stored = room
            .transact()
            .encode_state_as_update_v1(&yrs::StateVector::default());
        let reopened = DeckSession::open_from_update_with_source(&stored, DEMO, 999).unwrap();
        assert_eq!(
            reopened.snapshot().unwrap(),
            deck,
            "schedule {schedule}: room"
        );
    }
}

/// A peer update naming one client's structs twice from the same clock (a
/// second section for the client over the first) is refused. yrs placed the
/// struct twice and freed the first copy while the text still linked it.
#[test]
fn an_update_repeating_a_clients_structs_is_refused() {
    let left = DeckSession::open(DEMO, 901).unwrap();
    let right = DeckSession::open(DEMO, 902).unwrap();
    let deck = left.snapshot().unwrap();
    let story = &deck.slides[0]
        .shapes
        .iter()
        .find_map(|shape| shape.text_stories.first())
        .unwrap()
        .id;
    left.insert_text(&ctx(), story, 0, "abc", &TextStyle::default())
        .unwrap();
    let update = left
        .encode_diff_v1(&right.encode_state_vector_v1())
        .unwrap();
    // One client's section, then the delete set.
    let delete_set = Update::decode_v1(&update).unwrap().delete_set().encode_v1();
    assert_eq!(update[0], 1);
    let section = &update[1..update.len() - delete_set.len()];
    let repeated = [&[2], section, section, &delete_set].concat();

    let before = right.encode_state_as_update_v1();
    assert!(Update::decode_v1(&repeated).is_err());
    assert!(right.apply_update_v1(&repeated).is_err());
    assert_eq!(right.encode_state_as_update_v1(), before);
    right.apply_update_v1(&update).unwrap();
    assert_eq!(right.snapshot().unwrap(), left.snapshot().unwrap());
}
