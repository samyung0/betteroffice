//! Undo of a deleted slide or shape keeps what a peer did inside it meanwhile
//! (upstream #1506; Capy's own fix, Epo 2026-11-04, Q549 A).
//!
//! Each update takes the path it takes in a room: the writer's host document, the room
//! document (gc on), and for every other peer its host document and its engine. A room
//! relays what the writer told when its own collection took part of it, and a host hands
//! its engine the frame as it arrived.
use std::cell::RefCell;
use std::rc::Rc;

use pptx_edit::{
    DeckSession, EditCtx, ShapeDraft, ShapeRect, TextStyle, UpdateOrigin, UpdateSubscription,
};
use yrs::updates::decoder::Decode;
use yrs::{ClientID, Doc, OffsetKind, Options, ReadTxn, StateVector, Transact, Update};

const DECK: &[u8] = include_bytes!("../../../poc/fixtures/deck-50.pptx");

fn relay_doc(client: u64, seed: &[u8]) -> Doc {
    let mut options = Options::with_client_id(ClientID::new(client));
    options.offset_kind = OffsetKind::Utf16;
    options.cleanup_formatting = false;
    let doc = Doc::with_options(options);
    doc.transact_mut()
        .apply_update(Update::decode_v1(seed).unwrap())
        .unwrap();
    doc
}

/// The room's apply: what it relays for `update`. A resync's step 2 names its writer's whole
/// delete set, so only what landed makes the room relay it as sent.
fn room_relay(room: &Doc, update: &[u8], step2: bool) -> Option<Vec<u8>> {
    let mut txn = room.transact_mut();
    txn.apply_update(Update::decode_v1(update).unwrap())
        .unwrap();
    txn.commit();
    let landed = !txn.insert_set().intersect(txn.delete_set()).is_empty();
    let deletions = !txn.redundant_deletes().is_empty();
    if landed || (deletions && !step2) {
        return Some(yrs::merge_updates_v1([update]).unwrap());
    }
    if txn.delete_set().is_empty() && txn.after_state() == txn.before_state() {
        return None;
    }
    Some(txn.encode_update_v1())
}

struct Peer {
    session: DeckSession,
    outbox: Rc<RefCell<Vec<Vec<u8>>>>,
    host: Doc,
    online: std::cell::Cell<bool>,
    /// Frames on their way: this peer has not heard them yet.
    delayed: RefCell<Option<Vec<Vec<u8>>>>,
    /// How much of the outbox the host document holds: it learns an edit as it is made.
    hosted: std::cell::Cell<usize>,
    _sub: UpdateSubscription,
}

impl Peer {
    fn open(seed: &[u8], client: u64) -> Self {
        let session = DeckSession::open_from_update_with_source(seed, DECK, client).unwrap();
        let outbox = Rc::new(RefCell::new(Vec::new()));
        let sink = Rc::clone(&outbox);
        let sub = session
            .observe_update_v1(move |event| {
                if event.origin == UpdateOrigin::Local {
                    sink.borrow_mut().push(event.update);
                }
            })
            .unwrap();
        Self {
            session,
            outbox,
            host: relay_doc(client + 1000, seed),
            online: std::cell::Cell::new(true),
            delayed: RefCell::new(None),
            hosted: std::cell::Cell::new(0),
            _sub: sub,
        }
    }

    fn host_own_edits(&self) {
        let outbox = self.outbox.borrow();
        for update in &outbox[self.hosted.get()..] {
            self.host
                .transact_mut()
                .apply_update(Update::decode_v1(update).unwrap())
                .unwrap();
        }
        self.hosted.set(outbox.len());
    }

    fn receive(&self, frame: &[u8]) {
        if let Some(delayed) = self.delayed.borrow_mut().as_mut() {
            delayed.push(frame.to_vec());
            return;
        }
        self.host_own_edits();
        self.host
            .transact_mut()
            .apply_update(Update::decode_v1(frame).unwrap())
            .unwrap();
        self.session.apply_update_v1(frame).unwrap();
    }
}

struct Net {
    seed: Vec<u8>,
    room: Doc,
    peers: Vec<Peer>,
}

impl Net {
    fn new(seed: Vec<u8>, peers: usize) -> Self {
        let room = relay_doc(9000, &seed);
        let peers = (0..peers)
            .map(|index| Peer::open(&seed, 501 + index as u64))
            .collect();
        Self { seed, room, peers }
    }

    fn deck(peers: usize) -> Self {
        Self::new(
            DeckSession::open(DECK, 101)
                .unwrap()
                .encode_state_as_update_v1(),
            peers,
        )
    }

    fn peer(&self, index: usize) -> &DeckSession {
        &self.peers[index].session
    }

    /// Sends what peer `index` wrote; it reaches the others and may make them write.
    fn send(&self, index: usize) {
        self.peers[index].host_own_edits();
        self.peers[index].hosted.set(0);
        let outgoing: Vec<_> = self.peers[index].outbox.borrow_mut().drain(..).collect();
        for update in outgoing {
            let Some(frame) = room_relay(&self.room, &update, false) else {
                continue;
            };
            for (other, peer) in self.peers.iter().enumerate() {
                if other != index && peer.online.get() {
                    peer.receive(&frame);
                }
            }
        }
    }

    /// A socket comes back: the reply to the room's step 1, which the room sends first: all the
    /// room lacks from the peer's host document, with its whole delete set. It leaves before
    /// the peer hears what the room answers.
    fn step2(&self, index: usize) {
        let peer = &self.peers[index];
        peer.host_own_edits();
        peer.hosted.set(0);
        peer.outbox.borrow_mut().clear();
        let known = self.room.transact().state_vector();
        let update = peer.host.transact().encode_diff_v1(&known);
        let Some(frame) = room_relay(&self.room, &update, true) else {
            return;
        };
        for (other, peer) in self.peers.iter().enumerate() {
            if other != index {
                peer.receive(&frame);
            }
        }
    }

    /// What reaches peer `index` from now on stays on its way until [Self::hear].
    fn delay(&self, index: usize) {
        *self.peers[index].delayed.borrow_mut() = Some(Vec::new());
    }

    fn hear(&self, index: usize) {
        let frames = self.peers[index]
            .delayed
            .borrow_mut()
            .take()
            .unwrap_or_default();
        for frame in frames {
            self.peers[index].receive(&frame);
        }
    }

    /// Until no peer has anything left to send.
    fn settle(&self) {
        for index in 0..self.peers.len() {
            self.hear(index);
        }
        for _ in 0..8 {
            if self
                .peers
                .iter()
                .all(|peer| peer.outbox.borrow().is_empty())
            {
                return;
            }
            for index in 0..self.peers.len() {
                self.send(index);
            }
        }
        panic!("peers keep writing");
    }

    /// A socket comes back: the room answers with its own encoding of what the peer lacks.
    fn reconnect(&self, index: usize) {
        let peer = &self.peers[index];
        peer.online.set(true);
        let missed = self
            .room
            .transact()
            .encode_diff_v1(&peer.host.transact().state_vector());
        peer.receive(&missed);
    }

    /// Everyone reads alike, and a joiner reads the same from the room's state.
    fn assert_converged(&self) {
        self.settle();
        let expected = self.peer(0).snapshot().unwrap();
        for peer in &self.peers[1..] {
            assert_eq!(peer.session.snapshot().unwrap(), expected);
        }
        let state = self
            .room
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        let joiner = DeckSession::open_from_update_with_source(&state, DECK, 900).unwrap();
        assert_eq!(joiner.snapshot().unwrap(), expected);
        let _ = &self.seed;
    }
}

/// The first slide's first shape with text.
fn target(session: &DeckSession) -> (String, String, String) {
    let snapshot = session.snapshot().unwrap();
    snapshot
        .slides
        .iter()
        .find_map(|slide| {
            slide
                .shapes
                .iter()
                .find(|shape| !shape.text_stories.is_empty())
                .map(|shape| {
                    (
                        slide.id.clone(),
                        shape.id.clone(),
                        shape.text_stories[0].id.clone(),
                    )
                })
        })
        .unwrap()
}

fn delete(session: &DeckSession, slide: bool, slide_id: &str, shape_id: &str) {
    let context = EditCtx::local("left");
    if slide {
        session.delete_slide(&context, slide_id).unwrap();
    } else {
        session.remove_shape(&context, slide_id, shape_id).unwrap();
    }
    session.add_undo_barrier();
}

fn position(session: &DeckSession, shape_id: &str) -> (i64, i64) {
    let snapshot = session.snapshot().unwrap();
    let shape = snapshot
        .slides
        .iter()
        .flat_map(|slide| slide.shapes.iter())
        .find(|shape| shape.id == shape_id)
        .unwrap();
    (shape.x, shape.y)
}

fn text(session: &DeckSession, story: &str) -> String {
    session.story(story).unwrap().plain_text()
}

/// Both arrival orders at the room, for slide and shape, with the undoer's client id below and
/// above the peer's (an Undo restores a step's items in client order).
#[test]
fn undo_brings_a_deleted_slide_or_shape_back_with_the_peers_text_and_move() {
    for (slide, peer_first, undoer) in (0..8).map(|bits| (bits & 1 == 0, bits & 2 == 0, bits >> 2))
    {
        {
            let net = Net::deck(2);
            let (left, right) = (net.peer(undoer), net.peer(1 - undoer));
            let (slide_id, shape_id, story) = target(left);
            let before = text(left, &story);
            let end = right.story(&story).unwrap().length - 1;
            let context = EditCtx::local("right");
            delete(left, slide, &slide_id, &shape_id);
            right
                .insert_text(&context, &story, end, " REMOTE", &TextStyle::default())
                .unwrap();
            right
                .move_shape(&context, &slide_id, &shape_id, 1_234_567, 2_345_678)
                .unwrap();
            let order = if peer_first {
                [1 - undoer, undoer]
            } else {
                [undoer, 1 - undoer]
            };
            net.send(order[0]);
            net.send(order[1]);
            assert!(left.undo());
            net.assert_converged();
            for session in [left, right] {
                assert_eq!(text(session, &story), format!("{before} REMOTE"));
                assert_eq!(position(session, &shape_id), (1_234_567, 2_345_678));
            }
        }
    }
}

/// What the peer deleted and what it typed both stand; its paragraph break and its bold too.
#[test]
fn undo_keeps_the_peers_deletions_and_formatting_inside_the_shape() {
    let net = Net::deck(2);
    let (left, right) = (net.peer(0), net.peer(1));
    let (slide_id, shape_id, story) = target(left);
    let before = text(left, &story);
    assert!(before.chars().count() > 4);
    let context = EditCtx::local("right");
    delete(left, false, &slide_id, &shape_id);
    net.delay(1);
    net.send(0);
    // The peer has not heard yet: it removes the first two units, breaks the
    // paragraph after what is now the second, and types a bold word.
    let hold = right.story(&story).unwrap();
    assert!(hold.length > 4);
    right.delete_text(&context, &story, 0, 2).unwrap();
    right.insert_paragraph_break(&context, &story, 2).unwrap();
    right
        .insert_text(
            &context,
            &story,
            3,
            "BOLD",
            &TextStyle {
                bold: Some(true),
                ..TextStyle::default()
            },
        )
        .unwrap();
    let expected = right.story(&story).unwrap();
    net.send(1);
    assert!(left.undo());
    net.assert_converged();
    let restored = left.story(&story).unwrap();
    assert_eq!(restored.plain_text(), expected.plain_text());
    assert_eq!(restored.paragraphs, expected.paragraphs);
}

fn draft(text: &str) -> ShapeDraft {
    ShapeDraft {
        name: text.to_owned(),
        rect: ShapeRect {
            x: 100_000,
            y: 200_000,
            width: 3_000_000,
            height: 800_000,
        },
        text: text.to_owned(),
        style: TextStyle::default(),
    }
}

/// A shape added in the session has no source file to read its position from.
#[test]
fn undo_keeps_a_peers_move_of_a_shape_with_no_source() {
    let source = DeckSession::open(DECK, 101).unwrap();
    let slide_id = source.snapshot().unwrap().slides[0].id.clone();
    let shape_id = source
        .add_text_box(&EditCtx::local("source"), &slide_id, &draft("Note"))
        .unwrap()
        .shape_id;
    let net = Net::new(source.encode_state_as_update_v1(), 2);
    let (left, right) = (net.peer(0), net.peer(1));
    delete(left, false, &slide_id, &shape_id);
    right
        .move_shape(
            &EditCtx::local("right"),
            &slide_id,
            &shape_id,
            1_234_567,
            2_345_678,
        )
        .unwrap();
    net.send(0);
    net.send(1);
    assert!(left.undo());
    net.assert_converged();
    assert_eq!(position(left, &shape_id), (1_234_567, 2_345_678));
}

/// The new shape's order entry lands in the slide's deleted order array.
#[test]
fn undo_keeps_a_shape_a_peer_added_to_the_deleted_slide() {
    let net = Net::deck(2);
    let (left, right) = (net.peer(0), net.peer(1));
    let (slide_id, shape_id, _) = target(left);
    delete(left, true, &slide_id, &shape_id);
    let added = right
        .add_text_box(&EditCtx::local("right"), &slide_id, &draft("Added"))
        .unwrap()
        .shape_id;
    net.send(0);
    net.send(1);
    assert!(left.undo());
    net.assert_converged();
    assert_eq!(position(left, &added), (100_000, 200_000));
}

/// The peer's typing reaches the undoing replica only after its Undo.
#[test]
fn typing_that_arrives_after_the_undo_lands_in_the_restored_shape() {
    for slide in [true, false] {
        let net = Net::deck(2);
        let (left, right) = (net.peer(0), net.peer(1));
        let (slide_id, shape_id, story) = target(left);
        let before = text(left, &story);
        let end = right.story(&story).unwrap().length - 1;
        delete(left, slide, &slide_id, &shape_id);
        assert!(left.undo());
        right
            .insert_text(
                &EditCtx::local("right"),
                &story,
                end,
                " LATE",
                &TextStyle::default(),
            )
            .unwrap();
        // The room hears the peer last.
        net.send(0);
        net.send(1);
        net.assert_converged();
        assert_eq!(text(left, &story), format!("{before} LATE"));
    }
}

/// Undo, Redo, Undo: the peer's text comes back each time, what arrived between too.
#[test]
fn redo_and_undo_again_keep_the_peers_text() {
    for undoer in [0, 1] {
        let net = Net::deck(2);
        let (left, right) = (net.peer(undoer), net.peer(1 - undoer));
        let (slide_id, shape_id, story) = target(left);
        let before = text(left, &story);
        let end = right.story(&story).unwrap().length - 1;
        delete(left, true, &slide_id, &shape_id);
        right
            .insert_text(
                &EditCtx::local("right"),
                &story,
                end,
                " REMOTE",
                &TextStyle::default(),
            )
            .unwrap();
        net.send(undoer);
        net.send(1 - undoer);
        for _ in 0..3 {
            assert!(left.undo());
            net.assert_converged();
            assert_eq!(text(right, &story), format!("{before} REMOTE"));
            assert!(left.redo());
            net.assert_converged();
            assert!(right.story(&story).is_err());
        }
        // Typing that reaches the undoer between its Redo and its next Undo: the peer had
        // not heard of the first Undo either, so it types in the first story.
        let net = Net::deck(2);
        let (left, right) = (net.peer(undoer), net.peer(1 - undoer));
        delete(left, true, &slide_id, &shape_id);
        assert!(left.undo());
        assert!(left.redo());
        right
            .insert_text(
                &EditCtx::local("right"),
                &story,
                end,
                " BETWEEN",
                &TextStyle::default(),
            )
            .unwrap();
        net.send(undoer);
        net.send(1 - undoer);
        assert!(left.undo());
        net.assert_converged();
        assert_eq!(text(left, &story), format!("{before} BETWEEN"));
    }
}

/// The peer moved the slide in the deck, or the shape in the paint order, while it was
/// deleted: printed, and both replicas agree. An order entry the Undo restores keeps its old
/// place, so the slide shows at the earlier of the two; a later move is not kept.
#[test]
fn undo_of_a_deleted_slide_the_peer_moved_converges() {
    for (later, undoer) in [(true, 0), (false, 0), (true, 1), (false, 1)] {
        let net = Net::deck(2);
        let (left, right) = (net.peer(undoer), net.peer(1 - undoer));
        let context = EditCtx::local("user");
        let ids = left.slide_ids().unwrap();
        let (moved, to_index) = if later { (&ids[1], 5) } else { (&ids[5], 1) };
        left.delete_slide(&context, moved).unwrap();
        left.add_undo_barrier();
        right.move_slide(&context, moved, to_index).unwrap();
        net.send(undoer);
        net.send(1 - undoer);
        assert!(left.undo());
        net.assert_converged();
        let shown = left.slide_ids().unwrap().iter().position(|id| id == moved);
        println!("slide moved to {to_index}, undoer {undoer}: shows at {shown:?}");
        assert_eq!(shown, Some(1));
    }
}

#[test]
fn undo_of_a_deleted_shape_the_peer_moved_in_paint_order_converges() {
    let net = Net::deck(2);
    let (left, right) = (net.peer(0), net.peer(1));
    let context = EditCtx::local("user");
    let snapshot = left.snapshot().unwrap();
    let slide = snapshot
        .slides
        .iter()
        .find(|slide| slide.shapes.len() >= 2)
        .unwrap();
    let shape = &slide.shapes[0];
    left.remove_shape(&context, &slide.id, &shape.id).unwrap();
    left.add_undo_barrier();
    right
        .bring_to_front(&context, &slide.id, &shape.id)
        .unwrap();
    net.send(0);
    net.send(1);
    assert!(left.undo());
    net.assert_converged();
    let read = left.snapshot().unwrap();
    let shapes = &read
        .slides
        .iter()
        .find(|s| s.id == slide.id)
        .unwrap()
        .shapes;
    let shown = shapes.iter().position(|s| s.id == shape.id);
    println!(
        "shape brought to front of {}: shows at {shown:?}",
        shapes.len()
    );
    assert_eq!(shown, Some(0));
}

/// Two peers typed; a third joins from the room's state after the Undo.
#[test]
fn two_peers_typing_both_come_back_for_a_joiner_too() {
    let net = Net::deck(3);
    let (left, right, third) = (net.peer(0), net.peer(1), net.peer(2));
    let (slide_id, shape_id, story) = target(left);
    let before = text(left, &story);
    let end = right.story(&story).unwrap().length - 1;
    delete(left, false, &slide_id, &shape_id);
    let context = EditCtx::local("peer");
    right
        .insert_text(&context, &story, 0, "AAA ", &TextStyle::default())
        .unwrap();
    third
        .insert_text(&context, &story, end, " ZZZ", &TextStyle::default())
        .unwrap();
    net.send(0);
    net.send(1);
    net.send(2);
    assert!(left.undo());
    net.assert_converged();
    assert_eq!(text(third, &story), format!("AAA {before} ZZZ"));
}

/// Both delete; each may undo: one copy of the content, nothing doubled.
#[test]
fn both_delete_and_both_undo_show_one_copy() {
    let net = Net::deck(2);
    let (left, right) = (net.peer(0), net.peer(1));
    let (slide_id, shape_id, story) = target(left);
    let before = text(left, &story);
    delete(left, true, &slide_id, &shape_id);
    delete(right, true, &slide_id, &shape_id);
    net.send(0);
    net.send(1);
    assert!(left.undo());
    net.settle();
    let _ = right.undo();
    net.assert_converged();
    assert_eq!(text(left, &story), before);
    let slides = left.snapshot().unwrap().slides;
    assert_eq!(
        slides.iter().filter(|slide| slide.id == slide_id).count(),
        1
    );
}

/// The peer's own Undo after the restore changes nothing it can no longer reach and
/// everyone still reads alike.
#[test]
fn the_peers_undo_after_the_restore_stays_converged() {
    let net = Net::deck(2);
    let (left, right) = (net.peer(0), net.peer(1));
    let (slide_id, shape_id, story) = target(left);
    let before = text(left, &story);
    let end = right.story(&story).unwrap().length - 1;
    delete(left, false, &slide_id, &shape_id);
    right
        .insert_text(
            &EditCtx::local("right"),
            &story,
            end,
            " REMOTE",
            &TextStyle::default(),
        )
        .unwrap();
    net.send(0);
    net.send(1);
    assert!(left.undo());
    net.settle();
    let _ = right.undo();
    net.assert_converged();
    assert_eq!(text(right, &story), format!("{before} REMOTE"));
}

/// The limit: a socket down while the peer's typing passes. The room's answer holds no
/// text, so the Undo restores without it, as before; everyone still reads alike.
#[test]
fn an_undoer_offline_during_the_typing_restores_without_it() {
    let net = Net::deck(2);
    let (left, right) = (net.peer(0), net.peer(1));
    let (slide_id, shape_id, story) = target(left);
    let before = text(left, &story);
    let end = right.story(&story).unwrap().length - 1;
    delete(left, false, &slide_id, &shape_id);
    net.delay(1);
    net.send(0);
    net.peers[0].online.set(false);
    right
        .insert_text(
            &EditCtx::local("right"),
            &story,
            end,
            " REMOTE",
            &TextStyle::default(),
        )
        .unwrap();
    net.send(1);
    net.reconnect(0);
    assert!(left.undo());
    net.assert_converged();
    assert_eq!(text(left, &story), before);
}

/// The peer's socket is down while it types and the shape goes: its step 2 reaches the undoer
/// as it sent it, and the older deletion that step 2 names again stays deleted.
#[test]
fn typing_in_a_reconnects_step_2_comes_back() {
    let net = Net::deck(2);
    let (left, right) = (net.peer(0), net.peer(1));
    let (slide_id, shape_id, story) = target(left);
    let context = EditCtx::local("right");
    right.delete_text(&context, &story, 0, 1).unwrap();
    let kept = text(right, &story);
    net.send(1);
    delete(left, false, &slide_id, &shape_id);
    net.delay(1);
    net.send(0);
    let end = right.story(&story).unwrap().length - 1;
    right
        .insert_text(&context, &story, end, " REMOTE", &TextStyle::default())
        .unwrap();
    net.step2(1);
    net.hear(1);
    assert!(left.undo());
    net.assert_converged();
    assert_eq!(
        text(left, &story),
        format!("{} REMOTE", kept.trim_end_matches('\n'))
    );
}

/// Saved bytes carry the peer's text.
#[test]
fn the_saved_file_holds_the_peers_text() {
    let net = Net::deck(2);
    let (left, right) = (net.peer(0), net.peer(1));
    let (slide_id, shape_id, story) = target(left);
    let end = right.story(&story).unwrap().length - 1;
    delete(left, true, &slide_id, &shape_id);
    right
        .insert_text(
            &EditCtx::local("right"),
            &story,
            end,
            " REMOTE",
            &TextStyle::default(),
        )
        .unwrap();
    net.send(0);
    net.send(1);
    assert!(left.undo());
    net.assert_converged();
    let saved = left.save().unwrap();
    let parsed = pptx_parse::parse_pptx(&saved).unwrap();
    let found = format!("{:?}", parsed.slides);
    assert!(found.contains("REMOTE"));
}

/// What the room stores and sends a joiner does not grow with what the fix keeps: the
/// kept text lives in the undoing engine alone until its Undo writes it.
#[test]
fn the_room_state_holds_nothing_of_the_kept_text_before_the_undo() {
    let net = Net::deck(2);
    let (left, right) = (net.peer(0), net.peer(1));
    let (slide_id, shape_id, story) = target(left);
    let end = right.story(&story).unwrap().length - 1;
    delete(left, true, &slide_id, &shape_id);
    net.delay(1);
    net.send(0);
    let before = net
        .room
        .transact()
        .encode_state_as_update_v1(&StateVector::default())
        .len();
    right
        .insert_text(
            &EditCtx::local("right"),
            &story,
            end,
            &"REMOTE".repeat(100),
            &TextStyle::default(),
        )
        .unwrap();
    net.send(1);
    let after = net
        .room
        .transact()
        .encode_state_as_update_v1(&StateVector::default());
    println!(
        "room state {before} B, after the peer's 600-char typing {} B",
        after.len()
    );
    assert!(after.len() < before + 64);
    assert!(!after.windows(6).any(|window| window == b"REMOTE"));
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

/// One edit on what `session` shows: deletes, typing, moves, a new shape, Undo and Redo.
fn random_step(rng: &mut Rng, session: &DeckSession) {
    let context = EditCtx::local("user");
    let snapshot = session.snapshot().unwrap();
    if snapshot.slides.is_empty() {
        session.undo();
        return;
    }
    let slide = &snapshot.slides[rng.below(snapshot.slides.len().min(4))];
    let shape = (!slide.shapes.is_empty()).then(|| &slide.shapes[rng.below(slide.shapes.len())]);
    match rng.below(12) {
        0 if snapshot.slides.len() > 1 => {
            session.delete_slide(&context, &slide.id).unwrap();
        }
        1 | 2 => {
            if let Some(shape) = shape {
                session
                    .remove_shape(&context, &slide.id, &shape.id)
                    .unwrap();
            }
        }
        3..=5 => {
            if let Some(story) = shape.and_then(|shape| shape.text_stories.first()) {
                let length = session.story(&story.id).unwrap().length;
                let at = rng.below(length as usize) as u32;
                session
                    .insert_text(&context, &story.id, at, "ab", &TextStyle::default())
                    .unwrap();
            }
        }
        6 => {
            if let Some(story) = shape.and_then(|shape| shape.text_stories.first()) {
                let length = session.story(&story.id).unwrap().length;
                if length > 3 {
                    session.delete_text(&context, &story.id, 0, 1).unwrap();
                }
            }
        }
        7 => {
            if let Some(shape) = shape {
                session
                    .move_shape(
                        &context,
                        &slide.id,
                        &shape.id,
                        rng.below(1000) as i64,
                        rng.below(1000) as i64,
                    )
                    .unwrap();
            }
        }
        8 => {
            session
                .add_text_box(&context, &slide.id, &draft("New"))
                .unwrap();
        }
        9 => {
            let to = rng.below(snapshot.slides.len()) as u32;
            session.move_slide(&context, &slide.id, to).unwrap();
        }
        10 => {
            session.undo();
        }
        _ => {
            session.redo();
        }
    }
    session.add_undo_barrier();
}

/// One random schedule through the room's path, three peers, updates sent in random order:
/// no engine refuses a frame or a read, everyone reads alike, and so does a joiner.
fn random_schedule(schedule: u64) {
    let mut rng = Rng(schedule.wrapping_mul(0x9E37_79B9_7F4A_7C15) | 1);
    let net = Net::deck(3);
    for _ in 0..8 {
        for _ in 0..4 {
            let peer = rng.below(3);
            random_step(&mut rng, net.peer(peer));
        }
        for _ in 0..3 {
            net.send(rng.below(3));
        }
    }
    net.assert_converged();
}

/// `CONCURRENT_DELETE_SCHEDULES=2000` runs the longer scan.
#[test]
fn random_schedules_through_the_room_converge() {
    let schedules: u64 =
        std::env::var("CONCURRENT_DELETE_SCHEDULES").map_or(25, |n| n.parse().unwrap());
    let mut failures = std::collections::BTreeMap::<String, Vec<u64>>::new();
    for schedule in 0..schedules {
        let outcome = std::panic::catch_unwind(|| random_schedule(schedule));
        if let Err(panic) = outcome {
            let error = panic
                .downcast_ref::<String>()
                .cloned()
                .unwrap_or_else(|| "panic".to_owned());
            let shape: String = error
                .chars()
                .filter(|c| !c.is_ascii_digit())
                .take(120)
                .collect();
            failures.entry(shape).or_default().push(schedule);
        }
    }
    let summary: Vec<_> = failures
        .iter()
        .map(|(shape, list)| format!("{} x {shape} (first {})", list.len(), list[0]))
        .collect();
    assert!(failures.is_empty(), "{summary:?}");
}
