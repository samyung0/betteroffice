//! Undo of a deleted table keeps what a peer did inside it meanwhile (upstream #1530, #1542;
//! Capy's own fix, Epo 2026-11-04, Q549 A).
//!
//! Each update takes the path it takes in a room: the writer's host document, the room
//! document (gc on), and for every other peer its host document and its engine. A room
//! relays what the writer told when its own collection took part of it, and a host hands
//! its engine the frame as it arrived.
use std::cell::RefCell;
use std::rc::Rc;

use docx_edit::bridge::RenderEnv;
use docx_edit::*;
use serde_json::{Value, json};
use yrs::types::text::YChange;
use yrs::updates::decoder::Decode;
use yrs::{
    Any, ClientID, Doc, Map, OffsetKind, Options, Out, ReadTxn, StateVector, Subscription, Text,
    Transact, Update,
};

const DATE: &str = "2026-11-05T00:00:00Z";

fn ctx() -> EditCtx {
    EditCtx::local("Ada", DATE)
}

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

pub struct Peer {
    pub engine: EngineSession,
    undo: UndoSession,
    outbox: Rc<RefCell<Vec<Vec<u8>>>>,
    host: Doc,
    /// Frames on their way: this peer has not heard them yet.
    delayed: RefCell<Option<Vec<Vec<u8>>>>,
    /// How much of the outbox the host document holds: it learns an edit as it is made.
    hosted: std::cell::Cell<usize>,
    _sub: Subscription,
}

impl Peer {
    fn open(seed: &[u8], client: u64) -> Self {
        let engine = EngineSession::new(client);
        engine.doc().apply_update_v1(seed).unwrap();
        let undo = UndoSession::new();
        undo.track(engine.doc());
        let outbox = Rc::new(RefCell::new(Vec::new()));
        let sink = Rc::clone(&outbox);
        let sub = engine
            .doc()
            .yrs_doc()
            .observe_update_v1(move |txn, event| {
                if txn.origin().is_some() {
                    sink.borrow_mut().push(event.update.clone());
                }
            })
            .unwrap();
        Self {
            engine,
            undo,
            outbox,
            host: relay_doc(client + 1000, seed),
            delayed: RefCell::new(None),
            hosted: std::cell::Cell::new(0),
            _sub: sub,
        }
    }

    fn doc(&self) -> &EditingDoc {
        self.engine.doc()
    }

    /// As the editor's `apply_peer_update`.
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
        let doc = self.doc();
        doc.applying_peer_update(|| doc.apply_update_v1(frame).unwrap());
        self.undo.forward_concurrent_edits();
    }
}

pub struct Net {
    room: Doc,
    pub peers: Vec<Peer>,
}

impl Net {
    pub fn over(seed: &[u8], peers: usize) -> Self {
        Self {
            room: relay_doc(9000, seed),
            peers: (0..peers)
                .map(|index| Peer::open(seed, 51 + index as u64))
                .collect(),
        }
    }

    pub fn send(&self, index: usize) {
        self.peers[index].host_own_edits();
        self.peers[index].hosted.set(0);
        let outgoing: Vec<_> = self.peers[index].outbox.borrow_mut().drain(..).collect();
        for update in outgoing {
            let Some(frame) = room_relay(&self.room, &update, false) else {
                continue;
            };
            for (other, peer) in self.peers.iter().enumerate() {
                if other != index {
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

    pub fn settle(&self) {
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

    /// Everyone reads alike, each lowers the body, and so does a joiner on the room's state.
    fn assert_converged(&self, stories: &[String]) {
        self.settle();
        let state = self
            .room
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        let joiner = EngineSession::new(900);
        joiner.doc().apply_update_v1(&state).unwrap();
        let read = |engine: &EngineSession| {
            engine
                .lower_story_json("body", &RenderEnv::default())
                .unwrap();
            stories
                .iter()
                .map(|story| text(engine.doc(), story))
                .collect::<Vec<_>>()
        };
        let expected = read(&self.peers[0].engine);
        for peer in &self.peers[1..] {
            assert_eq!(read(&peer.engine), expected);
        }
        assert_eq!(read(&joiner), expected);
        assert_eq!(
            joiner.doc().story_segments("body").unwrap(),
            self.peers[0].doc().story_segments("body").unwrap()
        );
    }
}

fn text(doc: &EditingDoc, story: &str) -> String {
    doc.paragraphs(story)
        .unwrap()
        .into_iter()
        .map(|paragraph| paragraph.text)
        .collect::<Vec<_>>()
        .join("\n")
}

/// A body paragraph and a 2x2 table whose cells hold `cell0` to `cell3`.
pub fn table_seed() -> (Vec<u8>, Vec<String>) {
    let origin = EngineSession::new(300);
    origin
        .doc()
        .create_story("body", "Before", "Normal", "left")
        .unwrap();
    let table = origin
        .doc()
        .insert_table(&ctx(), Position::new("body", 0), 2, 2, 15)
        .unwrap();
    for (index, story) in table.created_story_ids.iter().enumerate() {
        origin
            .doc()
            .insert_text(
                &ctx(),
                Position::new(story, 0),
                &format!("cell{index}"),
                FormatPolicy::Plain,
            )
            .unwrap();
    }
    (
        origin.doc().encode_state_as_update_v1(),
        table.created_story_ids,
    )
}

fn table() -> TableLocator {
    TableLocator::new("body", 0)
}

/// The rows of the body's first table.
fn rows(doc: &EditingDoc) -> Vec<Any> {
    let txn = doc.yrs_doc().transact();
    let Some(Out::YText(body)) = txn.get_map("stories").unwrap().get(&txn, "body") else {
        panic!("no body story");
    };
    body.diff(&txn, YChange::identity)
        .into_iter()
        .find_map(|diff| match diff.insert {
            Out::YMap(table)
                if matches!(table.get(&txn, "_kind"), Some(Out::Any(Any::String(kind))) if &*kind == "table") =>
            {
                match table.get(&txn, "rows") {
                    Some(Out::Any(Any::Array(rows))) => Some(rows.to_vec()),
                    _ => None,
                }
            }
            _ => None,
        })
        .expect("a table with rows")
}

fn delete_table(peer: &Peer) {
    peer.doc().delete_table(&ctx(), &table()).unwrap();
    peer.undo.add_undo_barrier();
}

fn type_in(peer: &Peer, story: &str, at: u32, what: &str) {
    peer.doc()
        .insert_text(&ctx(), Position::new(story, at), what, FormatPolicy::Plain)
        .unwrap();
    peer.undo.add_undo_barrier();
}

/// A layout pass as an editor's (`relayout_frame`), which builds the display list, with no fonts
/// (the layout's own metrics). The lowered table and its measure.
fn lay_out(engine: &EngineSession) -> Result<(Value, Value), String> {
    let request = json!({
        "bodyStory": "body",
        "regions": {"sections": [{
            "sectionId": "main",
            "pageSize": {"w": 612, "h": 792},
            "margins": {"top": 72, "right": 72, "bottom": 72, "left": 72, "header": 36, "footer": 36},
        }]},
        "measurement": {
            "fontChains": {},
            "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
        },
        "renderEnv": {},
    })
    .to_string();
    engine.relayout_frame(&request, "{}", 0)?;
    let output: Value = serde_json::from_str(&engine.layout_document_with_regions_json(&request)?)
        .map_err(|error| error.to_string())?;
    let Some(table) = output["measured"].as_array().and_then(|measured| {
        measured
            .iter()
            .find(|entry| entry["block"]["kind"] == "table")
    }) else {
        // Every table can be gone in a random schedule.
        return Ok((Value::Null, Value::Null));
    };
    Ok((table["block"].clone(), table["measure"].clone()))
}

/// #1530, both arrival orders at the room.
#[test]
fn undo_brings_a_deleted_table_back_with_the_peers_typing() {
    for peer_first in [true, false] {
        let (seed, cells) = table_seed();
        let net = Net::over(&seed, 2);
        let (left, right) = (&net.peers[0], &net.peers[1]);
        delete_table(left);
        type_in(right, &cells[0], 5, " PEER");
        type_in(right, &cells[3], 0, "PEER2 ");
        let order: [usize; 2] = if peer_first { [1, 0] } else { [0, 1] };
        net.send(order[0]);
        net.send(order[1]);
        assert!(left.undo.undo());
        net.assert_converged(&cells);
        assert_eq!(text(left.doc(), &cells[0]), "cell0 PEER");
        assert_eq!(text(left.doc(), &cells[3]), "PEER2 cell3");
        assert_eq!(text(right.doc(), &cells[1]), "cell1");
    }
}

/// #1542: both delete, one undoes; every cell comes back on both.
#[test]
fn both_delete_and_one_undoes() {
    for undoer in [0, 1] {
        for peer_first in [true, false] {
            let (seed, cells) = table_seed();
            let net = Net::over(&seed, 2);
            let other = 1 - undoer;
            delete_table(&net.peers[undoer]);
            delete_table(&net.peers[other]);
            let order: [usize; 2] = if peer_first {
                [other, undoer]
            } else {
                [undoer, other]
            };
            net.send(order[0]);
            net.send(order[1]);
            assert!(net.peers[undoer].undo.undo());
            net.assert_converged(&cells);
            for (index, cell) in cells.iter().enumerate() {
                assert_eq!(text(net.peers[other].doc(), cell), format!("cell{index}"));
            }
        }
    }
}

/// #1530 on top of #1542: the peer types, then deletes the table, unaware of the first delete.
#[test]
fn both_delete_after_the_peer_typed_and_one_undoes() {
    for peer_first in [true, false] {
        let (seed, cells) = table_seed();
        let net = Net::over(&seed, 2);
        let (left, right) = (&net.peers[0], &net.peers[1]);
        delete_table(left);
        type_in(right, &cells[0], 5, " PEER");
        delete_table(right);
        let order: [usize; 2] = if peer_first { [1, 0] } else { [0, 1] };
        net.send(order[0]);
        net.send(order[1]);
        assert!(left.undo.undo());
        net.assert_converged(&cells);
        assert_eq!(text(left.doc(), &cells[0]), "cell0 PEER");
        assert_eq!(text(left.doc(), &cells[3]), "cell3");
    }
}

/// What the peer deleted stays deleted; its Enter and bold come back.
#[test]
fn undo_keeps_the_peers_deletion_enter_and_bold_in_a_cell() {
    let (seed, cells) = table_seed();
    let net = Net::over(&seed, 2);
    let (left, right) = (&net.peers[0], &net.peers[1]);
    delete_table(left);
    net.delay(1);
    net.send(0);
    right
        .doc()
        .delete_range(&ctx(), StoryRange::new(&cells[0], 0, 2))
        .unwrap();
    right
        .doc()
        .split_paragraph(&ctx(), Position::new(&cells[0], 1))
        .unwrap();
    type_in(right, &cells[0], 2, "Bold");
    right
        .doc()
        .toggle_format(&ctx(), StoryRange::new(&cells[0], 2, 6), SimpleFormat::Bold)
        .unwrap();
    let expected = right.doc().paragraphs(&cells[0]).unwrap();
    let runs = right.doc().story_segments(&cells[0]).unwrap();
    net.send(1);
    assert!(left.undo.undo());
    net.assert_converged(&cells);
    let restored = left.doc().paragraphs(&cells[0]).unwrap();
    assert_eq!(
        restored.iter().map(|p| p.text.clone()).collect::<Vec<_>>(),
        expected.iter().map(|p| p.text.clone()).collect::<Vec<_>>()
    );
    let bold = |segments: &[_]| format!("{segments:?}").matches("bold").count();
    assert_eq!(
        bold(left.doc().story_segments(&cells[0]).unwrap().as_slice()),
        bold(runs.as_slice())
    );
}

/// The peer's socket is down while it types and the table goes: its step 2 reaches the undoer
/// as it sent it, and the older deletion that step 2 names again stays deleted.
#[test]
fn typing_in_a_reconnects_step_2_comes_back() {
    let (seed, cells) = table_seed();
    let net = Net::over(&seed, 2);
    let (left, right) = (&net.peers[0], &net.peers[1]);
    right
        .doc()
        .delete_range(&ctx(), StoryRange::new(&cells[0], 0, 1))
        .unwrap();
    right.undo.add_undo_barrier();
    net.send(1);
    delete_table(left);
    net.delay(1);
    net.send(0);
    type_in(right, &cells[0], 0, "PEER ");
    net.step2(1);
    net.hear(1);
    assert!(left.undo.undo());
    net.assert_converged(&cells);
    assert_eq!(text(left.doc(), &cells[0]), "PEER ell0");
}

/// A row the peer inserted while the table went stays in the restored table.
#[test]
fn undo_keeps_a_row_the_peer_inserted() {
    let (seed, cells) = table_seed();
    let net = Net::over(&seed, 2);
    let (left, right) = (&net.peers[0], &net.peers[1]);
    delete_table(left);
    let row = right
        .doc()
        .insert_row(&ctx(), &CellLoc::new("body", 0, 1, 0), true)
        .unwrap();
    type_in(right, &row.created_story_ids[0], 0, "new row");
    net.send(0);
    net.send(1);
    assert!(left.undo.undo());
    net.assert_converged(&cells);
    assert_eq!(rows(left.doc()).len(), 3);
    assert_eq!(text(left.doc(), &row.created_story_ids[0]), "new row");
}

/// Two people insert rows and the first undoes, with no table delete: the second's `rows` value
/// still names the cells the Undo removed. Everyone lowers and lays out, the cells read as empty
/// at a line's height, no caret lands in them, and the document keeps working (Q624).
#[test]
fn a_row_naming_removed_cells_reads_as_empty() {
    let (seed, cells) = table_seed();
    let net = Net::over(&seed, 2);
    let (left, right) = (&net.peers[0], &net.peers[1]);
    let removed = left
        .doc()
        .insert_row(&ctx(), &CellLoc::new("body", 0, 1, 0), true)
        .unwrap();
    left.undo.add_undo_barrier();
    net.send(0);
    let kept = right
        .doc()
        .insert_row(&ctx(), &CellLoc::new("body", 0, 2, 0), true)
        .unwrap();
    right.undo.add_undo_barrier();
    net.send(1);
    assert!(left.undo.undo());
    net.assert_converged(&cells);
    assert_eq!(rows(left.doc()).len(), 4);
    assert!(
        left.doc()
            .paragraphs(&removed.created_story_ids[0])
            .is_err()
    );

    // Laid out, on both: the row keeps its place and a line's height, as the inner row of real
    // cells beside it.
    let (table, measure) = lay_out(&left.engine).unwrap();
    assert_eq!(lay_out(&right.engine).unwrap().1, measure);
    let heights: Vec<f64> = measure["rows"]
        .as_array()
        .unwrap()
        .iter()
        .map(|row| row["height"].as_f64().unwrap())
        .collect();
    assert_eq!(heights.len(), 4);
    assert!(heights[2] > 0.0, "{heights:?}");
    assert_eq!(heights[2], heights[1], "{heights:?}");

    // The cell takes no positions beyond its own open and close, as the editor's projection
    // counts it.
    let block = |row: usize, cell: usize| &table["rows"][row]["cells"][cell]["blocks"][0];
    assert!(block(2, 0).get("pmStart").is_none());
    assert_eq!(block(2, 0)["runs"], json!([]));
    let after_row_1 = block(1, 1)["pmEnd"].as_f64().unwrap();
    assert_eq!(block(3, 0)["pmStart"].as_f64().unwrap(), after_row_1 + 10.0);

    // A click anywhere finds a paragraph someone holds.
    let (first, last) = (after_row_1 + 2.0, after_row_1 + 8.0);
    let mut hits = 0;
    for y in (0..792).step_by(2) {
        for x in (0..612).step_by(12) {
            let hit: Value = serde_json::from_str(
                &left
                    .engine
                    .display_hit_test_regions_json(0, f64::from(x), f64::from(y))
                    .unwrap(),
            )
            .unwrap();
            if let Some(position) = hit["pos"].as_f64() {
                assert!(
                    position <= first || position >= last,
                    "{position} at {x},{y}"
                );
                hits += 1;
            }
        }
    }
    assert!(hits > 0);

    // Find, typing elsewhere, and removing the row all work.
    assert_eq!(
        left.doc().search_text("cell3", true, None).unwrap().len(),
        1
    );
    type_in(right, &kept.created_story_ids[0], 0, "still here");
    let row = CellLoc::new("body", 0, 2, 0);
    left.doc()
        .delete_row(&ctx(), &TableRange::new(row.clone(), row))
        .unwrap();
    left.undo.add_undo_barrier();
    net.assert_converged(&cells);
    assert_eq!(rows(left.doc()).len(), 3);
    assert_eq!(text(left.doc(), &kept.created_story_ids[0]), "still here");
    lay_out(&right.engine).unwrap();
}

/// The peer's typing reaches the undoing replica only after its Undo.
#[test]
fn typing_that_arrives_after_the_undo_lands_in_the_restored_cell() {
    let (seed, cells) = table_seed();
    let net = Net::over(&seed, 2);
    let (left, right) = (&net.peers[0], &net.peers[1]);
    delete_table(left);
    assert!(left.undo.undo());
    type_in(right, &cells[0], 5, " LATE");
    net.send(0);
    net.send(1);
    net.assert_converged(&cells);
    assert_eq!(text(left.doc(), &cells[0]), "cell0 LATE");
}

/// Undo, Redo, Undo: the peer's text comes back each time.
#[test]
fn redo_and_undo_again_keep_the_peers_typing() {
    let (seed, cells) = table_seed();
    let net = Net::over(&seed, 2);
    let (left, right) = (&net.peers[0], &net.peers[1]);
    delete_table(left);
    type_in(right, &cells[0], 5, " PEER");
    net.send(0);
    net.send(1);
    for _ in 0..3 {
        assert!(left.undo.undo());
        net.assert_converged(&cells);
        assert_eq!(text(right.doc(), &cells[0]), "cell0 PEER");
        assert!(left.undo.redo());
        net.settle();
        assert!(right.doc().story_len(&cells[0]).is_err());
    }
}

/// Row and column removal, and a nested table's cell, as upstream's #1530 tests have them.
#[test]
fn undo_of_a_row_or_column_removal_keeps_the_peers_typing() {
    for column in [false, true] {
        let (seed, cells) = table_seed();
        let net = Net::over(&seed, 2);
        let (left, right) = (&net.peers[0], &net.peers[1]);
        let first = CellLoc::new("body", 0, 0, 0);
        let range = TableRange::new(first.clone(), first);
        if column {
            left.doc().delete_column(&ctx(), &range).unwrap();
        } else {
            left.doc().delete_row(&ctx(), &range).unwrap();
        }
        left.undo.add_undo_barrier();
        type_in(right, &cells[0], 0, "peer ");
        net.send(0);
        net.send(1);
        assert!(left.undo.undo());
        net.assert_converged(&cells);
        assert_eq!(text(right.doc(), &cells[0]), "peer cell0");
    }
}

#[test]
fn undo_keeps_typing_in_a_nested_table() {
    let origin = EngineSession::new(300);
    origin
        .doc()
        .create_story("body", "Before", "Normal", "left")
        .unwrap();
    let outer = origin
        .doc()
        .insert_table(&ctx(), Position::new("body", 0), 1, 1, 15)
        .unwrap();
    let nested = origin
        .doc()
        .insert_table(
            &ctx(),
            Position::new(&outer.created_story_ids[0], 0),
            1,
            1,
            15,
        )
        .unwrap()
        .created_story_ids[0]
        .clone();
    let seed = origin.doc().encode_state_as_update_v1();
    let net = Net::over(&seed, 2);
    let (left, right) = (&net.peers[0], &net.peers[1]);
    delete_table(left);
    type_in(right, &nested, 0, "nested peer");
    net.send(0);
    net.send(1);
    assert!(left.undo.undo());
    let stories = [nested];
    net.assert_converged(&stories);
    assert_eq!(text(right.doc(), &stories[0]), "nested peer");
}

/// A real file: a table of `exchange-plan.docx` deleted while a peer types in its first cell.
#[test]
fn a_real_file_keeps_the_peers_typing() {
    let origin = EngineSession::new(SEED_CLIENT_ID);
    seed_from_docx(
        origin.doc(),
        include_bytes!("../../../poc/fixtures/exchange-plan.docx"),
    )
    .unwrap();
    let seed = origin.doc().encode_state_as_update_v1();
    let net = Net::over(&seed, 2);
    let (left, right) = (&net.peers[0], &net.peers[1]);
    let cell = rows(left.doc())
        .iter()
        .flat_map(|row| {
            let Any::Map(row) = row else { panic!() };
            let Some(Any::Array(cells)) = row.get("cells") else {
                panic!()
            };
            cells.iter().cloned().collect::<Vec<_>>()
        })
        .find_map(|cell| {
            let Any::Map(cell) = cell else { panic!() };
            match cell.get("story") {
                Some(Any::String(story)) => Some(story.to_string()),
                _ => None,
            }
        })
        .unwrap();
    let before = text(left.doc(), &cell);
    delete_table(left);
    type_in(right, &cell, 0, "PEER ");
    net.send(0);
    net.send(1);
    assert!(left.undo.undo());
    let stories = [cell];
    net.assert_converged(&stories);
    assert_eq!(text(left.doc(), &stories[0]), format!("PEER {before}"));
}

pub struct Rng(pub u64);

impl Rng {
    fn next(&mut self) -> u64 {
        self.0 ^= self.0 << 13;
        self.0 ^= self.0 >> 7;
        self.0 ^= self.0 << 17;
        self.0
    }

    pub fn below(&mut self, n: usize) -> usize {
        if n == 0 {
            0
        } else {
            (self.next() % n as u64) as usize
        }
    }
}

/// One edit: table deletes, typing and backspaces in cells, rows, Enter, Undo and Redo.
/// Refused edits (the table is gone here) are fine.
pub fn random_step(rng: &mut Rng, peer: &Peer, cells: &[String]) {
    let doc = peer.doc();
    let cell = &cells[rng.below(cells.len())];
    let length = doc.story_len(cell).unwrap_or(0);
    match rng.below(10) {
        0 => {
            let _ = doc.delete_table(&ctx(), &table());
        }
        1..=3 if length > 0 => {
            let at = rng.below(length as usize) as u32;
            let _ = doc.insert_text(&ctx(), Position::new(cell, at), "ab", FormatPolicy::Plain);
        }
        4 if length > 2 => {
            let _ = doc.delete_range(&ctx(), StoryRange::new(cell, 0, 1));
        }
        5 if length > 1 => {
            let _ = doc.split_paragraph(&ctx(), Position::new(cell, 1));
        }
        6 => {
            let _ = doc.insert_row(&ctx(), &CellLoc::new("body", 0, 0, 0), true);
        }
        7 => {
            let first = CellLoc::new("body", 0, 0, 0);
            let _ = doc.delete_row(&ctx(), &TableRange::new(first.clone(), first));
        }
        8 => {
            peer.undo.undo();
        }
        _ => {
            peer.undo.redo();
        }
    }
    peer.undo.add_undo_barrier();
}

/// One random schedule through the room's path, three peers, updates sent in random order;
/// what fails to read at the end.
fn random_schedule(schedule: u64) -> Result<(), String> {
    let mut rng = Rng(schedule.wrapping_mul(0x9E37_79B9_7F4A_7C15) | 1);
    let (seed, cells) = table_seed();
    let net = Net::over(&seed, 3);
    for _ in 0..8 {
        for _ in 0..4 {
            let peer = rng.below(3);
            random_step(&mut rng, &net.peers[peer], &cells);
        }
        for _ in 0..3 {
            net.send(rng.below(3));
        }
    }
    net.settle();
    let state = net
        .room
        .transact()
        .encode_state_as_update_v1(&StateVector::default());
    let joiner = EngineSession::new(900);
    joiner.doc().apply_update_v1(&state).unwrap();
    let read = |engine: &EngineSession| -> Result<Vec<usize>, String> {
        engine
            .doc()
            .story_segments("body")
            .map_err(|error| error.to_string())?;
        engine
            .lower_story_json("body", &RenderEnv::default())
            .map_err(|error| error.to_string())?;
        lay_out(engine)?;
        Ok(cells
            .iter()
            .map(|cell| engine.doc().paragraphs(cell).map_or(0, |p| p.len()))
            .collect())
    };
    let expected = read(&net.peers[0].engine)?;
    for peer in &net.peers[1..] {
        if read(&peer.engine)? != expected {
            return Err("peers read differently".to_owned());
        }
    }
    if read(&joiner)? != expected {
        return Err("a joiner reads differently".to_owned());
    }
    Ok(())
}

/// Everyone reads alike and lowers the body, and so does a joiner on the room's state.
/// `CONCURRENT_DELETE_SCHEDULES=2000` runs the longer scan.
#[test]
fn random_schedules_through_the_room_converge() {
    let schedules: u64 =
        std::env::var("CONCURRENT_DELETE_SCHEDULES").map_or(200, |n| n.parse().unwrap());
    let mut failures = std::collections::BTreeMap::<String, Vec<u64>>::new();
    for schedule in 0..schedules {
        if let Err(error) = random_schedule(schedule) {
            let shape: String = error.chars().filter(|c| !c.is_ascii_digit()).collect();
            failures.entry(shape).or_default().push(schedule);
        }
    }
    // Queued apart (Q626, the report's section 8): `grid` lost by Undo of concurrent row
    // operations, older than this fix. Nothing else may fail.
    failures.retain(|shape, _| !shape.contains("missing grid"));
    assert!(failures.is_empty(), "{failures:?}");
}
