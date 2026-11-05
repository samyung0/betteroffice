//! What the service reads after an Undo brought back deleted content with a peer's edits
//! (Q549): its seed, its reads and its export hold them, on real files, from the state a room
//! stored. The state before the Undo holds nothing of the kept text.

use docx_edit::{
    EditCtx as DocxCtx, EngineSession, FormatPolicy, Position, TableLocator, UndoSession,
};
use office_service::{Checkpoint, Determinism, Format, export, inspect, seed, sha256_hex};
use pptx_edit::{DeckSession, EditCtx, TextStyle};
use yrs::updates::decoder::Decode;
use yrs::{
    Any, ClientID, Doc, Map, OffsetKind, Options, Out, ReadTxn, StateVector, Text, Transact, Update,
};

const DECK: &[u8] = include_bytes!("../../../poc/fixtures/deck-50.pptx");
const PLAN: &[u8] = include_bytes!("../../../poc/fixtures/exchange-plan.docx");
const PPTX_TEXT: &str = "QFIVEFOURNINE";

const DETERMINISM: Determinism = Determinism {
    seed: "0000000000000000000000000000000000000000000000000000000000000000",
    now: "2026-11-05T00:00:00.000Z",
};

fn room_of(seed: &[u8]) -> Doc {
    let mut options = Options::with_client_id(ClientID::new(9000));
    options.offset_kind = OffsetKind::Utf16;
    options.cleanup_formatting = false;
    let doc = Doc::with_options(options);
    doc.transact_mut()
        .apply_update(Update::decode_v1(seed).unwrap())
        .unwrap();
    doc
}

fn apply(room: &Doc, update: &[u8]) {
    room.transact_mut()
        .apply_update(Update::decode_v1(update).unwrap())
        .unwrap();
}

fn state(room: &Doc) -> Vec<u8> {
    room.transact()
        .encode_state_as_update_v1(&StateVector::default())
}

/// Every string of a package's XML parts, as one string.
fn package_text(bytes: &[u8]) -> String {
    let parts = ooxml_opc::unzip_parts(bytes).unwrap();
    let mut text = String::new();
    for (name, part) in parts {
        if name.ends_with(".xml") {
            text.push_str(&String::from_utf8_lossy(&part));
        }
    }
    text
}

fn checkpoint<'a>(format: Format, sha: &'a str, state: &'a [u8]) -> Checkpoint<'a> {
    Checkpoint {
        format,
        schema_version: 1,
        base_sha256: sha,
        state,
    }
}

#[test]
fn a_deck_exports_the_peers_text_after_the_undo() {
    let seed_state = seed(Format::Pptx, DECK).unwrap();
    // The fix changes no seed (`shared/docx-pptx-storage.test.ts`'s golden).
    assert_eq!(
        sha256_hex(&seed_state),
        "02d15d15807bf30eb01e862fc707951d8ed5e94dcd6855da69d457d558951928"
    );
    let sha = sha256_hex(DECK);
    let room = room_of(&seed_state);
    let left = DeckSession::open_from_update_with_source(&seed_state, DECK, 501).unwrap();
    let right = DeckSession::open_from_update_with_source(&seed_state, DECK, 502).unwrap();
    let snapshot = left.snapshot().unwrap();
    let (slide, shape, story) = snapshot
        .slides
        .iter()
        .find_map(|slide| {
            slide
                .shapes
                .iter()
                .find(|shape| !shape.text_stories.is_empty())
                .map(|shape| (&slide.id, &shape.id, &shape.text_stories[0].id))
        })
        .unwrap();
    let context = EditCtx::local("user");
    let base_state = left.encode_state_vector_v1();
    left.delete_slide(&context, slide).unwrap();
    left.add_undo_barrier();
    let delete = left.encode_diff_v1(&base_state).unwrap();
    apply(&room, &delete);
    let (before_typing, grew_by_typing);
    {
        let end = right.story(story).unwrap().length - 1;
        let base = right.encode_state_vector_v1();
        right
            .insert_text(&context, story, end, PPTX_TEXT, &TextStyle::default())
            .unwrap();
        right
            .move_shape(&context, slide, shape, 1_234_567, 2_345_678)
            .unwrap();
        let typing = right.encode_diff_v1(&base).unwrap();
        before_typing = state(&room).len();
        apply(&room, &typing);
        let stored = state(&room);
        grew_by_typing = stored.len() - before_typing;
        assert!(
            !stored
                .windows(PPTX_TEXT.len())
                .any(|window| window == PPTX_TEXT.as_bytes())
        );
        left.apply_update_v1(&typing).unwrap();
    }
    println!(
        "deck-50: room state {before_typing} B; the peer's landed edits add {grew_by_typing} B"
    );

    let before_undo = left.encode_state_vector_v1();
    assert!(left.undo());
    let undo = left.encode_diff_v1(&before_undo).unwrap();
    println!("deck-50: the Undo's update {} B", undo.len());
    apply(&room, &undo);
    right.apply_update_v1(&undo).unwrap();
    assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());

    let stored = state(&room);
    println!(
        "deck-50: seed {} B, room state after the Undo {} B",
        seed_state.len(),
        stored.len()
    );
    let exported = export(DECK, checkpoint(Format::Pptx, &sha, &stored), DETERMINISM).unwrap();
    assert!(package_text(&exported).contains(PPTX_TEXT));
    let entries = format!(
        "{:?}",
        inspect(DECK, checkpoint(Format::Pptx, &sha, &stored)).unwrap()
    );
    assert!(entries.contains(PPTX_TEXT));
}

/// A table's first cell story.
fn first_cell(engine: &EngineSession) -> String {
    let txn = engine.doc().yrs_doc().transact();
    let Some(Out::YText(body)) = txn.get_map("stories").unwrap().get(&txn, "body") else {
        panic!("no body");
    };
    body.diff(&txn, yrs::types::text::YChange::identity)
        .into_iter()
        .find_map(|diff| match diff.insert {
            Out::YMap(table)
                if matches!(table.get(&txn, "_kind"), Some(Out::Any(Any::String(kind))) if &*kind == "table") =>
            {
                let Some(Out::Any(Any::Array(rows))) = table.get(&txn, "rows") else {
                    panic!("no rows");
                };
                let Any::Map(row) = &rows[0] else { panic!() };
                let Some(Any::Array(cells)) = row.get("cells") else {
                    panic!()
                };
                let Any::Map(cell) = &cells[0] else { panic!() };
                match cell.get("story") {
                    Some(Any::String(story)) => Some(story.to_string()),
                    _ => None,
                }
            }
            _ => None,
        })
        .unwrap()
}

#[test]
fn a_document_exports_the_peers_text_after_the_undo() {
    let seed_state = seed(Format::Docx, PLAN).unwrap();
    assert_eq!(
        sha256_hex(&seed_state),
        "5d5d37b03bbe545798fa8dd857fe63ac35c1091c6a2be961e01994bc14644c77"
    );
    let sha = sha256_hex(PLAN);
    let room = room_of(&seed_state);
    let open = |client| {
        let engine = EngineSession::new(client);
        engine.doc().apply_update_v1(&seed_state).unwrap();
        engine
    };
    let (left, right) = (open(51), open(52));
    let undo = UndoSession::new();
    undo.track(left.doc());
    let ctx = DocxCtx::local("Ada", "2026-11-05T00:00:00Z");
    let cell = first_cell(&left);

    let base = left.doc().encode_state_vector_v1();
    left.doc()
        .delete_table(&ctx, &TableLocator::new("body", 0))
        .unwrap();
    undo.add_undo_barrier();
    let delete = left.doc().encode_diff_v1(&base).unwrap();
    apply(&room, &delete);

    let base = right.doc().encode_state_vector_v1();
    right
        .doc()
        .insert_text(
            &ctx,
            Position::new(&cell, 0),
            "QPEERTEXT",
            FormatPolicy::Plain,
        )
        .unwrap();
    let typing = right.doc().encode_diff_v1(&base).unwrap();
    let before_typing = state(&room).len();
    apply(&room, &typing);
    let stored = state(&room);
    println!(
        "exchange-plan: room state {before_typing} B; the peer's landed text adds {} B",
        stored.len() - before_typing
    );
    assert!(!stored.windows(9).any(|window| window == b"QPEERTEXT"));
    left.doc()
        .applying_peer_update(|| left.doc().apply_update_v1(&typing).unwrap());

    let base = left.doc().encode_state_vector_v1();
    assert!(undo.undo());
    let restored = left.doc().encode_diff_v1(&base).unwrap();
    println!("exchange-plan: the Undo's update {} B", restored.len());
    apply(&room, &restored);
    right
        .doc()
        .applying_peer_update(|| right.doc().apply_update_v1(&restored).unwrap());

    let stored = state(&room);
    println!(
        "exchange-plan: seed {} B, room state after the Undo {} B",
        seed_state.len(),
        stored.len()
    );
    let exported = export(PLAN, checkpoint(Format::Docx, &sha, &stored), DETERMINISM).unwrap();
    assert!(package_text(&exported).contains("QPEERTEXT"));
}
