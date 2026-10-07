//! Save effects on real documents (#6): an edit that only shifts entries
//! stores no effect (Epo, 2026-10-07, b2ce56a3), a real reorder still moves.

use docx_edit::{ChangeTarget, EditCtx, StoryRange};

use crate::common::sha256_hex;
use crate::types::{Checkpoint, Format, NetEffect, Operation};
use crate::{baseline, compare_baselines, save_effects, seed};

const STORIES: &[u8] = include_bytes!("../tests/fixtures/stories.docx");
const DECK: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

fn checkpoint<'a>(format: Format, sha: &'a str, state: &'a [u8]) -> Checkpoint<'a> {
    Checkpoint {
        format,
        schema_version: 1,
        base_sha256: sha,
        state,
    }
}

fn operations(effects: &[NetEffect]) -> Vec<Operation> {
    effects.iter().map(|effect| effect.operation).collect()
}

/// The states of `edit` applied in turn to the seeded DOCX session.
fn docx_states(edits: &[&dyn Fn(&docx_edit::EditingDoc)]) -> Vec<Vec<u8>> {
    let state = seed(Format::Docx, STORIES).expect("seeds");
    let session = crate::docx::DocxSession::open(STORIES, Some(&state)).expect("opens");
    let doc = session.engine().doc();
    edits
        .iter()
        .map(|edit| {
            edit(doc);
            doc.encode_state_as_update_v1()
        })
        .collect()
}

#[test]
fn accepting_a_deletion_before_objects_shifts_them_and_stores_no_effect() {
    let ctx = EditCtx::local("Reviewer", "2026-10-07T00:00:00Z");
    let states = docx_states(&[
        // "Intro " deleted as a suggestion: the paragraph reads without it,
        // the objects after it keep their offsets.
        &|doc| {
            doc.replace_range(&ctx.clone().suggesting(), StoryRange::new("body", 0, 6), "")
                .expect("suggests");
        },
        // Accepting it removes the deleted text: later objects shift by six.
        &|doc| {
            doc.accept_change(&ctx, &ChangeTarget::All)
                .expect("accepts");
        },
    ]);
    let sha = sha256_hex(STORIES);
    let suggested =
        baseline(STORIES, checkpoint(Format::Docx, &sha, &states[0])).expect("baseline");
    let accepted = checkpoint(Format::Docx, &sha, &states[1]);
    let all = compare_baselines(&suggested, &baseline(STORIES, accepted).expect("baseline"));
    assert!(!all.is_empty());
    assert!(all.iter().all(|effect| effect.operation == Operation::Move));
    assert_eq!(
        save_effects(STORIES, &suggested, accepted).expect("effects"),
        []
    );
}

#[test]
fn an_empty_paragraph_inserted_before_others_reports_only_itself() {
    let ctx = EditCtx::local("", "");
    let states = docx_states(&[&|_| {}, &|doc| {
        // At the end of the first paragraph (after its note reference): a
        // new empty paragraph; every later entry shifts by one.
        let end = doc
            .paragraphs("body")
            .expect("paragraphs")
            .first()
            .map(|paragraph| doc.paragraph_mark_position(&paragraph.para_id))
            .expect("a paragraph")
            .expect("its mark");
        doc.split_paragraph(&ctx, end, None).expect("splits");
    }]);
    let sha = sha256_hex(STORIES);
    let before = baseline(STORIES, checkpoint(Format::Docx, &sha, &states[0])).expect("baseline");
    let after = checkpoint(Format::Docx, &sha, &states[1]);
    let all = compare_baselines(&before, &baseline(STORIES, after).expect("baseline"));
    assert!(all.iter().any(|effect| effect.operation == Operation::Move));
    let saved = save_effects(STORIES, &before, after).expect("effects");
    assert!(!saved.is_empty());
    assert!(
        saved
            .iter()
            .all(|effect| effect.operation != Operation::Move)
    );
    let unchanged: Vec<NetEffect> = all
        .into_iter()
        .filter(|effect| effect.operation != Operation::Move)
        .collect();
    assert_eq!(saved, unchanged);
}

#[test]
fn a_shape_brought_to_the_front_still_moves() {
    let state = seed(Format::Pptx, DECK).expect("seeds");
    let deck =
        pptx_edit::DeckSession::open_from_update_with_source(&state, DECK, 7).expect("opens");
    let snapshot = crate::pptx::deck_json(&deck).expect("snapshot");
    let slides = snapshot
        .get("slides")
        .and_then(crate::js::J::as_arr)
        .expect("slides");
    let (slide, shapes) = slides
        .iter()
        .find_map(|slide| {
            let shapes = slide.get("shapes")?.as_arr()?;
            (shapes.len() >= 3).then_some((slide, shapes))
        })
        .expect("a slide with three shapes");
    let id = |value: &crate::js::J| {
        value
            .get("id")
            .and_then(crate::js::J::as_str)
            .expect("id")
            .to_owned()
    };
    deck.bring_to_front(&pptx_edit::EditCtx::local(""), &id(slide), &id(&shapes[0]))
        .expect("reorders");
    let edited = deck.encode_state_as_update_v1();
    let sha = sha256_hex(DECK);
    let before = baseline(DECK, checkpoint(Format::Pptx, &sha, &state)).expect("baseline");
    let saved =
        save_effects(DECK, &before, checkpoint(Format::Pptx, &sha, &edited)).expect("effects");
    assert!(
        saved
            .iter()
            .any(|effect| effect.operation == Operation::Move && effect.id == id(&shapes[0]))
    );
    assert!(
        operations(&saved)
            .iter()
            .all(|operation| *operation == Operation::Move)
    );
}
