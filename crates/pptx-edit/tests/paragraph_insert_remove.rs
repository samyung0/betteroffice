//! A paragraph removed whole leaves its neighbours' ids and properties, and
//! one inserted at a story's start takes a fresh id: the agent's commands
//! name paragraphs by id within one call.

use pptx_edit::{DeckSession, EditCtx, EditError, ParagraphSnapshot};

const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
/// "Office files," then "without the office.".
const STORY: &str = "story:slide:0:256:shape:2:0";

fn context() -> EditCtx {
    EditCtx::local("test")
}

/// The demo deck with its first story's second paragraph centred.
fn open() -> DeckSession {
    let session = DeckSession::open(FIXTURE, 7).unwrap();
    let second = "Office files,".encode_utf16().count() as u32 + 1;
    session
        .set_paragraph_alignment(&context(), STORY, second, second, Some("ctr"))
        .unwrap();
    session
}

fn paragraphs(session: &DeckSession) -> Vec<ParagraphSnapshot> {
    session.story(STORY).unwrap().paragraphs
}

#[test]
fn a_removed_paragraph_leaves_its_neighbours_as_they_were() {
    let session = open();
    let mut expected = paragraphs(&session);
    let first = expected.remove(0);
    session
        .remove_paragraph(&context(), STORY, &first.id)
        .unwrap();
    assert_eq!(paragraphs(&session), expected);
    assert_eq!(
        session.story(STORY).unwrap().plain_text(),
        "without the office."
    );
    let reopened = DeckSession::open(&session.save().unwrap(), 8).unwrap();
    assert_eq!(
        reopened.story(STORY).unwrap().plain_text(),
        "without the office."
    );
}

#[test]
fn the_last_paragraph_and_an_unknown_one_are_refused() {
    let session = open();
    let before = paragraphs(&session);
    let last = before[1].id.clone();
    assert!(matches!(
        session.remove_paragraph(&context(), STORY, &last),
        Err(EditError::LastParagraph(id)) if id == last
    ));
    assert!(matches!(
        session.remove_paragraph(&context(), STORY, "nope"),
        Err(EditError::ParagraphNotFound(id)) if id == "nope"
    ));
    assert_eq!(paragraphs(&session), before);
}

#[test]
fn a_paragraph_inserted_at_the_start_takes_a_fresh_id_and_the_first_ones_properties() {
    let session = open();
    let before = paragraphs(&session);
    session
        .remove_paragraph(&context(), STORY, &before[0].id)
        .unwrap();
    // The centred paragraph is first now.
    let id = session
        .insert_paragraph_at_start(&context(), STORY)
        .unwrap();
    let after = paragraphs(&session);
    assert_eq!(after.len(), 2);
    assert_ne!(id, before[1].id);
    assert_eq!(after[0].id, id);
    assert_eq!(after[0].alignment.as_deref(), Some("ctr"));
    assert!(after[0].runs.iter().all(|run| run.text.is_empty()));
    assert_eq!(after[1], before[1]);
    let reopened = DeckSession::open(&session.save().unwrap(), 8).unwrap();
    assert_eq!(
        reopened.story(STORY).unwrap().plain_text(),
        "\nwithout the office."
    );
}
