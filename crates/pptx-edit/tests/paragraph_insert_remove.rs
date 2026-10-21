//! A paragraph removed whole leaves its neighbours' ids and properties, and
//! one inserted before another takes a fresh id: the agent's commands name
//! paragraphs by id within one call.

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

/// Before the first paragraph and before the second, a new paragraph takes
/// a fresh id and the level and alignment of the paragraph it precedes,
/// which keeps its id; it reopens saved.
#[test]
fn a_paragraph_inserted_before_another_takes_a_fresh_id_and_its_properties() {
    let session = open();
    let second = "Office files,".encode_utf16().count() as u32 + 1;
    session
        .change_paragraph_level(&context(), STORY, second, second, 1, None)
        .unwrap();
    let before = paragraphs(&session);
    assert_eq!((before[0].level, before[1].level), (0, 1));
    let first = session
        .insert_paragraph_before(&context(), STORY, &before[0].id)
        .unwrap();
    let middle = session
        .insert_paragraph_before(&context(), STORY, &before[1].id)
        .unwrap();
    let after = paragraphs(&session);
    assert_eq!(after.len(), 4);
    assert_eq!(
        (after[0].id.as_str(), after[2].id.as_str()),
        (first.as_str(), middle.as_str())
    );
    assert_eq!(
        (after[1].clone(), after[3].clone()),
        (before[0].clone(), before[1].clone())
    );
    for (new, old) in [(&after[0], &before[0]), (&after[2], &before[1])] {
        assert_ne!(new.id, old.id);
        assert_eq!((new.level, &new.alignment), (old.level, &old.alignment));
        assert!(new.runs.iter().all(|run| run.text.is_empty()));
    }
    let reopened = DeckSession::open(&session.save().unwrap(), 8).unwrap();
    assert_eq!(
        reopened.story(STORY).unwrap().plain_text(),
        "\nOffice files,\n\nwithout the office."
    );
    assert!(matches!(
        session.insert_paragraph_before(&context(), STORY, "nope"),
        Err(EditError::ParagraphNotFound(id)) if id == "nope"
    ));
}
