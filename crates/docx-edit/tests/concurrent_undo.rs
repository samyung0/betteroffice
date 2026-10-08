//! Undo against another person's later edit of the same paragraph property
//! (Epo 2026-10-08 narrow Undo rule): Undo never brings back a value another
//! person later removed.

use docx_edit::{EditCtx, EditingDoc, ParaSelector, UndoCaptureMode};
use yrs::Any;

fn ctx() -> EditCtx {
    EditCtx::local("test", "2026-01-01T00:00:00Z")
}

fn sync(from: &EditingDoc, to: &EditingDoc) {
    let diff = from.encode_diff_v1(&to.encode_state_vector_v1()).unwrap();
    to.apply_update_v1(&diff).unwrap();
}

fn indent(doc: &EditingDoc) -> Option<Any> {
    doc.paragraphs("body").unwrap()[0]
        .properties
        .get("indentLeft")
        .cloned()
}

/// Alice presses Increase Indent twice; Bob presses it once, then Decrease
/// Indent back to the margin, which removes the property. Alice's Ctrl+Z
/// leaves the paragraph at the margin, as Yjs does: the removal is Bob's
/// later, deliberate edit, not a value concurrent with Alice's.
#[test]
fn undo_keeps_a_paragraph_another_person_outdented_to_the_margin() {
    let seed = EditingDoc::new(500);
    let paragraph = seed
        .create_story("body", "indented", "Normal", "left")
        .unwrap();
    let state = seed.encode_state_as_update_v1();
    let (alice, bob) = (EditingDoc::new(601), EditingDoc::new(602));
    alice.apply_update_v1(&state).unwrap();
    bob.apply_update_v1(&state).unwrap();
    let mut undo = alice.undo_manager();
    undo.set_capture_mode(UndoCaptureMode::Manual);
    let selector = ParaSelector::One(paragraph);

    alice.increase_indent(&ctx(), &selector, None).unwrap();
    undo.add_undo_barrier();
    alice.increase_indent(&ctx(), &selector, None).unwrap();
    undo.add_undo_barrier();
    sync(&alice, &bob);
    bob.increase_indent(&ctx(), &selector, None).unwrap();
    while indent(&bob).is_some() {
        bob.decrease_indent(&ctx(), &selector, None).unwrap();
    }
    sync(&bob, &alice);
    assert_eq!(indent(&alice), None);

    undo.undo();
    sync(&alice, &bob);
    assert_eq!(indent(&alice), None);
    assert_eq!(indent(&bob), None);
}

/// The narrow Undo rule skips only a concurrent value that lost: the
/// paragraph is indented once; Alice and Bob press Increase Indent at the
/// same time and Bob's wins; Bob then outdents to the margin. Bob's winning
/// value and his later ones merge into one removed block that keeps the
/// winner's origin. Alice's Ctrl+Z leaves the paragraph at the margin, as
/// Yjs does.
#[test]
fn undo_keeps_a_paragraph_outdented_after_a_concurrent_indent_won() {
    let seed = EditingDoc::new(500);
    let paragraph = seed
        .create_story("body", "indented", "Normal", "left")
        .unwrap();
    let selector = ParaSelector::One(paragraph);
    seed.increase_indent(&ctx(), &selector, None).unwrap();
    let state = seed.encode_state_as_update_v1();
    let (alice, bob) = (EditingDoc::new(601), EditingDoc::new(602));
    alice.apply_update_v1(&state).unwrap();
    bob.apply_update_v1(&state).unwrap();
    let mut undo = alice.undo_manager();
    undo.set_capture_mode(UndoCaptureMode::Manual);

    alice.increase_indent(&ctx(), &selector, None).unwrap();
    undo.add_undo_barrier();
    bob.increase_indent(&ctx(), &selector, None).unwrap();
    sync(&alice, &bob);
    sync(&bob, &alice);
    assert_eq!(indent(&alice), Some(Any::Number(1440.0)));
    while indent(&bob).is_some() {
        bob.decrease_indent(&ctx(), &selector, None).unwrap();
    }
    sync(&bob, &alice);

    undo.undo();
    sync(&alice, &bob);
    assert_eq!(indent(&alice), None);
    assert_eq!(indent(&bob), None);
}
