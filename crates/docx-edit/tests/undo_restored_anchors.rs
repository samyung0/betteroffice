//! An Undo restores deleted text as new items and links each original to its
//! copy, on its own replica only. A comment or bookmark boundary that names
//! an original then resolves into the copy there and next to the deleted
//! original everywhere else. The replica that restored re-anchors such
//! boundaries, including ones that reach it later, so every peer agrees.

use docx_edit::*;
use yrs::{Any, ReadTxn, Transact};

const DATE: &str = "2026-10-08T00:00:00Z";

fn ctx() -> EditCtx {
    EditCtx::local("Ada", DATE)
}

/// Applies every pair's state diff the way an editor applies a peer's update
/// until a pass changes nothing; the passes that took, or `None`.
fn settle(peers: &[EditingDoc], limit: usize) -> Option<usize> {
    let vectors = |peers: &[EditingDoc]| -> Vec<yrs::StateVector> {
        peers
            .iter()
            .map(|peer| peer.yrs_doc().transact().state_vector())
            .collect()
    };
    for pass in 1..=limit {
        let before = vectors(peers);
        for from in peers {
            for to in peers {
                if !std::ptr::eq(from, to) {
                    sync(from, to);
                }
            }
        }
        if vectors(peers) == before {
            return Some(pass);
        }
    }
    None
}

fn sync(from: &EditingDoc, to: &EditingDoc) {
    let update = from.encode_diff_v1(&to.encode_state_vector_v1()).unwrap();
    to.applying_peer_update(|| to.apply_update_v1(&update).unwrap());
}

fn peers_of(base: &EditingDoc) -> Vec<EditingDoc> {
    let state = base.encode_state_as_update_v1();
    [701, 702, 703]
        .into_iter()
        .map(|client| {
            let doc = EditingDoc::new(client);
            doc.apply_update_v1(&state).unwrap();
            doc
        })
        .collect()
}

fn comment_body() -> Any {
    Any::from_json(
        r#"[{"type":"paragraph","content":[{"type":"run","content":[{"type":"text","text":"a comment"}]}]}]"#,
    )
    .unwrap()
}

fn ranges(doc: &EditingDoc, comment: &str) -> Vec<(String, u32, u32)> {
    doc.resolve_comment(comment)
        .unwrap()
        .into_iter()
        .map(|range| (range.story.clone(), range.start, range.end))
        .collect()
}

/// xml-mixed seed 344 of the override-seed spike's lockstep, reduced: one
/// peer deletes the text a comment starts in and undoes it while another
/// removes the comment and undoes that, bringing the comment back with the
/// anchors it had before the first Undo re-anchored them.
#[test]
fn a_comment_restored_with_old_anchors_resolves_alike_on_every_peer() {
    let base = EditingDoc::new(1);
    base.create_story(
        "body",
        "Hello brave new world, and then some more words to read.",
        "Normal",
        "left",
    )
    .unwrap();
    let comment = base
        .add_comment(
            &[StoryRange::new("body", 6, 21)],
            "Ada",
            DATE,
            comment_body(),
        )
        .unwrap();
    let peers = peers_of(&base);
    let deleter = UndoSession::new();
    deleter.track(&peers[2]);
    let remover = UndoSession::new();
    remover.track(&peers[1]);

    peers[2]
        .delete_range(&ctx(), StoryRange::new("body", 0, 30))
        .unwrap();
    deleter.add_undo_barrier();
    assert!(deleter.undo());
    peers[1]
        .apply_raw_ops(
            "body",
            vec![RawOp::RemoveComment {
                id: comment.clone(),
            }],
            &ctx(),
        )
        .unwrap();
    remover.add_undo_barrier();
    assert!(remover.undo());

    assert!(settle(&peers, 10).is_some(), "peers settle");
    let first = ranges(&peers[0], &comment);
    assert_eq!(first, [("body".to_owned(), 6, 21)]);
    for peer in &peers[1..] {
        assert_eq!(ranges(peer, &comment), first);
    }
}
