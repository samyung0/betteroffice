//! Peers rename the paragraph ids concurrent edits duplicated, and the
//! renames must settle. A mark an earlier rename gave `{client}.{clock}` of
//! its own item owns that id: renamed again, it would get the same id back,
//! and every peer would echo that write forever.

use docx_edit::*;
use yrs::{ReadTxn, Transact};

fn ctx() -> EditCtx {
    EditCtx::local("Ada", "2026-10-08T00:00:00Z")
}

/// Applies every pair's state diff the way an editor applies a peer's update
/// until a pass changes nothing; the passes that took, or `None` after `limit`.
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
                if std::ptr::eq(from, to) {
                    continue;
                }
                let update = from.encode_diff_v1(&to.encode_state_vector_v1()).unwrap();
                to.applying_peer_update(|| to.apply_update_v1(&update).unwrap());
            }
        }
        if vectors(peers) == before {
            return Some(pass);
        }
    }
    None
}

/// The body's paragraphs as (id, offset of its first character, text).
fn paragraphs(doc: &EditingDoc) -> Vec<(ParagraphId, u32, String)> {
    let mut offset = 0;
    doc.paragraphs("body")
        .unwrap()
        .into_iter()
        .map(|paragraph| {
            let start = offset;
            offset += paragraph.text.encode_utf16().count() as u32 + 1;
            (paragraph.para_id, start, paragraph.text)
        })
        .collect()
}

fn sync(from: &EditingDoc, to: &EditingDoc) {
    let update = from.encode_diff_v1(&to.encode_state_vector_v1()).unwrap();
    to.applying_peer_update(|| to.apply_update_v1(&update).unwrap());
}

#[test]
fn a_mark_keeps_the_id_a_rename_gave_it_and_peers_settle() {
    let base = EditingDoc::new(1);
    base.create_story(
        "body",
        "Alpha beta gamma delta epsilon zeta eta theta",
        "Normal",
        "left",
    )
    .unwrap();
    let state = base.encode_state_as_update_v1();
    let peers: Vec<EditingDoc> = [701, 702, 703]
        .into_iter()
        .map(|client| {
            let doc = EditingDoc::new(client);
            doc.apply_update_v1(&state).unwrap();
            doc
        })
        .collect();
    let undo = UndoSession::new();
    undo.track(&peers[0]);

    // 702 and 703 split the paragraph at once: both new marks take its id,
    // and 703's loses it, taking `703.<clock>` of its own item.
    peers[1]
        .split_paragraph(&ctx(), Position::new("body", 6))
        .unwrap();
    peers[2]
        .split_paragraph(&ctx(), Position::new("body", 17))
        .unwrap();
    assert!(settle(&peers, 6).is_some(), "the first renames settle");
    let (owned, start, text) = paragraphs(&peers[0])
        .into_iter()
        .find(|(id, _, _)| id.starts_with("703."))
        .expect("a mark renamed to its own item id");

    // 701 splits that paragraph: its new mark takes the id.
    undo.add_undo_barrier();
    let middle = start + text.encode_utf16().count() as u32 / 2;
    peers[0]
        .split_paragraph(&ctx(), Position::new("body", middle))
        .unwrap();
    undo.add_undo_barrier();
    sync(&peers[0], &peers[1]);

    // At once, 702 splits the first half again (its new mark takes the id
    // in turn) and 701 undoes its split (the owner gets the id back).
    let (_, first, _) = paragraphs(&peers[1])
        .into_iter()
        .find(|(id, _, _)| *id == owned)
        .expect("the first half carries the id");
    peers[1]
        .split_paragraph(&ctx(), Position::new("body", first + 2))
        .unwrap();
    assert!(undo.undo());

    let passes = settle(&peers, 10);
    assert!(
        passes.is_some(),
        "peers keep renaming: {:?}",
        paragraphs(&peers[0])
            .into_iter()
            .map(|(id, _, _)| id)
            .collect::<Vec<_>>()
    );
    let ids: Vec<ParagraphId> = paragraphs(&peers[0])
        .into_iter()
        .map(|(id, _, _)| id)
        .collect();
    let distinct: std::collections::HashSet<&ParagraphId> = ids.iter().collect();
    assert_eq!(distinct.len(), ids.len(), "{ids:?}");
    assert!(ids.contains(&owned), "the owner keeps {owned}: {ids:?}");
    for peer in &peers[1..] {
        let other: Vec<ParagraphId> = paragraphs(peer).into_iter().map(|(id, _, _)| id).collect();
        assert_eq!(other, ids);
    }
}
