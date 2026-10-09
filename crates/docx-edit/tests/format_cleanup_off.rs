//! A remote update leaves redundant format items alone (Epo 2026-10-08,
//! D3). yrs's formatting cleanup deleted them on the replica that applied the
//! update and never told the others: one peer bolds a word while another
//! deletes it, the bolding peer then drops the markers around the gap, the
//! deleting peer keeps them, and the replicas hold different documents (the
//! override-seed spike's lockstep and its D3 measurement).

use docx_edit::*;
use yrs::block::ItemContent;
use yrs::{Map, ReadTxn, Transact};

fn ctx() -> EditCtx {
    EditCtx::local("Ada", "2026-10-08T00:00:00Z")
}

fn sync(from: &EditingDoc, to: &EditingDoc) {
    let update = from.encode_diff_v1(&to.encode_state_vector_v1()).unwrap();
    to.applying_peer_update(|| to.apply_update_v1(&update).unwrap());
}

/// Every live item of every story: its text, embed or format.
fn live_items(doc: &EditingDoc) -> Vec<String> {
    let txn = doc.yrs_doc().transact();
    let stories = txn.get_map("stories").unwrap();
    let mut items = Vec::new();
    for (name, story) in stories.iter(&txn) {
        let yrs::Out::YText(story) = story else {
            continue;
        };
        let mut item = <yrs::TextRef as AsRef<yrs::branch::Branch>>::as_ref(&story).start();
        while let Some(at) = item {
            item = at.right();
            if at.is_deleted() {
                continue;
            }
            items.push(match at.content() {
                ItemContent::Format(key, value) => format!("{name}: format {key}={value}"),
                ItemContent::String(text) => format!("{name}: text {}", text.as_str()),
                _ => format!("{name}: embed"),
            });
        }
    }
    items
}

#[test]
fn peers_keep_the_same_format_items_after_a_format_meets_a_delete() {
    let base = EditingDoc::new(1);
    base.create_story("body", "Hello brave new world.", "Normal", "left")
        .unwrap();
    let state = base.encode_state_as_update_v1();
    let peers: Vec<EditingDoc> = [701, 702]
        .into_iter()
        .map(|client| {
            let doc = EditingDoc::new(client);
            doc.apply_update_v1(&state).unwrap();
            doc
        })
        .collect();
    for round in 0..3 {
        let word = format!("w{round}");
        peers[0]
            .insert_text(&ctx(), Position::new("body", 0), &word, FormatPolicy::Inherit)
            .unwrap();
        sync(&peers[0], &peers[1]);
        // At once: one peer bolds the word, the other deletes it.
        peers[0]
            .toggle_format(&ctx(), StoryRange::new("body", 0, 2), SimpleFormat::Bold)
            .unwrap();
        peers[1]
            .delete_range(&ctx(), StoryRange::new("body", 0, 2))
            .unwrap();
        sync(&peers[0], &peers[1]);
        sync(&peers[1], &peers[0]);
    }
    assert_eq!(
        peers[0].story_segments("body").unwrap(),
        peers[1].story_segments("body").unwrap()
    );
    let (bolder, deleter) = (live_items(&peers[0]), live_items(&peers[1]));
    assert!(
        bolder.iter().any(|item| item.contains("format bold")),
        "the markers stay: {bolder:?}"
    );
    assert_eq!(bolder, deleter, "both peers hold the same items");
}
