//! Cell stories that two replicas create concurrently stay apart, so one replica removing its
//! table never takes a story another replica's table still names (upstream #1454).

use docx_edit::bridge::RenderEnv;
use docx_edit::*;
use yrs::types::text::YChange;
use yrs::{Any, Map, Out, ReadTxn, Text, Transact};

const DATE: &str = "2026-10-28T00:00:00Z";

fn ctx() -> EditCtx {
    EditCtx::local("Ada", DATE)
}

fn open(state: &[u8], client: u64) -> EngineSession {
    let engine = EngineSession::new(client);
    engine.doc().apply_update_v1(state).unwrap();
    engine
}

fn exchange(left: &EditingDoc, right: &EditingDoc) {
    let from_left = left.encode_state_as_update_v1();
    let from_right = right.encode_state_as_update_v1();
    right.apply_update_v1(&from_left).unwrap();
    left.apply_update_v1(&from_right).unwrap();
}

/// The cell stories of each body table, in body order.
fn tables(doc: &EditingDoc) -> Vec<Vec<String>> {
    let txn = doc.yrs_doc().transact();
    let Some(Out::YText(body)) = txn.get_map("stories").unwrap().get(&txn, "body") else {
        panic!("no body story");
    };
    body.diff(&txn, YChange::identity)
        .into_iter()
        .filter_map(|diff| match diff.insert {
            Out::YMap(table)
                if matches!(table.get(&txn, "_kind"), Some(Out::Any(Any::String(kind))) if &*kind == "table") =>
            {
                let Some(Out::Any(Any::Array(rows))) = table.get(&txn, "rows") else {
                    panic!("table without rows");
                };
                Some(
                    rows.iter()
                        .flat_map(|row| {
                            let Any::Map(row) = row else { panic!("row is not a map") };
                            let Some(Any::Array(cells)) = row.get("cells") else {
                                panic!("row without cells");
                            };
                            cells
                                .iter()
                                .map(|cell| match cell {
                                    Any::Map(cell) => match cell.get("story") {
                                        Some(Any::String(story)) => story.to_string(),
                                        other => panic!("cell without story: {other:?}"),
                                    },
                                    other => panic!("cell is not a map: {other:?}"),
                                })
                                .collect::<Vec<_>>()
                        })
                        .collect(),
                )
            }
            _ => None,
        })
        .collect()
}

fn text(doc: &EditingDoc, story: &str) -> String {
    doc.paragraphs(story)
        .unwrap()
        .into_iter()
        .map(|paragraph| paragraph.text)
        .collect()
}

#[test]
fn concurrent_tables_keep_their_own_cells_and_survive_the_others_removal() {
    let origin = EditingDoc::new(300);
    origin
        .create_story("body", "Before", "Normal", "left")
        .unwrap();
    let state = origin.encode_state_as_update_v1();
    let (left, right) = (open(&state, 51), open(&state, 52));
    for (engine, typed) in [(&left, "Left"), (&right, "Right")] {
        let table = engine
            .doc()
            .insert_table(&ctx(), Position::new("body", 0), 2, 2, 15)
            .unwrap();
        assert!(
            table
                .created_story_ids
                .iter()
                .all(|story| story.ends_with(&format!("~{}", engine.doc().client_id())))
        );
        engine
            .doc()
            .insert_text(
                &ctx(),
                Position::new(&table.created_story_ids[0], 0),
                typed,
                FormatPolicy::Plain,
            )
            .unwrap();
    }
    exchange(left.doc(), right.doc());
    let merged = tables(left.doc());
    assert_eq!(merged, tables(right.doc()));
    assert_eq!(merged.len(), 2);
    assert!(merged[0].iter().all(|story| !merged[1].contains(story)));
    let mut first_cells: Vec<_> = merged
        .iter()
        .map(|cells| text(left.doc(), &cells[0]))
        .collect();
    first_cells.sort();
    assert_eq!(first_cells, ["Left", "Right"]);

    let removed = merged
        .iter()
        .position(|cells| text(left.doc(), &cells[0]) == "Left")
        .unwrap();
    left.doc()
        .delete_table(&ctx(), &TableLocator::new("body", removed as u32))
        .unwrap();
    exchange(left.doc(), right.doc());
    for engine in [&left, &right] {
        let remaining = tables(engine.doc());
        assert_eq!(remaining.len(), 1);
        assert_eq!(text(engine.doc(), &remaining[0][0]), "Right");
        engine
            .lower_story_json("body", &RenderEnv::default())
            .unwrap();
    }
}

#[test]
fn edit_born_cells_carry_their_replicas_client_id() {
    let origin = EditingDoc::new(300);
    origin
        .create_story("body", "Before", "Normal", "left")
        .unwrap();
    origin
        .insert_table(&ctx(), Position::new("body", 0), 1, 1, 15)
        .unwrap();
    let cell = &tables(&origin)[0][0];
    assert_eq!(cell, "body:t0:r0c0~300");
    let replica = open(&origin.encode_state_as_update_v1(), 51);
    let added = replica
        .doc()
        .insert_row(&ctx(), &CellLoc::new("body", 0, 0, 0), true)
        .unwrap();
    assert_eq!(added.created_story_ids, ["body:t0:r1c0~51"]);
}
