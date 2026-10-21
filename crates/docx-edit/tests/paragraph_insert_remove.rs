//! A paragraph removed whole leaves its neighbours' ids and properties, and
//! one inserted at a story's start takes a fresh id: the agent's commands
//! name paragraphs by id within one call.

use docx_edit::*;
use yrs::Any;

fn ctx() -> EditCtx {
    EditCtx::local("Ada", "2026-10-21T00:00:00Z")
}

/// A body of "Alpha", "Beta" and "Gamma", Gamma centred.
fn three() -> EditingDoc {
    let doc = EditingDoc::new(1);
    doc.create_story("body", "AlphaBetaGamma", "Normal", "left")
        .unwrap();
    doc.split_paragraph(&ctx(), Position::new("body", 5))
        .unwrap();
    doc.split_paragraph(&ctx(), Position::new("body", 10))
        .unwrap();
    let gamma = doc.paragraphs("body").unwrap()[2].para_id.clone();
    doc.set_paragraph_attrs(
        &ctx(),
        &ParaSelector::One(gamma),
        &ParaAttrDelta {
            alignment: Patch::Set("center".into()),
            ..ParaAttrDelta::default()
        },
    )
    .unwrap();
    doc
}

/// Each body paragraph as (id, text, alignment).
fn body(doc: &EditingDoc) -> Vec<(String, String, Option<Any>)> {
    doc.paragraphs("body")
        .unwrap()
        .into_iter()
        .map(|paragraph| {
            let alignment = paragraph.properties.get("alignment").cloned();
            (paragraph.para_id, paragraph.text, alignment)
        })
        .collect()
}

#[test]
fn a_removed_paragraph_leaves_its_neighbours_as_they_were() {
    for removed in [0, 1] {
        let doc = three();
        let mut expected = body(&doc);
        let (id, ..) = expected.remove(removed);
        doc.remove_paragraph(&ctx(), "body", &id).unwrap();
        assert_eq!(body(&doc), expected, "removing paragraph {removed}");
    }
}

#[test]
fn the_final_paragraph_and_an_unknown_one_are_refused() {
    let doc = three();
    let before = body(&doc);
    let gamma = before[2].0.clone();
    assert_eq!(
        doc.remove_paragraph(&ctx(), "body", &gamma),
        Err(OpError::CannotMergeFinalParagraph(gamma))
    );
    assert_eq!(
        doc.remove_paragraph(&ctx(), "body", "nope"),
        Err(OpError::UnknownPara("nope".into()))
    );
    assert_eq!(body(&doc), before);
}

#[test]
fn a_paragraph_inserted_at_the_start_takes_a_fresh_id_and_the_first_ones_properties() {
    let doc = three();
    let gamma = body(&doc)[2].0.clone();
    doc.remove_paragraph(&ctx(), "body", &body(&doc)[0].0)
        .unwrap();
    doc.remove_paragraph(&ctx(), "body", &body(&doc)[0].0)
        .unwrap();
    // Gamma, centred, is the only paragraph left.
    let receipt = doc.insert_paragraph_at_start(&ctx(), "body").unwrap();
    let center = Some(Any::from("center"));
    assert_eq!(
        body(&doc),
        [
            (receipt.first_para_id.clone(), String::new(), center.clone()),
            (gamma.clone(), "Gamma".into(), center),
        ]
    );
    assert_eq!(receipt.second_para_id, gamma);
    assert_ne!(receipt.first_para_id, gamma);
}
