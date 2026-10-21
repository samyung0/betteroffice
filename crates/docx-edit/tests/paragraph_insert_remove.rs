//! A paragraph removed whole leaves its neighbours' ids and properties, and
//! one inserted before another takes a fresh id: the agent's commands name
//! paragraphs by id within one call.

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

/// Before the first paragraph and before a later one, a new paragraph takes
/// a fresh id and the properties of the paragraph it precedes, which keeps
/// its own id; the paragraph before it is untouched.
#[test]
fn a_paragraph_inserted_before_another_takes_a_fresh_id_and_its_properties() {
    let doc = three();
    let before = body(&doc);
    let (alpha, gamma) = (before[0].0.clone(), before[2].0.clone());
    let first = doc.insert_paragraph_before(&ctx(), "body", &alpha).unwrap();
    let second = doc.insert_paragraph_before(&ctx(), "body", &gamma).unwrap();
    let center = Some(Any::from("center"));
    let alignment = before[0].2.clone();
    assert_eq!(
        body(&doc),
        [
            (first.first_para_id.clone(), String::new(), alignment),
            before[0].clone(),
            before[1].clone(),
            (second.first_para_id.clone(), String::new(), center),
            before[2].clone(),
        ]
    );
    assert_eq!(
        (first.second_para_id, second.second_para_id),
        (alpha, gamma)
    );
    assert_eq!(
        doc.insert_paragraph_before(&ctx(), "body", "nope"),
        Err(OpError::UnknownPara("nope".into()))
    );
}

/// A page break opening a paragraph's slot stays when that paragraph goes,
/// opening the next one's.
#[test]
fn a_removed_paragraph_leaves_the_page_break_it_opened_with() {
    let doc = three();
    let before = body(&doc);
    // "Alpha" ends at 5 and its mark is 5; "Beta" starts at 6.
    doc.insert_embed(&ctx(), Position::new("body", 6), "pageBreak", vec![])
        .unwrap();
    doc.remove_paragraph(&ctx(), "body", &before[1].0).unwrap();
    assert_eq!(body(&doc), [before[0].clone(), before[2].clone()]);
    let kinds: Vec<String> = doc
        .story_segments("body")
        .unwrap()
        .into_iter()
        .map(|segment| match segment.content {
            SegmentContent::Text(text) => text,
            SegmentContent::Pilcrow(_) => "¶".into(),
            SegmentContent::OtherEmbed { kind, .. } => kind,
        })
        .collect();
    assert_eq!(kinds, ["Alpha", "¶", "pageBreak", "Gamma", "¶"]);
}
