//! Two people editing the same text at once: character formatting, paragraph
//! styles, lists, footnotes and pictures converge on every peer.

use std::collections::BTreeMap;
use std::sync::Arc;

use docx_edit::*;
use yrs::Any;

const DATE: &str = "2026-10-13T00:00:00Z";
const FOOTNOTE_DOCX: &[u8] = include_bytes!("fixtures/footnote-anchor.docx");

fn ctx(author: &str) -> EditCtx {
    EditCtx::local(author, DATE)
}

/// Alice and Bob, each holding `seed`'s state.
fn peers(seed: &EditingDoc) -> (EditingDoc, EditingDoc) {
    let state = seed.encode_state_as_update_v1();
    let (alice, bob) = (EditingDoc::new(601), EditingDoc::new(602));
    alice.apply_update_v1(&state).unwrap();
    bob.apply_update_v1(&state).unwrap();
    (alice, bob)
}

/// Both ways, twice, so renames made while applying reach the other peer.
fn sync(alice: &EditingDoc, bob: &EditingDoc) {
    for _ in 0..2 {
        for (from, to) in [(alice, bob), (bob, alice)] {
            let update = from.encode_diff_v1(&to.encode_state_vector_v1()).unwrap();
            to.applying_peer_update(|| to.apply_update_v1(&update).unwrap());
        }
    }
}

fn assert_converged(alice: &EditingDoc, bob: &EditingDoc, story: &str) {
    assert_eq!(
        alice.story_segments(story).unwrap(),
        bob.story_segments(story).unwrap()
    );
    assert_eq!(
        alice.paragraphs(story).unwrap(),
        bob.paragraphs(story).unwrap()
    );
}

/// Attributes of each text segment, keyed by its text.
fn runs(doc: &EditingDoc, story: &str) -> Vec<(String, BTreeMap<String, Any>)> {
    doc.story_segments(story)
        .unwrap()
        .into_iter()
        .filter_map(|segment| match segment.content {
            SegmentContent::Text(text) => Some((text, segment.attributes)),
            _ => None,
        })
        .collect()
}

fn on(attrs: &BTreeMap<String, Any>, key: &str) -> bool {
    matches!(attrs.get(key), Some(value) if *value != Any::Null)
}

/// Alice bolds "quick brown" while Bob italicises "brown fox" and types
/// inside "quick": both peers end with the same runs, "brown" bold and
/// italic, and Bob's word in the text.
#[test]
fn bold_and_italic_over_overlapping_ranges_converge_with_both_marks() {
    let seed = EditingDoc::new(500);
    seed.create_story("body", "The quick brown fox", "Normal", "left")
        .unwrap();
    let (alice, bob) = peers(&seed);

    alice
        .toggle_format(
            &ctx("Alice"),
            StoryRange::new("body", 4, 15),
            SimpleFormat::Bold,
        )
        .unwrap();
    bob.toggle_format(
        &ctx("Bob"),
        StoryRange::new("body", 10, 19),
        SimpleFormat::Italic,
    )
    .unwrap();
    bob.insert_text(
        &ctx("Bob"),
        Position::new("body", 6),
        "i",
        FormatPolicy::Inherit,
    )
    .unwrap();
    sync(&alice, &bob);

    assert_converged(&alice, &bob, "body");
    assert_eq!(
        alice.paragraphs("body").unwrap()[0].text,
        "The quiick brown fox"
    );
    let marks: Vec<(String, bool, bool)> = runs(&alice, "body")
        .into_iter()
        .map(|(text, attrs)| (text, on(&attrs, "bold"), on(&attrs, "italic")))
        .collect();
    assert_eq!(
        marks,
        [
            ("The ".into(), false, false),
            ("quiick ".into(), true, false),
            ("brown".into(), true, true),
            (" fox".into(), false, true),
        ]
    );
}

/// Alice applies Heading 1 while Bob applies Title to the same paragraph.
/// Each style resets the other's properties with explicit nulls, so both
/// peers keep one style with exactly that style's values, not a mix.
#[test]
fn concurrent_paragraph_styles_converge_on_one_styles_values() {
    let seed = EditingDoc::new(500);
    let paragraph = seed
        .create_story("body", "Chapter one", "Normal", "left")
        .unwrap();
    let (alice, bob) = peers(&seed);
    let selector = ParaSelector::One(paragraph);
    // The paragraph's left alignment is Normal's.
    let previous: BTreeMap<String, BTreeMap<String, Any>> = [(
        "Normal".to_owned(),
        [("alignment".to_owned(), Any::from("left"))].into(),
    )]
    .into();
    let heading: BTreeMap<String, Any> = [
        ("spaceBefore".to_owned(), Any::Number(480.0)),
        ("keepNext".to_owned(), Any::Bool(true)),
        ("outlineLevel".to_owned(), Any::Number(0.0)),
    ]
    .into();
    let title: BTreeMap<String, Any> = [
        ("alignment".to_owned(), Any::from("center")),
        ("spaceAfter".to_owned(), Any::Number(80.0)),
        ("contextualSpacing".to_owned(), Any::Bool(true)),
    ]
    .into();

    alice
        .apply_paragraph_style(&ctx("Alice"), &selector, "Heading1", &heading, &previous)
        .unwrap();
    bob.apply_paragraph_style(&ctx("Bob"), &selector, "Title", &title, &previous)
        .unwrap();
    sync(&alice, &bob);

    assert_converged(&alice, &bob, "body");
    let properties = &alice.paragraphs("body").unwrap()[0].properties;
    let (winner, loser) = match properties.get("pStyle") {
        Some(Any::String(style)) if style.as_ref() == "Heading1" => (&heading, &title),
        Some(Any::String(style)) if style.as_ref() == "Title" => (&title, &heading),
        other => panic!("unexpected style {other:?}"),
    };
    for (key, value) in winner {
        assert_eq!(properties.get(key), Some(value), "{key}");
    }
    for key in loser.keys() {
        assert!(!on(properties, key), "{key} kept from the other style");
    }
}

fn list(num_id: f64, bullet: bool) -> ParaAttrDelta {
    ParaAttrDelta {
        other: [
            (
                "numPr".to_owned(),
                Some(Any::Map(Arc::new(
                    [
                        ("numId".to_owned(), Any::Number(num_id)),
                        ("ilvl".to_owned(), Any::Number(0.0)),
                    ]
                    .into(),
                ))),
            ),
            ("listIsBullet".to_owned(), Some(Any::Bool(bullet))),
            (
                "listNumFmt".to_owned(),
                Some(if bullet {
                    Any::Null
                } else {
                    Any::from("decimal")
                }),
            ),
        ]
        .into(),
        ..ParaAttrDelta::default()
    }
}

/// Alice makes both paragraphs a bulleted list while Bob makes the second a
/// numbered list and presses Enter in the first: both peers agree on every
/// paragraph's list, and each list paragraph's numbering and bullet flag come
/// from the same toolbar command (one wins the whole delta, never a mix).
#[test]
fn concurrent_list_toggles_and_enter_converge() {
    let seed = EditingDoc::new(500);
    let first = seed
        .create_story("body", "Apples", "Normal", "left")
        .unwrap();
    let second = seed
        .split_paragraph(&ctx("Seed"), Position::new("body", 6))
        .unwrap()
        .second_para_id;
    seed.insert_text(
        &ctx("Seed"),
        Position::new("body", 7),
        "Pears",
        FormatPolicy::Inherit,
    )
    .unwrap();
    let (alice, bob) = peers(&seed);

    alice
        .set_paragraph_attrs(
            &ctx("Alice"),
            &ParaSelector::Many(vec![first.clone(), second.clone()]),
            &list(1.0, true),
        )
        .unwrap();
    bob.set_paragraph_attrs(
        &ctx("Bob"),
        &ParaSelector::One(second.clone()),
        &list(2.0, false),
    )
    .unwrap();
    bob.split_paragraph(&ctx("Bob"), Position::new("body", 3))
        .unwrap();
    sync(&alice, &bob);

    assert_converged(&alice, &bob, "body");
    let paragraphs = alice.paragraphs("body").unwrap();
    let texts: Vec<&str> = paragraphs.iter().map(|p| p.text.as_str()).collect();
    assert_eq!(texts, ["App", "les", "Pears"]);
    let num_id = |properties: &BTreeMap<String, Any>| match properties.get("numPr") {
        Some(Any::Map(numbering)) => numbering.get("numId").cloned(),
        _ => None,
    };
    assert!(num_id(&paragraphs[2].properties).is_some());
    for paragraph in &paragraphs {
        let bullet = paragraph.properties.get("listIsBullet") == Some(&Any::Bool(true));
        if let Some(Any::Number(id)) = num_id(&paragraph.properties) {
            assert_eq!(bullet, id == 1.0, "{paragraph:?}");
        }
    }
}

/// Alice types at the end of a footnote while Bob types at its start and
/// presses Enter in it: both peers hold the same two note paragraphs with
/// both people's text.
#[test]
fn concurrent_typing_and_enter_in_one_footnote_converge() {
    let seed = EditingDoc::new(500);
    seed_from_docx(&seed, FOOTNOTE_DOCX).unwrap();
    let (alice, bob) = peers(&seed);
    let note = "fn:2";
    assert_eq!(alice.paragraphs(note).unwrap()[0].text, " Note body text");

    let end = alice.story_len(note).unwrap() - 1;
    alice
        .insert_text(
            &ctx("Alice"),
            Position::new(note, end),
            " (Alice)",
            FormatPolicy::Inherit,
        )
        .unwrap();
    bob.insert_text(
        &ctx("Bob"),
        Position::new(note, 0),
        "Bob:",
        FormatPolicy::Inherit,
    )
    .unwrap();
    bob.split_paragraph(&ctx("Bob"), Position::new(note, 14))
        .unwrap();
    sync(&alice, &bob);

    assert_converged(&alice, &bob, note);
    let texts: Vec<String> = alice
        .paragraphs(note)
        .unwrap()
        .into_iter()
        .map(|paragraph| paragraph.text)
        .collect();
    assert_eq!(texts, ["Bob: Note body", " text (Alice)"]);
}

fn image_payload(doc: &EditingDoc, embed_id: &str) -> BTreeMap<String, Any> {
    doc.story_segments("body")
        .unwrap()
        .into_iter()
        .find_map(|segment| match segment.content {
            SegmentContent::OtherEmbed { kind, payload }
                if kind == "image" && payload.get("embedId") == Some(&Any::from(embed_id)) =>
            {
                Some(payload)
            }
            _ => None,
        })
        .expect("image embed")
}

/// Alice resizes a picture while Bob sets its wrapping and rotates it and
/// types before it: both peers keep the picture with Alice's size and Bob's
/// wrap and rotation (each geometry entry merges on its own).
#[test]
fn concurrent_picture_resize_wrap_and_typing_keep_every_change() {
    let seed = EditingDoc::new(500);
    seed.create_story("body", "See the figure.", "Normal", "left")
        .unwrap();
    seed.insert_embed(
        &ctx("Seed"),
        Position::new("body", 8),
        "image",
        vec![
            ("embedId".to_owned(), Any::from("picture-1")),
            ("widthEmu".to_owned(), Any::Number(914_400.0)),
            ("heightEmu".to_owned(), Any::Number(457_200.0)),
            ("wrap".to_owned(), Any::from("inline")),
        ],
    )
    .unwrap();
    let (alice, bob) = peers(&seed);

    alice
        .set_embed_attrs_by_id(
            &ctx("Alice"),
            "picture-1",
            vec![
                ("widthEmu".to_owned(), Any::Number(1_828_800.0)),
                ("heightEmu".to_owned(), Any::Number(914_400.0)),
            ],
        )
        .unwrap();
    bob.set_embed_attrs_by_id(
        &ctx("Bob"),
        "picture-1",
        vec![
            ("wrap".to_owned(), Any::from("square")),
            ("transform".to_owned(), Any::from("rotate(90deg)")),
        ],
    )
    .unwrap();
    bob.insert_text(
        &ctx("Bob"),
        Position::new("body", 0),
        "Please ",
        FormatPolicy::Inherit,
    )
    .unwrap();
    sync(&alice, &bob);

    assert_converged(&alice, &bob, "body");
    let payload = image_payload(&alice, "picture-1");
    assert_eq!(payload.get("widthEmu"), Some(&Any::Number(1_828_800.0)));
    assert_eq!(payload.get("heightEmu"), Some(&Any::Number(914_400.0)));
    assert_eq!(payload.get("wrap"), Some(&Any::from("square")));
    assert_eq!(payload.get("transform"), Some(&Any::from("rotate(90deg)")));
    assert_eq!(
        alice.paragraphs("body").unwrap()[0].text,
        "Please See the figure."
    );
}
