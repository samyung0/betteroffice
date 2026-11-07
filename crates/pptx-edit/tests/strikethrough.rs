#[path = "../../pptx-parse/tests/common/strikethrough.rs"]
mod fixture;

use pptx_edit::{DeckSession, EditCtx, TextStyle, TextStylePatch};

fn authored_properties(strike: &str) -> String {
    format!(
        r#"<a:rPr lang="en-GB" strike="{strike}"><a:ea typeface="Yu Gothic"/><a:cs typeface="Arial"/></a:rPr>"#
    )
}

fn authored_deck(strike: &str) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(&fixture::deck()).unwrap();
    for (path, bytes) in &mut parts {
        if path == "ppt/slides/slide1.xml" {
            *bytes = String::from_utf8(bytes.clone())
                .unwrap()
                .replace(
                    "<a:r><a:t>P</a:t></a:r>",
                    &format!(
                        "<a:r><a:t>left</a:t></a:r><a:r>{}<a:t>author</a:t></a:r>",
                        authored_properties(strike)
                    ),
                )
                .into_bytes();
        }
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn slide_xml(bytes: &[u8]) -> String {
    String::from_utf8(
        ooxml_opc::unzip_parts(bytes)
            .unwrap()
            .into_iter()
            .find(|(path, _)| path == "ppt/slides/slide1.xml")
            .unwrap()
            .1,
    )
    .unwrap()
}

fn assert_authored_properties(bytes: &[u8], strike: &str, expected_text: &str, added: &str) {
    let xml = slide_xml(bytes);
    let mut text = String::new();
    let mut added_seen = added.is_empty();
    for run in xml.split("<a:r>").skip(1) {
        let run = run.split("</a:r>").next().unwrap();
        if !run.contains(r#"lang="en-GB""#) {
            continue;
        }
        let (properties, rest) = run.split_once("<a:t>").unwrap();
        assert_eq!(properties.replace(added, ""), authored_properties(strike));
        added_seen |= properties.contains(added);
        text.push_str(rest.strip_suffix("</a:t>").unwrap());
    }
    assert_eq!(text, expected_text);
    assert!(added_seen);
}

#[test]
fn deleting_inside_authored_strike_preserves_run_properties() {
    for strike in ["noStrike", "sngStrike", "dblStrike"] {
        let session = DeckSession::open(&authored_deck(strike), 6414).unwrap();
        let snapshot = session.snapshot().unwrap();
        let story_id = &snapshot.slides[0].shapes[3].text_stories[0].id;
        session
            .delete_text(&EditCtx::local("strike"), story_id, 6, 7)
            .unwrap();
        assert_authored_properties(&session.save().unwrap(), strike, "auhor", "");
    }
}

#[test]
fn bold_format_preserves_authored_strike_properties() {
    for strike in ["noStrike", "sngStrike", "dblStrike"] {
        let session = DeckSession::open(&authored_deck(strike), 6415).unwrap();
        let snapshot = session.snapshot().unwrap();
        let story_id = &snapshot.slides[0].shapes[3].text_stories[0].id;
        session
            .format_text(
                &EditCtx::local("strike"),
                story_id,
                4,
                10,
                &TextStylePatch {
                    bold: Some(true),
                    ..TextStylePatch::default()
                },
            )
            .unwrap();
        assert_authored_properties(&session.save().unwrap(), strike, "author", r#" b="1""#);
    }
}

#[test]
fn inserting_into_neighbour_preserves_authored_strike_properties() {
    for strike in ["noStrike", "sngStrike", "dblStrike"] {
        let session = DeckSession::open(&authored_deck(strike), 6416).unwrap();
        let snapshot = session.snapshot().unwrap();
        let story_id = &snapshot.slides[0].shapes[3].text_stories[0].id;
        session
            .insert_text(
                &EditCtx::local("strike"),
                story_id,
                2,
                "X",
                &TextStyle::default(),
            )
            .unwrap();
        assert_authored_properties(&session.save().unwrap(), strike, "author", "");
    }
}

#[test]
fn changing_a_sibling_run_keeps_authored_strike_markup() {
    let session = DeckSession::open(&fixture::deck(), 6408).unwrap();
    let snapshot = session.snapshot().unwrap();
    let story_id = &snapshot.slides[0].shapes[0].text_stories[0].id;
    session
        .format_text(
            &EditCtx::local("strike"),
            story_id,
            1,
            2,
            &TextStylePatch {
                strike: Some("sngStrike".to_owned()),
                ..TextStylePatch::default()
            },
        )
        .unwrap();
    let xml = slide_xml(&session.save().unwrap());
    assert!(
        xml.contains(
            r#"<a:r><a:rPr lang="en-US" strike="sngStrike" dirty="0"/><a:t>A</a:t></a:r>"#
        )
    );
    assert!(xml.contains(r#"<a:r><a:rPr strike="noStrike"/><a:t>C</a:t></a:r>"#));
    let before = session.encode_state_as_update_v1();
    assert!(
        session
            .format_text(
                &EditCtx::local("strike"),
                story_id,
                0,
                1,
                &TextStylePatch {
                    strike: Some("invalid".to_owned()),
                    ..TextStylePatch::default()
                }
            )
            .is_err()
    );
    assert_eq!(session.encode_state_as_update_v1(), before);
}

#[test]
fn changing_another_mark_preserves_authored_no_strike() {
    let mut parts = ooxml_opc::unzip_parts(&fixture::deck()).unwrap();
    for (path, bytes) in &mut parts {
        if path == "ppt/slides/slide1.xml" {
            *bytes = String::from_utf8(bytes.clone())
                .unwrap()
                .replace(
                    "<a:t>P</a:t>",
                    r#"<a:t>P</a:t></a:r><a:r><a:rPr strike="noStrike"/><a:t>off</a:t>"#,
                )
                .into_bytes();
        }
    }
    let session = DeckSession::open(&ooxml_opc::rezip_parts(&parts).unwrap(), 6410).unwrap();
    let snapshot = session.snapshot().unwrap();
    let story_id = &snapshot.slides[0].shapes[3].text_stories[0].id;
    session
        .format_text(
            &EditCtx::local("strike"),
            story_id,
            1,
            2,
            &TextStylePatch {
                bold: Some(true),
                ..TextStylePatch::default()
            },
        )
        .unwrap();
    let reopened = DeckSession::open(&session.save().unwrap(), 6411).unwrap();
    let story = reopened.story(story_id).unwrap();
    assert_eq!(
        story.paragraphs[0]
            .runs
            .iter()
            .map(|run| (run.text.as_str(), run.style.strike.as_deref()))
            .collect::<Vec<_>>(),
        [
            ("P", None),
            ("o", Some("noStrike")),
            ("ff", Some("noStrike")),
            ("Q", Some("sngStrike"))
        ]
    );
}

/// The fixture's first box strikes level 1 twice through its list style; its run "C" turns that
/// off with `strike="noStrike"`. Text typed into that run with no strike of its own stays
/// un-struck, as PowerPoint and Google Slides keep the run's off mark.
#[test]
fn typing_into_an_explicitly_unstruck_run_keeps_it_unstruck() {
    for replace in [false, true] {
        let session = DeckSession::open(&fixture::deck(), 6443).unwrap();
        let story_id = session.snapshot().unwrap().slides[0].shapes[0].text_stories[0]
            .id
            .clone();
        let context = EditCtx::local("strike");
        if replace {
            // Replaces "C" itself.
            session
                .replace_text(&context, &story_id, 2, 3, "XY", &TextStyle::default())
                .unwrap();
        } else {
            // Types after "C", before the plain "D".
            session
                .insert_text(&context, &story_id, 3, "XY", &TextStyle::default())
                .unwrap();
        }
        let saved = session.save().unwrap();
        let reopened = DeckSession::open(&saved, 6444).unwrap();
        let story = reopened.snapshot().unwrap().slides[0].shapes[0].text_stories[0].clone();
        let typed = story.paragraphs[0]
            .runs
            .iter()
            .find(|run| run.text.contains("XY"))
            .unwrap();
        assert_eq!(typed.style.strike.as_deref(), Some("noStrike"), "{replace}");
        let xml = slide_xml(&saved);
        let run = xml
            .split("<a:r>")
            .skip(1)
            .find(|run| run.contains("XY</a:t>"))
            .unwrap();
        assert!(run.contains(r#"strike="noStrike""#), "{run}");
        // "D", the plain run after it, still inherits the list style's strike.
        let plain = story.paragraphs[0]
            .runs
            .iter()
            .find(|run| run.text.ends_with('D'))
            .unwrap();
        assert_eq!(plain.style.strike, None, "{replace}");
    }
}
