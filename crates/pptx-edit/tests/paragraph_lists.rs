//! Lists, levels, spacing, vertical alignment and the run formatting Google
//! Slides has, each edited, saved and reopened. PowerPoint's markup is the
//! oracle: `a:buChar`/`a:buAutoNum`/`a:buNone` with `a:buFont`, `marL`/`indent`/
//! `lvl`, `a:lnSpc`/`a:spcBef`/`a:spcAft`, `a:bodyPr@anchor`, `a:rPr@strike`
//! and `a:highlight`.

use pptx_edit::{
    Bullet, BulletFont, DeckSession, EditCtx, LineSpacing, ParagraphSpacing, TextStyle,
    TextStylePatch,
};
use pptx_parse::{ShapeNode, TextBody};

const DEMO: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
const LIST_STYLES: &[u8] =
    include_bytes!("../../pptx-render/tests/fixtures/list-style-bullets.pptx");

fn context() -> EditCtx {
    EditCtx::local("lists")
}

fn discs() -> Vec<Bullet> {
    ["●", "○", "■"]
        .into_iter()
        .map(|value| Bullet::Character {
            value: value.to_owned(),
        })
        .collect()
}

fn decimal() -> Vec<Bullet> {
    ["arabicPeriod", "alphaLcPeriod", "romanLcPeriod"]
        .into_iter()
        .map(|scheme| Bullet::AutoNumber {
            scheme: scheme.to_owned(),
            start_at: 1,
            restart: false,
        })
        .collect()
}

/// (slide id, shape id, story id) of the first shape called `name`.
fn story_of(session: &DeckSession, name: &str) -> (String, String, String) {
    session
        .snapshot()
        .unwrap()
        .slides
        .iter()
        .find_map(|slide| {
            slide
                .shapes
                .iter()
                .find(|shape| shape.name == name)
                .map(|shape| {
                    (
                        slide.id.clone(),
                        shape.id.clone(),
                        shape.text_stories[0].id.clone(),
                    )
                })
        })
        .unwrap()
}

fn story_length(session: &DeckSession, story_id: &str) -> u32 {
    session.story(story_id).unwrap().length
}

/// The saved text body of the first shape called `name`, as PowerPoint reads it.
fn saved_body(bytes: &[u8], name: &str) -> TextBody {
    let package = pptx_parse::parse_pptx(bytes).unwrap();
    package
        .slides
        .iter()
        .flat_map(|slide| &slide.shapes)
        .find_map(|node| match node {
            ShapeNode::Shape(shape) if shape.base.name == name => shape.text.clone(),
            _ => None,
        })
        .unwrap()
}

/// The `p:sp` XML of the first shape called `name` in slide part `part`.
fn shape_xml(bytes: &[u8], part: &str, name: &str) -> String {
    let parts = ooxml_opc::unzip_parts(bytes).unwrap();
    let xml = parts
        .into_iter()
        .find(|(path, _)| path == part)
        .map(|(_, data)| String::from_utf8(data).unwrap())
        .unwrap();
    let marker = format!("name=\"{name}\"");
    let at = xml.find(&marker).unwrap();
    let start = xml[..at].rfind("<p:sp>").unwrap();
    let end = at + xml[at..].find("</p:sp>").unwrap();
    xml[start..end].to_owned()
}

fn texts(body: &TextBody) -> Vec<String> {
    body.paragraphs
        .iter()
        .map(|paragraph| paragraph.runs.iter().map(|run| run.text.as_str()).collect())
        .collect()
}

#[test]
fn bulleting_a_text_box_writes_powerpoint_list_markup() {
    let session = DeckSession::open(DEMO, 11).unwrap();
    let (_, _, story) = story_of(&session, "Title");
    let before = texts(&saved_body(DEMO, "Title"));
    session
        .set_paragraph_list(
            &context(),
            &story,
            0,
            story_length(&session, &story),
            Some(&discs()),
        )
        .unwrap();

    let saved = session.save().unwrap();
    let body = saved_body(&saved, "Title");
    assert_eq!(texts(&body), before);
    for paragraph in &body.paragraphs {
        let properties = &paragraph.properties;
        assert_eq!(
            properties.bullet,
            Some(Bullet::Character {
                value: "●".to_owned()
            })
        );
        assert_eq!(
            properties.bullet_font,
            Some(BulletFont::Typeface("Arial".to_owned()))
        );
        assert_eq!(properties.margin_left, Some(342_900));
        assert_eq!(properties.indent, Some(-342_900));
        assert_eq!(properties.level, 0);
    }
    let xml = shape_xml(&saved, "ppt/slides/slide1.xml", "Title");
    assert!(
        xml.contains(r#"<a:buFont typeface="Arial"/><a:buChar char="●"/>"#),
        "{xml}"
    );

    // Reopened, the list is the file's own: the editor shows the same paragraphs.
    let reopened = DeckSession::open(&saved, 12).unwrap();
    let (_, _, story) = story_of(&reopened, "Title");
    let paragraphs = reopened.story(&story).unwrap().paragraphs;
    assert!(paragraphs.iter().all(|paragraph| {
        paragraph.bullet_json.as_deref() == Some(r#"{"type":"character","value":"●"}"#)
    }));
}

#[test]
fn numbering_writes_bu_auto_num_in_the_text_font() {
    let session = DeckSession::open(DEMO, 13).unwrap();
    let (_, _, story) = story_of(&session, "Title");
    session
        .set_paragraph_list(&context(), &story, 0, 0, Some(&decimal()))
        .unwrap();
    let saved = session.save().unwrap();
    let body = saved_body(&saved, "Title");
    assert_eq!(
        body.paragraphs[0].properties.bullet,
        Some(Bullet::AutoNumber {
            scheme: "arabicPeriod".to_owned(),
            start_at: 1,
            restart: false
        })
    );
    assert_eq!(
        body.paragraphs[0].properties.bullet_font,
        Some(BulletFont::FollowText)
    );
    // A caret selects only its own paragraph.
    assert_eq!(body.paragraphs[1].properties.bullet, None);
    let xml = shape_xml(&saved, "ppt/slides/slide1.xml", "Title");
    assert!(
        xml.contains(r#"<a:buFontTx/><a:buAutoNum type="arabicPeriod"/>"#),
        "{xml}"
    );
}

#[test]
fn indenting_a_list_item_moves_its_level_margin_and_marker() {
    let session = DeckSession::open(DEMO, 14).unwrap();
    let (_, _, story) = story_of(&session, "Title");
    let end = story_length(&session, &story);
    session
        .set_paragraph_list(&context(), &story, 0, end, Some(&discs()))
        .unwrap();
    let second = session.story(&story).unwrap().paragraphs[0]
        .runs
        .iter()
        .map(|run| run.text.encode_utf16().count() as u32)
        .sum::<u32>()
        + 1;
    session
        .change_paragraph_level(&context(), &story, second, second, 1, Some(&discs()))
        .unwrap();

    let saved = session.save().unwrap();
    let body = saved_body(&saved, "Title");
    let indented = &body.paragraphs[1].properties;
    assert_eq!(indented.level, 1);
    assert_eq!(indented.margin_left, Some(457_200 + 342_900));
    assert_eq!(indented.indent, Some(-342_900));
    assert_eq!(
        indented.bullet,
        Some(Bullet::Character {
            value: "○".to_owned()
        })
    );
    assert_eq!(body.paragraphs[0].properties.level, 0);

    // Reopened, the margin is the file's: outdenting shifts it back.
    let reopened = DeckSession::open(&saved, 15).unwrap();
    let (_, _, story) = story_of(&reopened, "Title");
    reopened
        .change_paragraph_level(&context(), &story, second, second, -1, Some(&discs()))
        .unwrap();
    let body = saved_body(&reopened.save().unwrap(), "Title");
    let outdented = &body.paragraphs[1].properties;
    assert_eq!(outdented.level, 0);
    assert_eq!(outdented.margin_left, Some(342_900));
    assert_eq!(
        outdented.bullet,
        Some(Bullet::Character {
            value: "●".to_owned()
        })
    );
}

#[test]
fn levels_stop_at_zero_and_eight() {
    let session = DeckSession::open(DEMO, 16).unwrap();
    let (_, _, story) = story_of(&session, "Title");
    session
        .change_paragraph_level(&context(), &story, 0, 0, -1, None)
        .unwrap();
    assert_eq!(session.story(&story).unwrap().paragraphs[0].level, 0);
    for _ in 0..12 {
        session
            .change_paragraph_level(&context(), &story, 0, 0, 1, None)
            .unwrap();
    }
    assert_eq!(session.story(&story).unwrap().paragraphs[0].level, 8);
    let body = saved_body(&session.save().unwrap(), "Title");
    assert_eq!(body.paragraphs[0].properties.level, 8);
}

#[test]
fn indenting_a_plain_paragraph_without_a_margin_changes_only_its_level() {
    let session = DeckSession::open(LIST_STYLES, 17).unwrap();
    let (_, _, story) = story_of(&session, "Inherited bullets");
    session
        .change_paragraph_level(&context(), &story, 0, 0, 1, Some(&discs()))
        .unwrap();
    let saved = session.save().unwrap();
    let xml = shape_xml(&saved, "ppt/slides/slide1.xml", "Inherited bullets");
    assert!(
        xml.contains(r#"<a:pPr lvl="1"/><a:r><a:rPr/><a:t>First level"#),
        "{xml}"
    );
}

#[test]
fn removing_an_inherited_bullet_writes_bu_none_without_a_hanging_indent() {
    let session = DeckSession::open(LIST_STYLES, 18).unwrap();
    let (_, _, story) = story_of(&session, "Inherited bullets");
    session
        .set_paragraph_list(&context(), &story, 0, 0, None)
        .unwrap();
    let saved = session.save().unwrap();
    let body = saved_body(&saved, "Inherited bullets");
    let first = &body.paragraphs[0].properties;
    assert_eq!(first.bullet, Some(Bullet::None));
    assert_eq!(first.margin_left, Some(0));
    assert_eq!(first.indent, Some(0));
    let xml = shape_xml(&saved, "ppt/slides/slide1.xml", "Inherited bullets");
    assert!(
        xml.contains(r#"<a:pPr indent="0" marL="0"><a:buNone/></a:pPr>"#),
        "{xml}"
    );
    // The other paragraphs keep their exact markup.
    assert_eq!(
        body.paragraphs[1..],
        saved_body(LIST_STYLES, "Inherited bullets").paragraphs[1..]
    );
}

#[test]
fn a_bullet_replaces_the_file_marker_and_its_symbol_font() {
    let session = DeckSession::open(LIST_STYLES, 19).unwrap();
    let (_, _, story) = story_of(&session, "Inherited bullets");
    let length = story_length(&session, &story);
    session
        .set_paragraph_list(&context(), &story, length - 1, length - 1, Some(&discs()))
        .unwrap();
    let xml = shape_xml(
        &session.save().unwrap(),
        "ppt/slides/slide1.xml",
        "Inherited bullets",
    );
    // `a:buFontTx` gave way to the bullet's face, in schema order.
    assert!(
        xml.contains(
            r#"<a:pPr indent="-342900" marL="342900"><a:buClrTx/><a:buSzTx/><a:buFont typeface="Arial"/><a:buChar char="●"/></a:pPr>"#
        ),
        "{xml}"
    );
}

#[test]
fn line_and_paragraph_spacing_round_trip_in_schema_order() {
    let session = DeckSession::open(DEMO, 20).unwrap();
    let (_, _, story) = story_of(&session, "Title");
    session
        .set_paragraph_list(&context(), &story, 0, 0, Some(&discs()))
        .unwrap();
    session
        .set_paragraph_spacing(
            &context(),
            &story,
            0,
            0,
            &ParagraphSpacing {
                line: Some(LineSpacing::Percent { value: 1.5 }),
                before: Some(LineSpacing::Points { value: 10.0 }),
                after: Some(LineSpacing::Points { value: 0.0 }),
            },
        )
        .unwrap();
    let saved = session.save().unwrap();
    let body = saved_body(&saved, "Title");
    let properties = &body.paragraphs[0].properties;
    assert_eq!(
        properties.line_spacing,
        Some(LineSpacing::Percent { value: 1.5 })
    );
    assert_eq!(
        properties.space_before,
        Some(LineSpacing::Points { value: 10.0 })
    );
    assert_eq!(
        properties.space_after,
        Some(LineSpacing::Points { value: 0.0 })
    );
    let xml = shape_xml(&saved, "ppt/slides/slide1.xml", "Title");
    assert!(
        xml.contains(
            r#"<a:lnSpc><a:spcPct val="150000"/></a:lnSpc><a:spcBef><a:spcPts val="1000"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buFont typeface="Arial"/><a:buChar char="●"/>"#
        ),
        "{xml}"
    );
    // Only the paragraph the caret is in changed.
    assert_eq!(body.paragraphs[1].properties.line_spacing, None);
}

#[test]
fn a_text_box_anchor_round_trips_and_rejects_junk() {
    let session = DeckSession::open(DEMO, 21).unwrap();
    let (slide, shape, _) = story_of(&session, "Title");
    assert!(
        session
            .set_text_anchor(&context(), &slide, &shape, Some("middle"))
            .is_err()
    );
    session
        .set_text_anchor(&context(), &slide, &shape, Some("b"))
        .unwrap();
    let saved = session.save().unwrap();
    assert_eq!(saved_body(&saved, "Title").anchor.as_deref(), Some("b"));
    // Restoring the file's anchor saves the file's markup again.
    session
        .set_text_anchor(&context(), &slide, &shape, None)
        .unwrap();
    assert_eq!(
        saved_body(&session.save().unwrap(), "Title").anchor,
        saved_body(DEMO, "Title").anchor
    );
}

#[test]
fn strikethrough_and_highlight_round_trip_on_the_run() {
    let session = DeckSession::open(DEMO, 22).unwrap();
    let (_, _, story) = story_of(&session, "Subtitle");
    session
        .format_text(
            &context(),
            &story,
            0,
            5,
            &TextStylePatch {
                strike: Some("sngStrike".to_owned()),
                highlight: Some("#FFFF00".to_owned()),
                ..TextStylePatch::default()
            },
        )
        .unwrap();
    assert!(
        session
            .format_text(
                &context(),
                &story,
                0,
                5,
                &TextStylePatch {
                    strike: Some("wavy".to_owned()),
                    ..TextStylePatch::default()
                },
            )
            .is_err()
    );
    let saved = session.save().unwrap();
    let body = saved_body(&saved, "Subtitle");
    let run = &body.paragraphs[0].runs[0];
    assert_eq!(run.text.chars().count(), 5);
    assert_eq!(run.properties.strike.as_deref(), Some("sngStrike"));
    assert_eq!(
        run.properties
            .highlight
            .as_ref()
            .and_then(|color| color.rgb.as_deref()),
        Some("FFFF00")
    );
    let xml = shape_xml(&saved, "ppt/slides/slide1.xml", "Subtitle");
    assert!(xml.contains(r#"strike="sngStrike""#), "{xml}");
    assert!(
        xml.contains(r#"<a:highlight><a:srgbClr val="FFFF00"/></a:highlight><a:latin"#),
        "{xml}"
    );
    assert!(
        !body.paragraphs[0].runs[1..]
            .iter()
            .any(|run| run.properties.strike.is_some())
    );

    // Reopened, the formatting is seeded from the file and saving is a no-op.
    let reopened = DeckSession::open(&saved, 23).unwrap();
    let (_, _, story) = story_of(&reopened, "Subtitle");
    let style = &reopened.story(&story).unwrap().paragraphs[0].runs[0].style;
    assert_eq!(style.strike.as_deref(), Some("sngStrike"));
    assert_eq!(style.highlight.as_deref(), Some("#FFFF00"));
    assert_eq!(
        reopened.save().unwrap(),
        pptx_parse::write_pptx(reopened.package()).unwrap()
    );
}

#[test]
fn clearing_formatting_removes_every_run_attribute_an_edit_sets() {
    let session = DeckSession::open(DEMO, 24).unwrap();
    let (_, _, story) = story_of(&session, "Subtitle");
    session
        .format_text(
            &context(),
            &story,
            0,
            5,
            &TextStylePatch {
                bold: Some(true),
                italic: Some(true),
                strike: Some("sngStrike".to_owned()),
                highlight: Some("#00FF00".to_owned()),
                baseline_pct: Some(30.0),
                ..TextStylePatch::default()
            },
        )
        .unwrap();
    let length = story_length(&session, &story);
    session
        .clear_text_formatting(&context(), &story, 0, length - 1, None)
        .unwrap();
    let story_snapshot = session.story(&story).unwrap();
    assert!(
        story_snapshot
            .paragraphs
            .iter()
            .flat_map(|paragraph| &paragraph.runs)
            .all(|run| run.style == TextStyle::default())
    );
    let body = saved_body(&session.save().unwrap(), "Subtitle");
    for run in &body.paragraphs[0].runs {
        let properties = &run.properties;
        assert_eq!(properties.bold, None);
        assert_eq!(properties.italic, None);
        assert_eq!(properties.strike, None);
        assert_eq!(properties.highlight, None);
        assert_eq!(properties.baseline_pct, None);
        assert_eq!(properties.font_size_pt, None);
        assert_eq!(properties.color, None);
        assert_eq!(properties.font_family, None);
    }
}

#[test]
fn invalid_lists_and_spacing_are_refused_and_change_nothing() {
    let session = DeckSession::open(DEMO, 25).unwrap();
    let (_, _, story) = story_of(&session, "Title");
    let before = session.encode_state_as_update_v1();
    let refused = [
        vec![],
        vec![Bullet::None],
        vec![Bullet::Character {
            value: " ".to_owned(),
        }],
        vec![Bullet::AutoNumber {
            scheme: "emoji".to_owned(),
            start_at: 1,
            restart: false,
        }],
        vec![Bullet::AutoNumber {
            scheme: "arabicPeriod".to_owned(),
            start_at: 0,
            restart: false,
        }],
    ];
    for levels in refused {
        assert!(
            session
                .set_paragraph_list(&context(), &story, 0, 0, Some(&levels))
                .is_err(),
            "{levels:?}"
        );
    }
    for spacing in [
        LineSpacing::Percent { value: -1.0 },
        LineSpacing::Percent { value: 200.0 },
        LineSpacing::Points { value: f64::NAN },
        LineSpacing::Points { value: 2_000.0 },
    ] {
        assert!(
            session
                .set_paragraph_spacing(
                    &context(),
                    &story,
                    0,
                    0,
                    &ParagraphSpacing {
                        line: Some(spacing),
                        ..ParagraphSpacing::default()
                    },
                )
                .is_err()
        );
    }
    assert_eq!(session.encode_state_as_update_v1(), before);
}

#[test]
fn a_peer_writing_an_out_of_range_margin_is_refused() {
    let left = DeckSession::open(DEMO, 26).unwrap();
    let right = DeckSession::open(DEMO, 27).unwrap();
    let (_, _, story) = story_of(&left, "Title");
    let remote_vector = left.encode_state_vector_v1();
    {
        use yrs::{Map, Out, ReadTxn, Text, Transact};
        let doc = right.yrs_doc();
        let mut txn = doc.transact_mut_with(27_u64);
        let stories = txn.get_map("pptx:stories").unwrap();
        let text = stories
            .get(&txn, &story)
            .unwrap()
            .cast::<yrs::TextRef>()
            .unwrap();
        let pilcrow = text
            .diff(&txn, yrs::types::text::YChange::identity)
            .into_iter()
            .find_map(|diff| match diff.insert {
                Out::YMap(map) => Some(map),
                _ => None,
            })
            .unwrap();
        pilcrow.insert(&mut txn, "marginLeft", 9e12_f64);
    }
    let update = right.encode_diff_v1(&remote_vector).unwrap();
    assert!(left.apply_update_v1(&update).is_err());
    assert_eq!(left.story(&story).unwrap().paragraphs[0].margin_left, None);
}

/// Exchanges both peers' pending updates.
fn sync(left: &DeckSession, right: &DeckSession) {
    let to_right = left
        .encode_diff_v1(&right.encode_state_vector_v1())
        .unwrap();
    let to_left = right
        .encode_diff_v1(&left.encode_state_vector_v1())
        .unwrap();
    right.apply_update_v1(&to_right).unwrap();
    left.apply_update_v1(&to_left).unwrap();
}

/// A split keeps the paragraph's mark on its second half and copies its
/// properties to a new first half, so a concurrent property edit lands on the
/// second half only: the same as concurrent alignment does today.
#[test]
fn a_list_and_a_concurrent_split_converge() {
    let left = DeckSession::open(DEMO, 31).unwrap();
    let right = DeckSession::open(DEMO, 32).unwrap();
    let (_, _, story) = story_of(&left, "Title");
    let end = story_length(&left, &story);
    left.set_paragraph_list(&EditCtx::local("left"), &story, 0, end, Some(&discs()))
        .unwrap();
    // The other peer splits the first paragraph and types into the new one.
    right
        .replace_text(
            &EditCtx::local("right"),
            &story,
            6,
            6,
            "\nnew",
            &TextStyle::default(),
        )
        .unwrap();
    sync(&left, &right);

    assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
    let saved = left.save().unwrap();
    assert_eq!(saved, right.save().unwrap());
    let body = saved_body(&saved, "Title");
    assert_eq!(
        texts(&body),
        ["Office", "new files,", "without the office."]
    );
    let bullets: Vec<_> = body
        .paragraphs
        .iter()
        .map(|paragraph| paragraph.properties.bullet.clone())
        .collect();
    let disc = Some(Bullet::Character {
        value: "●".to_owned(),
    });
    assert_eq!(bullets, [None, disc.clone(), disc]);
}

#[test]
fn concurrent_indent_and_list_removal_converge() {
    let left = DeckSession::open(DEMO, 33).unwrap();
    let right = DeckSession::open(DEMO, 34).unwrap();
    let (_, _, story) = story_of(&left, "Title");
    let end = story_length(&left, &story);
    left.set_paragraph_list(&EditCtx::local("left"), &story, 0, end, Some(&discs()))
        .unwrap();
    sync(&left, &right);

    left.change_paragraph_level(&EditCtx::local("left"), &story, 0, 0, 1, Some(&discs()))
        .unwrap();
    right
        .set_paragraph_list(&EditCtx::local("right"), &story, 0, 0, None)
        .unwrap();
    sync(&left, &right);

    let snapshot = left.story(&story).unwrap();
    assert_eq!(snapshot, right.story(&story).unwrap());
    assert_eq!(left.save().unwrap(), right.save().unwrap());
    // Both edits survive key by key: the level moved and the list went.
    let first = &snapshot.paragraphs[0];
    assert_eq!(first.level, 1);
    assert_eq!(first.bullet_json.as_deref(), Some(r#"{"type":"none"}"#));
}

#[test]
fn undo_restores_the_file_markup_after_a_list_edit() {
    let session = DeckSession::open(DEMO, 35).unwrap();
    let (_, _, story) = story_of(&session, "Title");
    session
        .set_paragraph_list(&context(), &story, 0, 0, Some(&discs()))
        .unwrap();
    session.add_undo_barrier();
    session
        .set_paragraph_spacing(
            &context(),
            &story,
            0,
            0,
            &ParagraphSpacing {
                line: Some(LineSpacing::Percent { value: 2.0 }),
                ..ParagraphSpacing::default()
            },
        )
        .unwrap();
    assert!(session.undo());
    assert!(session.undo());
    assert_eq!(
        session.save().unwrap(),
        pptx_parse::write_pptx(session.package()).unwrap()
    );
}
