//! What the editor draws after list, spacing, anchor and run edits: the same
//! markup PowerPoint reads back from the saved file.

use pptx_edit::{
    Bullet, DeckSession, EditCtx, LineSpacing, ParagraphSpacing, TextStyle, TextStylePatch,
};
use pptx_render::{
    ListKind, PositionedTextLine, Primitive, SlideRenderer, Strike, SurfaceDisplayList, TextAnchor,
    TextParagraph,
};

const DEMO: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
const LIST_STYLES: &[u8] = include_bytes!("fixtures/list-style-bullets.pptx");
const LECTURE: &[u8] = include_bytes!("../../../poc/fixtures/lecture.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn renderer() -> SlideRenderer {
    let mut renderer = SlideRenderer::new();
    for bold in [false, true] {
        renderer.register_font("Arial", bold, false, FONT).unwrap();
    }
    renderer
}

fn context() -> EditCtx {
    EditCtx::local("render")
}

fn discs() -> Vec<Bullet> {
    vec![Bullet::Character {
        value: "●".to_owned(),
    }]
}

fn numbers() -> Vec<Bullet> {
    vec![Bullet::AutoNumber {
        scheme: "arabicPeriod".to_owned(),
        start_at: 1,
        restart: false,
    }]
}

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

fn layout(session: &DeckSession, caret: Option<(&str, usize)>) -> SurfaceDisplayList {
    layout_slide(session, 0, caret)
}

fn layout_slide(
    session: &DeckSession,
    slide: usize,
    caret: Option<(&str, usize)>,
) -> SurfaceDisplayList {
    renderer()
        .layout_scoped_slide_at_caret(
            session.package(),
            &session.slide_scope(slide).unwrap(),
            caret,
        )
        .unwrap()
        .display_list
}

fn text_box<'a>(
    list: &'a SurfaceDisplayList,
    story: &str,
) -> (&'a [TextParagraph], &'a [PositionedTextLine], TextAnchor) {
    list.primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::TextBox {
                story_id: Some(id),
                paragraphs,
                lines,
                anchor,
                ..
            } if id == story => Some((paragraphs.as_slice(), lines.as_slice(), *anchor)),
            _ => None,
        })
        .unwrap()
}

/// The marker drawn before each paragraph's first line: the runs that cover
/// no story text.
fn markers(session: &DeckSession, story: &str, lines: &[PositionedTextLine]) -> Vec<String> {
    let mut start = 0;
    session
        .story(story)
        .unwrap()
        .paragraphs
        .iter()
        .map(|paragraph| {
            let line = lines.iter().find(|line| line.start == start).unwrap();
            start += paragraph
                .runs
                .iter()
                .map(|run| run.text.encode_utf16().count() as u32)
                .sum::<u32>()
                + 1;
            line.runs
                .iter()
                .take_while(|run| run.start == run.end && !run.text.is_empty())
                .map(|run| run.text.as_str())
                .collect()
        })
        .collect()
}

#[test]
fn an_edited_bullet_is_drawn_and_reported() {
    let session = DeckSession::open(DEMO, 41).unwrap();
    let (_, _, story) = story_of(&session, "Title");
    let before = layout(&session, None);
    let (paragraphs, lines, _) = text_box(&before, &story);
    assert!(paragraphs.iter().all(|paragraph| paragraph.list.is_none()));
    let unlisted_x = lines[0].runs[0].x;
    assert_eq!(markers(&session, &story, lines), ["", ""]);

    let length = session.story(&story).unwrap().length;
    session
        .set_paragraph_list(&context(), &story, 0, length, Some(&discs()), &[])
        .unwrap();
    let after = layout(&session, None);
    let (paragraphs, lines, _) = text_box(&after, &story);
    assert!(
        paragraphs
            .iter()
            .all(|paragraph| paragraph.list == Some(ListKind::Bullet))
    );
    assert_eq!(markers(&session, &story, lines), ["●", "●"]);
    // The marker hangs at the old left edge and the text moves past it.
    let marker = &lines[0].runs[0];
    assert!(
        (marker.x - unlisted_x).abs() < 0.5,
        "{} {unlisted_x}",
        marker.x
    );
    assert!(lines[0].runs.last().unwrap().x > marker.x + marker.width);
}

#[test]
fn removing_an_inherited_bullet_hides_its_marker() {
    let session = DeckSession::open(LIST_STYLES, 42).unwrap();
    let (_, _, story) = story_of(&session, "Inherited bullets");
    let before = layout(&session, None);
    let (paragraphs, _, _) = text_box(&before, &story);
    assert_eq!(paragraphs[0].list, Some(ListKind::Bullet));
    session
        .set_paragraph_list(&context(), &story, 0, 0, None, &[])
        .unwrap();
    let after = layout(&session, None);
    let (paragraphs, lines, _) = text_box(&after, &story);
    assert_eq!(paragraphs[0].list, None);
    assert_eq!(paragraphs[1].list, Some(ListKind::Bullet));
    assert_eq!(markers(&session, &story, lines)[..1], [""]);
}

#[test]
fn the_empty_item_holding_the_caret_shows_the_marker_it_would_take() {
    let session = DeckSession::open(DEMO, 43).unwrap();
    let (_, _, story) = story_of(&session, "Title");
    let length = session.story(&story).unwrap().length;
    session
        .set_paragraph_list(&context(), &story, 0, length, Some(&numbers()), &[])
        .unwrap();
    // An empty item between the two: "1.", "", "2.".
    let first_end = session.story(&story).unwrap().paragraphs[0]
        .runs
        .iter()
        .map(|run| run.text.encode_utf16().count() as u32)
        .sum::<u32>();
    session
        .replace_text(
            &context(),
            &story,
            first_end,
            first_end,
            "\n",
            &TextStyle::default(),
        )
        .unwrap();

    let idle = layout(&session, None);
    assert_eq!(
        markers(&session, &story, text_box(&idle, &story).1),
        ["1.", "", "2."]
    );
    let typing = layout(&session, Some((story.as_str(), 1)));
    let (paragraphs, lines, _) = text_box(&typing, &story);
    assert_eq!(paragraphs[1].list, Some(ListKind::Number));
    // The empty item shows the number it would take; the next keeps its own.
    assert_eq!(markers(&session, &story, lines), ["1.", "2.", "2."]);
    // A caret in a non-empty item changes nothing.
    let on_text = layout(&session, Some((story.as_str(), 0)));
    assert_eq!(
        markers(&session, &story, text_box(&on_text, &story).1),
        ["1.", "", "2."]
    );
}

#[test]
fn spacing_anchor_strike_and_highlight_reach_the_display_list() {
    let session = DeckSession::open(DEMO, 44).unwrap();
    let (slide, shape, story) = story_of(&session, "Title");
    let before = layout(&session, None);
    let (_, lines, anchor) = text_box(&before, &story);
    assert_eq!(anchor, TextAnchor::Top);
    let gap = lines[1].y - lines[0].y;

    session
        .set_paragraph_spacing(
            &context(),
            &story,
            0,
            0,
            &ParagraphSpacing {
                after: Some(LineSpacing::Points { value: 30.0 }),
                ..ParagraphSpacing::default()
            },
        )
        .unwrap();
    session
        .set_text_anchor(&context(), &slide, &shape, Some("b"))
        .unwrap();
    session
        .format_text(
            &context(),
            &story,
            0,
            6,
            &TextStylePatch {
                strike: Some("sngStrike".to_owned()),
                highlight: Some("#FFFF00".to_owned()),
                ..TextStylePatch::default()
            },
        )
        .unwrap();
    let after = layout(&session, None);
    let (paragraphs, lines, anchor) = text_box(&after, &story);
    assert_eq!(anchor, TextAnchor::Bottom);
    assert_eq!(
        paragraphs[0].space_after,
        Some(LineSpacing::Points { value: 30.0 })
    );
    // 30 pt is 40 px.
    assert!(((lines[1].y - lines[0].y) - (gap + 40.0)).abs() < 0.5);
    let struck = &lines[0].runs[0];
    assert_eq!(struck.text, "Office");
    assert_eq!(struck.strike, Some(Strike::Single));
    assert_eq!(struck.highlight.as_deref(), Some("#FFFF00"));
    assert_eq!(lines[0].runs[1].strike, None);
    assert_eq!(lines[0].runs[1].highlight, None);
}

/// The x of each paragraph's marker and of its first text on its first line.
fn hangs(session: &DeckSession, story: &str, lines: &[PositionedTextLine]) -> Vec<(f32, f32)> {
    let mut start = 0;
    session
        .story(story)
        .unwrap()
        .paragraphs
        .iter()
        .map(|paragraph| {
            let line = lines.iter().find(|line| line.start == start).unwrap();
            start += paragraph
                .runs
                .iter()
                .map(|run| run.text.encode_utf16().count() as u32)
                .sum::<u32>()
                + 1;
            let marker = line
                .runs
                .iter()
                .find(|run| run.start == run.end && !run.text.is_empty())
                .unwrap();
            let text = line.runs.iter().find(|run| run.end > run.start).unwrap();
            (marker.x, text.x)
        })
        .collect()
}

#[test]
fn a_plain_paragraph_listed_with_items_lines_up_with_them_as_reopened() {
    // `lecture.pptx` slide 2: a plain paragraph (`marL` 457200, `indent` 0)
    // over two items (`marL` 457200, `indent` -317500).
    let session = DeckSession::open(LECTURE, 45).unwrap();
    let (_, _, story) = story_of(&session, "Google Shape;117;p26");
    let length = session.story(&story).unwrap().length;
    let square = vec![Bullet::Character {
        value: "■".to_owned(),
    }];
    session
        .set_paragraph_list(&context(), &story, 0, length, Some(&square), &[])
        .unwrap();
    let edited = layout_slide(&session, 1, None);
    let (_, lines, _) = text_box(&edited, &story);
    let hangs = hangs(&session, &story, lines);
    assert_eq!(hangs.len(), 3);
    assert!(hangs.iter().all(|hang| *hang == hangs[1]), "{hangs:?}");

    let reopened = DeckSession::open(&session.save().unwrap(), 46).unwrap();
    let again = layout_slide(&reopened, 1, None);
    assert_eq!(text_box(&again, &story).1, lines);
}
