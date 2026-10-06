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
const PLAIN_PARAGRAPH: &[u8] =
    include_bytes!("../../pptx-render/tests/fixtures/list-style-plain-paragraph.pptx");
const LECTURE: &[u8] = include_bytes!("../../../poc/fixtures/lecture.pptx");
/// `lecture.pptx` slide 2's body (a Google Slides export): a plain paragraph at
/// `marL="457200" indent="0"`, then two `●` items at `marL="457200" indent="-317500"`.
const LECTURE_BODY: &str = "Google Shape;117;p26";

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
            &[],
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
        .set_paragraph_list(&context(), &story, 0, 0, Some(&decimal()), &[])
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
        .set_paragraph_list(&context(), &story, 0, end, Some(&discs()), &[])
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
        .set_paragraph_list(&context(), &story, 0, 0, None, &[])
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
        .set_paragraph_list(
            &context(),
            &story,
            length - 1,
            length - 1,
            Some(&discs()),
            &[],
        )
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

/// The first list item of `lecture.pptx`'s slide 2 (a Google Slides export):
/// `marL="457200" indent="-317500"` with a 14 pt `●`.
fn lecture_list_item(session: &DeckSession) -> (String, u32) {
    let story = session
        .snapshot()
        .unwrap()
        .slides
        .iter()
        .flat_map(|slide| &slide.shapes)
        .flat_map(|shape| &shape.text_stories)
        .find(|story| story.plain_text().contains("Ad minim veniam"))
        .unwrap()
        .clone();
    let mut offset = 0;
    for paragraph in &story.paragraphs {
        let text: String = paragraph.runs.iter().map(|run| run.text.as_str()).collect();
        if text.starts_with("Ea c") {
            return (story.id, offset);
        }
        offset += text.encode_utf16().count() as u32 + 1;
    }
    panic!("no list item");
}

fn lecture_item_properties(bytes: &[u8]) -> pptx_parse::ParagraphProperties {
    pptx_parse::parse_pptx(bytes).unwrap().slides[1]
        .shapes
        .iter()
        .find_map(|node| match node {
            ShapeNode::Shape(shape) => shape.text.as_ref()?.paragraphs.iter().find(|paragraph| {
                paragraph
                    .runs
                    .iter()
                    .map(|run| run.text.as_str())
                    .collect::<String>()
                    .starts_with("Ea c")
            }),
            _ => None,
        })
        .unwrap()
        .properties
        .clone()
}

#[test]
fn a_new_style_on_an_existing_list_item_keeps_its_indents() {
    let before = lecture_item_properties(LECTURE);
    assert_eq!(
        (before.margin_left, before.indent),
        (Some(457_200), Some(-317_500))
    );
    let session = DeckSession::open(LECTURE, 28).unwrap();
    let (story, offset) = lecture_list_item(&session);
    session
        .set_paragraph_list(&context(), &story, offset, offset, Some(&discs()), &[])
        .unwrap();
    let restyled = lecture_item_properties(&session.save().unwrap());
    assert_eq!(
        (restyled.margin_left, restyled.indent),
        (Some(457_200), Some(-317_500))
    );
    assert_eq!(
        restyled.bullet,
        Some(Bullet::Character {
            value: "●".to_owned()
        })
    );
    assert_eq!(restyled.bullet_size, before.bullet_size);

    session
        .set_paragraph_list(&context(), &story, offset, offset, Some(&decimal()), &[])
        .unwrap();
    let numbered = lecture_item_properties(&session.save().unwrap());
    assert_eq!(
        (numbered.margin_left, numbered.indent),
        (Some(457_200), Some(-317_500))
    );
    assert!(matches!(numbered.bullet, Some(Bullet::AutoNumber { .. })));
    assert_eq!(numbered.bullet_font, Some(BulletFont::FollowText));
}

#[test]
fn an_inherited_list_item_named_by_the_caller_keeps_its_indents() {
    let session = DeckSession::open(LIST_STYLES, 29).unwrap();
    let (_, _, story) = story_of(&session, "Inherited bullets");
    let first = session.story(&story).unwrap().paragraphs[0].id.clone();
    session
        .set_paragraph_list(&context(), &story, 0, 0, Some(&discs()), &[first])
        .unwrap();
    let saved = session.save().unwrap();
    let properties = &saved_body(&saved, "Inherited bullets").paragraphs[0].properties;
    assert_eq!((properties.margin_left, properties.indent), (None, None));
    assert_eq!(
        properties.bullet,
        Some(Bullet::Character {
            value: "●".to_owned()
        })
    );
}

fn squares() -> Vec<Bullet> {
    vec![Bullet::Character {
        value: "■".to_owned(),
    }]
}

/// Each paragraph's start offset in the story.
fn paragraph_offsets(session: &DeckSession, story: &str) -> Vec<u32> {
    let mut offset = 0;
    session
        .story(story)
        .unwrap()
        .paragraphs
        .iter()
        .map(|paragraph| {
            let start = offset;
            offset += paragraph
                .runs
                .iter()
                .map(|run| run.text.encode_utf16().count() as u32)
                .sum::<u32>()
                + 1;
            start
        })
        .collect()
}

fn indents(body: &TextBody) -> Vec<(Option<i64>, Option<i64>)> {
    body.paragraphs
        .iter()
        .map(|paragraph| {
            (
                paragraph.properties.margin_left,
                paragraph.properties.indent,
            )
        })
        .collect()
}

/// `marL`/`indent` as the editor lays each paragraph out (an edit's value,
/// else its file paragraph's) must be what the saved file holds.
fn assert_saved_as_edited(session: &DeckSession, story: &str, source: &TextBody, saved: &TextBody) {
    let edited: Vec<_> = session
        .story(story)
        .unwrap()
        .paragraphs
        .iter()
        .map(|paragraph| {
            let file = paragraph
                .template_index(story)
                .map(|index| &source.paragraphs[index].properties);
            (
                paragraph
                    .margin_left
                    .or(file.and_then(|file| file.margin_left)),
                paragraph.indent.or(file.and_then(|file| file.indent)),
            )
        })
        .collect();
    assert_eq!(edited, indents(saved));
}

#[test]
fn a_plain_paragraph_listed_with_items_takes_their_indents() {
    let source = saved_body(LECTURE, LECTURE_BODY);
    assert_eq!(
        indents(&source),
        [
            (Some(457_200), Some(0)),
            (Some(457_200), Some(-317_500)),
            (Some(457_200), Some(-317_500)),
        ]
    );
    let session = DeckSession::open(LECTURE, 50).unwrap();
    let (_, _, story) = story_of(&session, LECTURE_BODY);
    let end = story_length(&session, &story);
    session
        .set_paragraph_list(&context(), &story, 0, end, Some(&squares()), &[])
        .unwrap();

    let saved = session.save().unwrap();
    let body = saved_body(&saved, LECTURE_BODY);
    assert_eq!(texts(&body), texts(&source));
    assert_eq!(indents(&body), [(Some(457_200), Some(-317_500)); 3]);
    assert!(body.paragraphs.iter().all(|paragraph| {
        paragraph.properties.bullet
            == Some(Bullet::Character {
                value: "■".to_owned(),
            })
    }));
    // The items keep their 14 pt markers.
    assert_eq!(
        body.paragraphs[1].properties.bullet_size,
        source.paragraphs[1].properties.bullet_size
    );
    let xml = shape_xml(&saved, "ppt/slides/slide2.xml", LECTURE_BODY);
    assert_eq!(
        xml.matches(r#"<a:pPr algn="l" indent="-317500" marL="457200" rtl="0">"#)
            .count(),
        3,
        "{xml}"
    );
    assert_saved_as_edited(&session, &story, &source, &body);
    let reopened = DeckSession::open(&saved, 51).unwrap();
    let (_, _, story) = story_of(&reopened, LECTURE_BODY);
    assert_saved_as_edited(&reopened, &story, &body, &body);
}

#[test]
fn a_paragraph_listed_with_an_inherited_indent_item_inherits_too() {
    let source = saved_body(LIST_STYLES, "Inherited bullets");
    let session = DeckSession::open(LIST_STYLES, 52).unwrap();
    let (_, _, story) = story_of(&session, "Inherited bullets");
    // The first item made plain first: an edit's `marL="0" indent="0"`.
    session
        .set_paragraph_list(&context(), &story, 0, 0, None, &[])
        .unwrap();
    let ids: Vec<_> = session
        .story(&story)
        .unwrap()
        .paragraphs
        .iter()
        .map(|paragraph| paragraph.id.clone())
        .collect();
    // "Second level" and "Text follows bullet" inherit their markers; the
    // level-2 paragraph has none.
    let listed = [ids[1].clone(), ids[3].clone()];
    let end = story_length(&session, &story);
    session.add_undo_barrier();
    session
        .set_paragraph_list(&context(), &story, 0, end, Some(&discs()), &listed)
        .unwrap();

    let saved = session.save().unwrap();
    let body = saved_body(&saved, "Inherited bullets");
    assert_eq!(texts(&body), texts(&source));
    assert_eq!(
        indents(&body),
        [
            // Level 0 inherits as "Text follows bullet" does.
            (None, None),
            (None, None),
            // No level-2 item: PowerPoint's hanging indent at that level.
            (Some(2 * 457_200 + 342_900), Some(-342_900)),
            (None, None),
        ]
    );
    let bullets: Vec<_> = body
        .paragraphs
        .iter()
        .map(|paragraph| paragraph.properties.bullet.clone())
        .collect();
    let character = |value: &str| {
        Some(Bullet::Character {
            value: value.to_owned(),
        })
    };
    assert_eq!(
        bullets,
        [
            character("●"),
            character("○"),
            character("■"),
            character("●")
        ]
    );
    let xml = shape_xml(&saved, "ppt/slides/slide1.xml", "Inherited bullets");
    assert!(
        xml.contains(
            r#"<a:pPr><a:buFont typeface="Arial"/><a:buChar char="●"/></a:pPr><a:r><a:rPr/><a:t>First level"#
        ),
        "{xml}"
    );
    assert_saved_as_edited(&session, &story, &source, &body);

    // One Undo brings the removed keys back with the plain paragraph.
    assert!(session.undo());
    let first = &session.story(&story).unwrap().paragraphs[0];
    assert_eq!((first.margin_left, first.indent), (Some(0), Some(0)));
    assert_eq!(first.bullet_json.as_deref(), Some(r#"{"type":"none"}"#));
}

#[test]
fn powerpoints_plain_paragraph_takes_an_inherited_items_laid_out_indents() {
    // `marL="0" indent="0"` with `a:buNone` beside items inheriting the
    // master's `bodyStyle` 228600/-228600; saving cannot drop the file's zeros.
    let source = saved_body(PLAIN_PARAGRAPH, "Inherited bullets");
    assert_eq!(indents(&source)[0], (Some(0), Some(0)));
    assert_eq!(indents(&source)[3], (None, None));
    let session = DeckSession::open(PLAIN_PARAGRAPH, 57).unwrap();
    let (_, _, story) = story_of(&session, "Inherited bullets");
    let ids: Vec<_> = session
        .story(&story)
        .unwrap()
        .paragraphs
        .iter()
        .map(|paragraph| paragraph.id.clone())
        .collect();
    let end = story_length(&session, &story);
    session
        .set_paragraph_list(
            &context(),
            &story,
            0,
            end,
            Some(&discs()),
            &[ids[1].clone(), ids[3].clone()],
        )
        .unwrap();

    let saved = session.save().unwrap();
    let body = saved_body(&saved, "Inherited bullets");
    assert_eq!(texts(&body), texts(&source));
    assert_eq!(
        indents(&body),
        [
            (Some(228_600), Some(-228_600)),
            (None, None),
            (Some(2 * 457_200 + 342_900), Some(-342_900)),
            (None, None),
        ]
    );
    let xml = shape_xml(&saved, "ppt/slides/slide1.xml", "Inherited bullets");
    assert!(
        xml.contains(
            r#"<a:pPr indent="-228600" marL="228600"><a:buFont typeface="Arial"/><a:buChar char="●"/></a:pPr><a:r><a:rPr/><a:t>First level"#
        ),
        "{xml}"
    );
    assert_saved_as_edited(&session, &story, &source, &body);
    let reopened = DeckSession::open(&saved, 58).unwrap();
    let (_, _, story) = story_of(&reopened, "Inherited bullets");
    assert_saved_as_edited(&reopened, &story, &body, &body);
}

#[test]
fn a_paragraph_inheriting_another_layout_paragraph_takes_the_items_indents() {
    // Layout paragraphs are looked up by position: "Plain line" inherits
    // 114300/-114300 from the first, "Bulleted line" 457200/-228600 from the second.
    let session = DeckSession::open(PLAIN_PARAGRAPH, 59).unwrap();
    let (_, _, story) = story_of(&session, "Positioned bullets");
    let end = story_length(&session, &story);
    session
        .set_paragraph_list(&context(), &story, 0, end, Some(&discs()), &[])
        .unwrap();
    let saved = session.save().unwrap();
    let body = saved_body(&saved, "Positioned bullets");
    assert_eq!(
        indents(&body),
        [(Some(457_200), Some(-228_600)), (None, None)]
    );
    let source = saved_body(PLAIN_PARAGRAPH, "Positioned bullets");
    assert_saved_as_edited(&session, &story, &source, &body);
}

/// The text body of the first table cell of the graphic frame called `name`.
fn saved_cell(bytes: &[u8], name: &str) -> TextBody {
    pptx_parse::parse_pptx(bytes)
        .unwrap()
        .slides
        .iter()
        .flat_map(|slide| &slide.shapes)
        .find_map(|node| match node {
            ShapeNode::GraphicFrame(frame) if frame.base.name == name => match &frame.data {
                pptx_parse::GraphicFrameData::Table(table) => {
                    Some(table.rows[0].cells[0].text.clone())
                }
                _ => None,
            },
            _ => None,
        })
        .unwrap()
}

#[test]
fn a_table_cell_paragraph_takes_its_items_cell_list_style_indents() {
    // The cell's own `a:lstStyle` sets 171450/-171450; no placeholder or
    // master style reaches a cell.
    let session = DeckSession::open(PLAIN_PARAGRAPH, 60).unwrap();
    let (_, _, story) = story_of(&session, "List table");
    let end = story_length(&session, &story);
    session
        .set_paragraph_list(&context(), &story, 0, end, Some(&discs()), &[])
        .unwrap();
    let saved = session.save().unwrap();
    let body = saved_cell(&saved, "List table");
    assert_eq!(
        indents(&body),
        [(Some(171_450), Some(-171_450)), (None, None)]
    );
    let source = saved_cell(PLAIN_PARAGRAPH, "List table");
    assert_saved_as_edited(&session, &story, &source, &body);
    let reopened = DeckSession::open(&saved, 61).unwrap();
    let (_, _, story) = story_of(&reopened, "List table");
    assert_saved_as_edited(&reopened, &story, &body, &body);
}

/// `lecture.pptx` with slide 2's XML passed through `edit`.
fn lecture_with(edit: impl Fn(String) -> String) -> Vec<u8> {
    let parts: Vec<_> = ooxml_opc::unzip_parts(LECTURE)
        .unwrap()
        .into_iter()
        .map(|(path, data)| match path.as_str() {
            "ppt/slides/slide2.xml" => {
                let xml = edit(String::from_utf8(data).unwrap());
                (path, xml.into_bytes())
            }
            _ => (path, data),
        })
        .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn an_out_of_range_item_indent_is_clamped_before_a_peer_gets_it() {
    // A parseable but out-of-schema `marL="-50000"` on the first item.
    let bytes = lecture_with(|xml| {
        xml.replacen(
            r#"<a:pPr marL="457200" lvl="0" indent="-317500""#,
            r#"<a:pPr marL="-50000" lvl="0" indent="-317500""#,
            1,
        )
    });
    let left = DeckSession::open(&bytes, 62).unwrap();
    let right = DeckSession::open(&bytes, 63).unwrap();
    let (_, _, story) = story_of(&left, LECTURE_BODY);
    let end = story_length(&left, &story);
    left.set_paragraph_list(&context(), &story, 0, end, Some(&squares()), &[])
        .unwrap();
    // The story still loads and a peer accepts the update.
    let first = &left.story(&story).unwrap().paragraphs[0];
    assert_eq!((first.margin_left, first.indent), (Some(0), Some(-317_500)));
    sync(&left, &right);
    assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
    let body = saved_body(&right.save().unwrap(), LECTURE_BODY);
    assert_eq!(
        (
            body.paragraphs[0].properties.margin_left,
            body.paragraphs[0].properties.indent
        ),
        (Some(0), Some(-317_500))
    );
}

#[test]
fn the_first_item_at_a_level_gives_its_indents_not_the_nearest() {
    // Slide 2 reordered to item, item, plain, the second item at `marL` 685800.
    let bytes = lecture_with(|xml| {
        let plain = xml
            .find(r#"<a:p><a:pPr marL="457200" lvl="0" indent="0""#)
            .unwrap();
        let items = plain
            + xml[plain..]
                .find(r#"<a:p><a:pPr marL="457200" lvl="0" indent="-317500""#)
                .unwrap();
        let end = xml.find("</p:txBody></p:sp><p:pic>").unwrap();
        let second = items
            + 1
            + xml[items + 1..]
                .find(r#"<a:p><a:pPr marL="457200" lvl="0" indent="-317500""#)
                .unwrap();
        let shifted = xml[second..end].replacen(r#"marL="457200""#, r#"marL="685800""#, 1);
        format!(
            "{}{}{}{}{}",
            &xml[..plain],
            &xml[items..second],
            shifted,
            &xml[plain..items],
            &xml[end..]
        )
    });
    let source = saved_body(&bytes, LECTURE_BODY);
    assert_eq!(
        indents(&source),
        [
            (Some(457_200), Some(-317_500)),
            (Some(685_800), Some(-317_500)),
            (Some(457_200), Some(0)),
        ]
    );
    let session = DeckSession::open(&bytes, 64).unwrap();
    let (_, _, story) = story_of(&session, LECTURE_BODY);
    let end = story_length(&session, &story);
    session
        .set_paragraph_list(&context(), &story, 0, end, Some(&squares()), &[])
        .unwrap();
    let body = saved_body(&session.save().unwrap(), LECTURE_BODY);
    assert_eq!(indents(&body)[2], (Some(457_200), Some(-317_500)));
}

#[test]
fn newly_listed_paragraphs_take_the_indents_of_an_item_at_their_level() {
    let source = saved_body(LECTURE, LECTURE_BODY);
    let session = DeckSession::open(LECTURE, 53).unwrap();
    let (_, _, story) = story_of(&session, LECTURE_BODY);
    let offsets = paragraph_offsets(&session, &story);
    // The last item and the plain paragraph move to level 1 (`marL` 914400).
    for offset in [offsets[2], offsets[0]] {
        session
            .change_paragraph_level(&context(), &story, offset, offset, 1, Some(&discs()))
            .unwrap();
    }
    let end = story_length(&session, &story);
    session
        .set_paragraph_list(&context(), &story, 0, end, Some(&discs()), &[])
        .unwrap();

    let saved = session.save().unwrap();
    let body = saved_body(&saved, LECTURE_BODY);
    assert_eq!(texts(&body), texts(&source));
    // The plain paragraph takes the level-1 item's edited `marL` and file
    // `indent`, not the level-0 item's.
    assert_eq!(
        indents(&body),
        [
            (Some(914_400), Some(-317_500)),
            (Some(457_200), Some(-317_500)),
            (Some(914_400), Some(-317_500)),
        ]
    );
    let levels: Vec<_> = body
        .paragraphs
        .iter()
        .map(|paragraph| paragraph.properties.level)
        .collect();
    assert_eq!(levels, [1, 0, 1]);
    assert_saved_as_edited(&session, &story, &source, &body);
}

#[test]
fn listing_with_items_and_a_concurrent_indent_converge() {
    let source = saved_body(LECTURE, LECTURE_BODY);
    let left = DeckSession::open(LECTURE, 54).unwrap();
    let right = DeckSession::open(LECTURE, 55).unwrap();
    let (_, _, story) = story_of(&left, LECTURE_BODY);
    let end = story_length(&left, &story);
    left.set_paragraph_list(
        &EditCtx::local("left"),
        &story,
        0,
        end,
        Some(&squares()),
        &[],
    )
    .unwrap();
    // The other peer moves the plain paragraph a level in meanwhile.
    right
        .change_paragraph_level(&EditCtx::local("right"), &story, 0, 0, 1, None)
        .unwrap();
    sync(&left, &right);

    assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
    let saved = left.save().unwrap();
    assert_eq!(saved, right.save().unwrap());
    let body = saved_body(&saved, LECTURE_BODY);
    assert_eq!(texts(&body), texts(&source));
    let first = &body.paragraphs[0].properties;
    // Both edits survive key by key: the level, the marker and the item's
    // hanging indent; `marL` is whichever peer's write wins.
    assert_eq!(first.level, 1);
    assert_eq!(
        first.bullet,
        Some(Bullet::Character {
            value: "■".to_owned()
        })
    );
    assert_eq!(first.indent, Some(-317_500));
    assert!(matches!(first.margin_left, Some(457_200 | 914_400)));
    assert_saved_as_edited(&left, &story, &source, &body);
    let reopened = DeckSession::open(&saved, 56).unwrap();
    let (_, _, story) = story_of(&reopened, LECTURE_BODY);
    assert_saved_as_edited(&reopened, &story, &body, &body);
}

#[test]
fn line_and_paragraph_spacing_round_trip_in_schema_order() {
    let session = DeckSession::open(DEMO, 20).unwrap();
    let (_, _, story) = story_of(&session, "Title");
    session
        .set_paragraph_list(&context(), &story, 0, 0, Some(&discs()), &[])
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
                .set_paragraph_list(&context(), &story, 0, 0, Some(&levels), &[])
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
    left.set_paragraph_list(&EditCtx::local("left"), &story, 0, end, Some(&discs()), &[])
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
    left.set_paragraph_list(&EditCtx::local("left"), &story, 0, end, Some(&discs()), &[])
        .unwrap();
    sync(&left, &right);

    left.change_paragraph_level(&EditCtx::local("left"), &story, 0, 0, 1, Some(&discs()))
        .unwrap();
    right
        .set_paragraph_list(&EditCtx::local("right"), &story, 0, 0, None, &[])
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
        .set_paragraph_list(&context(), &story, 0, 0, Some(&discs()), &[])
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

#[test]
fn a_peer_writing_a_junk_strike_or_highlight_is_refused() {
    for (key, value) in [("strike", "wavy"), ("highlight", "red")] {
        let left = DeckSession::open(DEMO, 36).unwrap();
        let right = DeckSession::open(DEMO, 37).unwrap();
        let (_, _, story) = story_of(&left, "Subtitle");
        let remote_vector = left.encode_state_vector_v1();
        {
            use yrs::{Map, ReadTxn, Text, Transact};
            let doc = right.yrs_doc();
            let mut txn = doc.transact_mut_with(37_u64);
            let text = txn
                .get_map("pptx:stories")
                .unwrap()
                .get(&txn, &story)
                .unwrap()
                .cast::<yrs::TextRef>()
                .unwrap();
            text.format(
                &mut txn,
                0,
                3,
                yrs::types::Attrs::from([(std::sync::Arc::from(key), yrs::Any::from(value))]),
            );
        }
        let update = right.encode_diff_v1(&remote_vector).unwrap();
        assert!(left.apply_update_v1(&update).is_err(), "{key}={value}");
    }
}

/// Both peers hold the same story and save the same, readable file.
fn assert_converged(left: &DeckSession, right: &DeckSession) -> Vec<u8> {
    assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
    let saved = left.save().unwrap();
    assert_eq!(saved, right.save().unwrap());
    DeckSession::open(&saved, 99).unwrap();
    saved
}

#[test]
fn clearing_formatting_races_a_peers_highlight_and_typing() {
    let left = DeckSession::open(DEMO, 38).unwrap();
    let right = DeckSession::open(DEMO, 39).unwrap();
    let (_, _, story) = story_of(&left, "Subtitle");
    left.clear_text_formatting(&EditCtx::local("left"), &story, 0, 20, None)
        .unwrap();
    right
        .format_text(
            &EditCtx::local("right"),
            &story,
            5,
            15,
            &TextStylePatch {
                highlight: Some("#FFFF00".to_owned()),
                ..TextStylePatch::default()
            },
        )
        .unwrap();
    right
        .insert_text(
            &EditCtx::local("right"),
            &story,
            10,
            "typed",
            &TextStyle {
                bold: Some(true),
                ..TextStyle::default()
            },
        )
        .unwrap();
    sync(&left, &right);
    assert_converged(&left, &right);
    let text = left.story(&story).unwrap().plain_text();
    assert!(text.contains("typed"), "{text}");
}

#[test]
fn spacing_races_a_split_of_its_paragraph() {
    let left = DeckSession::open(DEMO, 40).unwrap();
    let right = DeckSession::open(DEMO, 41).unwrap();
    let (_, _, story) = story_of(&left, "Title");
    left.set_paragraph_spacing(
        &EditCtx::local("left"),
        &story,
        0,
        0,
        &ParagraphSpacing {
            line: Some(LineSpacing::Percent { value: 1.5 }),
            ..ParagraphSpacing::default()
        },
    )
    .unwrap();
    right
        .replace_text(
            &EditCtx::local("right"),
            &story,
            6,
            6,
            "\n",
            &TextStyle::default(),
        )
        .unwrap();
    sync(&left, &right);
    let saved = assert_converged(&left, &right);
    // As concurrent alignment: the split's second half carries the edit.
    let spacing: Vec<_> = saved_body(&saved, "Title")
        .paragraphs
        .iter()
        .map(|paragraph| paragraph.properties.line_spacing)
        .collect();
    assert_eq!(spacing[1], Some(LineSpacing::Percent { value: 1.5 }));
}

#[test]
fn an_anchor_races_a_peers_move_of_the_box() {
    let left = DeckSession::open(DEMO, 42).unwrap();
    let right = DeckSession::open(DEMO, 43).unwrap();
    let (slide, shape, _) = story_of(&left, "Title");
    left.set_text_anchor(&EditCtx::local("left"), &slide, &shape, Some("b"))
        .unwrap();
    right
        .move_shape(&EditCtx::local("right"), &slide, &shape, 100_000, 200_000)
        .unwrap();
    sync(&left, &right);
    let saved = assert_converged(&left, &right);
    assert_eq!(saved_body(&saved, "Title").anchor.as_deref(), Some("b"));
    let moved = left
        .snapshot()
        .unwrap()
        .slides
        .iter()
        .flat_map(|slide| &slide.shapes)
        .find(|candidate| candidate.id == shape)
        .map(|candidate| (candidate.x, candidate.y));
    assert_eq!(moved, Some((100_000, 200_000)));
}
