use pptx_edit::{DeckSession, EditCtx, TextStyle};

const DECK: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/table-basic.pptx");
const STORY: &str = "story:slide:0:256:shape:0:table:0:0";

#[test]
fn editing_a_cell_leaves_its_cell_properties_in_place() {
    let session = DeckSession::open(DECK, 34101).unwrap();
    session
        .insert_text(
            &EditCtx::local("test"),
            STORY,
            0,
            "New ",
            &TextStyle::default(),
        )
        .unwrap();
    let saved = session.save().unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let slide = parts
        .iter()
        .find(|(path, _)| path == "ppt/slides/slide1.xml")
        .map(|(_, bytes)| String::from_utf8(bytes.clone()).unwrap())
        .unwrap();

    assert!(slide.contains("New Header spans two"));
    assert_eq!(
        slide
            .matches(r#"<a:tcPr anchor="ctr" marL="91440" marR="91440"/>"#)
            .count(),
        7
    );
    assert!(slide.contains(r#"<a:solidFill><a:srgbClr val="DDEBF7"/></a:solidFill>"#));
    assert!(slide.contains(r#"<a:lnB w="19050">"#));
    assert!(slide.contains(r#"<a:gridCol w="1828800"/>"#));
    assert!(slide.contains(r#"<a:tc rowSpan="2">"#));
}

#[test]
fn every_cell_is_addressable_in_row_major_order() {
    let session = DeckSession::open(DECK, 34102).unwrap();
    let snapshot = session.snapshot().unwrap();
    let stories: Vec<_> = snapshot.slides[0].shapes[0]
        .text_stories
        .iter()
        .map(|story| story.id.clone())
        .collect();
    assert_eq!(stories.len(), 9);
    assert_eq!(stories[1], "story:slide:0:256:shape:0:table:0:1");
    assert_eq!(stories[8], "story:slide:0:256:shape:0:table:2:2");
    assert_eq!(session.story(&stories[3]).unwrap().plain_text(), "Tall");
}
