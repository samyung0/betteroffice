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

fn cell(row: usize, column: usize) -> String {
    format!("story:slide:0:256:shape:0:table:{row}:{column}")
}

fn saved_slide(session: &DeckSession) -> String {
    let parts = ooxml_opc::unzip_parts(&session.save().unwrap()).unwrap();
    parts
        .into_iter()
        .find(|(path, _)| path == "ppt/slides/slide1.xml")
        .map(|(_, bytes)| String::from_utf8(bytes).unwrap())
        .unwrap()
}

#[test]
fn two_peers_editing_table_cells_at_once_converge() {
    let left = DeckSession::open(DECK, 34103).unwrap();
    let right = DeckSession::open(DECK, 34104).unwrap();
    let baseline = left.encode_state_vector_v1();
    let style = TextStyle::default();
    let left_ctx = EditCtx::local("left");
    let right_ctx = EditCtx::local("right");

    // Both type into the header cell, one at each end; each also edits a cell
    // of their own.
    left.insert_text(&left_ctx, &cell(0, 0), 0, "Left ", &style)
        .unwrap();
    left.insert_text(&left_ctx, &cell(1, 1), 6, " row", &style)
        .unwrap();
    right
        .insert_text(&right_ctx, &cell(0, 0), 16, " right", &style)
        .unwrap();
    right
        .replace_text(&right_ctx, &cell(2, 2), 0, 4, "Final", &style)
        .unwrap();

    let from_left = left.encode_diff_v1(&baseline).unwrap();
    let from_right = right.encode_diff_v1(&baseline).unwrap();
    left.apply_update_v1(&from_right).unwrap();
    right.apply_update_v1(&from_left).unwrap();

    assert_eq!(
        left.encode_state_vector_v1(),
        right.encode_state_vector_v1()
    );
    assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
    let text = |row, column| left.story(&cell(row, column)).unwrap().plain_text();
    assert_eq!(text(0, 0), "Left Header spans two right");
    assert_eq!(text(1, 1), "Middle row");
    assert_eq!(text(2, 2), "Final");
    assert_eq!(text(1, 0), "Tall");

    let slide = saved_slide(&left);
    assert_eq!(slide, saved_slide(&right));
    for expected in ["Left Header spans two right", "Middle row", "Final"] {
        assert!(slide.contains(expected), "{expected} is missing: {slide}");
    }
    assert!(slide.contains(r#"<a:tc gridSpan="2">"#));
    assert!(slide.contains(r#"<a:tc rowSpan="2">"#));
    assert!(slide.contains(r#"<a:lnB w="19050">"#));
}
