//! A slide's layout placeholders (`DeckSession::add_layout_placeholders`,
//! the chat agent's `add_slide`, Epo 2026-10-20) save as `p:ph` copies of
//! the layout's and reopen as those placeholders at the layout's box; the
//! editor's own added shapes save as before.

use std::collections::BTreeMap;

use pptx_edit::{DeckSession, EditCtx, PresetShapeDraft, ShapeDraft, ShapeRect, TextStyle};
use sha2::{Digest, Sha256};

const DEMO: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
/// PowerPoint's default layouts: "Title Slide" places its placeholders
/// itself, "Title and Content" inherits the master's.
const DECK: &[u8] = include_bytes!("../../../poc/fixtures/deck-50.pptx");

/// A reopened placeholder: its type, index, box and text.
type Placed = (Option<String>, Option<u32>, [i64; 4], String);

fn context() -> EditCtx {
    EditCtx::local("test")
}

fn parts(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    ooxml_opc::unzip_parts(bytes).unwrap().into_iter().collect()
}

/// The editor's text box and preset shape write no `p:ph`: the saved slide
/// is byte for byte what the save wrote before placeholders could be added
/// (the digest is the same with the placeholder change reverted).
#[test]
fn added_shapes_without_a_placeholder_save_the_same_bytes() {
    let session = DeckSession::open(DEMO, 7).unwrap();
    let slide = session.snapshot().unwrap().slides[0].id.clone();
    session
        .add_text_box(
            &context(),
            &slide,
            &ShapeDraft {
                name: "Note".to_owned(),
                rect: ShapeRect {
                    x: 100_000,
                    y: 200_000,
                    width: 3_000_000,
                    height: 500_000,
                },
                text: "Saved note".to_owned(),
                style: TextStyle {
                    bold: Some(true),
                    ..TextStyle::default()
                },
            },
        )
        .unwrap();
    session
        .add_shape(
            &context(),
            &slide,
            &PresetShapeDraft {
                name: "Badge".to_owned(),
                geometry: "roundRect".to_owned(),
                rect: ShapeRect {
                    x: 400_000,
                    y: 400_000,
                    width: 1_000_000,
                    height: 1_000_000,
                },
                fill: Some("#00AAFF".to_owned()),
            },
        )
        .unwrap();
    let saved = parts(&session.save().unwrap());
    let xml = &saved["ppt/slides/slide1.xml"];
    assert!(!String::from_utf8_lossy(xml).contains("<p:ph"));
    let digest: String = Sha256::digest(xml)
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect();
    assert_eq!(
        digest,
        "376aada64bd38986d3d4e61d8c0f1cd7824ad4127ca8c74aa51bb153342fdd9b"
    );
}

/// Each layout's text placeholders (not its date, footer and slide number)
/// become empty text shapes tied to them; saved and reopened, they read as
/// those placeholders at the layout's box, else the master's.
#[test]
fn layout_placeholders_save_as_placeholders_and_reopen_at_their_box() {
    let session = DeckSession::open(DECK, 7).unwrap();
    let mut added = Vec::new();
    for layout in ["slideLayout1", "slideLayout2"] {
        let path = format!("ppt/slideLayouts/{layout}.xml");
        let slide = session
            .insert_slide(&context(), 50 + added.len() as u32, Some(&path))
            .unwrap()
            .slide_id;
        let shapes = session.add_layout_placeholders(&context(), &slide).unwrap();
        assert_eq!(shapes.len(), 2);
        session
            .replace_text(
                &context(),
                &format!("story:{}:0", shapes[0]),
                0,
                0,
                &format!("{layout} title"),
                &TextStyle::default(),
            )
            .unwrap();
        added.push(slide);
    }
    let saved = session.save().unwrap();
    let files = parts(&saved);
    let slide_xml = |text: &str| {
        files
            .values()
            .map(|bytes| String::from_utf8_lossy(bytes).into_owned())
            .find(|xml| xml.contains(&format!(">{text}</a:t>")))
            .unwrap()
    };
    let first = slide_xml("slideLayout1 title");
    assert!(first.contains(r#"<p:ph type="ctrTitle"/>"#), "{first}");
    assert!(
        first.contains(r#"<p:ph idx="1" type="subTitle"/>"#),
        "{first}"
    );
    let second = slide_xml("slideLayout2 title");
    assert!(second.contains(r#"<p:ph type="title"/>"#) && second.contains(r#"<p:ph idx="1"/>"#));

    let reopened = DeckSession::open(&saved, 8).unwrap().snapshot().unwrap();
    let boxes: Vec<Vec<Placed>> = reopened.slides[50..]
        .iter()
        .map(|slide| {
            slide
                .shapes
                .iter()
                .map(|shape| {
                    let placeholder = shape.placeholder.clone().unwrap();
                    (
                        placeholder.placeholder_type,
                        placeholder.index,
                        [shape.x, shape.y, shape.width, shape.height],
                        shape.text_stories[0].plain_text(),
                    )
                })
                .collect()
        })
        .collect();
    assert_eq!(
        boxes,
        [
            vec![
                (
                    Some("ctrTitle".to_owned()),
                    None,
                    [685_800, 2_130_425, 7_772_400, 1_470_025],
                    "slideLayout1 title".to_owned()
                ),
                (
                    Some("subTitle".to_owned()),
                    Some(1),
                    [1_371_600, 3_886_200, 6_400_800, 1_752_600],
                    String::new()
                ),
            ],
            vec![
                (
                    Some("title".to_owned()),
                    None,
                    [457_200, 274_638, 8_229_600, 1_143_000],
                    "slideLayout2 title".to_owned()
                ),
                (
                    None,
                    Some(1),
                    [457_200, 1_600_200, 8_229_600, 4_525_963],
                    String::new()
                ),
            ],
        ]
    );
}
