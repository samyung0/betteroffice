//! The chat agent's PPTX slide and formatting commands (Epo 2026-10-20):
//! `add_slide` with its layout's text placeholders, `delete_slide` naming
//! the slide's text, `move_slide`, `add_text_box` inside the slide, and a
//! paragraph's `format_text` and `set_alignment`.

use office_service::{
    Alignment, Checkpoint, Command, CommandResult, Determinism, EditCode, Error, Format,
    InspectQuery, TextFormat, apply_commands, export, inspect_scoped, seed, sha256_hex,
};
use serde_json::{Value, json};

/// 50 slides on "Title and Content" (PowerPoint's default layouts).
const DECK: &[u8] = include_bytes!("../../../poc/fixtures/deck-50.pptx");
const LAYOUT: &str = "ppt/slideLayouts/slideLayout";

struct Deck {
    base: Vec<u8>,
    sha: String,
    state: Vec<u8>,
}

impl Deck {
    fn new() -> Self {
        let state = seed(Format::Pptx, DECK).unwrap();
        Self {
            sha: sha256_hex(DECK),
            base: DECK.to_vec(),
            state,
        }
    }

    fn checkpoint(&self) -> Checkpoint<'_> {
        Checkpoint {
            format: Format::Pptx,
            schema_version: 1,
            base_sha256: &self.sha,
            state: &self.state,
        }
    }

    fn apply(&mut self, commands: &[Command]) -> Result<CommandResult, Error> {
        let result = apply_commands(&self.base, self.checkpoint(), commands)?;
        self.state = result.state.clone();
        Ok(result)
    }

    fn read(&self, slide: Option<u32>) -> Value {
        inspect_scoped(
            &self.base,
            self.checkpoint(),
            &InspectQuery {
                slide,
                start: Some(0),
                count: Some(200),
                ..InspectQuery::default()
            },
        )
        .unwrap()
    }

    /// The slides' ids in deck order.
    fn slides(&self) -> Vec<String> {
        self.read(None)["slides"]
            .as_array()
            .unwrap()
            .iter()
            .map(|slide| slide["id"].as_str().unwrap().to_owned())
            .collect()
    }

    /// A slide's text as `delete_slide` expects it.
    fn slide_text(&self, number: u32) -> String {
        self.read(Some(number))["paragraphs"]
            .as_array()
            .unwrap()
            .iter()
            .map(|paragraph| paragraph["text"].as_str().unwrap())
            .collect::<Vec<_>>()
            .join("\n")
    }

    /// The exported slide whose text holds `text`, as XML.
    fn slide_xml(&self, text: &str) -> String {
        let saved = export(
            &self.base,
            self.checkpoint(),
            Determinism {
                seed: &"0".repeat(64),
                now: "2026-10-22T00:00:00.000Z",
            },
        )
        .unwrap();
        ooxml_opc::unzip_parts(&saved)
            .unwrap()
            .into_iter()
            .map(|(_, bytes)| String::from_utf8_lossy(&bytes).into_owned())
            .find(|xml| xml.starts_with("<?xml") && xml.contains("<p:sld ") && xml.contains(text))
            .unwrap_or_else(|| panic!("no exported slide holds {text:?}"))
    }
}

fn add_slide(after: Option<&str>, before: Option<&str>, layout: Option<&str>) -> Command {
    Command::AddSlide {
        after_slide: after.map(str::to_owned),
        before_slide: before.map(str::to_owned),
        layout: layout.map(str::to_owned),
    }
}

fn invalid(message: &str) -> Error {
    Error::Edit {
        code: EditCode::InvalidInput,
        message: message.into(),
        details: None,
    }
}

fn shapes(read: &Value) -> Vec<(String, Value)> {
    read["shapes"]
        .as_array()
        .unwrap()
        .iter()
        .map(|shape| {
            (
                shape["id"].as_str().unwrap().to_owned(),
                shape["placeholder"].clone(),
            )
        })
        .collect()
}

/// A new slide takes the layout of the slide it follows (else the one it
/// precedes) or the one named, with that layout's title and text
/// placeholders, empty and ready for `set_shape_text`; a blank layout
/// gives none.
#[test]
fn add_slide_places_a_slide_with_its_layouts_text_placeholders() {
    let mut deck = Deck::new();
    let ids = deck.slides();
    deck.apply(&[
        add_slide(Some(&ids[0]), None, None),
        add_slide(None, Some(&ids[0]), Some(&format!("{LAYOUT}1.xml"))),
        add_slide(None, None, Some(&format!("{LAYOUT}7.xml"))),
    ])
    .unwrap();
    let list = deck.read(None);
    assert_eq!(list["total"], 53);
    let layouts: Vec<&Value> = list["slides"]
        .as_array()
        .unwrap()
        .iter()
        .map(|slide| &slide["layout"])
        .collect();
    assert_eq!(
        [layouts[0], layouts[1], layouts[2], layouts[52]],
        [
            &json!("Title Slide"),
            &json!("Title and Content"),
            &json!("Title and Content"),
            &json!("Blank")
        ]
    );
    assert_eq!(deck.slides()[1], ids[0]);

    let title_slide = deck.read(Some(1));
    assert_eq!(
        shapes(&title_slide)
            .into_iter()
            .map(|(_, placeholder)| placeholder)
            .collect::<Vec<_>>(),
        [json!("ctrTitle"), json!("subTitle")]
    );
    assert_eq!(
        title_slide["shapes"][0]["box"],
        json!([54.0, 167.8, 612.0, 115.8])
    );
    assert_eq!(deck.read(Some(53))["shapes"], json!([]));
    let content = deck.read(Some(3));
    let (title, _) = &shapes(&content)[0];
    assert_eq!(
        content["shapes"][0]["box"],
        json!([36.0, 21.6, 648.0, 90.0])
    );
    deck.apply(&[Command::SetShapeText {
        shape_id: title.clone(),
        expected_text: String::new(),
        paragraphs: vec!["Added".into()],
    }])
    .unwrap();
    assert!(
        deck.slide_xml(">Added</a:t>")
            .contains(r#"<p:ph type="title"/>"#)
    );
}

#[test]
fn add_slide_refuses_an_unknown_layout_and_two_anchors() {
    let mut deck = Deck::new();
    let ids = deck.slides();
    assert_eq!(
        deck.apply(&[add_slide(None, None, Some("ppt/slideLayouts/nope.xml"))])
            .unwrap_err(),
        Error::Edit {
            code: EditCode::UnavailableTarget,
            message:
                "the deck has no layout ppt/slideLayouts/nope.xml; the deck read lists its layouts"
                    .into(),
            details: None,
        }
    );
    assert_eq!(
        deck.apply(&[add_slide(Some(&ids[0]), Some(&ids[1]), None)])
            .unwrap_err(),
        invalid("add_slide names at most one of after_slide and before_slide")
    );
    assert_eq!(
        deck.apply(&[add_slide(Some("slide:99:1"), None, None)])
            .unwrap_err(),
        Error::Edit {
            code: EditCode::UnavailableTarget,
            message: "slide slide:99:1 is no longer in the deck".into(),
            details: None,
        }
    );
}

/// A delete names the slide's text, compared without blank lines and the
/// spaces around a line; a stale one carries the text as it reads now.
#[test]
fn delete_slide_removes_a_slide_reading_its_expected_text() {
    let mut deck = Deck::new();
    let ids = deck.slides();
    let text = deck.slide_text(2);
    let stale = deck
        .apply(&[Command::DeleteSlide {
            slide_id: ids[1].clone(),
            expected_text: "2. Less with behind".into(),
        }])
        .unwrap_err();
    assert_eq!(
        stale,
        Error::Edit {
            code: EditCode::StaleTarget,
            message: "the slide's text differs from expected_text".into(),
            details: Some(json!({ "block": { "id": ids[1], "type": "slide", "text": text } })),
        }
    );
    let edit = deck
        .apply(&[Command::DeleteSlide {
            slide_id: ids[1].clone(),
            expected_text: format!("\n  {}\n\n", text.replace('\n', " \n\n")),
        }])
        .unwrap();
    assert_eq!(edit.inverse, None);
    let after = deck.slides();
    assert_eq!((after.len(), &after[1]), (49, &ids[2]));
}

#[test]
fn a_deck_keeps_its_last_slide() {
    let mut deck = Deck::new();
    let ids = deck.slides();
    let commands: Vec<Command> = ids[1..]
        .iter()
        .enumerate()
        .map(|(index, id)| Command::DeleteSlide {
            slide_id: id.clone(),
            expected_text: deck.slide_text(index as u32 + 2),
        })
        .collect();
    for chunk in commands.chunks(10) {
        deck.apply(chunk).unwrap();
    }
    assert_eq!(
        deck.apply(&[Command::DeleteSlide {
            slide_id: ids[0].clone(),
            expected_text: deck.slide_text(1),
        }])
        .unwrap_err(),
        invalid("a deck keeps at least one slide")
    );
}

#[test]
fn move_slide_puts_a_slide_after_or_before_another() {
    let mut deck = Deck::new();
    let ids = deck.slides();
    deck.apply(&[
        Command::MoveSlide {
            slide_id: ids[0].clone(),
            after_slide: Some(ids[2].clone()),
            before_slide: None,
        },
        Command::MoveSlide {
            slide_id: ids[49].clone(),
            after_slide: None,
            before_slide: Some(ids[1].clone()),
        },
    ])
    .unwrap();
    assert_eq!(
        deck.slides()[..5],
        [
            ids[49].clone(),
            ids[1].clone(),
            ids[2].clone(),
            ids[0].clone(),
            ids[3].clone()
        ]
    );
    for (command, message) in [
        (
            Command::MoveSlide {
                slide_id: ids[1].clone(),
                after_slide: Some(ids[1].clone()),
                before_slide: None,
            },
            "a slide cannot move next to itself",
        ),
        (
            Command::MoveSlide {
                slide_id: ids[1].clone(),
                after_slide: None,
                before_slide: None,
            },
            "move_slide names after_slide or before_slide",
        ),
    ] {
        assert_eq!(deck.apply(&[command]).unwrap_err(), invalid(message));
    }
}

/// A text box lands on the slide at its box in points, with its paragraphs;
/// one off the slide is refused.
#[test]
fn add_text_box_places_text_inside_the_slide() {
    let mut deck = Deck::new();
    let slide = deck.slides()[0].clone();
    let text_box = |x: f64, w: f64| Command::AddTextBox {
        slide_id: slide.clone(),
        paragraphs: vec!["Remember".into(), "the units".into()],
        x,
        y: 450.0,
        w,
        h: 60.0,
    };
    deck.apply(&[text_box(36.0, 300.5)]).unwrap();
    let read = deck.read(Some(1));
    let added = read["shapes"].as_array().unwrap().last().unwrap().clone();
    assert_eq!(
        (&added["kind"], &added["box"]),
        (&json!("text"), &json!([36.0, 450.0, 300.5, 60.0]))
    );
    let texts: Vec<&Value> = read["paragraphs"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|paragraph| paragraph["shape"] == added["id"])
        .map(|paragraph| &paragraph["text"])
        .collect();
    assert_eq!(texts, [&json!("Remember"), &json!("the units")]);
    assert!(
        deck.slide_xml(">the units</a:t>")
            .contains(r#"<a:off x="457200" y="5715000"/>"#)
    );
    assert_eq!(
        deck.apply(&[text_box(500.0, 300.0)]).unwrap_err(),
        invalid(
            "the box (x 500, y 450, w 300, h 60 points) must lie on the 720 × 540 point slide, with a positive size"
        )
    );
}

/// A PPTX span takes bold, italic, underline, strikethrough, colour, size
/// and font, not DOCX's highlight; an alignment sets the paragraph's.
#[test]
fn format_text_and_set_alignment_write_a_paragraphs_runs_and_alignment() {
    let mut deck = Deck::new();
    let read = deck.read(Some(1));
    let title = read["paragraphs"][0]["id"].as_str().unwrap().to_owned();
    let format = TextFormat {
        bold: Some(true),
        underline: Some(true),
        strikethrough: Some(true),
        color: Some("1f4e79".into()),
        font_size: Some(28.0),
        font_family: Some("Arial".into()),
        ..TextFormat::default()
    };
    deck.apply(&[
        Command::FormatText {
            target_id: title.clone(),
            expected_text: "Mendel".into(),
            format,
        },
        Command::SetAlignment {
            target_id: title.clone(),
            alignment: Alignment::Center,
        },
    ])
    .unwrap();
    let xml = deck.slide_xml(">Mendel</a:t>");
    let run = &xml[xml[..xml.find(">Mendel</a:t>").unwrap()]
        .rfind("<a:r>")
        .unwrap()..];
    for attribute in [
        r#"b="1""#,
        r#"u="sng""#,
        r#"strike="sngStrike""#,
        r#"sz="2800""#,
        "1F4E79",
        r#"typeface="Arial""#,
    ] {
        assert!(
            run[..run.find("</a:r>").unwrap()].contains(attribute),
            "{attribute}: {run}"
        );
    }
    assert!(xml.contains(r#"algn="ctr""#));

    let highlight = TextFormat {
        highlight: Some("yellow".into()),
        ..TextFormat::default()
    };
    assert_eq!(
        deck.apply(&[Command::FormatText {
            target_id: title.clone(),
            expected_text: "Mendel".into(),
            format: highlight,
        }])
        .unwrap_err(),
        invalid(
            "highlight formats Word text; PowerPoint text takes bold, italic, underline, strikethrough, color, font_size and font_family"
        )
    );
    assert_eq!(
        deck.apply(&[Command::FormatText {
            target_id: title.clone(),
            expected_text: "Mendeleev".into(),
            format: TextFormat {
                italic: Some(true),
                ..TextFormat::default()
            },
        }])
        .unwrap_err(),
        Error::Edit {
            code: EditCode::StaleTarget,
            message: "expected text was not found in the paragraph".into(),
            details: Some(json!({ "block": {
                "id": title,
                "type": "paragraph",
                "text": "1. Look further moment me second Mendel being",
            } })),
        }
    );
}

#[test]
fn the_slide_commands_parse_from_their_wire_json() {
    let commands: Vec<Command> = serde_json::from_value(json!([
        { "type": "add_slide", "afterSlide": "s1", "layout": "l" },
        { "type": "delete_slide", "slideId": "s1", "expectedText": "a" },
        { "type": "move_slide", "slideId": "s1", "beforeSlide": "s2" },
        { "type": "add_text_box", "slideId": "s1", "paragraphs": ["x"], "x": 1, "y": 2.5, "w": 3, "h": 4 },
        { "type": "format_text", "targetId": "p", "expectedText": "a", "italic": true },
    ]))
    .unwrap();
    assert_eq!(
        commands,
        [
            add_slide(Some("s1"), None, Some("l")),
            Command::DeleteSlide {
                slide_id: "s1".into(),
                expected_text: "a".into(),
            },
            Command::MoveSlide {
                slide_id: "s1".into(),
                after_slide: None,
                before_slide: Some("s2".into()),
            },
            Command::AddTextBox {
                slide_id: "s1".into(),
                paragraphs: vec!["x".into()],
                x: 1.0,
                y: 2.5,
                w: 3.0,
                h: 4.0,
            },
            Command::FormatText {
                target_id: "p".into(),
                expected_text: "a".into(),
                format: TextFormat {
                    italic: Some(true),
                    ..TextFormat::default()
                },
            },
        ]
    );
}
