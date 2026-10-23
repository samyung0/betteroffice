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
        Self::of(DECK.to_vec())
    }

    fn of(base: Vec<u8>) -> Self {
        let state = seed(Format::Pptx, &base).unwrap();
        Self {
            sha: sha256_hex(&base),
            base,
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

    /// A slide's text as `delete_slide` expects it: its paragraphs, then
    /// its notes as the read shows them.
    fn slide_text(&self, number: u32) -> String {
        let read = self.read(Some(number));
        let mut lines: Vec<&str> = read["paragraphs"]
            .as_array()
            .unwrap()
            .iter()
            .map(|paragraph| paragraph["text"].as_str().unwrap())
            .collect();
        lines.extend(read["slide"]["notes"].as_str());
        lines.join("\n")
    }

    /// A person writes the slide's speaker notes in the editor.
    fn set_notes(&mut self, slide_id: &str, notes: &str) {
        let deck = pptx_edit::DeckSession::open_from_update_with_source(&self.state, &self.base, 9)
            .unwrap();
        deck.set_slide_notes(&pptx_edit::EditCtx::local("person"), slide_id, notes)
            .unwrap();
        self.state = deck.encode_state_as_update_v1();
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

/// A layout placeholder with no size of its own nor on the master refuses
/// the slide as invalid input, adding nothing (review M2).
#[test]
fn add_slide_refuses_a_layout_placeholder_with_no_size() {
    let strip = |xml: String| {
        let mut out = String::new();
        let mut rest = xml.as_str();
        while let Some(at) = rest.find("<a:xfrm>") {
            out.push_str(&rest[..at]);
            let end = rest[at..].find("</a:xfrm>").unwrap() + at + "</a:xfrm>".len();
            rest = &rest[end..];
        }
        out + rest
    };
    let parts: Vec<(String, Vec<u8>)> = ooxml_opc::unzip_parts(DECK)
        .unwrap()
        .into_iter()
        .map(|(name, bytes)| {
            if name != "ppt/slideMasters/slideMaster1.xml" {
                return (name, bytes);
            }
            (name, strip(String::from_utf8(bytes).unwrap()).into_bytes())
        })
        .collect();
    let mut deck = Deck::of(ooxml_opc::rezip_parts(&parts).unwrap());
    let ids = deck.slides();
    assert_eq!(
        deck.apply(&[add_slide(
            Some(&ids[0]),
            None,
            Some(&format!("{LAYOUT}2.xml"))
        )])
        .unwrap_err(),
        invalid("invalid shape geometry: the layout's placeholder \"Title 1\" has no size")
    );
    assert_eq!(deck.slides(), ids);
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
            message: "the slide's text or notes differ from expected_text".into(),
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
            text: "Mendel".into(),
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

    // A noncharacter a file cannot store (review L4).
    assert_eq!(
        deck.apply(&[Command::FormatText {
            target_id: title.clone(),
            text: "Mendel".into(),
            format: TextFormat {
                font_family: Some("Ari\u{fffe}al".into()),
                ..TextFormat::default()
            },
        }])
        .unwrap_err(),
        invalid("font_family is a font name of 1 to 31 characters")
    );
    let highlight = TextFormat {
        highlight: Some("yellow".into()),
        ..TextFormat::default()
    };
    assert_eq!(
        deck.apply(&[Command::FormatText {
            target_id: title.clone(),
            text: "Mendel".into(),
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
            text: "Mendeleev".into(),
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
        { "type": "format_text", "targetId": "p", "text": "a", "italic": true },
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
                text: "a".into(),
                format: TextFormat {
                    italic: Some(true),
                    ..TextFormat::default()
                },
            },
        ]
    );
}

/// Each PPTX command has its own per-call count (Epo 2026-10-22): the
/// commands before the first past one apply and the notice names the rest;
/// shape and text box text share one budget of paragraphs and characters.
#[test]
fn a_call_stops_at_its_first_command_past_a_pptx_limit() {
    let mut deck = Deck::new();
    let adds: Vec<Command> = (0..21).map(|_| add_slide(None, None, None)).collect();
    let result = deck.apply(&adds).unwrap();
    assert_eq!(
        result.notice.as_deref(),
        Some(
            "Not applied: the command at index 20 (counting from 0), past the limit of 20 slides added a call; send them in another call."
        )
    );
    assert_eq!(deck.read(None)["total"], 70);

    let mut commands: Vec<Command> = (1..=10)
        .map(|number| {
            let read = deck.read(Some(number));
            let body = read["shapes"][1]["id"].as_str().unwrap().to_owned();
            let expected: Vec<&str> = read["paragraphs"]
                .as_array()
                .unwrap()
                .iter()
                .filter(|paragraph| paragraph["shape"] == body.as_str())
                .map(|paragraph| paragraph["text"].as_str().unwrap())
                .collect();
            Command::SetShapeText {
                shape_id: body.clone(),
                expected_text: expected.join("\n"),
                paragraphs: vec!["x".repeat(100); 50],
            }
        })
        .collect();
    commands.push(Command::AddTextBox {
        slide_id: deck.slides()[0].clone(),
        paragraphs: vec!["y".into()],
        x: 0.0,
        y: 0.0,
        w: 100.0,
        h: 50.0,
    });
    let result = deck.apply(&commands).unwrap();
    assert_eq!(
        result.notice.as_deref(),
        Some(
            "Not applied: the command at index 10 (counting from 0), past the limit of 500 paragraphs of shape and text box text a call; send them in another call."
        )
    );
    assert_eq!(result.touched.len(), 10);
}

/// What a delete_slide's or set_shape_text's Undo re-creates is measured
/// before it applies: the slide whole, the shape's story.
#[test]
fn a_slide_delete_and_a_shape_write_measure_what_their_undo_recreates() {
    let mut deck = Deck::new();
    let ids = deck.slides();
    let text = deck.slide_text(1);
    let title = deck.read(Some(2))["shapes"][0]["id"]
        .as_str()
        .unwrap()
        .to_owned();
    let result = deck
        .apply(&[
            Command::DeleteSlide {
                slide_id: ids[0].clone(),
                expected_text: text.clone(),
            },
            Command::SetShapeText {
                shape_id: title,
                expected_text: "2. Less with behind problem areas".into(),
                paragraphs: vec!["New".into()],
            },
            Command::MoveSlide {
                slide_id: ids[2].clone(),
                after_slide: Some(ids[3].clone()),
                before_slide: None,
            },
        ])
        .unwrap();
    let [slide, story, moved] = result.touched[..] else {
        panic!("{:?}", result.touched);
    };
    assert!(
        slide.items > 30 && slide.content > text.len() as u64,
        "{slide:?}"
    );
    assert!(story.items > 1 && story.content >= 33, "{story:?}");
    assert_eq!((moved.items, moved.content), (0, 0));
}

/// An add_slide measures the slide it wrote, after it applies: what a
/// delete of that slide measures, since the add's Redo re-creates it whole.
#[test]
fn an_added_slide_measures_the_slide_it_wrote() {
    let mut deck = Deck::new();
    let first = deck.slides().remove(0);
    let layout = format!("{LAYOUT}2.xml");
    let added = deck
        .apply(&[add_slide(Some(&first), None, Some(&layout))])
        .unwrap()
        .touched[0];
    let deleted = deck
        .apply(&[Command::DeleteSlide {
            slide_id: deck.slides()[1].clone(),
            expected_text: deck.slide_text(2),
        }])
        .unwrap()
        .touched[0];
    assert!(added.items > 10, "{added:?}");
    assert_eq!(added, deleted);
}

/// A slide's speaker notes (Epo 2026-10-22): the slide read's first page
/// shows them, past 2,000 characters cut with a marker
/// counting the rest, the deck read none; delete_slide compares them as
/// shown, so notes edited after the read make it stale and unchanged ones
/// delete.
#[test]
fn delete_slide_names_the_notes_the_read_shows() {
    let mut deck = Deck::new();
    let ids = deck.slides();
    deck.set_notes(&ids[1], "Greet the room.\nThen the demo.");
    let read = deck.read(Some(2));
    assert_eq!(read["slide"]["notes"], "Greet the room.\nThen the demo.");
    let later = inspect_scoped(
        &deck.base,
        deck.checkpoint(),
        &InspectQuery {
            slide: Some(2),
            start: Some(1),
            count: Some(200),
            ..InspectQuery::default()
        },
    )
    .unwrap();
    assert_eq!(later["slide"].get("notes"), None);

    let long: String = (0..2_500)
        .map(|n| if n % 50 == 49 { '.' } else { 'x' })
        .collect();
    deck.set_notes(&ids[1], &long);
    let shown = format!("{}… [notes cut: 500 more characters]", &long[..2_000]);
    assert_eq!(deck.read(Some(2))["slide"]["notes"], shown);
    assert_eq!(deck.read(None)["slides"][1].get("notes"), None);

    // Edited after the read, in the part the read showed.
    let text = deck.slide_text(2);
    let edited = format!("y{}", &long[1..]);
    deck.set_notes(&ids[1], &edited);
    let stale = deck
        .apply(&[Command::DeleteSlide {
            slide_id: ids[1].clone(),
            expected_text: text,
        }])
        .unwrap_err();
    let Error::Edit { code, details, .. } = stale else {
        panic!("{stale:?}");
    };
    assert_eq!(code, EditCode::StaleTarget);
    let now = details.unwrap()["block"]["text"]
        .as_str()
        .unwrap()
        .to_owned();
    assert!(now.ends_with("… [notes cut: 500 more characters]") && now.contains("\nyxx"));

    deck.apply(&[Command::DeleteSlide {
        slide_id: ids[1].clone(),
        expected_text: deck.slide_text(2),
    }])
    .unwrap();
    assert_eq!(deck.slides().len(), 49);
}
