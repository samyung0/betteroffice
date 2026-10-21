//! The chat agent's PPTX read and `set_shape_text` (Epo 2026-10-20): the
//! deck read lists slides with their layouts and shapes, a slide read its
//! shapes and paragraphs; `set_shape_text` replaces a shape's whole text.

use office_service::{
    Checkpoint, Command, CommandResult, Determinism, EditCode, Error, Format, InspectQuery,
    apply_commands, export, inspect_scoped, seed, sha256_hex,
};
use serde_json::{Value, json};

/// 50 slides on "Title and Content", each a title and a body of bulleted
/// paragraphs; its placeholders spell out no position.
const DECK: &[u8] = include_bytes!("../../../poc/fixtures/deck-50.pptx");

struct Deck {
    base: Vec<u8>,
    sha: String,
    state: Vec<u8>,
}

impl Deck {
    fn new(base: Vec<u8>) -> Self {
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

    fn read(&self, slide: Option<u32>, start: u32, count: u32) -> Result<Value, Error> {
        inspect_scoped(
            &self.base,
            self.checkpoint(),
            &InspectQuery {
                slide,
                start: Some(start),
                count: Some(count),
                ..InspectQuery::default()
            },
        )
    }

    /// A slide's paragraphs as (id, shape, text).
    fn paragraphs(&self, slide: u32) -> Vec<(String, String, String)> {
        let read = self.read(Some(slide), 0, 200).unwrap();
        read["paragraphs"]
            .as_array()
            .unwrap()
            .iter()
            .map(|paragraph| {
                let text = |key: &str| paragraph[key].as_str().unwrap().to_owned();
                (text("id"), text("shape"), text("text"))
            })
            .collect()
    }

    /// A shape's paragraphs' text on a slide.
    fn shape_text(&self, slide: u32, shape: &str) -> Vec<String> {
        self.paragraphs(slide)
            .into_iter()
            .filter(|(_, owner, _)| owner == shape)
            .map(|(_, _, text)| text)
            .collect()
    }

    fn part(&self, name: &str) -> String {
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
            .find(|(path, _)| path == name)
            .map(|(_, bytes)| String::from_utf8(bytes).unwrap())
            .unwrap()
    }
}

fn set_text(shape: &str, expected: &str, paragraphs: &[&str]) -> Command {
    Command::SetShapeText {
        shape_id: shape.into(),
        expected_text: expected.into(),
        paragraphs: paragraphs.iter().map(|text| (*text).to_owned()).collect(),
    }
}

fn invalid(message: &str) -> Error {
    Error::Edit {
        code: EditCode::InvalidInput,
        message: message.into(),
        details: None,
    }
}

/// The deck read pages its slides, each with its layout and shapes (kind,
/// placeholder, first line), the first page listing the layouts too.
#[test]
fn the_deck_read_pages_slides_with_their_layouts_and_shapes() {
    let deck = Deck::new(DECK.to_vec());
    let page = deck.read(None, 0, 2).unwrap();
    assert_eq!(page["total"], 50);
    assert_eq!(page["nextStart"], 2);
    assert_eq!(
        (&page["width"], &page["height"]),
        (&json!(720.0), &json!(540.0))
    );
    let layouts = page["layouts"].as_array().unwrap();
    assert_eq!(layouts.len(), 11);
    assert!(layouts.contains(&json!({
        "id": "ppt/slideLayouts/slideLayout2.xml",
        "name": "Title and Content",
    })));
    let first = &page["slides"][0];
    assert_eq!(first["number"], 1);
    assert_eq!(first["layout"], "Title and Content");
    let shapes = first["shapes"].as_array().unwrap();
    assert_eq!(shapes.len(), 2);
    assert_eq!(shapes[0]["kind"], "text");
    assert_eq!(shapes[0]["placeholder"], "title");
    assert_eq!(
        shapes[0]["text"],
        "1. Look further moment me second Mendel being"
    );
    assert_eq!(shapes[1]["placeholder"], "obj");
    let line = shapes[1]["text"].as_str().unwrap();
    assert!(line.ends_with('…') && line.chars().count() == 81, "{line}");

    let last = deck.read(None, 48, 200).unwrap();
    assert_eq!(last["slides"][1]["number"], 50);
    assert_eq!(last["nextStart"], Value::Null);
    assert_eq!(last.get("layouts"), None);
    assert_eq!(
        deck.read(None, 0, 0).unwrap_err(),
        invalid("a PPTX read names a count of 1 to 200")
    );
}

/// A slide read lists its shapes, boxed in points (a placeholder that
/// spells out no position at its layout's), and pages its paragraphs in
/// shape order with their levels.
#[test]
fn a_slide_read_lists_shapes_and_pages_their_paragraphs() {
    let deck = Deck::new(DECK.to_vec());
    let read = deck.read(Some(2), 0, 3).unwrap();
    assert_eq!(read["slide"]["number"], 2);
    assert_eq!(
        read["slide"]["layout"],
        json!({ "id": "ppt/slideLayouts/slideLayout2.xml", "name": "Title and Content" })
    );
    assert_eq!(read["slides"], 50);
    let shapes = read["shapes"].as_array().unwrap();
    assert_eq!(shapes[0]["placeholder"], "title");
    assert_eq!(shapes[0]["name"].as_str().map(str::is_empty), Some(false));
    let boxed = shapes[0]["box"].as_array().unwrap();
    assert!(boxed.iter().all(|value| value.as_f64().unwrap() >= 0.0));
    assert!(boxed[2].as_f64().unwrap() > 0.0 && boxed[3].as_f64().unwrap() > 0.0);
    let total = deck.paragraphs(2).len();
    assert_eq!(read["total"], total);
    assert_eq!(read["nextStart"], 3);
    let paragraphs = read["paragraphs"].as_array().unwrap();
    assert_eq!(paragraphs.len(), 3);
    assert_eq!(paragraphs[0]["shape"], shapes[0]["id"]);
    assert_eq!(paragraphs[0]["text"], "2. Less with behind problem areas");
    assert_eq!(paragraphs[1]["shape"], shapes[1]["id"]);

    let rest = deck.read(Some(2), 3, 200).unwrap();
    assert_eq!(rest["paragraphs"].as_array().unwrap().len(), total - 3);
    assert_eq!(rest["nextStart"], Value::Null);
    assert_eq!(
        deck.read(Some(51), 0, 60).unwrap_err(),
        invalid("slide 51 is not in the deck, which has 50 slides")
    );
}

/// The kept paragraphs keep their level and bullet, added ones continue
/// the last, and removed ones go; the stale expectation refuses with the
/// shape's text as it reads now.
#[test]
fn set_shape_text_replaces_a_shapes_paragraphs_keeping_their_levels() {
    let mut deck = Deck::new(DECK.to_vec());
    let read = deck.read(Some(1), 0, 200).unwrap();
    let body = read["shapes"][1]["id"].as_str().unwrap().to_owned();
    // Levels 0, 0, 0, 1, 0, 0, 0, 1.
    let before = deck.shape_text(1, &body);
    assert_eq!(before.len(), 8);
    let expected = before.join("\n");

    assert_eq!(
        deck.apply(&[set_text(&body, "Hand effect", &["x"])])
            .unwrap_err(),
        Error::Edit {
            code: EditCode::StaleTarget,
            message: "the shape's text differs from expected_text".into(),
            details: Some(json!({
                "block": { "id": body, "type": "shape", "text": expected }
            })),
        }
    );

    let more = [
        "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
    ];
    let edit = deck.apply(&[set_text(&body, &expected, &more)]).unwrap();
    assert_eq!(edit.inverse, None);
    assert_eq!(deck.shape_text(1, &body), more);
    let slide = deck.part("ppt/slides/slide1.xml");
    let level_one = |text: &str| {
        let chunk = slide
            .split("</a:p>")
            .find(|chunk| chunk.contains(&format!(">{text}</a:t>")))
            .unwrap();
        chunk[chunk.rfind("<a:p>").unwrap()..].contains("lvl=\"1\"")
    };
    assert_eq!(
        more.map(level_one),
        [false, false, false, true, false, false, false, true, true]
    );

    deck.apply(&[set_text(&body, &more.join("\n"), &["Only", ""])])
        .unwrap();
    assert_eq!(deck.shape_text(1, &body), ["Only", ""]);
    let title = read["shapes"][0]["id"].as_str().unwrap();
    deck.apply(&[set_text(
        title,
        "1. Look further moment me second Mendel being",
        &["New title"],
    )])
    .unwrap();
    assert_eq!(deck.shape_text(1, title), ["New title"]);
    assert!(
        deck.part("ppt/slides/slide1.xml")
            .contains(">New title</a:t>")
    );
}

#[test]
fn set_shape_text_refuses_line_breaks_and_text_over_the_limits() {
    let mut deck = Deck::new(DECK.to_vec());
    let read = deck.read(Some(1), 0, 1).unwrap();
    let title = read["shapes"][0]["id"].as_str().unwrap().to_owned();
    let expected = "1. Look further moment me second Mendel being";
    let long = "x".repeat(2_501);
    let cases: [(Vec<&str>, &str); 4] = [
        (
            vec![],
            "paragraphs holds the shape's text, one string a paragraph; [\"\"] empties it",
        ),
        (
            vec!["a"; 51],
            "paragraphs holds 51 paragraphs; a shape takes at most 50 a command",
        ),
        (
            vec![long.as_str(), long.as_str()],
            "paragraphs hold 5002 characters; a shape takes at most 5000 a command",
        ),
        (
            vec!["a\nb"],
            "text must be one paragraph without line breaks; insert one paragraph a command",
        ),
    ];
    for (paragraphs, message) in cases {
        assert_eq!(
            deck.apply(&[set_text(&title, expected, &paragraphs)])
                .unwrap_err(),
            invalid(message)
        );
    }
    assert_eq!(
        deck.apply(&[set_text("slide:0:256:shape:9", "", &["x"])])
            .unwrap_err(),
        Error::Edit {
            code: EditCode::UnavailableTarget,
            message: "the shape is no longer in the deck".into(),
            details: None,
        }
    );
}

/// A table's cells are edited by paragraph, an equation's text not at all;
/// the read marks both.
#[test]
fn set_shape_text_refuses_tables_and_equations() {
    let shapes = concat!(
        r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="80" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>"#,
        r#"<p:xfrm><a:off x="457200" y="3200400"/><a:ext cx="3657600" cy="370840"/></p:xfrm>"#,
        r#"<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblGrid><a:gridCol w="1828800"/><a:gridCol w="1828800"/></a:tblGrid>"#,
        r#"<a:tr h="370840"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Cell A</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#,
        r#"<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Cell B</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#,
        r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">"#,
        r#"<mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">"#,
        r#"<p:sp><p:nvSpPr><p:cNvPr id="90" name="Equation"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>"#,
        r#"<p:spPr><a:xfrm><a:off x="457200" y="4114800"/><a:ext cx="1625600" cy="369332"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>"#,
        r#"<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a14:m><m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara></a14:m></a:p></p:txBody></p:sp>"#,
        r#"</mc:Choice><mc:Fallback>"#,
        r#"<p:sp><p:nvSpPr><p:cNvPr id="90" name="Equation"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>"#,
        r#"<p:spPr><a:xfrm><a:off x="457200" y="4114800"/><a:ext cx="1625600" cy="369332"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>"#,
        r#"<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>x squared</a:t></a:r></a:p></p:txBody></p:sp>"#,
        r#"</mc:Fallback></mc:AlternateContent>"#,
    );
    let parts: Vec<(String, Vec<u8>)> = ooxml_opc::unzip_parts(DECK)
        .unwrap()
        .into_iter()
        .map(|(name, bytes)| match name.as_str() {
            "ppt/slides/slide1.xml" => {
                let xml = String::from_utf8(bytes).unwrap();
                let xml = xml.replacen("</p:spTree>", &format!("{shapes}</p:spTree>"), 1);
                (name, xml.into_bytes())
            }
            _ => (name, bytes),
        })
        .collect();
    let mut deck = Deck::new(ooxml_opc::rezip_parts(&parts).unwrap());
    let read = deck.read(Some(1), 0, 200).unwrap();
    let shapes = read["shapes"].as_array().unwrap();
    let (table, equation) = (&shapes[2], &shapes[3]);
    assert_eq!(table["kind"], "table");
    assert_eq!(equation["equation"], true);
    let cells: Vec<&Value> = read["paragraphs"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|paragraph| paragraph["shape"] == table["id"])
        .collect();
    assert_eq!(
        (&cells[1]["text"], &cells[1]["row"], &cells[1]["column"]),
        (&json!("Cell B"), &json!(0), &json!(1))
    );

    assert_eq!(
        deck.apply(&[set_text(table["id"].as_str().unwrap(), "", &["x"])])
            .unwrap_err(),
        invalid(
            "a table's text is edited by paragraph: replace_text, insert_paragraph and delete_paragraph"
        )
    );
    let error = deck
        .apply(&[set_text(
            equation["id"].as_str().unwrap(),
            "x squared",
            &["y"],
        )])
        .unwrap_err();
    assert!(
        matches!(&error, Error::Edit { code: EditCode::InvalidInput, message, details: None }
            if message.ends_with("is an equation: its math can only be changed in PowerPoint")),
        "{error:?}"
    );
}

#[test]
fn set_shape_text_and_the_slide_read_parse_from_their_wire_json() {
    let command: Command = serde_json::from_value(json!({
        "type": "set_shape_text", "shapeId": "s", "expectedText": "a\nb", "paragraphs": ["c"],
    }))
    .unwrap();
    assert_eq!(command, set_text("s", "a\nb", &["c"]));
    let query: InspectQuery =
        serde_json::from_value(json!({ "slide": 3, "start": 0, "count": 60 })).unwrap();
    assert_eq!(query.slide, Some(3));
}
