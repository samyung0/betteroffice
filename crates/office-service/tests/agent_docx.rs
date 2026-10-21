//! The chat agent's DOCX commands, scoped read and `create` (Epo 2026-10-20):
//! paragraphs, lists, tables, text formatting and alignment over the starter
//! document (python-docx's default template, starters/NOTICE.md).

use office_service::{
    Alignment, Checkpoint, Command, CommandResult, Determinism, EditCode, Error, Format,
    InspectQuery, ListKind, TextFormat, apply_commands, create, export, inspect_scoped, seed,
    sha256_hex, starter,
};
use serde_json::{Value, json};

const DETERMINISM: Determinism = Determinism {
    seed: "0000000000000000000000000000000000000000000000000000000000000000",
    now: "2026-10-21T00:00:00.000Z",
};

/// A seeded document and its latest state.
struct Doc {
    base: Vec<u8>,
    sha: String,
    state: Vec<u8>,
}

impl Doc {
    fn new(base: Vec<u8>) -> Self {
        let state = seed(Format::Docx, &base).unwrap();
        Self {
            sha: sha256_hex(&base),
            base,
            state,
        }
    }

    /// The starter with `commands` applied, exported and reopened.
    fn created(commands: &[Command]) -> Self {
        Self::new(
            create(
                Format::Docx,
                starter(Format::Docx).unwrap(),
                commands,
                DETERMINISM,
            )
            .unwrap()
            .bytes,
        )
    }

    fn checkpoint(&self) -> Checkpoint<'_> {
        Checkpoint {
            format: Format::Docx,
            schema_version: 1,
            base_sha256: &self.sha,
            state: &self.state,
        }
    }

    /// Applies `commands`, keeping the state only when they land.
    fn apply(&mut self, commands: &[Command]) -> Result<CommandResult, Error> {
        let result = apply_commands(&self.base, self.checkpoint(), commands)?;
        self.state = result.state.clone();
        Ok(result)
    }

    fn read(&self, start: Option<u32>, count: u32) -> Result<Value, Error> {
        inspect_scoped(
            &self.base,
            self.checkpoint(),
            &InspectQuery {
                start,
                count: Some(count),
                ..InspectQuery::default()
            },
        )
    }

    /// Each body paragraph as `text|style|list`, a cell's prefixed with its
    /// row and column.
    fn lines(&self) -> Vec<String> {
        self.read(None, 200).unwrap()["paragraphs"]
            .as_array()
            .unwrap()
            .iter()
            .map(|paragraph| {
                let cell = match &paragraph["table"] {
                    Value::Null => String::new(),
                    table => format!("r{}c{} ", table["row"], table["column"]),
                };
                format!(
                    "{cell}{}|{}|{}",
                    paragraph["text"].as_str().unwrap(),
                    paragraph["style"].as_str().unwrap_or(""),
                    paragraph["list"].as_str().unwrap_or("")
                )
            })
            .collect()
    }

    /// The id of the paragraph reading `text`.
    fn id(&self, text: &str) -> String {
        self.paragraph(text)["id"].as_str().unwrap().to_owned()
    }

    fn paragraph(&self, text: &str) -> Value {
        self.read(None, 200).unwrap()["paragraphs"]
            .as_array()
            .unwrap()
            .iter()
            .find(|paragraph| paragraph["text"] == text)
            .unwrap_or_else(|| panic!("no paragraph reads {text:?}"))
            .clone()
    }

    /// `word/document.xml` of the exported file.
    fn xml(&self) -> String {
        part(
            &export(&self.base, self.checkpoint(), DETERMINISM).unwrap(),
            "word/document.xml",
        )
    }
}

fn part(package: &[u8], path: &str) -> String {
    ooxml_opc::unzip_parts(package)
        .unwrap()
        .into_iter()
        .find(|(name, _)| name == path)
        .map(|(_, bytes)| String::from_utf8(bytes).unwrap())
        .unwrap_or_default()
}

/// The starter with `paragraphs` (body XML) before its section properties.
fn with_body(paragraphs: &str) -> Vec<u8> {
    let parts: Vec<(String, Vec<u8>)> = ooxml_opc::unzip_parts(starter(Format::Docx).unwrap())
        .unwrap()
        .into_iter()
        .map(|(name, bytes)| {
            if name != "word/document.xml" {
                return (name, bytes);
            }
            let xml = String::from_utf8(bytes).unwrap();
            let xml = xml.replacen("<w:body>", &format!("<w:body>{paragraphs}"), 1);
            (name, xml.into_bytes())
        })
        .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn insert(after_id: Option<&str>, text: &str, style: Option<&str>) -> Command {
    Command::InsertParagraph {
        after_id: after_id.map(str::to_owned),
        before_id: None,
        text: text.into(),
        style: style.map(str::to_owned),
        list: None,
    }
}

fn listed(text: &str, list: ListKind) -> Command {
    Command::InsertParagraph {
        after_id: None,
        before_id: None,
        text: text.into(),
        style: None,
        list: Some(list),
    }
}

fn grid(rows: &[&[&str]]) -> Vec<Vec<String>> {
    rows.iter()
        .map(|row| row.iter().map(|cell| (*cell).to_owned()).collect())
        .collect()
}

fn table(after_id: Option<&str>, rows: &[&[&str]]) -> Command {
    Command::InsertTable {
        after_id: after_id.map(str::to_owned),
        rows: grid(rows),
    }
}

const PARAGRAPH_ANCHOR: &str = "a new paragraph goes in the document body: after_id and before_id name a body paragraph, not one of a table, header, footer or note";

fn refused(code: EditCode, message: &str) -> Error {
    Error::Edit {
        code,
        message: message.into(),
        details: None,
    }
}

fn invalid(message: &str) -> Error {
    refused(EditCode::InvalidInput, message)
}

/// The document `report` builds: a title, a heading, text, a two-item
/// bulleted list, a table and closing text.
fn report() -> Doc {
    Doc::created(&[
        insert(None, "Lab report", Some("Title")),
        insert(None, "Method", Some("Heading1")),
        insert(None, "We timed the pendulum.", None),
        listed("Ruler", ListKind::Bullet),
        listed("Stopwatch", ListKind::Bullet),
        table(None, &[&["Length", "Period"], &["0.5 m", "1.4 s"]]),
        insert(None, "Done.", None),
    ])
}

/// A new document from the starter holds what its commands wrote, in order,
/// from its first paragraph on; reopened, the read shows the styles, the
/// list and the table's cells before the paragraph that follows it, with no
/// blank paragraph between (Epo 2026-10-21), and the file carries no
/// python-docx core properties.
#[test]
fn create_builds_a_document_from_the_starter_and_it_reopens() {
    let doc = report();
    assert_eq!(
        doc.lines(),
        [
            "Lab report|Title|",
            "Method|Heading1|",
            "We timed the pendulum.||",
            "Ruler||bullet",
            "Stopwatch||bullet",
            "r0c0 Length||",
            "r0c1 Period||",
            "r1c0 0.5 m||",
            "r1c1 1.4 s||",
            "Done.||",
        ]
    );
    // The paragraph a table opens cannot take text commands.
    assert_eq!(
        doc.read(None, 200).unwrap()["paragraphs"][9]["locked"],
        "follows a table"
    );
    let core = part(
        &export(&doc.base, doc.checkpoint(), DETERMINISM).unwrap(),
        "docProps/core.xml",
    );
    assert!(
        core.contains("<cp:coreProperties")
            && !core.contains("python-docx")
            && !core.contains("2013"),
        "{core}"
    );
    let xml = doc.xml();
    assert!(xml.contains(r#"<w:pStyle w:val="Title"/>"#), "{xml}");
    assert!(xml.contains(r#"<w:pStyle w:val="Heading1"/>"#), "{xml}");
    assert_eq!(xml.matches("<w:tbl>").count(), 1, "{xml}");
    // The list buttons' bullet id follows the starter's nine w:num ids.
    assert_eq!(xml.matches(r#"<w:numId w:val="10"/>"#).count(), 2, "{xml}");
}

/// The starter seeds, exports and reopens as one empty paragraph, and the
/// first read lists its shown paragraph styles.
#[test]
fn the_docx_starter_reads_as_one_empty_paragraph_with_its_styles() {
    let doc = Doc::new(starter(Format::Docx).unwrap().to_vec());
    let read = doc.read(None, 60).unwrap();
    assert_eq!(read["total"], 1);
    assert_eq!(read["nextStart"], Value::Null);
    assert_eq!(read["paragraphs"][0]["text"], "");
    let styles = read["styles"].as_array().unwrap();
    assert_eq!(styles.len(), 28);
    assert_eq!(styles[0], json!({ "id": "Normal", "name": "Normal" }));
    assert!(styles.contains(&json!({ "id": "Heading1", "name": "heading 1" })));
    let reopened = Doc::new(export(&doc.base, doc.checkpoint(), DETERMINISM).unwrap());
    assert_eq!(reopened.lines(), ["||"]);
}

/// Inserts after one paragraph land in the order given, tables too.
#[test]
fn inserts_after_one_paragraph_keep_their_order() {
    let mut doc = report();
    let method = doc.id("Method");
    doc.apply(&[
        insert(Some(&method), "First.", None),
        table(Some(&method), &[&["a"]]),
        insert(Some(&method), "Second.", Some("Heading2")),
    ])
    .unwrap();
    assert_eq!(
        doc.lines()[..6],
        [
            "Lab report|Title|",
            "Method|Heading1|",
            "First.||",
            "r0c0 a||",
            "Second.|Heading2|",
            "We timed the pendulum.||",
        ]
    );
}

/// An inserted table keeps the blank paragraph after it only where Word
/// needs one: at the body's end and before another table (Epo 2026-10-21).
#[test]
fn an_inserted_table_keeps_a_blank_paragraph_only_where_word_needs_one() {
    let mut doc = Doc::created(&[insert(None, "Intro", None), table(None, &[&["x"]])]);
    assert_eq!(doc.lines(), ["Intro||", "r0c0 x||", "||"]);
    doc.apply(&[table(Some(&doc.id("Intro")), &[&["y"]])])
        .unwrap();
    assert_eq!(doc.lines(), ["Intro||", "r0c0 y||", "||", "r0c0 x||", "||"]);
    assert_eq!(doc.xml().matches("<w:tbl>").count(), 2);
}

#[test]
fn delete_paragraph_removes_a_plain_paragraph_reading_its_expected_text() {
    let mut doc = report();
    let text = doc.id("We timed the pendulum.");
    let ruler = doc.paragraph("Ruler");
    doc.apply(&[Command::DeleteParagraph {
        target_id: text,
        expected_text: "We timed the pendulum.".into(),
    }])
    .unwrap();
    assert!(!doc.lines().iter().any(|line| line.starts_with("We timed")));
    // The next paragraph keeps its own id and list.
    assert_eq!(doc.paragraph("Ruler"), ruler);

    // A paragraph after a table goes, the next one taking the table, unless
    // it ends the body (Word needs it there).
    doc.apply(&[insert(None, "More.", None)]).unwrap();
    doc.apply(&[Command::DeleteParagraph {
        target_id: doc.id("Done."),
        expected_text: "Done.".into(),
    }])
    .unwrap();
    assert_eq!(doc.lines()[7..], ["r1c1 1.4 s||", "More.||"]);
    assert_eq!(
        doc.apply(&[Command::DeleteParagraph {
            target_id: doc.id("More."),
            expected_text: "More.".into(),
        }])
        .unwrap_err(),
        invalid(
            "the paragraph after a table that ends the body or a cell cannot be deleted; Word needs it"
        )
    );
}

/// An insert after a list item continues the list as Enter does; its
/// `list` ends or sets the list instead.
#[test]
fn an_insert_continues_a_list_unless_its_list_says_otherwise() {
    let mut doc = report();
    let stopwatch = doc.id("Stopwatch");
    doc.apply(&[
        insert(Some(&stopwatch), "Timer", None),
        Command::InsertParagraph {
            after_id: Some(stopwatch),
            before_id: None,
            text: "Results follow.".into(),
            style: None,
            list: Some(ListKind::None),
        },
    ])
    .unwrap();
    assert_eq!(
        doc.lines()[3..7],
        [
            "Ruler||bullet",
            "Stopwatch||bullet",
            "Timer||bullet",
            "Results follow.||"
        ]
    );
}

#[test]
fn set_list_makes_bulleted_and_numbered_lists_and_ends_them() {
    let mut doc = report();
    let (ruler, stopwatch, done) = (doc.id("Ruler"), doc.id("Stopwatch"), doc.id("Done."));
    doc.apply(&[
        Command::SetList {
            ids: vec![ruler.clone(), stopwatch],
            list: ListKind::Numbered,
        },
        Command::SetList {
            ids: vec![done],
            list: ListKind::Bullet,
        },
    ])
    .unwrap();
    assert_eq!(doc.paragraph("Ruler")["list"], "numbered");
    assert_eq!(doc.paragraph("Stopwatch")["list"], "numbered");
    assert_eq!(doc.paragraph("Done.")["list"], "bullet");
    doc.apply(&[Command::SetList {
        ids: vec![ruler],
        list: ListKind::None,
    }])
    .unwrap();
    assert_eq!(doc.paragraph("Ruler")["list"], Value::Null);
    let reopened = Doc::new(export(&doc.base, doc.checkpoint(), DETERMINISM).unwrap());
    assert_eq!(
        [
            reopened.paragraph("Ruler")["list"].clone(),
            reopened.paragraph("Stopwatch")["list"].clone(),
            reopened.paragraph("Done.")["list"].clone(),
        ],
        [Value::Null, json!("numbered"), json!("bullet")]
    );
}

#[test]
fn insert_table_refuses_a_too_large_ragged_or_nested_table() {
    let mut doc = report();
    let method = doc.id("Method");
    let rule = "rows must be 1 to 20 rows of the same 1 to 10 cells";
    let rows: Vec<&[&str]> = vec![&["x"]; 21];
    assert_eq!(
        doc.apply(&[table(Some(&method), &rows)]).unwrap_err(),
        invalid(rule)
    );
    assert_eq!(
        doc.apply(&[table(Some(&method), &[&["x"; 11]])])
            .unwrap_err(),
        invalid(rule)
    );
    assert_eq!(
        doc.apply(&[table(Some(&method), &[&["a", "b"], &["c"]])])
            .unwrap_err(),
        invalid(rule)
    );
    let cell = doc.id("Period");
    assert_eq!(
        doc.apply(&[table(Some(&cell), &[&["x"]])]).unwrap_err(),
        invalid(
            "a new table goes in the document body: after_id names a body paragraph, not one of a table, header, footer or note"
        )
    );
}

#[test]
fn delete_table_removes_the_table_whose_cells_read_as_expected() {
    let mut doc = report();
    let first = doc.paragraph("Length");
    let id = first["table"]["id"].as_str().unwrap().to_owned();
    assert_eq!(id, first["id"]);
    let stale = doc
        .apply(&[Command::DeleteTable {
            id: id.clone(),
            expected: grid(&[&["Length", "Period"], &["0.5 m", "1.5 s"]]),
        }])
        .unwrap_err();
    assert_eq!(
        stale,
        Error::Edit {
            code: EditCode::StaleTarget,
            message: "the table's cells differ from expected; read it again".into(),
            details: Some(json!({ "block": {
                "id": id, "type": "table", "text": "Length\tPeriod\n0.5 m\t1.4 s",
                "grid": [["Length", "Period"], ["0.5 m", "1.4 s"]],
            }})),
        }
    );
    // Any of its cells' paragraphs names the table.
    doc.apply(&[Command::DeleteTable {
        id: doc.id("1.4 s"),
        expected: grid(&[&["Length", "Period"], &["0.5 m", "1.4 s"]]),
    }])
    .unwrap();
    assert!(!doc.lines().iter().any(|line| line.starts_with('r')));
    assert!(!doc.xml().contains("<w:tbl>"));
}

#[test]
fn format_text_formats_the_one_occurrence_of_its_expected_text() {
    let mut doc = report();
    let text = doc.id("We timed the pendulum.");
    doc.apply(&[Command::FormatText {
        target_id: text.clone(),
        expected_text: "pendulum".into(),
        format: TextFormat {
            bold: Some(true),
            italic: Some(true),
            underline: Some(true),
            strikethrough: Some(true),
            color: Some("#1f4e79".into()),
            highlight: Some("Yellow".into()),
            font_size: Some(14.0),
            font_family: Some("Arial".into()),
        },
    }])
    .unwrap();
    let xml = doc.xml();
    let run = xml
        .split("<w:r>")
        .find(|run| run.contains(">pendulum<"))
        .unwrap_or_else(|| panic!("{xml}"));
    for property in [
        "<w:b/>",
        "<w:i/>",
        r#"<w:u w:val="single"/>"#,
        "<w:strike/>",
        r#"<w:color w:val="1F4E79"/>"#,
        r#"<w:highlight w:val="yellow"/>"#,
        r#"<w:sz w:val="28"/>"#,
        r#"w:ascii="Arial""#,
    ] {
        assert!(run.contains(property), "{property} in {run}");
    }
    assert!(xml.contains(">We timed the <"), "{xml}");

    let details = Some(json!({ "block": {
        "id": text, "type": "paragraph", "text": "We timed the pendulum.",
    }}));
    let format = |expected: &str, format: TextFormat| Command::FormatText {
        target_id: text.clone(),
        expected_text: expected.into(),
        format,
    };
    let bold = TextFormat {
        bold: Some(true),
        ..TextFormat::default()
    };
    assert_eq!(
        doc.apply(&[format("clock", bold.clone())]).unwrap_err(),
        Error::Edit {
            code: EditCode::StaleTarget,
            message: "expected text was not found in the paragraph".into(),
            details: details.clone(),
        }
    );
    assert_eq!(
        doc.apply(&[format("e", bold)]).unwrap_err(),
        Error::Edit {
            code: EditCode::InvalidInput,
            message: "expected text occurs 4 times in the paragraph; include more context".into(),
            details,
        }
    );
    assert_eq!(
        doc.apply(&[format("We", TextFormat::default())])
            .unwrap_err(),
        invalid("format_text names no formatting")
    );
    assert_eq!(
        doc.apply(&[format(
            "We",
            TextFormat {
                color: Some("navy".into()),
                ..TextFormat::default()
            }
        )])
        .unwrap_err(),
        invalid(r##"color "navy" is not a hex colour like "#1F4E79""##)
    );
}

#[test]
fn set_alignment_aligns_one_paragraph() {
    let mut doc = report();
    doc.apply(&[
        Command::SetAlignment {
            target_id: doc.id("Lab report"),
            alignment: Alignment::Center,
        },
        Command::SetAlignment {
            target_id: doc.id("We timed the pendulum."),
            alignment: Alignment::Justify,
        },
    ])
    .unwrap();
    let xml = doc.xml();
    assert!(xml.contains(r#"<w:jc w:val="center"/>"#), "{xml}");
    assert!(xml.contains(r#"<w:jc w:val="both"/>"#), "{xml}");
}

/// A read pages through the paragraphs; only the first page lists the styles.
#[test]
fn scoped_reads_page_through_the_paragraphs() {
    let doc = report();
    let first = doc.read(None, 4).unwrap();
    assert_eq!(first["total"], 10);
    assert_eq!(first["nextStart"], 4);
    assert_eq!(first["paragraphs"].as_array().unwrap().len(), 4);
    assert!(first["styles"].is_array());
    let last = doc.read(Some(8), 4).unwrap();
    assert_eq!(last["nextStart"], Value::Null);
    assert_eq!(last["styles"], Value::Null);
    assert_eq!(
        last["paragraphs"]
            .as_array()
            .unwrap()
            .iter()
            .map(|paragraph| paragraph["text"].as_str().unwrap())
            .collect::<Vec<_>>(),
        ["1.4 s", "Done."]
    );
    assert_eq!(
        doc.read(None, 201).unwrap_err(),
        invalid("a DOCX read names a count of 1 to 200 paragraphs")
    );
}

/// The read leaves out tracked-deleted text and says why the text commands
/// refuse a paragraph (`locked`); a content control's paragraphs are marked,
/// and one cannot be an `after_id`.
#[test]
fn the_read_leaves_out_deleted_text_and_marks_what_cannot_be_edited() {
    let mut doc = Doc::new(with_body(concat!(
        r#"<w:p><w:r><w:t xml:space="preserve">Kept </w:t></w:r>"#,
        r#"<w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>"#,
        r#"<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>"#,
        r#"<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Inside</w:t></w:r></w:p></w:sdtContent></w:sdt>"#,
        r#"<w:p><w:r><w:t>After</w:t></w:r></w:p>"#,
        r#"<w:p><w:r><w:t>Plain</w:t></w:r></w:p>"#,
    )));
    let read = doc.read(None, 10).unwrap();
    let marks: Vec<Value> = read["paragraphs"]
        .as_array()
        .unwrap()
        .iter()
        .map(|paragraph| {
            json!([
                paragraph["text"],
                paragraph["locked"],
                paragraph["contentControl"]
            ])
        })
        .collect();
    assert_eq!(
        marks,
        [
            json!(["Kept ", "holds tracked changes", null]),
            // A field's shown text is not the paragraph's own.
            json!(["Page ", "holds a field", null]),
            json!(["Inside", null, true]),
            json!(["After", "follows a content control", null]),
            json!(["Plain", null, null]),
        ]
    );
    assert_eq!(
        doc.apply(&[insert(Some(&doc.id("Inside")), "x", None)])
            .unwrap_err(),
        invalid(PARAGRAPH_ANCHOR)
    );
    // What the read marks locked, the text commands refuse.
    assert_eq!(
        doc.apply(&[Command::DeleteParagraph {
            target_id: doc.id("After"),
            expected_text: "After".into(),
        }])
        .unwrap_err(),
        refused(
            EditCode::UnavailableTarget,
            "the paragraph follows a content control; it cannot be deleted"
        )
    );
}

/// Header, footer and note paragraphs read after the body, each marked with
/// its part, and take `replace_text` only (Epo 2026-10-21).
#[test]
fn headers_footers_and_notes_read_after_the_body_and_take_replace_text_only() {
    let mut doc = Doc::new(include_bytes!("fixtures/stories.docx").to_vec());
    let parts: Vec<(String, String)> = doc.read(None, 200).unwrap()["paragraphs"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|paragraph| !paragraph["part"].is_null())
        .map(|paragraph| {
            (
                paragraph["part"].as_str().unwrap().to_owned(),
                paragraph["text"].as_str().unwrap().to_owned(),
            )
        })
        .collect();
    assert_eq!(
        parts,
        [
            ("header (default)".to_owned(), "Header".to_owned()),
            ("footer (default)".to_owned(), "Footer".to_owned()),
            ("footnote 1".to_owned(), "A note".to_owned()),
        ]
    );
    assert!(doc.paragraph("End")["part"].is_null());
    let header = doc.id("Header");
    assert!(header.starts_with("hf:"), "{header}");
    let only = refused(
        EditCode::UnsupportedOperation,
        "header, footer and note paragraphs take replace_text only",
    );
    for command in [
        Command::SetAlignment {
            target_id: header.clone(),
            alignment: Alignment::Center,
        },
        Command::DeleteParagraph {
            target_id: doc.id("A note"),
            expected_text: "A note".into(),
        },
    ] {
        assert_eq!(doc.apply(&[command]).unwrap_err(), only);
    }
    assert_eq!(
        doc.apply(&[insert(Some(&header), "x", None)]).unwrap_err(),
        invalid(PARAGRAPH_ANCHOR)
    );
    doc.apply(&[Command::ReplaceText {
        target_id: header,
        expected_text: "Header".into(),
        text: "Lab 3".into(),
    }])
    .unwrap();
    assert_eq!(doc.paragraph("Lab 3")["part"], "header (default)");
}

/// A call holds at most 50 commands; the pipeline sends no more, so this
/// refusal guards other callers.
#[test]
fn a_call_holds_at_most_50_commands() {
    let mut doc = report();
    let alignments: Vec<Command> = (0..51)
        .map(|_| Command::SetAlignment {
            target_id: doc.id("Done."),
            alignment: Alignment::Left,
        })
        .collect();
    assert_eq!(
        doc.apply(&alignments).unwrap_err(),
        invalid("the call holds 51 commands; a call holds at most 50")
    );
    doc.apply(&alignments[..50]).unwrap();
}

/// A list a paragraph style gives reads as a list, and `set_list` `none`
/// ends it with numId 0.
#[test]
fn a_style_list_reads_as_a_list_and_none_ends_it() {
    let mut doc = Doc::created(&[
        insert(None, "Item", Some("ListBullet")),
        insert(None, "Step", Some("ListNumber")),
    ]);
    assert_eq!(
        doc.lines(),
        ["Item|ListBullet|bullet", "Step|ListNumber|numbered"]
    );
    doc.apply(&[Command::SetList {
        ids: vec![doc.id("Item")],
        list: ListKind::None,
    }])
    .unwrap();
    assert_eq!(
        doc.lines(),
        ["Item|ListBullet|", "Step|ListNumber|numbered"]
    );
    let xml = doc.xml();
    assert!(xml.contains(r#"<w:numId w:val="0"/>"#), "{xml}");
    let reopened = Doc::new(export(&doc.base, doc.checkpoint(), DETERMINISM).unwrap());
    assert_eq!(
        reopened.lines(),
        ["Item|ListBullet|", "Step|ListNumber|numbered"]
    );
}

/// The wire shapes Capy sends: snake_case `type`, camelCase fields.
#[test]
fn docx_commands_parse_from_their_wire_json() {
    let commands: Vec<Command> = serde_json::from_value(json!([
        { "type": "insert_paragraph", "afterId": "body:paragraph:1", "text": "Hi", "style": "Heading1", "list": "bullet" },
        { "type": "insert_paragraph", "text": "End" },
        { "type": "delete_paragraph", "targetId": "body:paragraph:1", "expectedText": "Hi" },
        { "type": "set_list", "ids": ["body:paragraph:1"], "list": "none" },
        { "type": "insert_table", "afterId": "body:paragraph:1", "rows": [["a", "b"]] },
        { "type": "delete_table", "id": "body:t0:r0c0:paragraph:2", "expected": [["a", "b"]] },
        { "type": "format_text", "targetId": "body:paragraph:1", "expectedText": "Hi", "bold": true, "fontSize": 12 },
        { "type": "set_alignment", "targetId": "body:paragraph:1", "alignment": "justify" },
    ]))
    .unwrap();
    assert_eq!(
        commands[0],
        Command::InsertParagraph {
            after_id: Some("body:paragraph:1".into()),
            before_id: None,
            text: "Hi".into(),
            style: Some("Heading1".into()),
            list: Some(ListKind::Bullet),
        }
    );
    assert_eq!(commands[1], insert(None, "End", None));
    assert_eq!(
        commands[6],
        Command::FormatText {
            target_id: "body:paragraph:1".into(),
            expected_text: "Hi".into(),
            format: TextFormat {
                bold: Some(true),
                font_size: Some(12.0),
                ..TextFormat::default()
            },
        }
    );
}

/// An XLSX command on a DOCX file is refused as unsupported.
#[test]
fn xlsx_commands_are_refused_on_a_docx() {
    let mut doc = report();
    assert_eq!(
        doc.apply(&[Command::AddSheet {
            name: "Data".into(),
            after: None,
        }])
        .unwrap_err(),
        refused(
            EditCode::UnsupportedOperation,
            "this command edits XLSX files"
        )
    );
}
