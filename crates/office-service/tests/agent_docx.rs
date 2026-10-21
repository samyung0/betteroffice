//! The chat agent's DOCX commands, scoped read and `create` (Epo 2026-10-20):
//! paragraphs, lists, tables, text formatting and alignment over the starter
//! document (python-docx's default template, starters/NOTICE.md).

use office_service::{
    Alignment, Checkpoint, Command, CommandResult, Determinism, EditCode, Error, Format,
    InspectQuery, ListKind, TextFormat, Touched, apply_commands, create, export, inspect_scoped,
    seed, sha256_hex, starter,
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

const PARAGRAPH_ANCHOR: &str = "a new paragraph goes in the document body: after_id and before_id name a body paragraph, not one of a table, content control, header, footer or note";

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
/// python-docx core properties, thumbnail or application identity.
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
    let exported = export(&doc.base, doc.checkpoint(), DETERMINISM).unwrap();
    let core = part(&exported, "docProps/core.xml");
    assert!(
        core.contains("<cp:coreProperties")
            && !core.contains("python-docx")
            && !core.contains("2013"),
        "{core}"
    );
    let app = part(&exported, "docProps/app.xml");
    assert!(
        !app.contains("Microsoft") && !app.contains("Normal.dotm"),
        "{app}"
    );
    let names: Vec<String> = ooxml_opc::unzip_parts(&exported)
        .unwrap()
        .into_iter()
        .map(|(name, _)| name)
        .collect();
    assert!(
        !names.iter().any(|name| name.contains("thumbnail")),
        "{names:?}"
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
            "a new table goes in the document body: after_id names a body paragraph, not one of a table, content control, header, footer or note"
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
    // A tracked paragraph's list and alignment stay as they are (Epo
    // 2026-09-25): Word's Reject All would undo them unseen.
    let tracked = refused(
        EditCode::UnavailableTarget,
        "the paragraph holds tracked changes; it cannot be changed",
    );
    let kept = doc.id("Kept ");
    assert_eq!(
        doc.apply(&[Command::SetList {
            ids: vec![kept.clone()],
            list: ListKind::Bullet,
        }])
        .unwrap_err(),
        tracked
    );
    assert_eq!(
        doc.apply(&[Command::SetAlignment {
            target_id: kept,
            alignment: Alignment::Center,
        }])
        .unwrap_err(),
        tracked
    );
}

/// Review H2 and r2 L2: a table whose cell holds what delete_paragraph
/// refuses is refused, a table or content control (`expected` names only
/// the cells' text), an equation, a field or a tracked change.
#[test]
fn delete_table_refuses_cells_holding_what_delete_paragraph_refuses() {
    let cell = |inner: &str| {
        format!(
            r#"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>{inner}<w:p/></w:tc></w:tr></w:tbl><w:p/>"#
        )
    };
    let cases = [
        (
            "Outer",
            concat!(
                r#"<w:p><w:r><w:t>Outer</w:t></w:r></w:p>"#,
                r#"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
            ),
            "holds a table or content control",
        ),
        (
            "Field",
            r#"<w:sdt><w:sdtPr><w:id w:val="9"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Field</w:t></w:r></w:p></w:sdtContent></w:sdt>"#,
            "holds a table or content control",
        ),
        (
            "Sum",
            r#"<w:p><w:r><w:t>Sum</w:t></w:r><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x+1</m:t></m:r></m:oMath></w:p>"#,
            "holds an equation",
        ),
        (
            "Page",
            r#"<w:p><w:r><w:t>Page</w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>"#,
            "",
        ),
        (
            "Added",
            r#"<w:p><w:ins w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Added</w:t></w:r></w:ins></w:p>"#,
            "holds tracked changes",
        ),
    ];
    let mut doc = Doc::new(with_body(
        &cases
            .iter()
            .map(|(_, inner, _)| cell(inner))
            .collect::<String>(),
    ));
    for (text, _, reason) in cases {
        let first = doc.read(None, 200).unwrap()["paragraphs"]
            .as_array()
            .unwrap()
            .iter()
            .find(|p| p["text"].as_str().unwrap().starts_with(text))
            .unwrap_or_else(|| panic!("no paragraph reads {text}"))
            .clone();
        let id = first["table"]["id"].as_str().unwrap().to_owned();
        let error = doc
            .apply(&[Command::DeleteTable {
                id,
                expected: grid(&[&[&format!(
                    "{text}
"
                )]]),
            }])
            .unwrap_err();
        let Error::Edit { code, message, .. } = &error else {
            panic!("{text}: {error:?}");
        };
        assert_eq!(*code, EditCode::UnavailableTarget, "{text}: {message}");
        assert!(
            message.starts_with("a cell of the table ")
                && message.contains(reason)
                && message.ends_with("; the table cannot be deleted"),
            "{text}: {message}"
        );
    }
}

/// A filled paragraph (a document's only one, empty) keeps its own tracked
/// change records (review L6).
#[test]
fn an_insert_into_the_only_empty_paragraph_keeps_its_tracked_change() {
    let mut doc = Doc::new(with_body(concat!(
        r#"<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="3" w:author="A" w:date="2026-01-01T00:00:00Z">"#,
        r#"<w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr></w:p>"#,
    )));
    doc.apply(&[insert(None, "Filled", None)]).unwrap();
    let xml = doc.xml();
    assert!(
        xml.contains(">Filled<") && xml.contains("<w:pPrChange"),
        "{xml}"
    );
}

/// `touched` is what the command's Undo re-creates, read before it applies
/// (the Editor's rule, 2026-10-22): a format the span's text items and
/// their format markers, a delete the paragraph's text items and its mark
/// but not the markers, which stay; Capy sizes the edit's Undo with it.
#[test]
fn touched_counts_what_the_undo_re_creates() {
    let run =
        |props: &str, text: &str| format!("<w:r><w:rPr>{props}</w:rPr><w:t>{text}</w:t></w:r>");
    let many: String = (0..12)
        .map(|n| {
            run(
                ["<w:b/>", "<w:i/>", "<w:u w:val=\"single\"/>"][n % 3],
                "abcd",
            )
        })
        .collect();
    let mut doc = Doc::new(with_body(&format!(
        "<w:p>{many}</w:p><w:p>{}</w:p><w:p><w:r><w:t>End</w:t></w:r></w:p>",
        run("", &"abcd".repeat(12)),
    )));
    let format = |id: String| Command::FormatText {
        target_id: id,
        expected_text: "abcd".repeat(12),
        format: TextFormat {
            color: Some("#1F4E79".into()),
            ..TextFormat::default()
        },
    };
    let touched = |doc: &mut Doc, commands: &[Command]| {
        let state = doc.state.clone();
        let result = doc.apply(commands).unwrap();
        doc.state = state;
        result.touched
    };
    let ids: Vec<String> = doc.read(None, 10).unwrap()["paragraphs"]
        .as_array()
        .unwrap()
        .iter()
        .map(|paragraph| paragraph["id"].as_str().unwrap().to_owned())
        .collect();
    let (many, one) = (ids[0].clone(), ids[1].clone());
    let formatted = touched(&mut doc, &[format(many.clone()), format(one.clone())]);
    // Twelve runs, each a text item opened by a format marker, against one.
    assert!(
        formatted[0].items >= formatted[1].items + 2 * 11,
        "{formatted:?}"
    );
    assert!(formatted[1].content >= 48, "{formatted:?}");
    let deleted = touched(
        &mut doc,
        &[Command::DeleteParagraph {
            target_id: many,
            expected_text: "abcd".repeat(12),
        }],
    );
    // Twelve text items and the mark, without the markers.
    assert!(
        deleted[0].items > 12 && deleted[0].items < formatted[0].items,
        "{deleted:?} {formatted:?}"
    );
    assert_eq!(
        touched(&mut doc, &[insert(None, "x", None)]),
        [Touched::default()]
    );
}

/// Review r2 H1: `touched` measures a `replace_text` in a header or a note
/// as apply finds it, so the text it takes out is priced.
#[test]
fn touched_measures_a_header_and_a_note_replacement() {
    let mut doc = Doc::new(include_bytes!("fixtures/stories.docx").to_vec());
    for text in ["Header", "A note"] {
        let state = doc.state.clone();
        let result = doc
            .apply(&[Command::ReplaceText {
                target_id: doc.id(text),
                expected_text: text.into(),
                text: "x".into(),
            }])
            .unwrap();
        doc.state = state;
        assert!(
            result.touched[0].items >= 1 && result.touched[0].content >= text.len() as u64 - 1,
            "{text}: {:?}",
            result.touched
        );
    }
}

/// Review r2 L1: a field's projected children (a TOC's links) sit in the
/// field's own paragraph, which holds the field, so delete_paragraph,
/// replace_text and format_text refuse it and no agent command makes
/// docx-edit rewrite what the field shows outside the edited range, which
/// `touched` would not count. The entries the result continues into hold no
/// children.
#[test]
fn a_fields_own_paragraph_takes_no_text_command() {
    let mut doc = Doc::new(with_body(concat!(
        r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \o "1-3" \h </w:instrText></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:hyperlink w:anchor="_Toc1"><w:r><w:t>Intro</w:t></w:r></w:hyperlink></w:p>"#,
        r#"<w:p><w:hyperlink w:anchor="_Toc2"><w:r><w:t>Method</w:t></w:r></w:hyperlink></w:p>"#,
        r#"<w:p><w:hyperlink w:anchor="_Toc3"><w:r><w:t>Results</w:t></w:r></w:hyperlink><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
        r#"<w:p><w:r><w:t>End</w:t></w:r></w:p>"#,
    )));
    assert_eq!(doc.paragraph("Intro")["locked"], "holds a field");
    let intro = doc.id("Intro");
    for command in [
        Command::DeleteParagraph {
            target_id: intro.clone(),
            expected_text: "Intro".into(),
        },
        Command::ReplaceText {
            target_id: intro.clone(),
            expected_text: "Intro".into(),
            text: "Overview".into(),
        },
        Command::FormatText {
            target_id: intro.clone(),
            expected_text: "Intro".into(),
            format: TextFormat {
                bold: Some(true),
                ..TextFormat::default()
            },
        },
    ] {
        let error = doc.apply(std::slice::from_ref(&command)).unwrap_err();
        assert!(
            matches!(
                error,
                Error::Edit {
                    code: EditCode::UnavailableTarget,
                    ..
                }
            ),
            "{command:?}: {error:?}"
        );
    }
    doc.apply(&[Command::DeleteParagraph {
        target_id: doc.id("Method"),
        expected_text: "Method".into(),
    }])
    .unwrap();
}

/// Review r2 L9: a call reads each story once and its index follows the
/// call's own replacements, so later commands find their paragraphs where
/// they now are, and the targets it reports are where a fresh read finds
/// them.
#[test]
fn a_calls_paragraph_index_follows_its_replacements() {
    let mut doc = Doc::new(with_body(concat!(
        "<w:p><w:r><w:t>Alpha</w:t></w:r></w:p>",
        "<w:p><w:r><w:t>Beta</w:t></w:r></w:p>",
        "<w:p><w:r><w:t>Gamma</w:t></w:r></w:p>",
        "<w:p><w:r><w:t>Delta</w:t></w:r></w:p>",
        "<w:p><w:r><w:t>End</w:t></w:r></w:p>",
    )));
    let replace = |id: String, expected: &str, text: &str| Command::ReplaceText {
        target_id: id,
        expected_text: expected.into(),
        text: text.into(),
    };
    let (a, b, c, d) = (
        doc.id("Alpha"),
        doc.id("Beta"),
        doc.id("Gamma"),
        doc.id("Delta"),
    );
    let result = doc
        .apply(&[
            replace(a, "Alpha", "Alpha, much longer now"),
            replace(c.clone(), "Gamma", "G"),
            Command::FormatText {
                target_id: c.clone(),
                expected_text: "G".into(),
                format: TextFormat {
                    bold: Some(true),
                    ..TextFormat::default()
                },
            },
            replace(d, "Delta", "Delta and more"),
            insert(Some(&b), "New", None),
            replace(b, "Beta", "Beta!"),
            replace(c, "G", "Gg"),
        ])
        .unwrap();
    assert_eq!(
        doc.lines(),
        [
            "Alpha, much longer now||",
            "Beta!||",
            "New||",
            "Gg||",
            "Delta and more||",
            "End||",
        ]
    );
    assert!(doc.xml().contains("<w:b/>"), "{}", doc.xml());
    let ids: Vec<String> = result
        .targets
        .iter()
        .map(|target| target.id.clone())
        .collect();
    let fresh = office_service::locate_targets(&doc.base, doc.checkpoint(), &ids).unwrap();
    assert_eq!(result.targets, fresh);
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

/// The engine counts a call's limits (Epo 2026-10-22): it applies the
/// commands up to the first that would pass one, 50 commands or a DOCX count
/// (30 formatted spans, 400 table cells, 20 inserted paragraphs), and names
/// the rest in the notice, unapplied; after a create the indexes start at 1.
#[test]
fn a_call_applies_up_to_its_first_command_past_a_limit() {
    let notice = |from: usize, to: usize, limit: &str| {
        let which = if from == to {
            format!("the command at index {from}")
        } else {
            format!("the commands at indexes {from} to {to}")
        };
        Some(format!(
            "Not applied: {which} (counting from 0), past the limit of {limit} a call; send them in another call."
        ))
    };
    let mut doc = report();
    let timed = doc.id("We timed the pendulum.");
    let bold: Vec<Command> = (0..31)
        .map(|_| Command::FormatText {
            target_id: timed.clone(),
            expected_text: "pendulum".into(),
            format: TextFormat {
                bold: Some(true),
                ..TextFormat::default()
            },
        })
        .collect();
    assert_eq!(
        doc.apply(&bold).unwrap().notice,
        notice(30, 30, "30 formatted spans")
    );
    assert!(doc.xml().contains("<w:b/>"));
    let align: Vec<Command> = (0..52)
        .map(|_| Command::SetAlignment {
            target_id: timed.clone(),
            alignment: Alignment::Left,
        })
        .collect();
    assert_eq!(
        doc.apply(&align).unwrap().notice,
        notice(50, 51, "50 commands")
    );
    let wide: Vec<&[&str]> = vec![&["x"; 10]; 20];
    let tables = [
        table(None, &wide),
        table(None, &wide),
        table(None, &[&["x"]]),
    ];
    assert_eq!(
        doc.apply(&tables).unwrap().notice,
        notice(2, 2, "400 table cells inserted or deleted")
    );
    assert_eq!(doc.xml().matches("<w:tbl>").count(), 3);
    let paragraphs: Vec<Command> = (0..21)
        .map(|n| insert(None, &format!("P{n}"), None))
        .collect();
    let created = create(
        Format::Docx,
        starter(Format::Docx).unwrap(),
        &paragraphs,
        DETERMINISM,
    )
    .unwrap();
    assert_eq!(created.notice, notice(21, 21, "20 inserted paragraphs"));
    let reopened = Doc::new(created.bytes);
    assert_eq!(reopened.lines().len(), 20);
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

/// An XLSX or PPTX command on a DOCX file is refused as unsupported,
/// naming the format it edits.
#[test]
fn xlsx_and_pptx_commands_are_refused_on_a_docx() {
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
    assert_eq!(
        doc.apply(&[Command::AddSlide {
            after_slide: None,
            before_slide: None,
            layout: None,
        }])
        .unwrap_err(),
        refused(
            EditCode::UnsupportedOperation,
            "this command edits PPTX files"
        )
    );
}
