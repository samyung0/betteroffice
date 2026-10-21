//! The chat agent's paragraph commands (Epo 2026-10-21): `insert_paragraph`
//! after or before a paragraph (a DOCX one in the body, with a paragraph
//! style of the file) and `delete_paragraph` of a paragraph reading its
//! expected text, never its story's last; the DOCX read names each
//! paragraph's style and lists the file's shown paragraph styles.

use office_service::{
    Checkpoint, Command, CommandResult, Determinism, EditCode, Error, Format, InspectQuery,
    apply_commands, export, inspect, inspect_scoped, seed, sha256_hex,
};
use serde_json::{Value, json};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const DECK: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
/// The demo deck's first story: "Office files," then "without the office.".
const DECK_STORY: [&str; 2] = ["Office files,", "without the office."];

/// A document and its latest state.
struct Doc {
    format: Format,
    base: Vec<u8>,
    sha: String,
    state: Vec<u8>,
}

impl Doc {
    fn new(format: Format, base: Vec<u8>) -> Self {
        let state = seed(format, &base).unwrap();
        Self {
            format,
            sha: sha256_hex(&base),
            base,
            state,
        }
    }

    fn checkpoint(&self) -> Checkpoint<'_> {
        Checkpoint {
            format: self.format,
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

    /// The editable paragraphs as (id, text).
    fn paragraphs(&self) -> Vec<(String, String)> {
        inspect(&self.base, self.checkpoint())
            .unwrap()
            .into_iter()
            .map(|entry| (entry.id, entry.value))
            .collect()
    }

    /// The editable paragraphs' text, a DOCX's of its body only.
    fn texts(&self) -> Vec<String> {
        self.paragraphs()
            .into_iter()
            .filter(|(id, _)| self.format != Format::Docx || id.starts_with("body:paragraph:"))
            .map(|(_, text)| text)
            .collect()
    }

    fn id(&self, text: &str) -> String {
        self.paragraphs()
            .into_iter()
            .find(|(_, value)| value == text)
            .unwrap_or_else(|| panic!("no paragraph reads {text:?}"))
            .0
    }

    /// `word/document.xml` of the exported file.
    fn xml(&self) -> String {
        self.part("word/document.xml")
    }

    /// A part of the exported file.
    fn part(&self, name: &str) -> String {
        let saved = export(
            &self.base,
            self.checkpoint(),
            Determinism {
                seed: &"0".repeat(64),
                now: "2026-10-21T00:00:00.000Z",
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

    /// The paragraph of the exported document whose text is `text`, as XML.
    fn paragraph_xml(&self, text: &str) -> String {
        paragraph_in(&self.xml(), "w", text)
    }
}

/// The paragraph (`<{ns}:p>`) of `xml` whose text is `text`.
fn paragraph_in(xml: &str, ns: &str, text: &str) -> String {
    let chunk = xml
        .split(&format!("</{ns}:p>"))
        .find(|chunk| chunk.contains(&format!(">{text}</{ns}:t>")))
        .unwrap_or_else(|| panic!("no exported paragraph reads {text:?}: {xml}"));
    let start = chunk
        .rfind(&format!("<{ns}:p "))
        .into_iter()
        .chain(chunk.rfind(&format!("<{ns}:p>")))
        .max()
        .unwrap();
    chunk[start..].to_owned()
}

/// A DOCX with shown paragraph styles Normal, Heading1 (next: Normal) and
/// Title, a hidden and a semi-hidden one and a character style, and a body
/// of: a title, a centred paragraph, a paragraph with a tracked insertion,
/// one ending a section, the next one, a table, the paragraph after it,
/// Word's page-break paragraph, a heading, a bulleted list item and a last
/// paragraph styled Normal.
fn docx() -> Vec<u8> {
    docx_with_styles("")
}

/// [`docx`] with `extra` style elements.
fn docx_with_styles(extra: &str) -> Vec<u8> {
    let p = |id: &str, ppr: &str, text: &str| {
        format!(
            r#"<w:p w14:paraId="{id}"><w:pPr>{ppr}</w:pPr><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>"#
        )
    };
    let body = [
        p("0000000A", r#"<w:pStyle w:val="Title"/>"#, "Lab report"),
        p("0000000B", r#"<w:jc w:val="center"/>"#, "We timed it."),
        r#"<w:p w14:paraId="0000000C"><w:r><w:t xml:space="preserve">Kept </w:t></w:r><w:ins w:id="1" w:author="Ada"><w:r><w:t>added</w:t></w:r></w:ins></w:p>"#.to_owned(),
        p(
            "0000000D",
            r#"<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>"#,
            "Section end",
        ),
        p("00000013", "", "Next section"),
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p w14:paraId="0000000E"><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#.to_owned(),
        p("0000000F", "", "After the table"),
        r#"<w:p w14:paraId="00000011"><w:r><w:br w:type="page"/></w:r></w:p>"#.to_owned(),
        p("00000012", r#"<w:pStyle w:val="Heading1"/>"#, "Chapter"),
        p(
            "00000014",
            r#"<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>"#,
            "Item one",
        ),
        p("00000010", r#"<w:pStyle w:val="Normal"/>"#, "Last"),
    ]
    .concat();
    let files = [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>"#.to_owned(),
        ),
        (
            "_rels/.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned(),
        ),
        (
            "word/_rels/document.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>"#.to_owned(),
        ),
        (
            "word/styles.xml",
            format!(
                r#"<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:rPr><w:b/></w:rPr></w:style><w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/></w:style><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/></w:style><w:style w:type="paragraph" w:styleId="Hidden"><w:name w:val="Hidden"/><w:hidden/></w:style><w:style w:type="paragraph" w:styleId="Quiet"><w:name w:val="Quiet"/><w:semiHidden/></w:style>{extra}</w:styles>"#
            ),
        ),
        (
            "word/numbering.xml",
            format!(
                r#"<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="-"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#
            ),
        ),
        (
            "word/document.xml",
            format!(
                r#"<w:document xmlns:w="{W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>{body}<w:sectPr/></w:body></w:document>"#
            ),
        ),
    ];
    let files: Vec<(String, Vec<u8>)> = files
        .into_iter()
        .map(|(path, xml)| (path.to_owned(), xml.into_bytes()))
        .collect();
    ooxml_opc::rezip_parts(&files).unwrap()
}

fn after(id: &str, text: &str) -> Command {
    Command::InsertParagraph {
        after_id: Some(id.to_owned()),
        before_id: None,
        text: text.into(),
        style: None,
    }
}

fn before(id: &str, text: &str) -> Command {
    Command::InsertParagraph {
        after_id: None,
        before_id: Some(id.to_owned()),
        text: text.into(),
        style: None,
    }
}

fn styled(command: Command, name: &str) -> Command {
    let Command::InsertParagraph {
        after_id,
        before_id,
        text,
        ..
    } = command
    else {
        unreachable!()
    };
    Command::InsertParagraph {
        after_id,
        before_id,
        text,
        style: Some(name.into()),
    }
}

fn delete(id: &str, expected: &str) -> Command {
    Command::DeleteParagraph {
        target_id: id.to_owned(),
        expected_text: expected.into(),
    }
}

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

fn read(doc: &Doc, start: Option<u32>, count: Option<u32>) -> Result<Value, Error> {
    inspect_scoped(
        &doc.base,
        doc.checkpoint(),
        &InspectQuery {
            start,
            count,
            ..InspectQuery::default()
        },
    )
}

/// Each entry as inspect lists it, its label naming its paragraph style
/// unless that is the default.
fn styled_entries(doc: &Doc) -> Vec<Value> {
    let mut entries = json!(inspect(&doc.base, doc.checkpoint()).unwrap());
    for entry in entries.as_array_mut().unwrap() {
        let style = match entry["value"].as_str().unwrap() {
            "Lab report" => "Title",
            "Chapter" => "Heading1",
            _ => continue,
        };
        let label = format!("{} ({style})", entry["label"].as_str().unwrap());
        entry["label"] = json!(label);
    }
    entries.as_array().unwrap().clone()
}

#[test]
fn the_docx_read_pages_the_paragraphs_and_lists_the_paragraph_styles_first() {
    let doc = Doc::new(Format::Docx, docx());
    let first = read(&doc, None, Some(2)).unwrap();
    let entries = styled_entries(&doc);
    assert_eq!(first["entries"], json!(entries[..2]));
    assert_eq!(first["entries"][0]["label"], "body, paragraph 1 (Title)");
    assert_eq!(first["total"], entries.len());
    assert_eq!(first["nextStart"], 2);
    assert_eq!(
        first["styles"],
        json!([
            { "id": "Normal", "name": "Normal" },
            { "id": "Heading1", "name": "heading 1" },
            { "id": "Title", "name": "Title" },
        ])
    );
    assert!(first.get("stylesMore").is_none());
    let rest = read(&doc, Some(2), Some(200)).unwrap();
    assert_eq!(rest["entries"], json!(entries[2..]));
    assert_eq!(rest["nextStart"], Value::Null);
    assert!(rest.get("styles").is_none());
    assert_eq!(
        read(&doc, None, None).unwrap_err(),
        invalid("a DOCX read names a count of 1 to 200 paragraphs")
    );
}

/// The first page lists at most 50 shown styles and counts the rest.
#[test]
fn the_docx_read_caps_its_style_list() {
    let extra: String = (0..60)
        .map(|n| {
            format!(
                r#"<w:style w:type="paragraph" w:styleId="Body{n}"><w:name w:val="Body {n}"/></w:style>"#
            )
        })
        .collect();
    let doc = Doc::new(Format::Docx, docx_with_styles(&extra));
    let first = read(&doc, None, Some(1)).unwrap();
    let styles = first["styles"].as_array().unwrap();
    assert_eq!(styles.len(), 50);
    assert_eq!(styles[49], json!({ "id": "Body46", "name": "Body 46" }));
    assert_eq!(first["stylesMore"], 13);
}

/// Inserts after one paragraph, and before one, land in call order, one
/// after a paragraph right after it even past one inserted before the next;
/// a new paragraph continues its anchor (the centring of "We timed it.", the
/// Title of "Lab report", which names no next style) unless it names a
/// style, and the paragraphs around it keep their ids. The inverse deletes
/// them.
#[test]
fn docx_inserts_land_after_and_before_in_call_order_and_undo() {
    let mut doc = Doc::new(Format::Docx, docx());
    let original = doc.paragraphs();
    let (title, timed) = (doc.id("Lab report"), doc.id("We timed it."));
    let edit = doc
        .apply(&[
            before(&timed, "Aim"),
            after(&title, "Sub"),
            after(&timed, "First."),
            styled(after(&timed, "Method"), "Heading1"),
            before(&title, "Draft"),
            before(&title, "v2"),
        ])
        .unwrap();
    assert_eq!(
        doc.texts()[..9],
        [
            "Draft",
            "v2",
            "Lab report",
            "Sub",
            "Aim",
            "We timed it.",
            "First.",
            "Method",
            "Section end"
        ]
    );
    assert_eq!(
        (doc.id("Lab report"), doc.id("We timed it.")),
        (title, timed)
    );
    let method = doc.paragraph_xml("Method");
    assert!(
        method.contains(r#"<w:pStyle w:val="Heading1"/>"#) && !method.contains("<w:jc"),
        "{method}"
    );
    let first = doc.paragraph_xml("First.");
    assert!(
        first.contains(r#"<w:jc w:val="center"/>"#) && !first.contains("<w:pStyle"),
        "{first}"
    );
    for text in ["Draft", "v2", "Sub", "Aim"] {
        let xml = doc.paragraph_xml(text);
        assert!(xml.contains(r#"<w:pStyle w:val="Title"/>"#), "{xml}");
    }
    assert!(
        doc.paragraph_xml("We timed it.")
            .contains(r#"<w:jc w:val="center"/>"#)
    );
    assert_eq!(edit.targets.len(), 6);

    doc.apply(edit.inverse.as_deref().unwrap()).unwrap();
    assert_eq!(doc.paragraphs(), original);
}

/// A paragraph inserted after the body's last one is the body's last, which
/// delete_paragraph refuses: the edit has no inverse.
#[test]
fn an_insert_after_the_last_paragraph_has_no_inverse() {
    let mut doc = Doc::new(Format::Docx, docx());
    let last = doc.id("Last");
    let edit = doc.apply(&[after(&last, "Done.")]).unwrap();
    assert_eq!(doc.texts().last().unwrap(), "Done.");
    assert_eq!(edit.inverse, None);
}

/// Without a style an insert continues its anchor as Enter at its end does
/// (Epo 2026-10-21): after a list item the list goes on, after a heading the
/// heading's next style (Normal) follows; before a paragraph the one before
/// it is the anchor. A named style wins: no list, the style named.
#[test]
fn a_docx_insert_without_a_style_continues_its_anchor_as_enter_does() {
    let mut doc = Doc::new(Format::Docx, docx());
    let (chapter, item) = (doc.id("Chapter"), doc.id("Item one"));
    doc.apply(&[
        before(&item, "Lead"),
        after(&item, "Item two"),
        after(&chapter, "Body"),
        styled(after(&item, "Heading"), "Heading1"),
    ])
    .unwrap();
    assert_eq!(
        doc.texts()[doc.texts().len() - 7..],
        [
            "Chapter", "Body", "Lead", "Item one", "Item two", "Heading", "Last"
        ]
    );
    let list = r#"<w:numId w:val="1"/>"#;
    let two = doc.paragraph_xml("Item two");
    assert!(two.contains(list) && !two.contains("<w:pStyle"), "{two}");
    for text in ["Body", "Lead"] {
        let xml = doc.paragraph_xml(text);
        assert!(!xml.contains("<w:pStyle") && !xml.contains(list), "{xml}");
    }
    let heading = doc.paragraph_xml("Heading");
    assert!(
        heading.contains(r#"<w:pStyle w:val="Heading1"/>"#) && !heading.contains(list),
        "{heading}"
    );
}

/// A split at a section's end leaves the section mark on the new paragraph,
/// as Enter does, and delete_paragraph refuses a section's end: an insert
/// after the section's last paragraph has no inverse. One before the next
/// section's first paragraph leaves the section mark where it was and
/// undoes.
#[test]
fn an_insert_next_to_a_section_end_keeps_the_section_and_undoes_where_it_can() {
    let mut doc = Doc::new(Format::Docx, docx());
    let original = doc.paragraphs();
    let end = doc.id("Section end");
    let edit = doc.apply(&[after(&end, "Coda")]).unwrap();
    assert_eq!(edit.inverse, None);
    assert!(doc.paragraph_xml("Coda").contains("<w:sectPr"));

    let mut doc = Doc::new(Format::Docx, docx());
    let next = doc.id("Next section");
    let edit = doc.apply(&[before(&next, "Opening")]).unwrap();
    assert!(doc.paragraph_xml("Section end").contains("<w:sectPr"));
    assert!(!doc.paragraph_xml("Opening").contains("<w:sectPr"));
    doc.apply(edit.inverse.as_deref().unwrap()).unwrap();
    assert_eq!(doc.paragraphs(), original);
}

/// Word's page-break paragraph seeds as an empty paragraph and a break
/// opening the next one; deleting that next paragraph keeps the break.
#[test]
fn docx_delete_after_a_page_break_keeps_the_break() {
    let mut doc = Doc::new(Format::Docx, docx());
    let chapter = doc.id("Chapter");
    doc.apply(&[delete(&chapter, "Chapter")]).unwrap();
    assert!(!doc.texts().iter().any(|text| text == "Chapter"));
    let xml = doc.xml();
    assert!(xml.contains(r#"w:type="page""#), "{xml}");
}

#[test]
fn docx_insert_refuses_unknown_styles_line_breaks_and_over_limit_input() {
    let mut doc = Doc::new(Format::Docx, docx());
    let timed = doc.id("We timed it.");
    let cases = [
        (
            vec![styled(after(&timed, "x"), "Heading 1")],
            invalid(
                r#"the document has no paragraph style "Heading 1"; use a style id inspect_office_document lists"#,
            ),
        ),
        (
            vec![styled(after(&timed, "x"), "Strong")],
            invalid(
                r#"the document has no paragraph style "Strong"; use a style id inspect_office_document lists"#,
            ),
        ),
        (
            vec![after(&timed, "one\ntwo")],
            invalid(
                "text must be one paragraph without line breaks; insert one paragraph a command",
            ),
        ),
        (
            vec![after(&timed, &"x".repeat(5_001))],
            invalid("text holds 5001 characters; a paragraph holds at most 5000"),
        ),
        (
            (0..21).map(|_| after(&timed, "x")).collect(),
            invalid("the call inserts 21 paragraphs; insert at most 20 a call"),
        ),
        (
            vec![Command::InsertParagraph {
                after_id: Some(timed.clone()),
                before_id: Some(timed.clone()),
                text: "x".into(),
                style: None,
            }],
            invalid("insert_paragraph names exactly one of after_id and before_id"),
        ),
        (
            vec![after(&doc.id("Cell"), "x")],
            invalid(
                "a new paragraph goes in the document body: after_id and before_id name a body paragraph, not one of a table, header, footer or note",
            ),
        ),
        (
            vec![after("body:paragraph:nope", "x")],
            refused(
                EditCode::UnavailableTarget,
                "the paragraph is no longer in the document",
            ),
        ),
    ];
    let before = doc.state.clone();
    for (commands, error) in cases {
        assert_eq!(doc.apply(&commands).unwrap_err(), error);
        assert_eq!(doc.state, before);
    }
    // Exactly at the bounds it lands.
    let mut bounded: Vec<Command> = (0..19).map(|_| after(&timed, "x")).collect();
    bounded.push(after(&timed, &"y".repeat(5_000)));
    doc.apply(&bounded).unwrap();
}

/// A delete names the paragraph's whole text; the paragraphs around a
/// deleted one keep their ids and formatting, the first one's included.
#[test]
fn docx_delete_removes_a_plain_paragraph_reading_its_expected_text() {
    let mut doc = Doc::new(Format::Docx, docx());
    let (title, timed) = (doc.id("Lab report"), doc.id("We timed it."));
    assert_eq!(
        doc.apply(&[delete(&title, "Lab")]).unwrap_err(),
        Error::Edit {
            code: EditCode::StaleTarget,
            message: "the paragraph's text differs from expected_text".into(),
            details: Some(json!({
                "block": { "id": title, "type": "paragraph", "text": "Lab report" }
            })),
        }
    );
    let edit = doc.apply(&[delete(&title, "Lab report")]).unwrap();
    assert_eq!(edit.inverse, None);
    assert_eq!(doc.texts()[0], "We timed it.");
    assert_eq!(doc.id("We timed it."), timed);
    let xml = doc.paragraph_xml("We timed it.");
    assert!(xml.contains(r#"<w:jc w:val="center"/>"#), "{xml}");
    assert!(!xml.contains("Title"), "{xml}");
}

#[test]
fn docx_delete_refuses_tracked_section_ending_and_last_paragraphs() {
    let mut doc = Doc::new(Format::Docx, docx());
    let tracked = format!("body:paragraph:{}", "0000000C");
    let cases = [
        (
            delete(&tracked, "Kept added"),
            refused(
                EditCode::UnavailableTarget,
                "the paragraph holds objects, fields or tracked changes, or follows a table or content control; it cannot be deleted",
            ),
        ),
        (
            delete(&doc.id("Section end"), "Section end"),
            invalid("the paragraph ends a section; deleting it would join two sections"),
        ),
        (
            delete(&doc.id("Last"), "Last"),
            invalid("a story's last paragraph cannot be deleted; replace its text instead"),
        ),
        (
            delete(&doc.id("Cell"), "Cell"),
            invalid("a story's last paragraph cannot be deleted; replace its text instead"),
        ),
    ];
    for (command, error) in cases {
        assert_eq!(doc.apply(&[command]).unwrap_err(), error);
    }
}

/// In a shape's text, inserts after and before a paragraph land in call
/// order with its neighbours' ids kept; the inverse deletes them; a delete
/// keeps the paragraph after it as it was.
#[test]
fn pptx_inserts_and_deletes_paragraphs_in_a_shapes_text() {
    let mut doc = Doc::new(Format::Pptx, DECK.to_vec());
    let original = doc.paragraphs();
    let (first, second) = (doc.id(DECK_STORY[0]), doc.id(DECK_STORY[1]));
    let edit = doc
        .apply(&[
            before(&second, "C"),
            after(&first, "A"),
            after(&first, "B"),
            before(&first, "Top"),
        ])
        .unwrap();
    let story: Vec<String> = doc
        .paragraphs()
        .into_iter()
        .map(|(_, text)| text)
        .skip_while(|text| text != "Top")
        .take(6)
        .collect();
    assert_eq!(story, ["Top", DECK_STORY[0], "A", "B", "C", DECK_STORY[1]]);
    assert_eq!(
        (doc.id(DECK_STORY[0]), doc.id(DECK_STORY[1])),
        (first.clone(), second.clone())
    );
    doc.apply(edit.inverse.as_deref().unwrap()).unwrap();
    assert_eq!(doc.paragraphs(), original);

    assert_eq!(
        doc.apply(&[delete(&first, "Office")]).unwrap_err(),
        Error::Edit {
            code: EditCode::StaleTarget,
            message: "the paragraph's text differs from expected_text".into(),
            details: Some(json!({
                "block": { "id": first, "type": "paragraph", "text": DECK_STORY[0] }
            })),
        }
    );
    assert_eq!(
        doc.apply(&[delete(&second, DECK_STORY[1])]).unwrap_err(),
        invalid("a shape's last paragraph cannot be deleted; replace its text instead")
    );
    let edit = doc.apply(&[delete(&first, DECK_STORY[0])]).unwrap();
    assert_eq!(edit.inverse, None);
    assert!(!doc.texts().iter().any(|text| text == DECK_STORY[0]));
    assert_eq!(doc.id(DECK_STORY[1]), second);
}

/// A new PPTX paragraph takes the level, bullet and alignment of the
/// paragraph it goes after or before.
#[test]
fn a_pptx_insert_takes_the_level_and_bullet_of_its_anchor() {
    let parts: Vec<(String, Vec<u8>)> = ooxml_opc::unzip_parts(DECK)
        .unwrap()
        .into_iter()
        .map(|(name, bytes)| match name.as_str() {
            "ppt/slides/slide1.xml" => {
                let xml = String::from_utf8(bytes).unwrap();
                let plain = r#"<a:pPr algn="l" lvl="0"></a:pPr>"#;
                let at = xml[..xml.find(DECK_STORY[1]).unwrap()]
                    .rfind(plain)
                    .unwrap();
                let xml = format!(
                    r#"{}<a:pPr marL="457200" lvl="1" algn="ctr"><a:buChar char="-"/></a:pPr>{}"#,
                    &xml[..at],
                    &xml[at + plain.len()..]
                );
                (name, xml.into_bytes())
            }
            _ => (name, bytes),
        })
        .collect();
    let mut doc = Doc::new(Format::Pptx, ooxml_opc::rezip_parts(&parts).unwrap());
    let (first, second) = (doc.id(DECK_STORY[0]), doc.id(DECK_STORY[1]));
    doc.apply(&[after(&first, "After"), before(&second, "Before")])
        .unwrap();
    let slide = doc.part("ppt/slides/slide1.xml");
    let level = |text: &str| {
        let xml = paragraph_in(&slide, "a", text);
        (xml.contains(r#"lvl="1""#), xml.contains("<a:buChar"))
    };
    assert_eq!(level(DECK_STORY[0]), (false, false));
    assert_eq!(level("After"), (false, false));
    assert_eq!(level("Before"), (true, true));
    assert_eq!(level(DECK_STORY[1]), (true, true));
}

/// An equation's text box (PowerPoint's spelling: the math in an `a14`
/// choice, a text box picturing it as the fallback) takes no new paragraph.
#[test]
fn pptx_insert_into_an_equation_is_refused_as_invalid_input() {
    let equation = concat!(
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
                let xml = xml.replacen("</p:spTree>", &format!("{equation}</p:spTree>"), 1);
                (name, xml.into_bytes())
            }
            _ => (name, bytes),
        })
        .collect();
    let mut doc = Doc::new(Format::Pptx, ooxml_opc::rezip_parts(&parts).unwrap());
    let id = doc.id("x squared");
    for command in [after(&id, "y"), before(&id, "y")] {
        let error = doc.apply(&[command]).unwrap_err();
        assert!(
            matches!(&error, Error::Edit { code: EditCode::InvalidInput, message, details: None }
                if message.ends_with("is an equation: its math can only be changed in PowerPoint")),
            "{error:?}"
        );
    }
}

#[test]
fn pptx_insert_takes_no_style() {
    let mut doc = Doc::new(Format::Pptx, DECK.to_vec());
    let first = doc.id(DECK_STORY[0]);
    assert_eq!(
        doc.apply(&[styled(after(&first, "x"), "Title")])
            .unwrap_err(),
        invalid(
            "style names a DOCX paragraph style; a new PPTX paragraph takes its neighbour's formatting"
        )
    );
}

#[test]
fn paragraph_commands_are_refused_on_a_workbook() {
    let workbook = office_service::starter(Format::Xlsx).unwrap().to_vec();
    let mut doc = Doc::new(Format::Xlsx, workbook);
    for command in [after("x", "y"), delete("x", "y")] {
        assert_eq!(
            doc.apply(&[command]).unwrap_err(),
            refused(
                EditCode::UnsupportedOperation,
                "this command edits DOCX and PPTX files"
            )
        );
    }
}

#[test]
fn paragraph_commands_parse_from_their_wire_json() {
    let commands: Vec<Command> = serde_json::from_value(json!([
        { "type": "insert_paragraph", "afterId": "body:paragraph:1", "text": "Hi", "style": "Heading1" },
        { "type": "insert_paragraph", "beforeId": "body:paragraph:1", "text": "" },
        { "type": "delete_paragraph", "targetId": "body:paragraph:1", "expectedText": "Hi" },
    ]))
    .unwrap();
    assert_eq!(
        commands,
        [
            styled(after("body:paragraph:1", "Hi"), "Heading1"),
            before("body:paragraph:1", ""),
            delete("body:paragraph:1", "Hi"),
        ]
    );
}
