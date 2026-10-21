//! A DOCX save writes back only what an edit changed (Epo 2026-10-21): a run
//! whose formatting the session left saves its source formatting rather than
//! the values the seed resolved from styles, and a table cell saves only the
//! borders and margins that differ from what the seed resolved. A style
//! changed in Word afterwards still restyles that text, cells keep taking the
//! table style's borders and margins, and an unedited save no longer writes
//! twice the source's document.xml.

use std::collections::{BTreeMap, HashMap};

use docx_edit::{
    CanonicalItem, CellLoc, EditCtx, EditingDoc, FormatPolicy, InlineFormatDelta, ParaSelector,
    Patch, Position, StoryRange, TableRange, project_story,
};
use office_service::{Checkpoint, Determinism, Format, export, seed, sha256_hex};
use serde_json::Value;
use yrs::{Any, Map, ReadTxn, Transact};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const OFFICE: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml";

/// A heading styled red and bold, a plain paragraph with an italic run, and a
/// two-by-two table whose borders and cell margins come from its style.
fn styled_package() -> Vec<u8> {
    let styles = format!(
        r#"<w:styles xmlns:w="{W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:color w:val="FF0000"/><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="table" w:styleId="Grid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>"#
    );
    let cell = |text: &str| {
        format!(
            r#"<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>"#
        )
    };
    let body = format!(
        r#"<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">Plain words </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>slanted</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblStyle w:val="Grid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr>{}{}</w:tr><w:tr>{}{}</w:tr></w:tbl><w:p/>"#,
        cell("A"),
        cell("B"),
        cell("C"),
        cell("D")
    );
    package(&styles, &body)
}

/// A package of `body` with `styles`.
fn package(styles: &str, body: &str) -> Vec<u8> {
    ooxml_opc::rezip_parts(&[
        (
            "[Content_Types].xml".to_owned(),
            format!(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{OFFICE}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="{OFFICE}.styles+xml"/></Types>"#
            )
            .into_bytes(),
        ),
        (
            "_rels/.rels".to_owned(),
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>"#
            )
            .into_bytes(),
        ),
        (
            "word/_rels/document.xml.rels".to_owned(),
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{R}/styles" Target="styles.xml"/></Relationships>"#
            )
            .into_bytes(),
        ),
        ("word/styles.xml".to_owned(), styles.as_bytes().to_vec()),
        (
            "word/document.xml".to_owned(),
            format!(
                r#"<w:document xmlns:w="{W}"><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"#
            )
            .into_bytes(),
        ),
    ])
    .unwrap()
}

fn opened(bytes: &[u8]) -> EditingDoc {
    let doc = EditingDoc::new(4101);
    doc.load_state_v1(&seed(Format::Docx, bytes).unwrap())
        .unwrap();
    doc
}

fn saved(base: &[u8], doc: &EditingDoc) -> Vec<u8> {
    let state = doc.encode_state_as_update_v1();
    let sha = sha256_hex(base);
    export(
        base,
        Checkpoint {
            format: Format::Docx,
            schema_version: 1,
            base_sha256: &sha,
            state: &state,
        },
        Determinism {
            seed: &"0".repeat(64),
            now: "2026-10-21T00:00:00.000Z",
        },
    )
    .unwrap()
}

fn part(package: &[u8], name: &str) -> String {
    ooxml_opc::unzip_parts(package)
        .unwrap()
        .into_iter()
        .find(|(path, _)| path == name)
        .map(|(_, xml)| String::from_utf8(xml).unwrap())
        .unwrap()
}

fn with_part(package: &[u8], name: &str, xml: String) -> Vec<u8> {
    let parts: Vec<(String, Vec<u8>)> = ooxml_opc::unzip_parts(package)
        .unwrap()
        .into_iter()
        .map(|(path, bytes)| {
            let bytes = if path == name {
                xml.clone().into_bytes()
            } else {
                bytes
            };
            (path, bytes)
        })
        .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// A character and the formatting attributes it renders with.
type Styled = (String, BTreeMap<String, Value>);

/// Every story's characters as they render.
fn look(doc: &EditingDoc) -> BTreeMap<String, Vec<Styled>> {
    let stories: Vec<String> = {
        let txn = doc.yrs_doc().transact();
        txn.get_map("stories")
            .map(|stories| stories.keys(&txn).map(str::to_owned).collect())
            .unwrap_or_default()
    };
    stories
        .into_iter()
        .map(|story| {
            let chars = project_story(doc, &story)
                .unwrap()
                .into_iter()
                .filter_map(|item| match item {
                    CanonicalItem::CharItem { ch, marks, .. } => Some((ch, marks)),
                    _ => None,
                })
                .collect();
            (story, chars)
        })
        .collect()
}

fn marks_of(doc: &EditingDoc, story: &str, text: &str) -> Vec<BTreeMap<String, Value>> {
    let chars = look(doc).remove(story).unwrap();
    let joined: String = chars.iter().map(|(ch, _)| ch.as_str()).collect();
    let start = joined.find(text).unwrap();
    chars
        .into_iter()
        .skip(joined[..start].chars().count())
        .take(text.chars().count())
        .map(|(_, marks)| marks)
        .collect()
}

/// Each run's `w:rPr` in `xml` in order, "" for a run without one.
fn run_properties(xml: &str) -> Vec<String> {
    let mut found = Vec::new();
    let mut at = 0;
    while let Some(offset) = xml[at..].find("<w:r") {
        let start = at + offset;
        at = start + 4;
        if !matches!(xml.as_bytes().get(at), Some(b'>' | b' ')) {
            continue;
        }
        let rest = &xml[start + xml[start..].find('>').unwrap() + 1..];
        found.push(if rest.starts_with("<w:rPr>") {
            rest[..rest.find("</w:rPr>").unwrap() + "</w:rPr>".len()].to_owned()
        } else {
            String::new()
        });
    }
    found
}

/// The cases both saves must write the same run properties for
/// (`packages/docx/src/yrs/sourceFormattingSave.test.ts` reads it too).
fn cross_port_cases() -> Value {
    serde_json::from_str(include_str!("fixtures/source-formatting-cases.json")).unwrap()
}

/// A case's edit: types `insert` at `at` of the body, or clears `clear`
/// (bold or italic) over `[start, end)` of it, as the toolbar's toggle does.
fn edit_case(doc: &EditingDoc, edit: &Value) {
    let at = |key: &str| edit[key].as_u64().unwrap() as u32;
    let ctx = EditCtx::local("", "");
    if let Some(text) = edit["insert"].as_str() {
        doc.insert_text(
            &ctx,
            Position::new("body", at("at")),
            text,
            FormatPolicy::Inherit,
        )
        .unwrap();
        return;
    }
    let delta = match edit["clear"].as_str().unwrap() {
        "bold" => InlineFormatDelta {
            bold: Patch::Clear,
            ..InlineFormatDelta::default()
        },
        "italic" => InlineFormatDelta {
            italic: Patch::Clear,
            ..InlineFormatDelta::default()
        },
        other => panic!("no case clears {other}"),
    };
    doc.format_range(
        &ctx,
        StoryRange::new("body", at("start"), at("end")),
        &delta,
    )
    .unwrap();
}

#[test]
fn both_saves_write_these_run_properties() {
    let cases = cross_port_cases();
    let styles = cases["styles"].as_str().unwrap();
    for case in cases["cases"].as_array().unwrap() {
        let base = package(styles, case["body"].as_str().unwrap());
        let doc = opened(&base);
        if let Some(edit) = case.get("edit") {
            edit_case(&doc, edit);
        }
        let written = run_properties(&part(&saved(&base, &doc), "word/document.xml"));
        assert_eq!(
            serde_json::to_value(&written).unwrap(),
            case["runs"],
            "{}",
            case["name"]
        );
    }
    for fixture in cases["fixtures"].as_array().unwrap() {
        let file = format!(
            "{}/../../{}",
            env!("CARGO_MANIFEST_DIR"),
            fixture["path"].as_str().unwrap()
        );
        let base = std::fs::read(file).unwrap();
        let written = run_properties(&part(&saved(&base, &opened(&base)), "word/document.xml"));
        assert_eq!(
            sha256_hex(written.join("\n").as_bytes()),
            fixture["runs"].as_str().unwrap(),
            "{}",
            fixture["name"]
        );
    }
}

#[test]
fn a_restyled_paragraph_saves_what_the_editor_shows() {
    let base = styled_package();
    let doc = opened(&base);
    let ctx = EditCtx::local("", "");
    let previous = BTreeMap::from([("Heading1".to_owned(), BTreeMap::<String, Any>::new())]);
    doc.apply_paragraph_style(
        &ctx,
        &ParaSelector::Range(StoryRange::new("body", 0, 1)),
        "Normal",
        &BTreeMap::new(),
        &previous,
    )
    .unwrap();
    // The editor keeps showing the heading's bold red 16 pt under Normal;
    // making the text follow the new style is the seed-side todo.
    let heading = |marks: &BTreeMap<String, Value>| {
        marks.contains_key("bold")
            && marks["textColor"]["rgb"] == "FF0000"
            && marks["fontSize"]["size"] == 32
    };
    assert!(marks_of(&doc, "body", "Title").iter().all(heading));
    let saved = saved(&base, &doc);
    assert!(
        marks_of(&opened(&saved), "body", "Title")
            .iter()
            .all(heading)
    );
}

#[test]
fn an_unedited_save_keeps_the_look_and_about_the_source_size() {
    // exchange-plan's TOC tab loses its run formatting on any save (before
    // this change too), so only its size is checked.
    for (name, base, same_look) in [
        (
            "wordprocessingml-comprehensive",
            &include_bytes!("../../../poc/fixtures/wordprocessingml-comprehensive.docx")[..],
            true,
        ),
        (
            "exchange-plan",
            include_bytes!("../../../poc/fixtures/exchange-plan.docx"),
            false,
        ),
        (
            "feature-rich",
            include_bytes!("../../../poc/fixtures/feature-rich.docx"),
            true,
        ),
        ("styled", &styled_package(), true),
    ] {
        let doc = opened(base);
        let saved = saved(base, &doc);
        let (before, after) = (look(&doc), look(&opened(&saved)));
        for (story, chars) in before.iter().filter(|_| same_look) {
            let reopened = &after[story];
            assert_eq!(chars.len(), reopened.len(), "{name} {story}");
            if let Some(at) = (0..chars.len()).find(|at| chars[*at] != reopened[*at]) {
                let text: String = chars[at.saturating_sub(20)..at + 1]
                    .iter()
                    .map(|(ch, _)| ch.as_str())
                    .collect();
                panic!(
                    "{name} {story} at {at} ({text:?}): {:?} became {:?}",
                    chars[at].1, reopened[at].1
                );
            }
        }
        let (source, written) = (
            part(base, "word/document.xml").len(),
            part(&saved, "word/document.xml").len(),
        );
        // Every run's resolved formatting made it 2.3x on feature-rich; the
        // inline package is mostly the save's fixed markup.
        assert!(
            name == "styled" || written * 100 <= source * 115,
            "{name}: {written} bytes for {source}"
        );
    }
}

#[test]
fn text_and_cells_the_session_left_follow_a_style_edited_in_word() {
    let base = styled_package();
    let saved = saved(&base, &opened(&base));
    let document = part(&saved, "word/document.xml");
    assert!(!document.contains("FF0000"), "{document}");
    assert!(!document.contains("<w:tcBorders>"), "{document}");
    assert!(!document.contains("<w:tcMar>"), "{document}");
    // Word, or our own engine, reads the edited style for that text.
    let restyled = part(&saved, "word/styles.xml").replace("FF0000", "0000FF");
    let reopened = opened(&with_part(&saved, "word/styles.xml", restyled));
    for marks in marks_of(&reopened, "body", "Title") {
        assert_eq!(marks["textColor"]["rgb"], "0000FF", "{marks:?}");
    }
}

#[test]
fn an_edited_run_and_cell_save_their_edits() {
    let base = styled_package();
    let doc = opened(&base);
    let ctx = EditCtx::local("", "");
    // "Plain" in the second paragraph: the title and its pilcrow come first.
    doc.format_range(
        &ctx,
        StoryRange::new("body", 6, 11),
        &InlineFormatDelta {
            bold: Patch::Set(true),
            ..InlineFormatDelta::default()
        },
    )
    .unwrap();
    let thick = Any::from_json(r#"{"style":"double","size":12,"color":{"rgb":"00FF00"}}"#).unwrap();
    let cell = CellLoc::new("body", 0, 0, 0);
    doc.set_cell_borders(
        &ctx,
        &TableRange::new(cell.clone(), cell),
        &HashMap::from([("top".to_owned(), thick)]),
    )
    .unwrap();
    let saved = saved(&base, &doc);
    let reopened = opened(&saved);
    assert!(
        marks_of(&reopened, "body", "Plain")
            .iter()
            .all(|marks| marks.contains_key("bold"))
    );
    assert!(
        marks_of(&reopened, "body", " words")
            .iter()
            .all(|marks| !marks.contains_key("bold"))
    );
    assert!(
        marks_of(&reopened, "body", "slanted")
            .iter()
            .all(|marks| marks.contains_key("italic"))
    );
    let document = part(&saved, "word/document.xml");
    // Only the edited side of the edited cell is written.
    assert_eq!(document.matches("<w:tcBorders>").count(), 1, "{document}");
    let borders = &document[document.find("<w:tcBorders>").unwrap()..];
    let borders = &borders[..borders.find("</w:tcBorders>").unwrap()];
    assert!(
        borders.contains(r#"<w:top w:val="double""#) && !borders.contains("<w:bottom"),
        "{borders}"
    );
}
