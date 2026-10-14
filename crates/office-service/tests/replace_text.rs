//! An agent's `replace_text` matches `expected_text` as a substring that
//! occurs exactly once in the target paragraph (as note edits do), replaces
//! only that span, and a refusal of the match carries the paragraph's current
//! text (Epo 2026-10-14). DOCX and PPTX share the rule; XLSX edits cells.

use office_service::{
    Checkpoint, Command, Determinism, EditCode, Entry, Error, Format, apply_commands, export,
    inspect, seed, sha256_hex,
};
use pptx_edit::{DeckSession, EditCtx, TextStyle};
use serde_json::json;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const DECK: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

/// The first paragraph of [`docx`]: its prefix is longer in UTF-8 than in
/// UTF-16 (CJK, an astral emoji, a combining accent).
const LIKED: &str = "日本😀e\u{301} I like cats a lot, a lot";

/// A DOCX whose first paragraph reads [`LIKED`] with "cats" bold, followed by
/// an empty paragraph.
fn docx() -> Vec<u8> {
    let run = |text: &str, bold: bool| {
        let bold = if bold { "<w:rPr><w:b/></w:rPr>" } else { "" };
        format!(r#"<w:r>{bold}<w:t xml:space="preserve">{text}</w:t></w:r>"#)
    };
    let files = [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_owned(),
        ),
        (
            "_rels/.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned(),
        ),
        (
            "word/document.xml",
            format!(
                r#"<w:document xmlns:w="{W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="0000000A">{}{}{}</w:p><w:p w14:paraId="0000000B"/><w:sectPr/></w:body></w:document>"#,
                run("日本😀e\u{301} I like ", false),
                run("cats", true),
                run(" a lot, a lot", false)
            ),
        ),
    ];
    let files: Vec<(String, Vec<u8>)> = files
        .into_iter()
        .map(|(path, xml)| (path.to_owned(), xml.into_bytes()))
        .collect();
    ooxml_opc::rezip_parts(&files).unwrap()
}

/// The text of an XML fragment, without its tags.
fn text_of(xml: &str) -> String {
    let mut in_tag = false;
    xml.chars()
        .filter(|&c| {
            match c {
                '<' => in_tag = true,
                '>' => in_tag = false,
                _ => return !in_tag,
            }
            false
        })
        .collect()
}

fn checkpoint<'a>(format: Format, sha: &'a str, state: &'a [u8]) -> Checkpoint<'a> {
    Checkpoint {
        format,
        schema_version: 1,
        base_sha256: sha,
        state,
    }
}

fn replace(target: &Entry, expected: &str, text: &str) -> Command {
    Command::ReplaceText {
        target_id: target.id.clone(),
        expected_text: expected.into(),
        text: text.into(),
    }
}

/// The text `id` reads in `state`.
fn value(base: &[u8], format: Format, state: &[u8], id: &str) -> String {
    let sha = sha256_hex(base);
    inspect(base, checkpoint(format, &sha, state))
        .unwrap()
        .into_iter()
        .find(|entry| entry.id == id)
        .unwrap()
        .value
}

/// The refusal of `expected` in `target`: its code, message and details.
fn refusal(base: &[u8], format: Format, state: &[u8], target: &Entry, expected: &str) -> Error {
    let sha = sha256_hex(base);
    apply_commands(
        base,
        checkpoint(format, &sha, state),
        &[replace(target, expected, "x")],
    )
    .unwrap_err()
}

/// Every refusal of the match names the rule and carries the paragraph as it
/// reads now, the way a note block's refusal does.
fn assert_refusals(base: &[u8], format: Format, state: &[u8], target: &Entry, repeated: &str) {
    let details = Some(json!({
        "block": { "id": target.id, "type": "paragraph", "text": target.value }
    }));
    let count = target.value.matches(repeated).count();
    assert!(count > 1, "{repeated:?} repeats in {:?}", target.value);
    assert_eq!(
        refusal(base, format, state, target, repeated),
        Error::Edit {
            code: EditCode::InvalidInput,
            message: format!(
                "expected text occurs {count} times in the paragraph; include more context"
            ),
            details: details.clone(),
        }
    );
    assert_eq!(
        refusal(base, format, state, target, "not in this paragraph"),
        Error::Edit {
            code: EditCode::StaleTarget,
            message: "expected text was not found in the paragraph".into(),
            details: details.clone(),
        }
    );
    assert_eq!(
        refusal(base, format, state, target, ""),
        Error::Edit {
            code: EditCode::InvalidInput,
            message:
                "expected_text must not be empty unless the paragraph is; include surrounding text"
                    .into(),
            details,
        }
    );
}

#[test]
fn docx_replace_text_replaces_its_one_occurrence_keeping_the_runs_formatting() {
    let base = docx();
    let sha = sha256_hex(&base);
    let state = seed(Format::Docx, &base).unwrap();
    let entries = inspect(&base, checkpoint(Format::Docx, &sha, &state)).unwrap();
    let liked = entries.iter().find(|entry| entry.value == LIKED).unwrap();
    assert_refusals(&base, Format::Docx, &state, liked, "a lot");

    let edit = apply_commands(
        &base,
        checkpoint(Format::Docx, &sha, &state),
        &[replace(liked, "cats", "dogs")],
    )
    .unwrap();
    assert_eq!(
        value(&base, Format::Docx, &edit.state, &liked.id),
        LIKED.replace("cats", "dogs")
    );
    // Only "cats" was rewritten: the bold run now holds "dogs" alone.
    let saved = export(
        &base,
        checkpoint(Format::Docx, &sha, &edit.state),
        Determinism {
            seed: &"0".repeat(64),
            now: "2026-10-14T00:00:00.000Z",
        },
    )
    .unwrap();
    let xml = ooxml_opc::unzip_parts(&saved)
        .unwrap()
        .into_iter()
        .find(|(path, _)| path == "word/document.xml")
        .map(|(_, bytes)| String::from_utf8(bytes).unwrap())
        .unwrap();
    let bold: Vec<String> = xml
        .split("<w:r>")
        .skip(1)
        .filter_map(|run| run.split("</w:r>").next())
        .filter(|run| run.contains("<w:b/>"))
        .map(text_of)
        .collect();
    assert_eq!(bold, ["dogs"], "{xml}");

    let undone = apply_commands(
        &base,
        checkpoint(Format::Docx, &sha, &edit.state),
        &edit.inverse,
    )
    .unwrap();
    assert_eq!(
        value(&base, Format::Docx, &undone.state, &liked.id),
        liked.value
    );
}

#[test]
fn an_empty_expected_text_fills_an_empty_paragraph_and_an_overlapping_one_counts_once() {
    let base = docx();
    let sha = sha256_hex(&base);
    let state = seed(Format::Docx, &base).unwrap();
    let entries = inspect(&base, checkpoint(Format::Docx, &sha, &state)).unwrap();
    let empty = entries.iter().find(|entry| entry.value.is_empty()).unwrap();
    let edit = apply_commands(
        &base,
        checkpoint(Format::Docx, &sha, &state),
        &[replace(empty, "", "aaa"), replace(empty, "aa", "b")],
    )
    .unwrap();
    // "aa" occurs once in "aaa", as note edits count it: the first two go.
    assert_eq!(value(&base, Format::Docx, &edit.state, &empty.id), "ba");
    let undone = apply_commands(
        &base,
        checkpoint(Format::Docx, &sha, &edit.state),
        &edit.inverse,
    )
    .unwrap();
    assert_eq!(value(&base, Format::Docx, &undone.state, &empty.id), "");
}

#[test]
fn pptx_replace_text_replaces_its_one_occurrence_in_the_style_it_replaces() {
    let sha = sha256_hex(DECK);
    let seeded = seed(Format::Pptx, DECK).unwrap();
    let target = inspect(DECK, checkpoint(Format::Pptx, &sha, &seeded))
        .unwrap()
        .into_iter()
        .find(|entry| !entry.value.is_empty())
        .unwrap();
    let story = target.position[..target.position.rfind(':').unwrap()].to_owned();
    let story = story.as_str();
    // The paragraph rewritten as "Plain Bold, Plain" with "Bold" bold.
    let deck = DeckSession::open_from_update_with_source(&seeded, DECK, 9998).unwrap();
    let snapshot = deck.story(story).unwrap();
    let start: u32 = snapshot
        .paragraphs
        .iter()
        .take_while(|paragraph| paragraph.id != target.id)
        .map(|paragraph| {
            paragraph
                .runs
                .iter()
                .map(|run| run.text.encode_utf16().count() as u32)
                .sum::<u32>()
                + 1
        })
        .sum();
    let ctx = EditCtx::local("test");
    let length = target.value.encode_utf16().count() as u32;
    deck.delete_text(&ctx, story, start, start + length)
        .unwrap();
    let plain = TextStyle::default();
    let bold = TextStyle {
        bold: Some(true),
        ..TextStyle::default()
    };
    deck.insert_text(&ctx, story, start, "Plain ", &plain)
        .unwrap();
    deck.insert_text(&ctx, story, start + 6, "Bold", &bold)
        .unwrap();
    deck.insert_text(&ctx, story, start + 10, ", Plain", &plain)
        .unwrap();
    let mixed = deck.encode_state_as_update_v1();
    let target = Entry {
        value: "Plain Bold, Plain".into(),
        ..target
    };
    let runs = |state: &[u8]| -> Vec<(String, bool)> {
        let deck = DeckSession::open_from_update_with_source(state, DECK, 9997).unwrap();
        deck.story(story)
            .unwrap()
            .paragraphs
            .into_iter()
            .find(|paragraph| paragraph.id == target.id)
            .unwrap()
            .runs
            .into_iter()
            .map(|run| (run.text, run.style.bold == Some(true)))
            .collect()
    };
    assert_eq!(value(DECK, Format::Pptx, &mixed, &target.id), target.value);
    assert_refusals(DECK, Format::Pptx, &mixed, &target, "Plain");

    let edit = apply_commands(
        DECK,
        checkpoint(Format::Pptx, &sha, &mixed),
        &[
            replace(&target, "Bold", "Cold"),
            replace(&target, "Plain C", "Plainly C"),
        ],
    )
    .unwrap();
    assert_eq!(
        runs(&edit.state),
        [
            ("Plainly ".to_owned(), false),
            ("Cold".to_owned(), true),
            (", Plain".to_owned(), false)
        ]
    );
    let undone = apply_commands(
        DECK,
        checkpoint(Format::Pptx, &sha, &edit.state),
        &edit.inverse,
    )
    .unwrap();
    assert_eq!(
        value(DECK, Format::Pptx, &undone.state, &target.id),
        target.value
    );
}
