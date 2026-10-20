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
        edit.inverse.as_deref().unwrap(),
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
        edit.inverse.as_deref().unwrap(),
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
        edit.inverse.as_deref().unwrap(),
    )
    .unwrap();
    assert_eq!(
        value(DECK, Format::Pptx, &undone.state, &target.id),
        target.value
    );
}

/// `base` with `from`, which occurs once in the part at `path`, replaced.
fn with_part(base: &[u8], path: &str, from: &str, to: &str) -> Vec<u8> {
    let parts: Vec<(String, Vec<u8>)> = ooxml_opc::unzip_parts(base)
        .unwrap()
        .into_iter()
        .map(|(name, bytes)| {
            if name != path {
                return (name, bytes);
            }
            let xml = String::from_utf8(bytes).unwrap();
            assert_eq!(xml.matches(from).count(), 1, "{from} in {path}");
            (name, xml.replace(from, to).into_bytes())
        })
        .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn invalid(message: &str) -> Error {
    Error::Edit {
        code: EditCode::InvalidInput,
        message: message.into(),
        details: None,
    }
}

/// An agent's replacement holding characters an Office file cannot store
/// applies without them, and the result names them for the tool result (Epo
/// 2026-10-15); a line break the DOCX paragraph cannot hold is still the
/// engine's edit refusal, naming the rule.
#[test]
fn replacement_text_loses_what_an_office_file_cannot_store_and_says_so() {
    let base = docx();
    let sha = sha256_hex(&base);
    let state = seed(Format::Docx, &base).unwrap();
    let entries = inspect(&base, checkpoint(Format::Docx, &sha, &state)).unwrap();
    let liked = entries.iter().find(|entry| entry.value == LIKED).unwrap();
    let edit = |text: &str| {
        apply_commands(
            &base,
            checkpoint(Format::Docx, &sha, &state),
            &[replace(liked, "cats", text)],
        )
    };
    let applied = edit("d\u{1}o\u{c}g\u{1}s\u{fffe}").unwrap();
    assert_eq!(
        value(&base, Format::Docx, &applied.state, &liked.id),
        LIKED.replace("cats", "dogs")
    );
    assert_eq!(
        applied.notice.as_deref(),
        Some("Removed 4 characters Office files cannot store (U+0001, U+000C, U+FFFE).")
    );
    assert_eq!(edit("dogs").unwrap().notice, None);
    assert_eq!(
        edit("do\rgs").unwrap_err(),
        invalid("inserted text may not contain paragraph or line breaks")
    );

    let sha = sha256_hex(DECK);
    let seeded = seed(Format::Pptx, DECK).unwrap();
    let target = inspect(DECK, checkpoint(Format::Pptx, &sha, &seeded))
        .unwrap()
        .into_iter()
        .find(|entry| !entry.value.is_empty())
        .unwrap();
    let applied = apply_commands(
        DECK,
        checkpoint(Format::Pptx, &sha, &seeded),
        &[replace(&target, &target.value, "a\u{1}b")],
    )
    .unwrap();
    assert_eq!(value(DECK, Format::Pptx, &applied.state, &target.id), "ab");
    assert_eq!(
        applied.notice.as_deref(),
        Some("Removed 1 character Office files cannot store (U+0001).")
    );
}

/// PowerPoint's spelling of an equation: the math in an `a14` choice, a text
/// box picturing it as the fallback.
const EQUATION: &str = concat!(
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

/// An agent edit of an equation's text is refused as the deck refuses it,
/// as `invalid_input` naming the rule.
#[test]
fn pptx_replace_text_in_an_equation_is_refused_as_invalid_input() {
    let base = with_part(
        DECK,
        "ppt/slides/slide1.xml",
        "</p:spTree>",
        &format!("{EQUATION}</p:spTree>"),
    );
    let sha = sha256_hex(&base);
    let seeded = seed(Format::Pptx, &base).unwrap();
    let target = inspect(&base, checkpoint(Format::Pptx, &sha, &seeded))
        .unwrap()
        .into_iter()
        .find(|entry| entry.value == "x squared")
        .unwrap();
    let error = apply_commands(
        &base,
        checkpoint(Format::Pptx, &sha, &seeded),
        &[replace(&target, "squared", "cubed")],
    )
    .unwrap_err();
    assert!(
        matches!(&error, Error::Edit { code: EditCode::InvalidInput, message, details: None }
            if message.ends_with("is an equation: its math can only be changed in PowerPoint")),
        "{error:?}"
    );
}

/// The Undo of an agent edit restores the paragraph as the file can store it:
/// a character the file cannot hold, there before the edit (a paste or an
/// import), is dropped instead of refusing the Undo.
#[test]
fn undo_restores_a_paragraph_without_a_character_the_file_cannot_store() {
    for (format, base, path, from, to, expected, text) in [
        (
            Format::Docx,
            docx(),
            "word/document.xml",
            "I like ",
            "I\u{1} like ",
            "cats",
            "dogs",
        ),
        (
            Format::Pptx,
            DECK.to_vec(),
            "ppt/slides/slide1.xml",
            "<a:t>Office files,</a:t>",
            "<a:t>Office\u{1} files,</a:t>",
            "files",
            "docs",
        ),
    ] {
        let base = with_part(&base, path, from, to);
        let sha = sha256_hex(&base);
        let state = seed(format, &base).unwrap();
        let target = inspect(&base, checkpoint(format, &sha, &state))
            .unwrap()
            .into_iter()
            .find(|entry| entry.value.contains('\u{1}'))
            .unwrap();
        let edit = apply_commands(
            &base,
            checkpoint(format, &sha, &state),
            &[replace(&target, expected, text)],
        )
        .unwrap();
        let undone = apply_commands(
            &base,
            checkpoint(format, &sha, &edit.state),
            edit.inverse.as_deref().unwrap(),
        )
        .unwrap();
        assert_eq!(
            value(&base, format, &undone.state, &target.id),
            target.value.replace('\u{1}', ""),
            "{path}"
        );
    }
}
