//! Lists made with the editor's bullet and number buttons save with numbering
//! definitions of their own (Epo 2026-10-13): the buttons write the next two
//! `w:num` ids after the source's (`docx_parse::numbering::list_button_num_ids`),
//! and the save defines each id a paragraph uses that the numbering lacks as
//! that paragraph's kind, in the numbering part (created when missing).

use std::collections::BTreeMap;
use std::sync::Arc;

use docx_edit::{EditCtx, EditingDoc, ParaAttrDelta, ParaSelector, ParagraphId};
use docx_parse::numbering::{list_button_num_ids, parse_numbering};
use docx_parse::xml::{ParseBudget, ParseLimits};
use office_service::{Checkpoint, Determinism, Format, export, rebase, seed, sha256_hex};
use yrs::Any;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const OFFICE: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml";

fn package(numbering: Option<&str>) -> Vec<u8> {
    named_package("numbering.xml", numbering)
}

/// A package whose numbering part, when it has one, is `word/{name}`.
fn named_package(name: &str, numbering: Option<&str>) -> Vec<u8> {
    let (override_xml, rel) = if numbering.is_some() {
        (
            format!(r#"<Override PartName="/word/{name}" ContentType="{OFFICE}.numbering+xml"/>"#),
            format!(r#"<Relationship Id="rIdNum" Type="{R}/numbering" Target="{name}"/>"#),
        )
    } else {
        Default::default()
    };
    let body = ["Apples", "First step", "Pears"]
        .iter()
        .enumerate()
        .map(|(index, text)| {
            format!(r#"<w:p w14:paraId="1000000{index}"><w:r><w:t>{text}</w:t></w:r></w:p>"#)
        })
        .collect::<String>();
    let mut parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            format!(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{OFFICE}.document.main+xml"/>{override_xml}</Types>"#
            ),
        ),
        (
            "_rels/.rels".to_owned(),
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>"#
            ),
        ),
        (
            "word/_rels/document.xml.rels".to_owned(),
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rel}</Relationships>"#
            ),
        ),
        (
            "word/document.xml".to_owned(),
            format!(
                r#"<w:document xmlns:w="{W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"#
            ),
        ),
    ];
    if let Some(numbering) = numbering {
        parts.push((format!("word/{name}"), numbering.to_owned()));
    }
    let parts: Vec<(String, Vec<u8>)> = parts
        .into_iter()
        .map(|(path, xml)| (path, xml.into_bytes()))
        .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// numbering.xml with numId 1 numbered and numId 2 bulleted, as Word may write.
fn word_numbering() -> String {
    let level = |format: &str, text: &str| {
        format!(
            r#"<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="{format}"/><w:lvlText w:val="{text}"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>"#
        )
    };
    format!(
        r#"<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="0">{}</w:abstractNum><w:abstractNum w:abstractNumId="1">{}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>"#,
        level("decimal", "%1."),
        level("bullet", "\u{f0b7}")
    )
}

fn open(state: &[u8], client: u64) -> EditingDoc {
    let doc = EditingDoc::new(client);
    doc.load_state_v1(state).unwrap();
    doc
}

fn paragraph(doc: &EditingDoc, index: usize) -> ParagraphId {
    doc.paragraphs("body").unwrap()[index].para_id.clone()
}

/// What the toolbar's bullet (`bullet`) or number button writes, with `num_id`.
fn list_button(doc: &EditingDoc, paragraph: ParagraphId, bullet: bool, num_id: f64) {
    let delta = ParaAttrDelta {
        other: [
            (
                "numPr".to_owned(),
                Some(Any::Map(Arc::new(
                    [
                        ("numId".to_owned(), Any::Number(num_id)),
                        ("ilvl".to_owned(), Any::Number(0.0)),
                    ]
                    .into(),
                ))),
            ),
            ("listIsBullet".to_owned(), Some(Any::Bool(bullet))),
            (
                "listNumFmt".to_owned(),
                Some(if bullet {
                    Any::Null
                } else {
                    Any::from("decimal")
                }),
            ),
            ("listMarker".to_owned(), Some(Any::Null)),
        ]
        .into(),
        ..ParaAttrDelta::default()
    };
    doc.set_paragraph_attrs(
        &EditCtx::local("", ""),
        &ParaSelector::One(paragraph),
        &delta,
    )
    .unwrap();
}

fn checkpoint<'a>(base_sha256: &'a str, state: &'a [u8]) -> Checkpoint<'a> {
    Checkpoint {
        format: Format::Docx,
        schema_version: 1,
        base_sha256,
        state,
    }
}

fn save(base: &[u8], doc: &EditingDoc) -> Vec<u8> {
    let state = doc.encode_state_as_update_v1();
    let sha = sha256_hex(base);
    export(
        base,
        checkpoint(&sha, &state),
        Determinism {
            seed: &"0".repeat(64),
            now: "2026-10-13T00:00:00.000Z",
        },
    )
    .unwrap()
}

/// The ids the buttons write over `base`, read from its numbering part.
fn button_ids(base: &[u8]) -> (f64, f64) {
    let limits = ParseLimits::default();
    let mut budget = ParseBudget::new(&limits);
    let xml = part(base, "word/numbering.xml");
    let numbering = parse_numbering(
        (!xml.is_empty()).then_some(xml.as_bytes()),
        "word/numbering.xml",
        &mut budget,
    )
    .unwrap();
    list_button_num_ids(Some(&numbering.definitions))
}

fn part(saved: &[u8], path: &str) -> String {
    ooxml_opc::unzip_parts(saved)
        .unwrap()
        .into_iter()
        .find(|(name, _)| name == path)
        .map(|(_, bytes)| String::from_utf8(bytes).unwrap())
        .unwrap_or_default()
}

/// Each paragraph's list as the saved file shows it once reopened.
fn reopened_lists(saved: &[u8]) -> Vec<String> {
    let doc = open(&seed(Format::Docx, saved).unwrap(), 4302);
    doc.paragraphs("body")
        .unwrap()
        .iter()
        .map(|paragraph| {
            let values: &BTreeMap<String, Any> = &paragraph.properties;
            match (
                values.get("numPr"),
                values.get("listIsBullet"),
                values.get("listNumFmt"),
            ) {
                (None | Some(Any::Null), _, _) => "none".to_owned(),
                (_, Some(Any::Bool(true)), _) => "bullet".to_owned(),
                (_, _, Some(Any::String(format))) => format.to_string(),
                other => format!("{other:?}"),
            }
        })
        .collect()
}

#[test]
fn toolbar_lists_save_with_their_own_definitions_where_the_file_has_no_numbering() {
    let base = package(None);
    let doc = open(&seed(Format::Docx, &base).unwrap(), 4301);
    list_button(&doc, paragraph(&doc, 0), true, 1.0);
    list_button(&doc, paragraph(&doc, 1), false, 2.0);
    let saved = save(&base, &doc);
    assert!(part(&saved, "word/numbering.xml").contains(r#"<w:num w:numId="1">"#));
    assert!(part(&saved, "[Content_Types].xml").contains(r#"PartName="/word/numbering.xml""#));
    assert!(part(&saved, "word/_rels/document.xml.rels").contains(r#"Target="numbering.xml""#));
    assert_eq!(reopened_lists(&saved), ["bullet", "decimal", "none"]);
}

#[test]
fn toolbar_bullets_save_as_bullets_where_the_files_num_id_1_is_numbered() {
    let base = package(Some(&word_numbering()));
    let doc = open(&seed(Format::Docx, &base).unwrap(), 4301);
    list_button(&doc, paragraph(&doc, 0), true, 3.0);
    list_button(&doc, paragraph(&doc, 1), false, 4.0);
    list_button(&doc, paragraph(&doc, 2), false, 1.0);
    let saved = save(&base, &doc);
    let numbering = part(&saved, "word/numbering.xml");
    // The file's own definitions stay as they were, ahead of the added ones.
    assert!(numbering.starts_with(&word_numbering()[..word_numbering().find("<w:num ").unwrap()]));
    assert!(numbering.contains(r#"<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num><w:num w:numId="4"><w:abstractNumId w:val="3"/></w:num></w:numbering>"#));
    assert_eq!(reopened_lists(&saved), ["bullet", "decimal", "decimal"]);
}

/// Alice makes a bulleted list while Bob makes a numbered one: one save
/// defines each list once.
#[test]
fn two_peers_toolbar_lists_save_each_definition_once() {
    let base = package(None);
    let seeded = seed(Format::Docx, &base).unwrap();
    let (alice, bob) = (open(&seeded, 4311), open(&seeded, 4312));
    list_button(&alice, paragraph(&alice, 0), true, 1.0);
    list_button(&bob, paragraph(&bob, 1), false, 2.0);
    list_button(&bob, paragraph(&bob, 2), true, 1.0);
    alice
        .apply_update_v1(&bob.encode_diff_v1(&alice.encode_state_vector_v1()).unwrap())
        .unwrap();
    let saved = save(&base, &alice);
    let numbering = part(&saved, "word/numbering.xml");
    assert_eq!(numbering.matches("<w:num ").count(), 2);
    assert_eq!(numbering.matches("<w:abstractNum ").count(), 2);
    assert_eq!(reopened_lists(&saved), ["bullet", "decimal", "bullet"]);
}

/// A publication captures one list; the other kind is made after the capture
/// and rebased onto the published file, then a new list is made over that
/// file. Every list saves as it shows, in both orders (the save defines the
/// buttons' other id with the first, so the next base's ids are new).
#[test]
fn toolbar_lists_keep_their_kind_across_a_publication() {
    for first_bullet in [true, false] {
        let base = package(None);
        let (bullet_id, number_id) = button_ids(&base);
        let id = |bullet: bool| if bullet { bullet_id } else { number_id };
        let doc = open(&seed(Format::Docx, &base).unwrap(), 4321);
        list_button(&doc, paragraph(&doc, 0), first_bullet, id(first_bullet));
        let captured = doc.encode_state_as_update_v1();
        let published = save(&base, &doc);
        list_button(&doc, paragraph(&doc, 1), !first_bullet, id(!first_bullet));
        let latest = doc.encode_state_as_update_v1();
        let sha = sha256_hex(&base);
        let rebased = rebase(
            &base,
            checkpoint(&sha, &captured),
            checkpoint(&sha, &latest),
            &published,
        )
        .unwrap();
        let next = open(&rebased.state, 4322);
        let (next_bullet, _) = button_ids(&published);
        assert!(next_bullet > bullet_id.max(number_id), "{next_bullet}");
        list_button(&next, paragraph(&next, 2), true, next_bullet);
        let kind = |bullet: bool| if bullet { "bullet" } else { "decimal" };
        assert_eq!(
            reopened_lists(&save(&published, &next)),
            [kind(first_bullet), kind(!first_bullet), "bullet"],
            "first bullet {first_bullet}"
        );
    }
}

/// The numbering part is found through the document's relationships. (The
/// parser itself reads only `word/numbering.xml`, so this checks the XML.)
#[test]
fn toolbar_lists_save_into_a_numbering_part_of_another_name() {
    let base = named_package("lists.xml", Some(&word_numbering()));
    let doc = open(&seed(Format::Docx, &base).unwrap(), 4331);
    list_button(&doc, paragraph(&doc, 0), true, 3.0);
    let saved = save(&base, &doc);
    let lists = part(&saved, "word/lists.xml");
    assert!(lists.starts_with(&word_numbering()[..word_numbering().find("<w:num ").unwrap()]));
    assert!(lists.contains(
        r#"<w:abstractNum w:abstractNumId="2"><w:multiLevelType w:val="hybridMultilevel"/>"#
    ));
    assert!(
        lists.contains(r#"<w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num></w:numbering>"#)
    );
    assert_eq!(part(&saved, "word/numbering.xml"), "");
    assert_eq!(
        part(&saved, "word/_rels/document.xml.rels")
            .matches("/numbering\"")
            .count(),
        1
    );
}
