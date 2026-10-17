//! Comments and notes without a `w:id` are left out on open (upstream #1070),
//! so seeding, an edited export and its projection never name them.

use docx_edit::{EditCtx, EditingDoc, FormatPolicy, Position};
use office_service::{Checkpoint, Determinism, Format, export, seed, sha256_hex};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const OFFICE: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml";

fn package() -> Vec<u8> {
    let body = r#"<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>Kept</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r><w:commentRangeStart/><w:r><w:t> lost</w:t></w:r><w:commentRangeEnd/><w:r><w:commentReference/></w:r><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:footnoteReference/></w:r></w:p>"#;
    let files = [
        (
            "[Content_Types].xml",
            format!(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{OFFICE}.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="{OFFICE}.comments+xml"/><Override PartName="/word/footnotes.xml" ContentType="{OFFICE}.footnotes+xml"/></Types>"#
            ),
        ),
        (
            "_rels/.rels",
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>"#
            ),
        ),
        (
            "word/_rels/document.xml.rels",
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="{R}/comments" Target="comments.xml"/><Relationship Id="rId3" Type="{R}/footnotes" Target="footnotes.xml"/></Relationships>"#
            ),
        ),
        (
            "word/document.xml",
            format!(
                r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>"#
            ),
        ),
        (
            "word/comments.xml",
            format!(
                r#"<w:comments xmlns:w="{W}"><w:comment w:id="1" w:author="Grace"><w:p><w:r><w:t>One</w:t></w:r></w:p></w:comment><w:comment w:author="Missing"><w:p><w:r><w:t>No id</w:t></w:r></w:p></w:comment></w:comments>"#
            ),
        ),
        (
            "word/footnotes.xml",
            format!(
                r#"<w:footnotes xmlns:w="{W}"><w:footnote w:id="1"><w:p><w:r><w:t>Note one</w:t></w:r></w:p></w:footnote><w:footnote><w:p><w:r><w:t>No id note</w:t></w:r></w:p></w:footnote></w:footnotes>"#
            ),
        ),
    ];
    let files: Vec<(String, Vec<u8>)> = files
        .into_iter()
        .map(|(path, xml)| (path.to_owned(), xml.into_bytes()))
        .collect();
    ooxml_opc::rezip_parts(&files).unwrap()
}

#[test]
fn id_less_comments_and_notes_are_left_out_of_seed_and_export() {
    let base = package();
    let seeded = seed(Format::Docx, &base).unwrap();
    let doc = EditingDoc::new(4101);
    doc.load_state_v1(&seeded).unwrap();
    doc.insert_text(
        &EditCtx::local("", ""),
        Position::new("body", 0),
        "Edited ",
        FormatPolicy::Inherit,
    )
    .unwrap();
    let state = doc.encode_state_as_update_v1();
    let sha = sha256_hex(&base);
    let saved = export(
        &base,
        Checkpoint {
            format: Format::Docx,
            schema_version: 1,
            base_sha256: &sha,
            state: &state,
        },
        Determinism {
            seed: &"0".repeat(64),
            now: "2026-10-17T00:00:00.000Z",
        },
    )
    .unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let part = |path: &str| {
        parts
            .iter()
            .find(|(name, _)| name == path)
            .map(|(_, bytes)| String::from_utf8(bytes.clone()).unwrap())
            .unwrap_or_else(|| panic!("no {path}"))
    };
    let document = part("word/document.xml");
    for text in ["Edited ", "Kept", " lost"] {
        assert!(document.contains(&format!(">{text}</w:t>")), "{document}");
    }
    // An id-less `w:commentReference` saves as authored, as before the sync.
    for marker in ["commentRangeStart", "commentRangeEnd", "footnoteReference"] {
        assert_eq!(
            document.matches(&format!("<w:{marker} ")).count(),
            1,
            "{marker}: {document}"
        );
        assert!(
            !document.contains(&format!("<w:{marker}/>")),
            "{marker}: {document}"
        );
    }
    let comments = part("word/comments.xml");
    assert!(
        comments.contains("One") && !comments.contains("No id"),
        "{comments}"
    );
    let footnotes = part("word/footnotes.xml");
    assert!(
        footnotes.contains("Note one") && !footnotes.contains("No id note"),
        "{footnotes}"
    );
}
