//! A picture Word stores only as a link (`a:blip r:link` to an external
//! relationship) has no bytes: effects list it without an asset instead of
//! failing every save of the file, saves it back as a link, and never fetches it.

use std::net::TcpListener;

use docx_edit::{EditCtx, EditingDoc, FormatPolicy, Position};
use office_service::{
    Checkpoint, Determinism, EffectKind, Format, baseline, export, inspect, save_effects, seed,
    sha256_hex,
};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const WP: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const EXAMPLE: &str = "https://example.com/linked.png";

fn linked_picture_docx(target: &str) -> Vec<u8> {
    let picture = format!(
        r#"<w:r><w:drawing><wp:inline xmlns:wp="{WP}" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Picture 1"/><a:graphic xmlns:a="{A}"><a:graphicData uri="{PIC}"><pic:pic xmlns:pic="{PIC}"><pic:nvPicPr><pic:cNvPr id="1" name="linked.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:link="rIdLink1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#
    );
    let files = [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_owned(),
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
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLink1" Type="{R}/image" Target="{target}" TargetMode="External"/></Relationships>"#
            ),
        ),
        (
            "word/document.xml",
            format!(
                r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:p><w:r><w:t>Before</w:t></w:r></w:p><w:p>{picture}</w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>"#
            ),
        ),
    ];
    let files: Vec<(String, Vec<u8>)> = files
        .into_iter()
        .map(|(path, xml)| (path.to_owned(), xml.into_bytes()))
        .collect();
    ooxml_opc::rezip_parts(&files).unwrap()
}

fn checkpoint<'a>(sha: &'a str, state: &'a [u8]) -> Checkpoint<'a> {
    Checkpoint {
        format: Format::Docx,
        schema_version: 1,
        base_sha256: sha,
        state,
    }
}

#[test]
fn a_link_only_picture_has_an_image_entry_without_an_asset() {
    let base = linked_picture_docx(EXAMPLE);
    let sha = sha256_hex(&base);
    let seeded = seed(Format::Docx, &base).unwrap();
    let entries = baseline(&base, checkpoint(&sha, &seeded)).unwrap();
    let image = entries
        .iter()
        .find(|entry| entry.kind == EffectKind::Image)
        .expect("the linked picture is listed");
    assert!(image.asset_ref.is_none(), "nothing to resolve: {image:?}");
}

#[test]
fn a_file_with_a_link_only_picture_reports_text_effects() {
    let base = linked_picture_docx(EXAMPLE);
    let sha = sha256_hex(&base);
    let seeded = seed(Format::Docx, &base).unwrap();
    let doc = EditingDoc::new(4101);
    doc.load_state_v1(&seeded).unwrap();
    doc.insert_text(
        &EditCtx::local("", ""),
        Position::new("body", 0),
        "Typed ",
        FormatPolicy::Inherit,
    )
    .unwrap();
    let edited = doc.encode_state_as_update_v1();
    let indexed = baseline(&base, checkpoint(&sha, &seeded)).unwrap();
    let effects = save_effects(&base, &indexed, checkpoint(&sha, &edited)).unwrap();
    assert!(
        effects.iter().any(|effect| effect.kind == EffectKind::Text
            && effect
                .after
                .as_deref()
                .is_some_and(|after| after.contains("Typed"))),
        "{effects:?}"
    );
    assert!(
        effects
            .iter()
            .all(|effect| effect.kind != EffectKind::Image),
        "{effects:?}"
    );

    let saved = export(
        &base,
        checkpoint(&sha, &edited),
        Determinism {
            seed: &"0".repeat(64),
            now: "2026-10-17T00:00:00.000Z",
        },
    )
    .unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let document = parts
        .iter()
        .find(|(path, _)| path == "word/document.xml")
        .map(|(_, bytes)| String::from_utf8_lossy(bytes).into_owned())
        .unwrap();
    assert!(document.contains(r#"r:link="rIdLink1""#), "{document}");
}

#[test]
fn a_link_only_picture_is_never_fetched() {
    let listener = TcpListener::bind("127.0.0.1:0").unwrap();
    listener.set_nonblocking(true).unwrap();
    let target = format!("http://{}/linked.png", listener.local_addr().unwrap());
    let base = linked_picture_docx(&target);
    let sha = sha256_hex(&base);
    let seeded = seed(Format::Docx, &base).unwrap();
    let indexed = baseline(&base, checkpoint(&sha, &seeded)).unwrap();
    let doc = EditingDoc::new(4101);
    doc.load_state_v1(&seeded).unwrap();
    doc.insert_text(
        &EditCtx::local("", ""),
        Position::new("body", 0),
        "Typed ",
        FormatPolicy::Inherit,
    )
    .unwrap();
    let edited = doc.encode_state_as_update_v1();
    save_effects(&base, &indexed, checkpoint(&sha, &edited)).unwrap();
    inspect(&base, checkpoint(&sha, &edited)).unwrap();
    export(
        &base,
        checkpoint(&sha, &edited),
        Determinism {
            seed: &"0".repeat(64),
            now: "2026-10-17T00:00:00.000Z",
        },
    )
    .unwrap();
    let accepted = listener.accept();
    assert!(
        matches!(&accepted, Err(error) if error.kind() == std::io::ErrorKind::WouldBlock),
        "the external target was contacted: {accepted:?}"
    );
}
