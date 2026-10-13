//! Markup the editor draws but cannot edit is saved as authored after an edit
//! to its paragraph (Epo 2026-09-25, keep the original XML).

use docx_edit::{EditCtx, EditingDoc, FormatPolicy, Position};
use office_service::{Checkpoint, Determinism, Format, export, seed, sha256_hex};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const DGM: &str = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const OFFICE: &str = "application/vnd.openxmlformats-officedocument";

/// (path, xml, content type, relationship id and type from document.xml).
type Part = (&'static str, String, String, Option<(&'static str, String)>);

fn package(body: &str, parts: &[Part]) -> Vec<u8> {
    let overrides: String = parts
        .iter()
        .map(|(path, _, kind, _)| format!(r#"<Override PartName="/{path}" ContentType="{kind}"/>"#))
        .collect();
    let rels: String = parts
        .iter()
        .filter_map(|(path, _, _, rel)| {
            rel.as_ref().map(|(id, kind)| {
                let target = path.trim_start_matches("word/");
                format!(r#"<Relationship Id="{id}" Type="{kind}" Target="{target}"/>"#)
            })
        })
        .collect();
    let mut files = vec![
        (
            "[Content_Types].xml".to_owned(),
            format!(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{OFFICE}.wordprocessingml.document.main+xml"/>{overrides}</Types>"#
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
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>"#
            ),
        ),
        (
            "word/document.xml".to_owned(),
            format!(
                r#"<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>"#
            ),
        ),
    ];
    files.extend(
        parts
            .iter()
            .map(|(path, xml, _, _)| ((*path).to_owned(), xml.clone())),
    );
    let files: Vec<(String, Vec<u8>)> = files
        .into_iter()
        .map(|(path, xml)| (path, xml.into_bytes()))
        .collect();
    ooxml_opc::rezip_parts(&files).unwrap()
}

/// Seeds `base`, types `typed` at the start of the first body paragraph and
/// exports; the saved package's parts by path.
fn type_and_export(base: &[u8], typed: &str) -> Vec<(String, String)> {
    let seeded = seed(Format::Docx, base).unwrap();
    let doc = EditingDoc::new(4101);
    doc.load_state_v1(&seeded).unwrap();
    doc.insert_text(
        &EditCtx::local("", ""),
        Position::new("body", 0),
        typed,
        FormatPolicy::Inherit,
    )
    .unwrap();
    let state = doc.encode_state_as_update_v1();
    let sha = sha256_hex(base);
    let saved = export(
        base,
        Checkpoint {
            format: Format::Docx,
            schema_version: 1,
            base_sha256: &sha,
            state: &state,
        },
        Determinism {
            seed: &"0".repeat(64),
            now: "2026-10-13T00:00:00.000Z",
        },
    )
    .unwrap();
    ooxml_opc::unzip_parts(&saved)
        .unwrap()
        .into_iter()
        .map(|(path, bytes)| (path, String::from_utf8(bytes).unwrap()))
        .collect()
}

fn part<'a>(parts: &'a [(String, String)], path: &str) -> &'a str {
    parts
        .iter()
        .find(|(name, _)| name == path)
        .map(|(_, xml)| xml.as_str())
        .unwrap_or_else(|| panic!("no {path}"))
}

fn run(text: &str) -> String {
    format!(r#"<w:r><w:t xml:space="preserve">{text}</w:t></w:r>"#)
}

fn diagram_parts() -> Vec<Part> {
    let point = |id: &str, label: &str| {
        format!(
            r#"<dgm:pt modelId="{id}"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:p><a:r><a:t>{label}</a:t></a:r></a:p></dgm:t></dgm:pt>"#
        )
    };
    let diagram =
        |path: &'static str, xml: String, kind: &str, id: &'static str, rel: &str| -> Part {
            (
                path,
                xml,
                format!("{OFFICE}.drawingml.{kind}+xml"),
                Some((id, format!("{R}/{rel}"))),
            )
        };
    vec![
        diagram(
            "word/diagrams/data1.xml",
            format!(
                r#"<dgm:dataModel xmlns:dgm="{DGM}" xmlns:a="{A}"><dgm:ptLst><dgm:pt modelId="0" type="doc"><dgm:prSet/><dgm:spPr/></dgm:pt>{}{}</dgm:ptLst><dgm:cxnLst><dgm:cxn modelId="3" srcId="0" destId="1" srcOrd="0" destOrd="0"/><dgm:cxn modelId="4" srcId="0" destId="2" srcOrd="1" destOrd="0"/></dgm:cxnLst><dgm:bg/><dgm:whole/></dgm:dataModel>"#,
                point("1", "Plan"),
                point("2", "Test")
            ),
            "diagramData",
            "rIdDm",
            "diagramData",
        ),
        diagram(
            "word/diagrams/layout1.xml",
            format!(
                r#"<dgm:layoutDef xmlns:dgm="{DGM}" uniqueId="urn:microsoft.com/office/officeart/2005/8/layout/process1"><dgm:title val=""/><dgm:layoutNode name="Name0"/></dgm:layoutDef>"#
            ),
            "diagramLayout",
            "rIdLo",
            "diagramLayout",
        ),
        diagram(
            "word/diagrams/quickStyle1.xml",
            format!(
                r#"<dgm:styleDef xmlns:dgm="{DGM}" uniqueId="urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1"/>"#
            ),
            "diagramStyle",
            "rIdQs",
            "diagramQuickStyle",
        ),
        diagram(
            "word/diagrams/colors1.xml",
            format!(
                r#"<dgm:colorsDef xmlns:dgm="{DGM}" uniqueId="urn:microsoft.com/office/officeart/2005/8/colors/accent1_2"/>"#
            ),
            "diagramColors",
            "rIdCs",
            "diagramColors",
        ),
    ]
}

/// A SmartArt frame, its four diagram parts and, with `drawn`, the drawing
/// Word caches for it: an edit to its paragraph saves the frame as authored
/// and keeps every part.
#[test]
fn a_smart_art_frame_saves_as_authored_after_an_edit_to_its_paragraph() {
    let frame = format!(
        r#"<w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="3657600" cy="914400"/><wp:docPr id="1" name="Diagram 1"/><a:graphic xmlns:a="{A}"><a:graphicData uri="{DGM}"><dgm:relIds xmlns:dgm="{DGM}" r:dm="rIdDm" r:lo="rIdLo" r:qs="rIdQs" r:cs="rIdCs"/></a:graphicData></a:graphic></wp:inline></w:drawing>"#
    );
    let cached: Part = (
        "word/diagrams/drawing1.xml",
        format!(
            r#"<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="{A}"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/><dsp:sp modelId="{{1}}"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1600000" cy="900000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></dsp:spPr><dsp:txBody><a:bodyPr/><a:p><a:r><a:t>Plan</a:t></a:r></a:p></dsp:txBody></dsp:sp></dsp:spTree></dsp:drawing>"#
        ),
        "application/vnd.ms-office.drawingml.diagramDrawing+xml".to_owned(),
        Some((
            "rIdDr",
            "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing".to_owned(),
        )),
    );
    for drawn in [false, true] {
        let mut parts = diagram_parts();
        if drawn {
            parts.push(cached.clone());
        }
        let body = format!(
            r#"<w:p w14:paraId="10000001">{}<w:r>{frame}</w:r>{}</w:p>"#,
            run("Steps: "),
            run(" done.")
        );
        let saved = type_and_export(&package(&body, &parts), "Two ");
        let document = part(&saved, "word/document.xml");
        assert!(document.contains("Two Steps: "), "{document}");
        assert!(document.contains(&frame), "drawn {drawn}: {document}");
        let rels = part(&saved, "word/_rels/document.xml.rels");
        for (path, xml, _, rel) in &parts {
            assert_eq!(part(&saved, path), xml, "{path}");
            assert!(rels.contains(&format!(r#"Id="{}""#, rel.as_ref().unwrap().0)));
        }
    }
}

/// A phonetic guide keeps its guide and base text, with their runs'
/// formatting, after an edit to its paragraph.
#[test]
fn a_phonetic_guide_saves_as_authored_after_an_edit_to_its_paragraph() {
    let ruby = r#"<w:ruby><w:rubyPr><w:rubyAlign w:val="distributeSpace"/><w:hps w:val="10"/><w:hpsRaise w:val="18"/><w:hpsBaseText w:val="21"/><w:lid w:val="ja-JP"/></w:rubyPr><w:rt><w:r><w:rPr><w:sz w:val="10"/></w:rPr><w:t>かんじ</w:t></w:r></w:rt><w:rubyBase><w:r><w:rPr><w:b/></w:rPr><w:t>漢字</w:t></w:r></w:rubyBase></w:ruby>"#;
    let body = format!(
        r#"<w:p w14:paraId="10000001">{}<w:r>{ruby}</w:r>{}</w:p>"#,
        run("読み: "),
        run(" です")
    );
    let saved = type_and_export(&package(&body, &[]), "Lesson ");
    let document = part(&saved, "word/document.xml");
    assert!(document.contains("Lesson 読み: "), "{document}");
    assert!(
        document.contains(&format!("<w:r>{ruby}</w:r>")),
        "{document}"
    );
}
