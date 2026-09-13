//! Write-back fidelity against a deck exercising markup the model does not
//! carry: connectors, inherited placeholder rects, hyperlinks, fields,
//! theme colours, notes slides, custom shows, and hostile inputs.

use std::collections::BTreeMap;

use pptx_edit::{DeckSession, EditCtx, EditError, TextStyle};

const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
<Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>
<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>"#;

const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"#;

const PRESENTATION_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
</Relationships>"#;

const SLIDE1: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
<p:spPr/>
<p:txBody><a:bodyPr/><a:lstStyle/>
<a:p><a:r><a:rPr b="1"/><a:t>Hello </a:t></a:r><a:r><a:rPr strike="sngStrike"><a:hlinkClick r:id="rId2"/></a:rPr><a:t>link</a:t></a:r></a:p>
<a:p><a:fld id="{D038279B-FC19-497E-A7D1-5ADD9CAF016F}" type="slidenum"><a:rPr lang="en-US"/><a:t>1</a:t></a:fld><a:r><a:rPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:rPr><a:t>Accent</a:t></a:r></a:p>
</p:txBody></p:sp>
<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="5" name="Connector"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>
<p:sp><p:nvSpPr><p:cNvPr id="3" name="Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="1000" y="2000"/><a:ext cx="3000" cy="4000"/></a:xfrm>
<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 20000"/><a:gd name="adj2" fmla="*/ missing 2 3"/></a:avLst></a:prstGeom></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr><a:noFill/></a:rPr><a:t>Box</a:t></a:r></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="4" name="Halfway"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm rot="1200000"><a:off x="123456" y="654321"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Halfway</a:t></a:r></a:p></p:txBody></p:sp>
</p:spTree></p:cSld></p:sld>"#;

const SLIDE1_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>
</Relationships>"#;

const SLIDE2: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr/>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="Second"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="30" cy="40"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Second</a:t></a:r></a:p></p:txBody></p:sp>
<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="3" name="Trailing Connector"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
<p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>
</p:spTree></p:cSld></p:sld>"#;

const SLIDE2_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>
</Relationships>"#;

const NOTES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:notes>"#;

const NOTES_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide2.xml"/>
</Relationships>"#;

const LAYOUT: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="title">
<p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr/>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm flipH="1" rot="2700000"><a:off x="762000" y="457200"/><a:ext cx="10668000" cy="1143000"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="3" name="Body Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="1000000" y="2000000"/><a:ext cx="8000000" cy="3000000"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp>
</p:spTree></p:cSld></p:sldLayout>"#;

const LAYOUT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"#;

const MASTER: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>
<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>"#;

const MASTER_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>"#;

const THEME: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Fixture">
<a:themeElements><a:clrScheme name="Fixture">
<a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>
<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme></a:themeElements></a:theme>"#;

fn presentation_xml(first_slide_id: u32) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
<p:sldIdLst><p:sldId id="{first_slide_id}" r:id="rId2"/><p:sldId id="257" r:id="rId3"/></p:sldIdLst>
<p:custShowLst><p:custShow name="Short" id="0"><p:sldLst><p:sld r:id="rId2"/><p:sld r:id="rId3"/></p:sldLst></p:custShow></p:custShowLst>
<p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>"#
    )
}

fn fixture(first_slide_id: u32) -> Vec<u8> {
    let parts: Vec<(String, Vec<u8>)> = [
        ("[Content_Types].xml", CONTENT_TYPES.to_owned()),
        ("_rels/.rels", ROOT_RELS.to_owned()),
        ("ppt/presentation.xml", presentation_xml(first_slide_id)),
        (
            "ppt/_rels/presentation.xml.rels",
            PRESENTATION_RELS.to_owned(),
        ),
        ("ppt/slides/slide1.xml", SLIDE1.to_owned()),
        ("ppt/slides/_rels/slide1.xml.rels", SLIDE1_RELS.to_owned()),
        ("ppt/slides/slide2.xml", SLIDE2.to_owned()),
        ("ppt/slides/_rels/slide2.xml.rels", SLIDE2_RELS.to_owned()),
        ("ppt/notesSlides/notesSlide1.xml", NOTES.to_owned()),
        (
            "ppt/notesSlides/_rels/notesSlide1.xml.rels",
            NOTES_RELS.to_owned(),
        ),
        ("ppt/slideLayouts/slideLayout1.xml", LAYOUT.to_owned()),
        (
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
            LAYOUT_RELS.to_owned(),
        ),
        ("ppt/slideMasters/slideMaster1.xml", MASTER.to_owned()),
        (
            "ppt/slideMasters/_rels/slideMaster1.xml.rels",
            MASTER_RELS.to_owned(),
        ),
        ("ppt/theme/theme1.xml", THEME.to_owned()),
    ]
    .into_iter()
    .map(|(path, body)| (path.to_owned(), body.into_bytes()))
    .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn open() -> DeckSession {
    DeckSession::open(&fixture(256), 11).unwrap()
}

fn context() -> EditCtx {
    EditCtx::local("fidelity")
}

fn parts(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    ooxml_opc::unzip_parts(bytes).unwrap().into_iter().collect()
}

fn part_text(parts: &BTreeMap<String, Vec<u8>>, path: &str) -> String {
    String::from_utf8(
        parts
            .get(path)
            .unwrap_or_else(|| panic!("missing {path}"))
            .clone(),
    )
    .unwrap()
}

/// Every internal relationship in every .rels part must resolve to a part.
fn assert_relationships_resolve(parts: &BTreeMap<String, Vec<u8>>) {
    for (path, bytes) in parts {
        let Some((directory, name)) = path.rsplit_once("/_rels/") else {
            continue;
        };
        let source_directory = directory;
        let text = String::from_utf8(bytes.clone()).unwrap();
        for chunk in text.split("<Relationship ").skip(1) {
            if chunk.contains("TargetMode=\"External\"") {
                continue;
            }
            let target = chunk
                .split("Target=\"")
                .nth(1)
                .and_then(|rest| rest.split('"').next())
                .unwrap_or_else(|| panic!("{path}: relationship without a target"));
            let mut segments: Vec<&str> = source_directory.split('/').collect();
            for segment in target.split('/') {
                match segment {
                    "." | "" => {}
                    ".." => {
                        segments.pop();
                    }
                    segment => segments.push(segment),
                }
            }
            let resolved = segments.join("/");
            assert!(
                parts.contains_key(&resolved),
                "{path} ({name}) references missing part {resolved}"
            );
        }
    }
}

fn box_shape(session: &DeckSession) -> (String, String) {
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let shape = slide
        .shapes
        .iter()
        .find(|shape| shape.name == "Box")
        .unwrap();
    (slide.id.clone(), shape.id.clone())
}

#[test]
fn a_connector_between_shapes_keeps_its_slot() {
    let session = open();
    let (slide_id, shape_id) = box_shape(&session);
    session
        .move_shape(&context(), &slide_id, &shape_id, 5000, 6000)
        .unwrap();

    let saved = parts(&session.save().unwrap());
    let slide = part_text(&saved, "ppt/slides/slide1.xml");
    let title = slide.find("Title").unwrap();
    let connector = slide.find("<p:cxnSp>").unwrap();
    let moved = slide.find("\"Box\"").unwrap();
    assert!(title < connector && connector < moved);
    assert!(slide.contains(r#"<a:off x="5000" y="6000"/>"#));
    assert_relationships_resolve(&saved);
}

#[test]
fn a_moved_inherited_placeholder_keeps_its_rect_rotation_and_flip() {
    let session = open();
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let title = slide
        .shapes
        .iter()
        .find(|shape| shape.name == "Title")
        .unwrap();
    session
        .move_shape(&context(), &slide.id, &title.id, 900_000, 1_000_000)
        .unwrap();

    let saved = parts(&session.save().unwrap());
    let slide = part_text(&saved, "ppt/slides/slide1.xml");
    assert!(slide.contains(r#"<a:off x="900000" y="1000000"/>"#));
    assert!(slide.contains(r#"<a:ext cx="10668000" cy="1143000"/>"#));
    assert!(slide.contains(r#"rot="2700000""#));
    assert!(slide.contains(r#"flipH="1""#));

    let reopened = DeckSession::open(&session.save().unwrap(), 12).unwrap();
    let snapshot = reopened.snapshot().unwrap();
    let title = snapshot.slides[0]
        .shapes
        .iter()
        .find(|shape| shape.name == "Title")
        .unwrap();
    assert_eq!((title.x, title.y), (900_000, 1_000_000));
    assert_eq!((title.width, title.height), (10_668_000, 1_143_000));
    assert_eq!(title.rotation_deg, 45.0);
    assert!(title.flip_h);
}

#[test]
fn a_resized_inherited_placeholder_keeps_its_inherited_offset() {
    let session = open();
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let title = slide
        .shapes
        .iter()
        .find(|shape| shape.name == "Title")
        .unwrap();
    session
        .resize_shape(&context(), &slide.id, &title.id, 5_000_000, 900_000)
        .unwrap();

    let saved = parts(&session.save().unwrap());
    let slide = part_text(&saved, "ppt/slides/slide1.xml");
    assert!(slide.contains(r#"<a:off x="762000" y="457200"/>"#));
    assert!(slide.contains(r#"<a:ext cx="5000000" cy="900000"/>"#));
}

#[test]
fn a_shape_added_to_an_emptied_slide_lands_after_the_group_properties() {
    let session = open();
    let snapshot = session.snapshot().unwrap();
    let slide = snapshot.slides[1].clone();
    session
        .remove_shape(&context(), &slide.id, &slide.shapes[0].id)
        .unwrap();
    session
        .add_text_box(
            &context(),
            &slide.id,
            &pptx_edit::ShapeDraft {
                name: "Fresh".to_owned(),
                rect: pptx_edit::ShapeRect {
                    x: 100,
                    y: 200,
                    width: 3_000,
                    height: 4_000,
                },
                text: "on an empty slide".to_owned(),
                style: TextStyle::default(),
            },
        )
        .unwrap();

    let saved = parts(&session.save().unwrap());
    let slide = part_text(&saved, "ppt/slides/slide2.xml");
    let non_visual = slide.find("<p:nvGrpSpPr>").unwrap();
    let group_properties = slide.find("<p:grpSpPr").unwrap();
    let added = slide.find("<p:sp>").unwrap();
    assert!(non_visual < group_properties && group_properties < added);
}

#[test]
fn junk_style_values_are_rejected_at_the_edit() {
    let session = open();
    let snapshot = session.snapshot().unwrap();
    let story_id = snapshot.slides[0].shapes[0].text_stories[0].id.clone();
    let styles = [
        TextStyle {
            color: Some("rebeccapurple".to_owned()),
            ..TextStyle::default()
        },
        TextStyle {
            underline: Some("wobbly".to_owned()),
            ..TextStyle::default()
        },
        TextStyle {
            font_size_pt: Some(0.001),
            ..TextStyle::default()
        },
    ];
    for style in styles {
        let error = session
            .insert_text(&context(), &story_id, 0, "X", &style)
            .unwrap_err();
        assert!(matches!(error, EditError::InvalidText(_)));
    }
    assert_eq!(parts(&session.save().unwrap()), parts(&fixture(256)));
}

#[test]
fn an_edit_inside_a_linked_run_keeps_the_link() {
    let session = open();
    let snapshot = session.snapshot().unwrap();
    let story_id = snapshot.slides[0]
        .shapes
        .iter()
        .find(|shape| shape.name == "Title")
        .unwrap()
        .text_stories[0]
        .id
        .clone();
    session
        .insert_text(&context(), &story_id, 8, "X", &TextStyle::default())
        .unwrap();

    let saved = session.save().unwrap();
    let slide = part_text(&parts(&saved), "ppt/slides/slide1.xml");
    assert!(slide.contains("<a:t>liXnk</a:t>"));
    assert!(slide.contains("hlinkClick"));
    assert!(slide.contains(r#"strike="sngStrike""#));

    let reopened = DeckSession::open(&saved, 12).unwrap();
    assert!(
        reopened
            .story(&story_id)
            .unwrap()
            .plain_text()
            .starts_with("Hello liXnk")
    );
}

#[test]
fn an_edit_before_a_hyperlink_keeps_links_fields_and_theme_colours() {
    let session = open();
    let snapshot = session.snapshot().unwrap();
    let story_id = snapshot.slides[0]
        .shapes
        .iter()
        .find(|shape| shape.name == "Title")
        .unwrap()
        .text_stories[0]
        .id
        .clone();
    session
        .insert_text(&context(), &story_id, 0, "Hi ", &TextStyle::default())
        .unwrap();

    let saved = session.save().unwrap();
    let slide = part_text(&parts(&saved), "ppt/slides/slide1.xml");
    assert!(slide.contains("hlinkClick"));
    assert!(slide.contains(r#"strike="sngStrike""#));
    assert!(slide.contains("<a:fld "));
    assert!(slide.contains(r#"schemeClr val="accent1""#));

    let reopened = DeckSession::open(&saved, 12).unwrap();
    let text = reopened.story(&story_id).unwrap().plain_text();
    assert_eq!(text, "Hi Hello link\n1Accent");
}

#[test]
fn deleting_a_slide_prunes_its_notes_and_custom_show_entry() {
    let session = open();
    let second = session.snapshot().unwrap().slides[1].id.clone();
    session.delete_slide(&context(), &second).unwrap();

    let saved = parts(&session.save().unwrap());
    assert!(!saved.contains_key("ppt/slides/slide2.xml"));
    assert!(!saved.contains_key("ppt/notesSlides/notesSlide1.xml"));
    assert!(!saved.contains_key("ppt/notesSlides/_rels/notesSlide1.xml.rels"));
    let content_types = part_text(&saved, "[Content_Types].xml");
    assert!(!content_types.contains("notesSlide1.xml"));
    assert!(!content_types.contains("slide2.xml"));
    let presentation = part_text(&saved, "ppt/presentation.xml");
    assert_eq!(presentation.matches("rId3").count(), 0);
    assert!(presentation.contains("custShowLst"));
    assert_relationships_resolve(&saved);
}

#[test]
fn illegal_control_characters_are_rejected_at_the_edit() {
    let session = open();
    let snapshot = session.snapshot().unwrap();
    let story_id = snapshot.slides[0].shapes[0].text_stories[0].id.clone();
    let error = session
        .insert_text(&context(), &story_id, 0, "a\u{1}b", &TextStyle::default())
        .unwrap_err();
    assert!(matches!(error, EditError::InvalidText(_)));
    assert_eq!(parts(&session.save().unwrap()), parts(&fixture(256)));
}

#[test]
fn a_carriage_return_survives_the_round_trip() {
    let session = open();
    let snapshot = session.snapshot().unwrap();
    let story_id = snapshot.slides[0].shapes[0].text_stories[0].id.clone();
    session
        .insert_text(&context(), &story_id, 0, "a\rb ", &TextStyle::default())
        .unwrap();

    let saved = session.save().unwrap();
    let reopened = DeckSession::open(&saved, 12).unwrap();
    assert!(
        reopened
            .story(&story_id)
            .unwrap()
            .plain_text()
            .starts_with("a\rb ")
    );
}

#[test]
fn unevaluated_adjustment_guides_survive_a_shape_adjust() {
    let session = open();
    let (slide_id, shape_id) = box_shape(&session);
    let mut adjustments = BTreeMap::new();
    adjustments.insert("adj".to_owned(), 0.25);
    session
        .set_shape_adjust(&context(), &slide_id, &shape_id, &adjustments)
        .unwrap();

    let saved = parts(&session.save().unwrap());
    let slide = part_text(&saved, "ppt/slides/slide1.xml");
    assert!(slide.contains(r#"fmla="val 25000" name="adj""#));
    assert!(slide.contains(r#"fmla="*/ missing 2 3" name="adj2""#));
}

#[test]
fn an_exhausted_slide_id_space_errors_instead_of_panicking() {
    let session = DeckSession::open(&fixture(4_294_967_295), 11).unwrap();
    session.insert_slide(&context(), 2, None).unwrap();
    let error = session.save().unwrap_err();
    assert!(matches!(error, EditError::Write(message) if message.contains("slide id")));
}

#[test]
fn a_colour_write_replaces_an_existing_no_fill() {
    let session = open();
    let snapshot = session.snapshot().unwrap();
    let story_id = snapshot.slides[0]
        .shapes
        .iter()
        .find(|shape| shape.name == "Box")
        .unwrap()
        .text_stories[0]
        .id
        .clone();
    session
        .format_text(
            &context(),
            &story_id,
            0,
            3,
            &pptx_edit::TextStylePatch {
                color: Some("#FF0000".to_owned()),
                ..pptx_edit::TextStylePatch::default()
            },
        )
        .unwrap();

    let saved = parts(&session.save().unwrap());
    let slide = part_text(&saved, "ppt/slides/slide1.xml");
    assert!(slide.contains(r#"srgbClr val="FF0000""#));
    assert!(!slide.contains("<a:noFill/></a:rPr>"));
}

#[test]
fn a_resize_keeps_a_partial_transforms_own_offset_and_rotation() {
    let session = open();
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let halfway = slide
        .shapes
        .iter()
        .find(|shape| shape.name == "Halfway")
        .unwrap();
    session
        .resize_shape(&context(), &slide.id, &halfway.id, 5_000_000, 900_000)
        .unwrap();

    let saved = parts(&session.save().unwrap());
    let slide = part_text(&saved, "ppt/slides/slide1.xml");
    assert!(slide.contains(r#"<a:off x="123456" y="654321"/>"#));
    assert!(slide.contains(r#"rot="1200000""#));
    assert!(slide.contains(r#"<a:ext cx="5000000" cy="900000"/>"#));
    assert!(!slide.contains(r#"<a:off x="1000000" y="2000000"/>"#));
}

#[test]
fn a_shape_added_after_a_trailing_connector_lands_on_top() {
    let session = open();
    let slide_id = session.snapshot().unwrap().slides[1].id.clone();
    session
        .add_text_box(
            &context(),
            &slide_id,
            &pptx_edit::ShapeDraft {
                name: "OnTop".to_owned(),
                rect: pptx_edit::ShapeRect {
                    x: 1,
                    y: 2,
                    width: 300,
                    height: 400,
                },
                text: "above the connector".to_owned(),
                style: TextStyle::default(),
            },
        )
        .unwrap();

    let saved = parts(&session.save().unwrap());
    let slide = part_text(&saved, "ppt/slides/slide2.xml");
    let second = slide.find("\"Second\"").unwrap();
    let connector = slide.find("<p:cxnSp>").unwrap();
    let added = slide.find("\"OnTop\"").unwrap();
    assert!(second < connector && connector < added);
}

#[test]
fn an_unknown_layout_is_rejected_at_insert_slide() {
    let session = open();
    let error = session
        .insert_slide(&context(), 2, Some("ppt/slideLayouts/nope.xml"))
        .unwrap_err();
    assert!(matches!(error, EditError::InvalidState(_)));
    assert_eq!(parts(&session.save().unwrap()), parts(&fixture(256)));
}

#[test]
fn a_session_opened_from_an_update_alone_refuses_to_save() {
    let source = fixture(256);
    let seeded = DeckSession::open(&source, 11).unwrap();
    let update = seeded.encode_state_as_update_v1();

    let bare = DeckSession::open_from_update(&update, 12).unwrap();
    assert!(matches!(bare.save(), Err(EditError::Write(_))));

    let attached = DeckSession::open_from_update_with_source(&update, &source, 13).unwrap();
    let saved = attached.save().unwrap();
    assert_eq!(parts(&saved), parts(&source));

    let mismatched = DeckSession::open_from_update_with_source(&update, b"garbage", 14);
    assert!(mismatched.is_err());
}

#[test]
fn rebase_keeps_later_edits_and_maps_only_the_remaining_changes() {
    let source = fixture(256);
    let session = DeckSession::open(&source, 11).unwrap();
    let initial = session.snapshot().unwrap();
    let first = &initial.slides[0];
    let second = &initial.slides[1];
    session.move_slide(&context(), &second.id, 0).unwrap();
    session
        .remove_shape(&context(), &first.id, &first.shapes[1].id)
        .unwrap();
    let added = session
        .add_text_box(
            &context(),
            &first.id,
            &pptx_edit::ShapeDraft {
                name: "Captured new shape".into(),
                rect: pptx_edit::ShapeRect {
                    x: 100,
                    y: 200,
                    width: 3000,
                    height: 4000,
                },
                text: "Captured".into(),
                style: TextStyle::default(),
            },
        )
        .unwrap();
    let captured = session.encode_state_as_update_v1();
    let published = session.save().unwrap();
    let zero =
        DeckSession::rebase_checkpoint(&source, &captured, &captured, &published, 21).unwrap();
    let zero_current =
        DeckSession::open_from_update_with_source(&zero.state, &published, 22).unwrap();
    let zero_indexed = DeckSession::open_from_update(&zero.indexed_state, 23).unwrap();
    assert_eq!(
        zero_current.snapshot().unwrap(),
        zero_indexed.snapshot().unwrap()
    );

    let snapshot = session.snapshot().unwrap();
    let added_shape = snapshot
        .slides
        .iter()
        .flat_map(|slide| &slide.shapes)
        .find(|shape| shape.id == added.shape_id)
        .unwrap();
    session
        .insert_text(
            &context(),
            &added_shape.text_stories[0].id,
            8,
            " later",
            &TextStyle::default(),
        )
        .unwrap();
    session
        .resize_shape(&context(), &first.id, &added.shape_id, 5000, 6000)
        .unwrap();
    session.move_slide(&context(), &first.id, 0).unwrap();
    session
        .remove_shape(&context(), &second.id, &second.shapes[0].id)
        .unwrap();
    let latest = session.encode_state_as_update_v1();
    let expected = parts(&session.save().unwrap());
    let rebased =
        DeckSession::rebase_checkpoint(&source, &captured, &latest, &published, 21).unwrap();
    let current =
        DeckSession::open_from_update_with_source(&rebased.state, &published, 22).unwrap();
    let baseline = DeckSession::open_from_update(&rebased.indexed_state, 23).unwrap();
    assert_eq!(parts(&current.save().unwrap()), expected);
    let current_snapshot = current.snapshot().unwrap();
    let indexed_snapshot = baseline.snapshot().unwrap();
    let current_shape = current_snapshot
        .slides
        .iter()
        .flat_map(|s| &s.shapes)
        .find(|s| s.name == "Captured new shape")
        .unwrap();
    let indexed_shape = indexed_snapshot
        .slides
        .iter()
        .flat_map(|s| &s.shapes)
        .find(|s| s.name == "Captured new shape")
        .unwrap();
    assert_eq!(current_shape.id, indexed_shape.id);
    assert_eq!(current_shape.source_id, indexed_shape.source_id);
    assert_eq!(
        current_shape.text_stories[0].paragraphs[0].id,
        indexed_shape.text_stories[0].paragraphs[0].id
    );
    assert_eq!(indexed_shape.text_stories[0].plain_text(), "Captured");
    assert_eq!(current_shape.text_stories[0].plain_text(), "Captured later");
    assert_eq!(indexed_shape.width, 3000);
    assert_eq!(current_shape.width, 5000);
    let removed = indexed_snapshot
        .slides
        .iter()
        .flat_map(|s| &s.shapes)
        .find(|s| s.name == "Second")
        .unwrap();
    assert!(removed.id.starts_with("rebase:removed:"));
    assert_eq!(indexed_snapshot.slides[1].id, current_snapshot.slides[0].id);
    current
        .insert_text(
            &context(),
            &current_shape.text_stories[0].id,
            0,
            "Again ",
            &TextStyle::default(),
        )
        .unwrap();
    let reopened = DeckSession::open(&current.save().unwrap(), 24).unwrap();
    assert!(
        reopened
            .snapshot()
            .unwrap()
            .slides
            .iter()
            .flat_map(|s| &s.shapes)
            .any(|s| s
                .text_stories
                .iter()
                .any(|story| story.plain_text() == "Again Captured later"))
    );
    assert!(DeckSession::rebase_checkpoint(&source, &captured, &latest, &source, 21).is_err());
    let next_capture = current.encode_state_as_update_v1();
    let next_source = current.save().unwrap();
    let next =
        DeckSession::rebase_checkpoint(&published, &next_capture, &next_capture, &next_source, 31)
            .unwrap();
    let next_current =
        DeckSession::open_from_update_with_source(&next.state, &next_source, 32).unwrap();
    assert_eq!(parts(&next_current.save().unwrap()), parts(&next_source));
    let next_indexed = DeckSession::open_from_update(&next.indexed_state, 33).unwrap();
    assert_eq!(
        next_current.snapshot().unwrap(),
        next_indexed.snapshot().unwrap()
    );
}

fn media_rebase_fixture() -> (Vec<u8>, Vec<u8>) {
    let mut entries = parts(&fixture(256));
    let picture = r#"<p:pic><p:nvPicPr><p:cNvPr id="7" name="Unique picture"><a:extLst><a:ext uri="opaque-marker"/></a:extLst></p:cNvPr><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId3"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="1000" cy="2000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>"#;
    entries.insert(
        "ppt/slides/slide2.xml".into(),
        SLIDE2
            .replace("</p:spTree>", &format!("{picture}</p:spTree>"))
            .into_bytes(),
    );
    entries.insert("ppt/slides/_rels/slide2.xml.rels".into(), SLIDE2_RELS.replace("</Relationships>", r#"<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/unique.png"/></Relationships>"#).into_bytes());
    entries.insert(
        "[Content_Types].xml".into(),
        CONTENT_TYPES
            .replace(
                "</Types>",
                r#"<Default Extension="png" ContentType="image/png"/></Types>"#,
            )
            .into_bytes(),
    );
    let media = (0..4096).map(|i| (i % 251) as u8).collect::<Vec<_>>();
    entries.insert("ppt/media/unique.png".into(), media.clone());
    (
        ooxml_opc::rezip_parts(&entries.into_iter().collect::<Vec<_>>()).unwrap(),
        media,
    )
}

#[test]
fn rebase_restores_saved_slide_and_picture_with_binary_media_and_opaque_parts() {
    use yrs::{Any, Map, Out, ReadTxn, Transact};
    for remove_slide in [true, false] {
        let (source, media) = media_rebase_fixture();
        let session = DeckSession::open(&source, 11).unwrap();
        let snapshot = session.snapshot().unwrap();
        let slide = &snapshot.slides[1];
        let picture = slide
            .shapes
            .iter()
            .find(|s| s.name == "Unique picture")
            .unwrap();
        if remove_slide {
            session.delete_slide(&context(), &slide.id).unwrap();
        } else {
            session
                .remove_shape(&context(), &slide.id, &picture.id)
                .unwrap();
        }
        session.add_undo_barrier();
        let captured = session.encode_state_as_update_v1();
        let published = session.save().unwrap();
        if remove_slide {
            assert!(!parts(&published).contains_key("ppt/media/unique.png"));
        }
        assert!(session.undo());
        session
            .resize_shape(&context(), &slide.id, &picture.id, 7000, 8000)
            .unwrap();
        let latest = session.encode_state_as_update_v1();
        let expected = parts(&session.save().unwrap());
        let rebased =
            DeckSession::rebase_checkpoint(&source, &captured, &latest, &published, 21).unwrap();
        let current =
            DeckSession::open_from_update_with_source(&rebased.state, &published, 22).unwrap();
        let actual = parts(&current.save().unwrap());
        assert_eq!(actual, expected);
        assert_eq!(actual.get("ppt/media/unique.png"), Some(&media));
        assert!(part_text(&actual, "ppt/slides/slide2.xml").contains("opaque-marker"));
        assert_relationships_resolve(&actual);
        let txn = current.yrs_doc().transact();
        let meta = txn.get_map("pptx:meta").unwrap();
        let Some(Out::Any(Any::Array(overlay))) = meta.get(&txn, "sourceOverlay") else {
            panic!("missing overlay")
        };
        let published_parts = parts(&published);
        for value in overlay.iter() {
            let Any::Array(fields) = value else {
                panic!("invalid overlay")
            };
            let Any::String(path) = &fields[0] else {
                panic!("invalid path")
            };
            assert_ne!(
                published_parts.get(path.as_ref()),
                actual.get(path.as_ref())
            );
            if path.as_ref() == "ppt/media/unique.png" {
                assert!(matches!(&fields[1], Any::Null));
            }
        }
        let indexed = DeckSession::open_from_update(&rebased.indexed_state, 23).unwrap();
        assert!(
            !indexed
                .snapshot()
                .unwrap()
                .slides
                .iter()
                .flat_map(|s| &s.shapes)
                .any(|s| s.name == "Unique picture")
        );
        assert!(
            current
                .snapshot()
                .unwrap()
                .slides
                .iter()
                .flat_map(|s| &s.shapes)
                .any(|s| s.name == "Unique picture" && s.width == 7000)
        );
    }
}

#[test]
fn rebase_maps_paragraph_restoration_without_replaying_captured_changes() {
    let source = fixture(256);
    let session = DeckSession::open(&source, 11).unwrap();
    let snapshot = session.snapshot().unwrap();
    let title = snapshot.slides[0]
        .shapes
        .iter()
        .find(|s| s.name == "Title")
        .unwrap();
    let story = &title.text_stories[0];
    let first_length = story.paragraphs[0]
        .runs
        .iter()
        .map(|r| r.text.encode_utf16().count() as u32)
        .sum::<u32>();
    session
        .delete_text(&context(), &story.id, 0, first_length)
        .unwrap();
    session
        .delete_paragraph_break(&context(), &story.id, 0)
        .unwrap();
    session.add_undo_barrier();
    let captured = session.encode_state_as_update_v1();
    let published = session.save().unwrap();
    assert!(session.undo());
    let latest = session.encode_state_as_update_v1();
    let rebased =
        DeckSession::rebase_checkpoint(&source, &captured, &latest, &published, 21).unwrap();
    let current =
        DeckSession::open_from_update_with_source(&rebased.state, &published, 22).unwrap();
    assert_eq!(parts(&current.save().unwrap()), parts(&source));
    let baseline = DeckSession::open_from_update(&rebased.indexed_state, 23).unwrap();
    let current_snapshot = current.snapshot().unwrap();
    let indexed_snapshot = baseline.snapshot().unwrap();
    let current_story = &current_snapshot.slides[0]
        .shapes
        .iter()
        .find(|s| s.name == "Title")
        .unwrap()
        .text_stories[0];
    let indexed_story = &indexed_snapshot.slides[0]
        .shapes
        .iter()
        .find(|s| s.name == "Title")
        .unwrap()
        .text_stories[0];
    assert_eq!(indexed_story.paragraphs.len(), 1);
    assert_eq!(current_story.paragraphs.len(), 2);
    assert_eq!(
        indexed_story.paragraphs[0].id,
        current_story.paragraphs[1].id
    );
    assert_ne!(
        indexed_story.paragraphs[0].id,
        current_story.paragraphs[0].id
    );
    assert!(
        part_text(&parts(&current.save().unwrap()), "ppt/slides/slide1.xml")
            .contains("a:hlinkClick")
    );
}
