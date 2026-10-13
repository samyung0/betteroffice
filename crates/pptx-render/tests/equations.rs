//! An equation as PowerPoint saves it: the math in an `a14` choice, which the
//! renderer does not read, and a fallback text box filled with a picture of
//! the equation, which it draws.

use pptx_edit::DeckSession;
use pptx_render::{Primitive, SlideRenderer};

const RELATIONSHIPS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const DRAWING: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PRESENTATION: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const SANS: &[u8] = include_bytes!("../../../packages/fonts/assets/LiberationSans-Regular.ttf");
const PNG: &[u8] = &[0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A, 1, 2, 3, 4];

const EQUATION: &str = r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14"><p:sp><p:nvSpPr><p:cNvPr id="5" name="TextBox 4"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="4114800"/><a:ext cx="1625600" cy="369332"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="none"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a14:m><m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr><m:oMath><m:sSup><m:sSupPr><m:ctrlPr><a:rPr lang="en-US" i="1"><a:latin typeface="Cambria Math"/></a:rPr></m:ctrlPr></m:sSupPr><m:e><m:r><a:rPr lang="en-US" i="1"><a:latin typeface="Cambria Math"/></a:rPr><m:t>x</m:t></m:r></m:e><m:sup><m:r><a:rPr lang="en-US" i="1"><a:latin typeface="Cambria Math"/></a:rPr><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath></m:oMathPara></a14:m><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></mc:Choice><mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="5" name="TextBox 4"/><p:cNvSpPr txBox="1"><a:spLocks noRot="1" noChangeAspect="1" noMove="1" noResize="1" noEditPoints="1" noAdjustHandles="1" noChangeArrowheads="1" noChangeShapeType="1" noTextEdit="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="4114800"/><a:ext cx="1625600" cy="369332"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"><a:noFill/></a:rPr><a:t> </a:t></a:r></a:p></p:txBody></p:sp></mc:Fallback></mc:AlternateContent>"#;

fn deck() -> Vec<u8> {
    let xml = |body: String| {
        format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>{body}"#).into_bytes()
    };
    let rels = |body: &str| {
        xml(format!(
            r#"<Relationships xmlns="{RELATIONSHIPS}">{body}</Relationships>"#
        ))
    };
    let tree = r#"<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"#;
    let namespaces = format!(r#"xmlns:a="{DRAWING}" xmlns:r="{OFFICE}" xmlns:p="{PRESENTATION}""#);
    let parts: Vec<(String, Vec<u8>)> = vec![
        (
            "[Content_Types].xml".into(),
            xml(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/></Types>"#.to_owned()),
        ),
        (
            "_rels/.rels".into(),
            rels(&format!(r#"<Relationship Id="rId1" Type="{OFFICE}/officeDocument" Target="ppt/presentation.xml"/>"#)),
        ),
        (
            "ppt/presentation.xml".into(),
            xml(format!(r#"<p:presentation {namespaces}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>"#)),
        ),
        (
            "ppt/_rels/presentation.xml.rels".into(),
            rels(&format!(r#"<Relationship Id="rId1" Type="{OFFICE}/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="{OFFICE}/slide" Target="slides/slide1.xml"/>"#)),
        ),
        (
            "ppt/slides/slide1.xml".into(),
            xml(format!(r#"<p:sld {namespaces}>{tree}{EQUATION}</p:spTree></p:cSld></p:sld>"#)),
        ),
        (
            "ppt/slides/_rels/slide1.xml.rels".into(),
            rels(&format!(r#"<Relationship Id="rId1" Type="{OFFICE}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="{OFFICE}/image" Target="../media/image1.png"/>"#)),
        ),
        ("ppt/media/image1.png".into(), PNG.to_vec()),
        (
            "ppt/slideLayouts/slideLayout1.xml".into(),
            xml(format!(r#"<p:sldLayout {namespaces} type="blank">{tree}</p:spTree></p:cSld></p:sldLayout>"#)),
        ),
        (
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels".into(),
            rels(&format!(r#"<Relationship Id="rId1" Type="{OFFICE}/slideMaster" Target="../slideMasters/slideMaster1.xml"/>"#)),
        ),
        (
            "ppt/slideMasters/slideMaster1.xml".into(),
            xml(format!(r#"<p:sldMaster {namespaces}>{tree}</p:spTree></p:cSld><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>"#)),
        ),
        (
            "ppt/slideMasters/_rels/slideMaster1.xml.rels".into(),
            rels(&format!(r#"<Relationship Id="rId1" Type="{OFFICE}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>"#)),
        ),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn a_powerpoint_equation_draws_its_fallback_picture_in_place() {
    let session = DeckSession::open(&deck(), 951).unwrap();
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, SANS).unwrap();
    let list = renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list;

    let images: Vec<_> = list
        .primitives
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::Image {
                name,
                x,
                y,
                w,
                h,
                asset_id,
                ..
            } => Some((name.as_str(), *x, *y, *w, *h, asset_id.as_deref())),
            _ => None,
        })
        .collect();
    assert_eq!(images.len(), 1, "{:?}", list.primitives);
    let (name, x, y, w, h, asset) = images[0];
    assert_eq!(name, "TextBox 4");
    assert_eq!(asset, Some("ppt/media/image1.png"));
    // EMU to CSS pixels: 9525 per pixel.
    assert_eq!((x, y), (48.0, 432.0));
    assert!((w - 1_625_600.0 / 9525.0).abs() < 0.01, "{w}");
    assert!((h - 369_332.0 / 9525.0).abs() < 0.01, "{h}");

    let drawn_text: String = list
        .primitives
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::TextBox { lines, .. } => Some(lines),
            _ => None,
        })
        .flatten()
        .flat_map(|line| line.runs.iter().map(|run| run.text.as_str()))
        .collect();
    assert_eq!(drawn_text.trim(), "", "the choice's math leaked into text");
}
