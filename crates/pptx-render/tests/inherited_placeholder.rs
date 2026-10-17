//! A placeholder without a full transform of its own draws at the rect and
//! orientation it inherits.

use pptx_edit::{DeckSession, ShapeSnapshot};
use pptx_render::{Primitive, SlideRenderer, Transform};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

const NS: &str = r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main""#;

const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
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
</Relationships>"#;

const SLIDE_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"#;

const LAYOUT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"#;

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

/// A drawn shape: its rect in EMU and the transform it is drawn with.
type Drawn = ((i64, i64, i64, i64), Transform);

fn sp(id: u32, name: &str, placeholder: &str, properties: &str) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="{name}"/><p:cNvSpPr/><p:nvPr>{placeholder}</p:nvPr></p:nvSpPr>{properties}<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr><a:latin typeface="Arial"/></a:rPr><a:t>{name}</a:t></a:r></a:p></p:txBody></p:sp>"#
    )
}

fn filled(transform: &str) -> String {
    format!(r#"<p:spPr>{transform}<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr>"#)
}

fn tree(shapes: &str) -> String {
    format!(
        r#"<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{shapes}</p:spTree></p:cSld>"#
    )
}

fn deck(slide: &str, layout: &str, master: &str) -> Vec<u8> {
    let presentation = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation {NS}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>"#
    );
    let slide = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld {NS}>{}</p:sld>"#,
        tree(slide)
    );
    let layout = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout {NS} type="obj">{}</p:sldLayout>"#,
        tree(layout)
    );
    let master = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster {NS}>{}<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>"#,
        tree(master)
    );
    let parts: Vec<(String, Vec<u8>)> = [
        ("[Content_Types].xml", CONTENT_TYPES.to_owned()),
        ("_rels/.rels", ROOT_RELS.to_owned()),
        ("ppt/presentation.xml", presentation),
        (
            "ppt/_rels/presentation.xml.rels",
            PRESENTATION_RELS.to_owned(),
        ),
        ("ppt/slides/slide1.xml", slide),
        ("ppt/slides/_rels/slide1.xml.rels", SLIDE_RELS.to_owned()),
        ("ppt/slideLayouts/slideLayout1.xml", layout),
        (
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
            LAYOUT_RELS.to_owned(),
        ),
        ("ppt/slideMasters/slideMaster1.xml", master),
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

fn emu(px: f32) -> i64 {
    (f64::from(px) * 9525.0).round() as i64
}

fn drawn(session: &DeckSession, shape_id: &str) -> Drawn {
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    let rendered = renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap();
    rendered
        .display_list
        .primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::Shape {
                shape_id: Some(id),
                x,
                y,
                w,
                h,
                transform,
                ..
            } if id == shape_id => Some(((emu(*x), emu(*y), emu(*w), emu(*h)), *transform)),
            _ => None,
        })
        .unwrap_or_else(|| panic!("{shape_id} was not drawn"))
}

fn shape(session: &DeckSession, index: usize) -> ShapeSnapshot {
    session.snapshot().unwrap().slides[0].shapes[index].clone()
}

/// A body placeholder over a turned and flipped layout body at
/// `952500,1905000 7620000x2857500`.
fn body_deck(properties: &str) -> Vec<u8> {
    deck(
        &sp(3, "Body", r#"<p:ph type="body"/>"#, properties),
        &sp(
            3,
            "Body Placeholder",
            r#"<p:ph type="body"/>"#,
            &filled(
                r#"<a:xfrm rot="1800000" flipH="1"><a:off x="952500" y="1905000"/><a:ext cx="7620000" cy="2857500"/></a:xfrm>"#,
            ),
        ),
        "",
    )
}

const LAYOUT_BODY: (i64, i64, i64, i64) = (952_500, 1_905_000, 7_620_000, 2_857_500);

#[test]
fn an_offset_only_transform_is_not_inherited_and_draws_with_its_own_orientation() {
    let session = DeckSession::open(
        &body_deck(r#"<p:spPr><a:xfrm><a:off x="123825" y="657225"/></a:xfrm></p:spPr>"#),
        1,
    )
    .unwrap();
    let body = shape(&session, 0);

    assert_eq!(
        drawn(&session, &body.id),
        (LAYOUT_BODY, Transform::default())
    );
}

/// A slide number and a date whose indices name the layout's content
/// placeholders, while the master's own slide number and date sit elsewhere.
fn footer_deck() -> Vec<u8> {
    deck(
        &[
            sp(
                4,
                "Number",
                r#"<p:ph type="sldNum" sz="quarter" idx="12"/>"#,
                "<p:spPr/>",
            ),
            sp(
                5,
                "Date",
                r#"<p:ph type="dt" sz="half" idx="10"/>"#,
                "<p:spPr/>",
            ),
        ]
        .concat(),
        &[
            sp(
                6,
                "Content 12",
                r#"<p:ph sz="quarter" idx="12"/>"#,
                &filled(
                    r#"<a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="1905000"/></a:xfrm>"#,
                ),
            ),
            sp(
                7,
                "Content 10",
                r#"<p:ph sz="half" idx="10"/>"#,
                &filled(
                    r#"<a:xfrm><a:off x="952500" y="3810000"/><a:ext cx="3810000" cy="1905000"/></a:xfrm>"#,
                ),
            ),
        ]
        .concat(),
        &[
            sp(
                8,
                "Slide Number",
                r#"<p:ph type="sldNum" sz="quarter" idx="4"/>"#,
                &filled(
                    r#"<a:xfrm><a:off x="8610600" y="6286500"/><a:ext cx="2743200" cy="381000"/></a:xfrm>"#,
                ),
            ),
            sp(
                9,
                "Date",
                r#"<p:ph type="dt" sz="half" idx="2"/>"#,
                &filled(
                    r#"<a:xfrm><a:off x="838200" y="6286500"/><a:ext cx="2743200" cy="381000"/></a:xfrm>"#,
                ),
            ),
        ]
        .concat(),
    )
}

const MASTER_NUMBER: (i64, i64, i64, i64) = (8_610_600, 6_286_500, 2_743_200, 381_000);
const MASTER_DATE: (i64, i64, i64, i64) = (838_200, 6_286_500, 2_743_200, 381_000);

#[test]
fn a_slide_number_and_date_draw_at_the_masters_rect() {
    let session = DeckSession::open(&footer_deck(), 1).unwrap();
    for (index, expected) in [(0, MASTER_NUMBER), (1, MASTER_DATE)] {
        let footer = shape(&session, index);
        assert_eq!(
            drawn(&session, &footer.id),
            (expected, Transform::default())
        );
    }
}

