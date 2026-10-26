//! Pictures and charts in the agent's reads and its by-id read (Epo
//! 2026-10-24, Q240, Q278-Q280, Q284): each format's read lists a picture's
//! alt text, pixel size, crop and whether the model can see it, and a
//! chart's series paired by `c:pt/@idx`; `ids` returns a picture's stored
//! bytes and crop (EMF, WMF, an SVG without its PNG and a linked picture
//! answer their alt text) and a chart's every point, at most limits.json's
//! ids a call and within its answer bytes.

use base64::Engine as _;
use office_service::{
    Checkpoint, Command, EditCode, Error, Format, InspectQuery, apply_commands, inspect_scoped,
    seed, sha256_hex,
};
use serde_json::{Value, json};

const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const C: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WP: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const SVG_EXT: &str = r#"<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rIdSvg"/></a:ext></a:extLst>"#;
/// The by-id answer's bytes (limits.json `inspectIds.answerBytes`).
const ANSWER_BYTES: usize = 8_384_512;

/// A PNG header naming `width` by `height` (nothing is decoded), padded to
/// `len` bytes.
fn png(width: u32, height: u32, len: usize) -> Vec<u8> {
    let mut bytes = b"\x89PNG\r\n\x1a\n\0\0\0\x0dIHDR".to_vec();
    bytes.extend(width.to_be_bytes());
    bytes.extend(height.to_be_bytes());
    bytes.extend([8, 6, 0, 0, 0]);
    bytes.resize(len.max(bytes.len()), 0);
    bytes
}

/// A JPEG header naming 400 by 300.
fn jpeg() -> Vec<u8> {
    vec![
        0xff, 0xd8, 0xff, 0xe0, 0, 4, 0, 0, 0xff, 0xc0, 0, 11, 8, 1, 44, 1, 144, 3, 1, 0x11, 0,
        0xff, 0xd9,
    ]
}

fn emf() -> Vec<u8> {
    let mut bytes = vec![1, 0, 0, 0];
    bytes.resize(40, 0);
    bytes.extend(b" EMF");
    bytes.resize(108, 0);
    bytes
}

const SVG: &[u8] = br#"<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"/>"#;

/// A column chart of `points` months whose values cache skips idx 5 and
/// holds `#N/A` at idx 7: the categories and values pair by idx.
fn chart_xml(points: u32) -> String {
    chart_xml_on("Sheet1", points)
}

/// [`chart_xml`] with its references on sheet `sheet`.
fn chart_xml_on(sheet: &str, points: u32) -> String {
    let categories: String = (0..points)
        .map(|idx| format!(r#"<c:pt idx="{idx}"><c:v>M{idx}</c:v></c:pt>"#))
        .collect();
    let values: String = (0..points)
        .filter(|idx| *idx != 5)
        .map(|idx| {
            let value = if idx == 7 {
                "#N/A".to_owned()
            } else {
                (idx * 10).to_string()
            };
            format!(r#"<c:pt idx="{idx}"><c:v>{value}</c:v></c:pt>"#)
        })
        .collect();
    format!(
        r#"<c:chartSpace xmlns:c="{C}" xmlns:a="{A}"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Monthly sales</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>{sheet}!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx><c:cat><c:strRef><c:f>{sheet}!$A$2:$A${end}</c:f><c:strCache><c:ptCount val="{points}"/>{categories}</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>{sheet}!$B$2:$B${end}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="{points}"/>{values}</c:numCache></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart></c:chartSpace>"#,
        end = points + 1
    )
}

/// Point `idx` of [`chart_xml`] as the reads give it: `[idx, category, value]`.
fn point(idx: u32) -> Value {
    let value = match idx {
        5 => Value::Null,
        7 => json!("#N/A"),
        _ => json!(f64::from(idx * 10)),
    };
    json!([idx, format!("M{idx}"), value])
}

fn zip(files: Vec<(&str, Vec<u8>)>) -> Vec<u8> {
    let files: Vec<(String, Vec<u8>)> = files
        .into_iter()
        .map(|(path, bytes)| (path.to_owned(), bytes))
        .collect();
    ooxml_opc::rezip_parts(&files).unwrap()
}

fn rels(entries: &[(&str, &str, &str)]) -> Vec<u8> {
    let body: String = entries
        .iter()
        .map(|(id, kind, target)| {
            let external = if target.starts_with("https:") {
                r#" TargetMode="External""#
            } else {
                ""
            };
            format!(r#"<Relationship Id="{id}" Type="{R}/{kind}" Target="{target}"{external}/>"#)
        })
        .collect();
    format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{body}</Relationships>"#).into_bytes()
}

fn content_types(overrides: &[(&str, &str)]) -> Vec<u8> {
    let overrides: String = overrides
        .iter()
        .map(|(part, kind)| format!(r#"<Override PartName="{part}" ContentType="{kind}"/>"#))
        .collect();
    format!(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="emf" ContentType="image/x-emf"/><Default Extension="svg" ContentType="image/svg+xml"/>{overrides}</Types>"#).into_bytes()
}

/// A seeded file and its latest state.
struct File {
    format: Format,
    base: Vec<u8>,
    sha: String,
    state: Vec<u8>,
}

impl File {
    fn new(format: Format, base: Vec<u8>) -> Self {
        let state = seed(format, &base).unwrap();
        Self {
            format,
            sha: sha256_hex(&base),
            base,
            state,
        }
    }

    fn checkpoint(&self) -> Checkpoint<'_> {
        Checkpoint {
            format: self.format,
            schema_version: 1,
            base_sha256: &self.sha,
            state: &self.state,
        }
    }

    fn read(&self, query: InspectQuery) -> Result<Value, Error> {
        inspect_scoped(&self.base, self.checkpoint(), &query)
    }

    fn ids(&self, ids: &[&str]) -> Result<Value, Error> {
        self.read(InspectQuery {
            ids: Some(ids.iter().map(|id| (*id).to_owned()).collect()),
            ..InspectQuery::default()
        })
    }
}

fn data(item: &Value) -> Vec<u8> {
    base64::engine::general_purpose::STANDARD
        .decode(item["data"].as_str().unwrap())
        .unwrap()
}

fn refusal(error: Error) -> (EditCode, String) {
    match error {
        Error::Edit { code, message, .. } => (code, message),
        other => panic!("{other:?}"),
    }
}

// ------------------------------------------------------------------ PPTX

fn picture_xml(id: u32, name: &str, descr: &str, blip: &str) -> String {
    format!(
        r#"<p:pic><p:nvPicPr><p:cNvPr id="{id}" name="{name}" descr="{descr}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill>{blip}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="{x}" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>"#,
        x = id * 100_000
    )
}

fn slide_xml(shapes: &str) -> Vec<u8> {
    format!(
        r#"<p:sld xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{shapes}</p:spTree></p:cSld></p:sld>"#
    )
    .into_bytes()
}

/// Slide 1: a cropped PNG, a JPEG, an EMF, an SVG alone, an SVG with its
/// PNG, a linked picture, a chart and a text box; slide 2: four 2.5 MB
/// pictures and a cropped one whose shape id is 0.
fn deck() -> File {
    let slide1 = [
        picture_xml(2, "Photo", "A red square", r#"<a:blip r:embed="rIdPng"/><a:srcRect l="10000" t="5000"/>"#),
        picture_xml(3, "Jpeg", "", r#"<a:blip r:embed="rIdJpg"/>"#),
        picture_xml(4, "Metafile", "Company logo", r#"<a:blip r:embed="rIdEmf"/>"#),
        picture_xml(5, "Vector", "A star", r#"<a:blip r:embed="rIdSvgOnly"/>"#),
        picture_xml(6, "Icon", "An icon", &format!(r#"<a:blip r:embed="rIdIcon">{SVG_EXT}</a:blip>"#)),
        picture_xml(7, "Linked", "Remote", r#"<a:blip r:link="rIdLink"/>"#),
        format!(r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="8" name="Chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="914400"/><a:ext cx="4572000" cy="2743200"/></p:xfrm><a:graphic><a:graphicData uri="{C}"><c:chart xmlns:c="{C}" r:id="rIdChart"/></a:graphicData></a:graphic></p:graphicFrame>"#),
        r#"<p:sp><p:nvSpPr><p:cNvPr id="9" name="Words"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>Hello</a:t></a:r></a:p></p:txBody></p:sp>"#.to_owned(),
    ]
    .concat();
    let mut slide2: String = (0..4)
        .map(|n| {
            picture_xml(
                10 + n,
                &format!("Big {n}"),
                "",
                &format!(r#"<a:blip r:embed="rIdBig{n}"/>"#),
            )
        })
        .collect();
    slide2 += &picture_xml(
        0,
        "Zero",
        "Old alt",
        r#"<a:blip r:embed="rIdBig0"/><a:srcRect l="20000"/>"#,
    );
    let mut files = vec![
        (
            "[Content_Types].xml",
            content_types(&[
                ("/ppt/presentation.xml", "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"),
                ("/ppt/slides/slide1.xml", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"),
                ("/ppt/slides/slide2.xml", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"),
                ("/ppt/charts/chart1.xml", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"),
            ]),
        ),
        ("_rels/.rels", rels(&[("rId1", "officeDocument", "ppt/presentation.xml")])),
        (
            "ppt/presentation.xml",
            format!(r#"<p:presentation xmlns:p="{P}" xmlns:r="{R}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/></p:presentation>"#).into_bytes(),
        ),
        (
            "ppt/_rels/presentation.xml.rels",
            rels(&[("rId1", "slide", "slides/slide1.xml"), ("rId2", "slide", "slides/slide2.xml")]),
        ),
        ("ppt/slides/slide1.xml", slide_xml(&slide1)),
        (
            "ppt/slides/_rels/slide1.xml.rels",
            rels(&[
                ("rIdPng", "image", "../media/red.png"),
                ("rIdJpg", "image", "../media/photo.jpg"),
                ("rIdEmf", "image", "../media/logo.emf"),
                ("rIdSvgOnly", "image", "../media/star.svg"),
                ("rIdIcon", "image", "../media/icon.png"),
                ("rIdSvg", "image", "../media/icon.svg"),
                ("rIdLink", "image", "https://example.com/remote.png"),
                ("rIdChart", "chart", "../charts/chart1.xml"),
            ]),
        ),
        ("ppt/slides/slide2.xml", slide_xml(&slide2)),
        (
            "ppt/slides/_rels/slide2.xml.rels",
            rels(&[
                ("rIdBig0", "image", "../media/big0.png"),
                ("rIdBig1", "image", "../media/big1.png"),
                ("rIdBig2", "image", "../media/big2.png"),
                ("rIdBig3", "image", "../media/big3.png"),
            ]),
        ),
        ("ppt/media/red.png", png(64, 32, 100)),
        ("ppt/media/photo.jpg", jpeg()),
        ("ppt/media/logo.emf", emf()),
        ("ppt/media/star.svg", SVG.to_vec()),
        ("ppt/media/icon.png", png(48, 48, 80)),
        ("ppt/media/icon.svg", SVG.to_vec()),
        ("ppt/charts/chart1.xml", chart_xml(20).into_bytes()),
    ];
    let big = [
        "ppt/media/big0.png",
        "ppt/media/big1.png",
        "ppt/media/big2.png",
        "ppt/media/big3.png",
    ];
    for (n, path) in big.into_iter().enumerate() {
        files.push((path, png(2000 + n as u32, 1500, 2_500_000)));
    }
    File::new(Format::Pptx, zip(files))
}

fn by_name<'a>(shapes: &'a Value, name: &str) -> &'a Value {
    shapes
        .as_array()
        .unwrap()
        .iter()
        .find(|shape| shape["name"] == name)
        .unwrap_or_else(|| panic!("{name} in {shapes}"))
}

fn slide_read(deck: &File, slide: u32) -> Value {
    deck.read(InspectQuery {
        slide: Some(slide),
        start: Some(0),
        count: Some(200),
        ..InspectQuery::default()
    })
    .unwrap()
}

/// The slide read lists each picture's alt text, pixel size, crop and
/// whether the model can see it (an EMF, an SVG alone and a linked picture
/// cannot; an SVG with its PNG can), and a chart's first points paired by idx.
#[test]
fn a_slide_read_lists_pictures_and_charts() {
    let deck = deck();
    let read = slide_read(&deck, 1);
    let shapes = &read["shapes"];
    let photo = by_name(shapes, "Photo");
    assert_eq!(photo["kind"], "picture");
    assert_eq!(photo["alt"], "A red square");
    assert_eq!(photo["pixels"], json!([64, 32]));
    assert_eq!(photo["viewable"], true);
    assert_eq!(
        photo["crop"],
        json!({ "left": 0.1, "top": 0.05, "right": 0.0, "bottom": 0.0 })
    );
    assert_eq!(by_name(shapes, "Jpeg")["pixels"], json!([400, 300]));
    let metafile = by_name(shapes, "Metafile");
    assert_eq!(
        (&metafile["viewable"], &metafile["reason"], &metafile["alt"]),
        (
            &json!(false),
            &json!("an EMF or WMF drawing"),
            &json!("Company logo")
        )
    );
    assert_eq!(
        by_name(shapes, "Vector")["reason"],
        "an SVG without the PNG copy Office stores beside it"
    );
    let icon = by_name(shapes, "Icon");
    assert_eq!(
        (&icon["viewable"], &icon["pixels"]),
        (&json!(true), &json!([48, 48]))
    );
    assert_eq!(
        by_name(shapes, "Linked")["reason"],
        "linked to a file outside this one"
    );
    let chart = by_name(shapes, "Chart");
    assert_eq!(chart["kind"], "chart");
    assert_eq!(chart["chartType"], "column");
    assert_eq!(chart["title"], "Monthly sales");
    let series = &chart["series"][0];
    assert_eq!(series["name"], "Revenue");
    assert_eq!(
        series["refs"],
        json!({ "name": "Sheet1!$B$1", "categories": "Sheet1!$A$2:$A$21", "values": "Sheet1!$B$2:$B$21" })
    );
    assert_eq!(series["count"], 20);
    assert_eq!(
        series["points"],
        json!((0..12).map(point).collect::<Vec<_>>())
    );
    assert_eq!(series["more"], 8);
    assert!(by_name(shapes, "Words").get("viewable").is_none());

    // The deck list names a picture's alt text and a chart's type and title.
    let list = deck
        .read(InspectQuery {
            start: Some(0),
            count: Some(10),
            ..InspectQuery::default()
        })
        .unwrap();
    let listed = &list["slides"][0]["shapes"];
    let listed_id = |name: &str| by_name(shapes, name)["id"].clone();
    let find = |id: &Value| {
        listed
            .as_array()
            .unwrap()
            .iter()
            .find(|shape| shape["id"] == *id)
            .unwrap()
            .clone()
    };
    let metafile = find(&listed_id("Metafile"));
    assert_eq!(
        (&metafile["alt"], &metafile["viewable"]),
        (&json!("Company logo"), &json!(false))
    );
    let chart = find(&listed_id("Chart"));
    assert_eq!(
        (&chart["chartType"], &chart["title"]),
        (&json!("column"), &json!("Monthly sales"))
    );
    assert!(chart.get("series").is_none());
}

/// By id: a picture's stored bytes and crop (an SVG with its PNG sends the
/// PNG), an EMF its alt text, a chart every point paired by idx.
#[test]
fn ids_return_picture_bytes_and_whole_chart_series() {
    let deck = deck();
    let read = slide_read(&deck, 1);
    let id = |name: &str| {
        by_name(&read["shapes"], name)["id"]
            .as_str()
            .unwrap()
            .to_owned()
    };
    // A bracketed id, as the read shows ids, is taken as the id.
    let bracketed = format!("[{}]", id("Photo"));
    let answer = deck
        .ids(&[&bracketed, &id("Icon"), &id("Metafile"), &id("Chart")])
        .unwrap();
    let items = answer["items"].as_array().unwrap();
    assert_eq!(items.len(), 4, "{answer}");
    assert!(answer.get("notice").is_none(), "{answer}");
    assert_eq!(items[0]["id"], id("Photo"));
    assert_eq!(items[0]["mimeType"], "image/png");
    assert_eq!(data(&items[0]), png(64, 32, 100));
    assert_eq!(items[0]["crop"]["left"], 0.1);
    assert_eq!(data(&items[1]), png(48, 48, 80));
    assert_eq!(items[1]["mimeType"], "image/png");
    assert_eq!(
        (
            &items[2]["viewable"],
            &items[2]["alt"],
            items[2].get("data")
        ),
        (&json!(false), &json!("Company logo"), None)
    );
    let series = &items[3]["series"][0];
    assert_eq!(
        series["points"],
        json!((0..20).map(point).collect::<Vec<_>>())
    );
    assert!(series.get("cut").is_none());

    // An SVG alone and a linked picture answer their alt text.
    let answer = deck.ids(&[&id("Vector"), &id("Linked")]).unwrap();
    assert_eq!(answer["items"][0]["alt"], "A star");
    assert_eq!(
        answer["items"][1]["reason"],
        "linked to a file outside this one"
    );
}

/// Security regression (the by-id read's limits, limits.json
/// `inspectIds`): at most 4 ids a call, the rest named; pictures past the
/// answer's bytes are named for another call and the answer stays within
/// them.
#[test]
fn ids_stop_at_four_a_call_and_at_the_answer_bytes() {
    let deck = deck();
    let first = slide_read(&deck, 1);
    let id = |name: &str| {
        by_name(&first["shapes"], name)["id"]
            .as_str()
            .unwrap()
            .to_owned()
    };
    let names = ["Photo", "Jpeg", "Icon", "Metafile", "Vector"];
    let ids: Vec<String> = names.iter().map(|name| id(name)).collect();
    let refs: Vec<&str> = ids.iter().map(String::as_str).collect();
    let answer = deck.ids(&refs).unwrap();
    assert_eq!(answer["items"].as_array().unwrap().len(), 4);
    assert_eq!(
        answer["notice"],
        format!(
            "Not read: {}, past the limit of 4 ids a call; read them in another call.",
            ids[4]
        )
    );

    let second = slide_read(&deck, 2);
    let big: Vec<String> = (0..4)
        .map(|n| {
            by_name(&second["shapes"], &format!("Big {n}"))["id"]
                .as_str()
                .unwrap()
                .to_owned()
        })
        .collect();
    let refs: Vec<&str> = big.iter().map(String::as_str).collect();
    let answer = deck.ids(&refs).unwrap();
    let items = answer["items"].as_array().unwrap();
    assert_eq!(items.len(), 2);
    assert_eq!(data(&items[0]), png(2000, 1500, 2_500_000));
    assert_eq!(data(&items[1]), png(2001, 1500, 2_500_000));
    assert_eq!(
        answer["notice"],
        format!(
            "Not sent: {}, {}: a read carries at most 6 MB of pictures; read them in another call.",
            big[2], big[3]
        )
    );
    assert!(serde_json::to_vec(&answer).unwrap().len() <= ANSWER_BYTES);
}

/// The ids found are answered and the rest named with why (Epo 2026-10-26,
/// Q381): an unknown id of a shape's form is gone, anything else is
/// malformed and shows the deck's ids, a text shape is no picture; blank
/// ids alone are refused.
#[test]
fn unknown_and_malformed_ids_are_named_beside_the_ids_found() {
    let deck = deck();
    let read = slide_read(&deck, 1);
    let id = |name: &str| {
        by_name(&read["shapes"], name)["id"]
            .as_str()
            .unwrap()
            .to_owned()
    };
    let photo = id("Photo");
    let gone = format!("{photo}9");
    let answer = deck
        .ids(&[&gone, &photo, "picture 1", &id("Words")])
        .unwrap();
    let items = answer["items"].as_array().unwrap();
    assert_eq!(items.len(), 1, "{answer}");
    assert_eq!(data(&items[0]), png(64, 32, 100));
    let notice = answer["notice"].as_str().unwrap();
    assert!(
        notice.starts_with(&format!("Not found: {gone} is no longer in the deck; picture 1 is a malformed picture or chart id; picture and chart ids read like {photo:?}, slide 1's")),
        "{notice}"
    );
    assert!(
        notice.contains(&format!(
            "{} is a text, not a picture or chart",
            id("Words")
        )),
        "{notice}"
    );
    assert_eq!(
        refusal(deck.ids(&[" ", "[]"]).unwrap_err()).0,
        EditCode::InvalidInput
    );
}

/// A picture added in the editor (source id 0 in the deck's state) takes no
/// alt text or crop from the file's picture of id 0, which keeps its own
/// (review r1 L5).
#[test]
fn an_added_picture_takes_nothing_from_a_file_picture() {
    let mut deck = deck();
    let session =
        pptx_edit::DeckSession::open_from_update_with_source(&deck.state, &deck.base, 7).unwrap();
    let slide = session.slide_ids().unwrap()[1].clone();
    session
        .add_picture(
            &pptx_edit::EditCtx::local("test"),
            &slide,
            &pptx_edit::PictureDraft {
                name: "Pasted".into(),
                rect: pptx_edit::ShapeRect {
                    x: 0,
                    y: 0,
                    width: 914_400,
                    height: 914_400,
                },
                content_type: "image/png".into(),
                media_bytes: png(30, 20, 200),
            },
        )
        .unwrap();
    deck.state = session.encode_state_as_update_v1();
    let shapes = slide_read(&deck, 2)["shapes"].clone();
    let zero = by_name(&shapes, "Zero");
    assert_eq!(
        (&zero["alt"], &zero["crop"]["left"]),
        (&json!("Old alt"), &json!(0.2))
    );
    let pasted = by_name(&shapes, "Pasted");
    assert_eq!(pasted["pixels"], json!([30, 20]));
    assert!(pasted.get("alt").is_none(), "{pasted}");
    assert!(pasted.get("crop").is_none(), "{pasted}");
}

// ------------------------------------------------------------------ DOCX

fn docx_picture(id: u32, descr: &str, blip: &str) -> String {
    format!(
        r#"<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="914400"/><wp:docPr id="{id}" name="Picture {id}" descr="{descr}"/><a:graphic xmlns:a="{A}"><a:graphicData uri="{PIC}"><pic:pic xmlns:pic="{PIC}"><pic:nvPicPr><pic:cNvPr id="{id}" name="p{id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill>{blip}<a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"#
    )
}

/// Paragraphs: text, a cropped PNG, an EMF, an SVG alone, an SVG with its
/// PNG, a linked picture, an equation, a chart.
fn document() -> File {
    let body = [
        r#"<w:p><w:r><w:t>Before</w:t></w:r></w:p>"#.to_owned(),
        docx_picture(1, "A red square", r#"<a:blip r:embed="rIdPng"/><a:srcRect l="10000"/>"#),
        docx_picture(2, "Company logo", r#"<a:blip r:embed="rIdEmf"/>"#),
        docx_picture(3, "A star", r#"<a:blip r:embed="rIdSvgOnly"/>"#),
        docx_picture(4, "An icon", &format!(r#"<a:blip r:embed="rIdIcon">{SVG_EXT}</a:blip>"#)),
        docx_picture(5, "Remote", r#"<a:blip r:link="rIdLink"/>"#),
        r#"<w:p><m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath><m:r><m:t>x=2</m:t></m:r></m:oMath></m:oMathPara></w:p>"#.to_owned(),
        format!(r#"<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="4572000" cy="2743200"/><wp:docPr id="6" name="Chart 1"/><a:graphic xmlns:a="{A}"><a:graphicData uri="{C}"><c:chart xmlns:c="{C}" r:id="rIdChart"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"#),
    ]
    .concat();
    let files = vec![
        (
            "[Content_Types].xml",
            content_types(&[
                ("/word/document.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"),
                ("/word/charts/chart1.xml", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"),
            ]),
        ),
        ("_rels/.rels", rels(&[("rId1", "officeDocument", "word/document.xml")])),
        (
            "word/_rels/document.xml.rels",
            rels(&[
                ("rIdPng", "image", "media/red.png"),
                ("rIdEmf", "image", "media/logo.emf"),
                ("rIdSvgOnly", "image", "media/star.svg"),
                ("rIdIcon", "image", "media/icon.png"),
                ("rIdSvg", "image", "media/icon.svg"),
                ("rIdLink", "image", "https://example.com/remote.png"),
                ("rIdChart", "chart", "charts/chart1.xml"),
            ]),
        ),
        (
            "word/document.xml",
            format!(r#"<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:wp="{WP}"><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"#).into_bytes(),
        ),
        ("word/media/red.png", png(64, 32, 100)),
        ("word/media/logo.emf", emf()),
        ("word/media/star.svg", SVG.to_vec()),
        ("word/media/icon.png", png(48, 48, 80)),
        ("word/media/icon.svg", SVG.to_vec()),
        ("word/charts/chart1.xml", chart_xml(20).into_bytes()),
    ];
    File::new(Format::Docx, zip(files))
}

fn page(document: &File) -> Vec<Value> {
    document
        .read(InspectQuery {
            start: Some(0),
            count: Some(200),
            ..InspectQuery::default()
        })
        .unwrap()["paragraphs"]
        .as_array()
        .unwrap()
        .clone()
}

/// The paged read lists each paragraph's pictures, charts and equations.
#[test]
fn a_docx_page_lists_pictures_charts_and_equation_text() {
    let document = document();
    let paragraphs = page(&document);
    let object = |n: usize| paragraphs[n]["objects"][0].clone();
    let photo = object(1);
    assert_eq!(photo["kind"], "picture");
    assert!(
        photo["id"].as_str().unwrap().starts_with("body:picture:"),
        "{photo}"
    );
    assert_eq!(
        (
            &photo["alt"],
            &photo["pixels"],
            &photo["viewable"],
            &photo["crop"]["left"]
        ),
        (
            &json!("A red square"),
            &json!([64, 32]),
            &json!(true),
            &json!(0.1)
        )
    );
    assert_eq!(object(2)["reason"], "an EMF or WMF drawing");
    assert_eq!(object(2)["alt"], "Company logo");
    assert_eq!(
        object(3)["reason"],
        "an SVG without the PNG copy Office stores beside it"
    );
    assert_eq!(
        (&object(4)["viewable"], &object(4)["pixels"]),
        (&json!(true), &json!([48, 48]))
    );
    assert_eq!(object(5)["reason"], "linked to a file outside this one");
    assert_eq!(object(6), json!({ "kind": "equation", "text": "x=2" }));
    let chart = object(7);
    assert!(
        chart["id"].as_str().unwrap().starts_with("body:chart:"),
        "{chart}"
    );
    assert_eq!(
        (&chart["chartType"], &chart["title"]),
        (&json!("column"), &json!("Monthly sales"))
    );
    assert_eq!(
        chart["series"][0]["points"],
        json!((0..12).map(point).collect::<Vec<_>>())
    );
    assert!(paragraphs[0].get("objects").is_none());
}

/// By id: a picture's bytes, a chart's every point; a paragraph id holding
/// one picture stands for it; a gone id and a malformed one are refused.
#[test]
fn docx_ids_return_pictures_and_charts() {
    let document = document();
    let paragraphs = page(&document);
    let id = |n: usize| {
        paragraphs[n]["objects"][0]["id"]
            .as_str()
            .unwrap()
            .to_owned()
    };
    let answer = document.ids(&[&id(1), &id(4), &id(2), &id(7)]).unwrap();
    let items = answer["items"].as_array().unwrap();
    assert_eq!(data(&items[0]), png(64, 32, 100));
    assert_eq!(items[0]["crop"]["left"], 0.1);
    assert_eq!(data(&items[1]), png(48, 48, 80));
    assert_eq!(
        (&items[2]["viewable"], &items[2]["alt"]),
        (&json!(false), &json!("Company logo"))
    );
    assert_eq!(
        items[3]["series"][0]["points"],
        json!((0..20).map(point).collect::<Vec<_>>())
    );

    let paragraph = paragraphs[1]["id"].as_str().unwrap();
    let answer = document.ids(&[paragraph]).unwrap();
    assert_eq!(data(&answer["items"][0]), png(64, 32, 100));

    // A gone id and a malformed one are named; the ids found answer (Q381).
    let photo = id(1);
    let (client, _) = photo.rsplit_once('.').unwrap();
    let gone = format!("{client}.999999");
    let answer = document.ids(&[&gone, "image1.png", &id(4)]).unwrap();
    assert_eq!(answer["items"].as_array().unwrap().len(), 1, "{answer}");
    assert_eq!(data(&answer["items"][0]), png(48, 48, 80));
    assert_eq!(
        answer["notice"],
        format!(
            "Not found: {gone} is no longer in the document; image1.png is a malformed picture or chart id; picture and chart ids read like {photo:?}: copy each id whole from inspect_office_document."
        )
    );
}

/// A page names each media and chart part many times; the read reads each
/// part once (review r1 M3), and so does a deck's slide list.
#[test]
fn a_read_reads_each_part_once() {
    let mut body = String::new();
    for n in 0..60 {
        body += &docx_picture(100 + n, "Same", r#"<a:blip r:embed="rIdPng"/>"#);
        body += &format!(
            r#"<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="4572000" cy="2743200"/><wp:docPr id="{}" name="Chart {n}"/><a:graphic xmlns:a="{A}"><a:graphicData uri="{C}"><c:chart xmlns:c="{C}" r:id="rIdChart"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"#,
            1000 + n
        );
    }
    let files = vec![
        (
            "[Content_Types].xml",
            content_types(&[
                ("/word/document.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"),
                ("/word/charts/chart1.xml", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"),
            ]),
        ),
        ("_rels/.rels", rels(&[("rId1", "officeDocument", "word/document.xml")])),
        (
            "word/_rels/document.xml.rels",
            rels(&[("rIdPng", "image", "media/red.png"), ("rIdChart", "chart", "charts/chart1.xml")]),
        ),
        (
            "word/document.xml",
            format!(r#"<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:wp="{WP}"><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"#).into_bytes(),
        ),
        ("word/media/red.png", png(64, 32, 100)),
        ("word/charts/chart1.xml", chart_xml(20).into_bytes()),
    ];
    let document = File::new(Format::Docx, zip(files));
    let before = office_service::part_reads();
    let listed = page(&document)
        .iter()
        .filter(|paragraph| paragraph.get("objects").is_some())
        .count();
    assert_eq!(listed, 120);
    assert_eq!(office_service::part_reads() - before, 2);

    let deck = deck();
    let before = office_service::part_reads();
    deck.read(InspectQuery {
        start: Some(0),
        count: Some(10),
        ..InspectQuery::default()
    })
    .unwrap();
    // Slide 1's red, photo, logo, star and icon headers and its chart (the
    // linked picture reads nothing), slide 2's four parts ("Zero" shares
    // big0); without the per-call cache, 11 (review r2 L4).
    assert_eq!(office_service::part_reads() - before, 10);
}

/// A drawn shape comes through as its text, an equation as its own.
#[test]
fn docx_shapes_and_equations_read_as_text() {
    let document = File::new(
        Format::Docx,
        include_bytes!(
            "../../betteroffice-docx/tests/corpus/fixtures/wordprocessingml-comprehensive.docx"
        )
        .to_vec(),
    );
    let objects: Vec<Value> = page(&document)
        .iter()
        .flat_map(|paragraph| paragraph["objects"].as_array().cloned().unwrap_or_default())
        .collect();
    assert!(
        objects.contains(&json!({ "kind": "equation", "text": "x=12" })),
        "{objects:?}"
    );
    assert!(
        objects.iter().any(|object| object["kind"] == "drawing"
            && object["text"]
                .as_str()
                .is_some_and(|text| text.contains("A wrapped-square shape with its own story"))),
        "{objects:?}"
    );
}

// ------------------------------------------------------------------ XLSX

/// Sheet 1's drawing: a cropped picture at B3, a group of two pictures, a
/// chart and a text box; sheet 2 has none.
fn workbook() -> File {
    let pic = |id: u32, descr: &str, rel: &str, crop: &str| {
        format!(
            r#"<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{id}" name="Picture {id}" descr="{descr}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="{rel}"/>{crop}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>"#
        )
    };
    let anchor = |col: u32, row: u32, body: &str| {
        format!(
            r#"<xdr:twoCellAnchor><xdr:from><xdr:col>{col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>{}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>{body}<xdr:clientData/></xdr:twoCellAnchor>"#,
            col + 3,
            row + 5
        )
    };
    let drawing = [
        anchor(1, 2, &pic(2, "A red square", "rIdPng", r#"<a:srcRect l="NaN" r="25000"/>"#)),
        anchor(
            5,
            2,
            &format!(
                r#"<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="3" name="Group"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr/>{}{}</xdr:grpSp>"#,
                pic(4, "First", "rIdJpg", ""),
                pic(5, "Second", "rIdEmf", "")
            ),
        ),
        anchor(1, 10, &format!(r#"<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="6" name="Chart"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="{C}"><c:chart xmlns:c="{C}" r:id="rIdChart"/></a:graphicData></a:graphic></xdr:graphicFrame>"#)),
        anchor(8, 10, r#"<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="7" name="Note"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr><xdr:spPr/><xdr:txBody><a:bodyPr/><a:p><a:r><a:t>Check Q2</a:t></a:r></a:p></xdr:txBody></xdr:sp>"#),
    ]
    .concat();
    let sheet = |rows: &str, extra: &str| {
        format!(r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="{R}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Month</t></is></c><c r="B1" t="inlineStr"><is><t>Revenue</t></is></c></row>{rows}</sheetData>{extra}</worksheet>"#).into_bytes()
    };
    // The cells the chart names, as its cache holds them.
    let rows: String = (0..20)
        .map(|idx| {
            let row = idx + 2;
            let value = match idx {
                5 => String::new(),
                7 => format!(r#"<c r="B{row}" t="e"><v>#N/A</v></c>"#),
                _ => format!(r#"<c r="B{row}"><v>{}</v></c>"#, idx * 10),
            };
            format!(r#"<row r="{row}"><c r="A{row}" t="inlineStr"><is><t>M{idx}</t></is></c>{value}</row>"#)
        })
        .collect();
    let files = vec![
        (
            "[Content_Types].xml",
            content_types(&[
                ("/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"),
                ("/xl/worksheets/sheet1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"),
                ("/xl/worksheets/sheet2.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"),
                ("/xl/drawings/drawing1.xml", "application/vnd.openxmlformats-officedocument.drawing+xml"),
                ("/xl/charts/chart1.xml", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"),
            ]),
        ),
        ("_rels/.rels", rels(&[("rId1", "officeDocument", "xl/workbook.xml")])),
        (
            "xl/workbook.xml",
            format!(r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="{R}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Notes" sheetId="2" r:id="rId2"/></sheets></workbook>"#).into_bytes(),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            rels(&[("rId1", "worksheet", "worksheets/sheet1.xml"), ("rId2", "worksheet", "worksheets/sheet2.xml")]),
        ),
        ("xl/worksheets/sheet1.xml", sheet(&rows, r#"<drawing r:id="rIdDrawing"/>"#)),
        ("xl/worksheets/_rels/sheet1.xml.rels", rels(&[("rIdDrawing", "drawing", "../drawings/drawing1.xml")])),
        ("xl/worksheets/sheet2.xml", sheet("", "")),
        (
            "xl/drawings/drawing1.xml",
            format!(r#"<xdr:wsDr xmlns:xdr="{XDR}" xmlns:a="{A}" xmlns:r="{R}">{drawing}</xdr:wsDr>"#).into_bytes(),
        ),
        (
            "xl/drawings/_rels/drawing1.xml.rels",
            rels(&[
                ("rIdPng", "image", "../media/red.png"),
                ("rIdJpg", "image", "../media/photo.jpg"),
                ("rIdEmf", "image", "../media/logo.emf"),
                ("rIdChart", "chart", "../charts/chart1.xml"),
            ]),
        ),
        ("xl/media/red.png", png(64, 32, 100)),
        ("xl/media/photo.jpg", jpeg()),
        ("xl/media/logo.emf", emf()),
        ("xl/charts/chart1.xml", chart_xml_on("Data", 20).into_bytes()),
    ];
    File::new(Format::Xlsx, zip(files))
}

fn sheets(book: &File) -> Value {
    book.read(InspectQuery::default()).unwrap()["sheets"].clone()
}

/// The sheet list gives each sheet's pictures (alt text, pixel size, crop,
/// anchor cell), charts and drawn shapes' text.
#[test]
fn the_sheet_list_shows_pictures_charts_and_shapes() {
    let book = workbook();
    let sheets = sheets(&book);
    let objects = sheets[0]["objects"].as_array().unwrap();
    assert_eq!(
        objects[0],
        json!({
            "id": "xl/drawings/drawing1.xml#0:0", "kind": "picture", "alt": "A red square",
            "cell": "B3", "viewable": true, "pixels": [64, 32],
            "crop": { "left": 0.0, "top": 0.0, "right": 0.25, "bottom": 0.0 },
        })
    );
    assert_eq!(objects[1]["id"], "xl/drawings/drawing1.xml#1:0");
    assert_eq!(objects[1]["pixels"], json!([400, 300]));
    assert_eq!(
        (
            &objects[2]["id"],
            &objects[2]["viewable"],
            &objects[2]["alt"]
        ),
        (
            &json!("xl/drawings/drawing1.xml#1:1"),
            &json!(false),
            &json!("Second")
        )
    );
    assert_eq!(objects[3]["id"], "xl/drawings/drawing1.xml#2:chart");
    assert_eq!(objects[3]["cell"], "B11");
    assert_eq!(
        objects[3]["series"][0]["points"],
        json!((0..12).map(point).collect::<Vec<_>>())
    );
    assert_eq!(
        objects[4],
        json!({ "kind": "drawing", "text": "Check Q2", "cell": "I11" })
    );
    assert!(sheets[1].get("objects").is_none());
}

/// The ids the sheet list shows are the workbook's own image ids, and the
/// by-id read, which takes pictures from the package's drawing parts without
/// opening the workbook, returns their bytes and the chart's every point.
/// The chart's references and cell follow a rename and an inserted row, as
/// the workbook holds them, and its points are the cells' now, an edited
/// one included (Epo 2026-10-26, Q381); a picture's cell moves with the
/// rows.
#[test]
fn xlsx_ids_round_trip_between_the_sheet_list_and_the_by_id_read() {
    let mut book = workbook();
    let workbook = betteroffice_xlsx::Workbook::open(&book.base).unwrap();
    let images: Vec<String> = workbook
        .embedded_images(betteroffice_xlsx::SheetId(0))
        .unwrap()
        .into_iter()
        .map(|image| image.id)
        .collect();
    let listed: Vec<String> = sheets(&book)[0]["objects"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|object| object["kind"] == "picture")
        .map(|object| object["id"].as_str().unwrap().to_owned())
        .collect();
    assert_eq!(listed, images);

    let edit = |book: &mut File, command: Command| {
        book.state = apply_commands(&book.base, book.checkpoint(), &[command])
            .unwrap()
            .state;
    };
    edit(
        &mut book,
        Command::RenameSheet {
            sheet: "Data".into(),
            expected_name: "Data".into(),
            name: "Sales".into(),
        },
    );
    edit(
        &mut book,
        Command::InsertRows {
            sheet: "Sales".into(),
            at: 1,
            count: 3,
        },
    );
    edit(
        &mut book,
        Command::SetCell {
            sheet: "Sales".into(),
            cell: "B5".into(),
            expected_value: "0".into(),
            value: "999".into(),
        },
    );
    // M0's value, now 999 in its cell where the file's cache holds 0.
    let points = |count: u32| {
        let mut points: Vec<Value> = (0..count).map(point).collect();
        points[0] = json!([0, "M0", 999.0]);
        points
    };
    let renamed = sheets(&book);
    assert_eq!(renamed[0]["name"], "Sales");
    let objects = renamed[0]["objects"].as_array().unwrap();
    assert_eq!(objects[0]["cell"], "B6");
    let chart = &objects[3];
    assert_eq!(chart["cell"], "B14");
    // The workbook writes a renamed sheet's references quoted.
    let refs = json!({ "name": "'Sales'!$B$4", "categories": "'Sales'!$A$5:$A$24", "values": "'Sales'!$B$5:$B$24" });
    assert_eq!(chart["series"][0]["refs"], refs);
    assert_eq!(chart["series"][0]["name"], "Revenue");
    assert_eq!(chart["series"][0]["points"], json!(points(12)));
    let ids: Vec<&str> = objects
        .iter()
        .filter_map(|object| object["id"].as_str())
        .collect();
    let answer = book.ids(&ids).unwrap();
    let items = answer["items"].as_array().unwrap();
    assert_eq!(data(&items[0]), png(64, 32, 100));
    assert_eq!(items[0]["crop"]["right"], 0.25);
    assert_eq!(data(&items[1]), jpeg());
    assert_eq!(
        (&items[2]["viewable"], &items[2]["alt"]),
        (&json!(false), &json!("Second"))
    );
    assert_eq!(items[3]["series"][0]["refs"], refs);
    assert_eq!(items[3]["series"][0]["points"], json!(points(20)));

    // A picture read by id with a state holding only a sheet order naming a
    // sheet the workbook does not hold, which it cannot open: the package
    // path reads no more of it.
    let state = std::mem::replace(
        &mut book.state,
        sheet_order_only(&["sheet:0", "sheet:missing"]),
    );
    assert_eq!(
        data(&book.ids(&[ids[0]]).unwrap()["items"][0]),
        png(64, 32, 100)
    );
    assert!(book.read(InspectQuery::default()).is_err());
    book.state = state;

    // Malformed, unknown and gone ids are named beside the ids found (Q381).
    let answer = book
        .ids(&[
            "Data!B3",
            "xl/drawings/drawing1.xml#9:0",
            "xl/drawings/drawing9.xml#0:0",
            ids[0],
        ])
        .unwrap();
    assert_eq!(answer["items"].as_array().unwrap().len(), 1);
    let notice = answer["notice"].as_str().unwrap();
    assert!(
        notice.starts_with("Not found: Data!B3 is a malformed picture or chart id; XLSX picture ids read like \"xl/drawings/drawing1.xml#0:0\""),
        "{notice}"
    );
    for unknown in [
        "xl/drawings/drawing1.xml#9:0",
        "xl/drawings/drawing9.xml#0:0",
    ] {
        assert!(
            notice.contains(&format!(
                "{unknown} names no picture or chart of the workbook; read the sheet list"
            )),
            "{notice}"
        );
    }

    // Its sheet deleted, a picture is gone from the by-id read as from the
    // sheet list (the state's sheets, not the package's).
    edit(
        &mut book,
        Command::DeleteSheet {
            sheet: "Sales".into(),
            expected_name: "Sales".into(),
        },
    );
    let listed = sheets(&book);
    assert_eq!(listed.as_array().unwrap().len(), 1);
    assert!(listed[0].get("objects").is_none());
    let answer = book.ids(&[ids[0]]).unwrap();
    assert_eq!(answer["items"], json!([]));
    assert_eq!(
        answer["notice"],
        format!(
            "Not found: {} is no longer in the workbook: its sheet was deleted; read the sheet list for the ids.",
            ids[0]
        )
    );
}

/// A drawing part that does not read is named in its sheet's notice and
/// skipped, the rest of the sheet list read (review r1 M1), and an id in it
/// is named by the by-id read.
#[test]
fn an_unreadable_drawing_is_named_and_skipped() {
    let good = workbook();
    let mut parts = ooxml_opc::unzip_parts(&good.base).unwrap();
    for (path, bytes) in &mut parts {
        if path == "xl/drawings/drawing1.xml" {
            bytes.truncate(bytes.len() / 2);
        }
    }
    let book = File::new(Format::Xlsx, ooxml_opc::rezip_parts(&parts).unwrap());
    let sheets = sheets(&book);
    assert_eq!(sheets.as_array().unwrap().len(), 2);
    assert!(sheets[0].get("objects").is_none());
    assert_eq!(
        sheets[0]["notice"],
        "Not listed: the pictures and charts of xl/drawings/drawing1.xml, a part that could not be read."
    );
    assert_eq!(
        book.ids(&["xl/drawings/drawing1.xml#0:0"]).unwrap()["notice"],
        "Not found: xl/drawings/drawing1.xml#0:0 is in a drawing part that could not be read."
    );
}

/// `zipped` with part `name`'s stored checksum made wrong in its local and
/// central headers, so only a read to the part's end finds it damaged.
fn wrong_checksum(mut zipped: Vec<u8>, name: &str) -> Vec<u8> {
    let name = name.as_bytes();
    let mut places = 0;
    for at in 0..zipped.len() - name.len() {
        if &zipped[at..at + name.len()] != name {
            continue;
        }
        let header = |back: usize, signature: &[u8; 4]| {
            at.checked_sub(back)
                .filter(|&start| &zipped[start..start + 4] == signature)
        };
        let crc = match (header(46, b"PK\x01\x02"), header(30, b"PK\x03\x04")) {
            (Some(central), _) => central + 16,
            (None, Some(local)) => local + 14,
            (None, None) => continue,
        };
        for byte in &mut zipped[crc..crc + 4] {
            *byte ^= 0xff;
        }
        places += 1;
    }
    assert_eq!(places, 2);
    zipped
}

/// [`workbook`] with `xl/media/red.png` holding `bytes` and a wrong stored
/// checksum in the package the by-id read takes it from (the state is the
/// sound package's).
fn workbook_with_damaged_png(bytes: Vec<u8>) -> File {
    let mut parts = ooxml_opc::unzip_parts(&workbook().base).unwrap();
    for (path, part) in &mut parts {
        if path == "xl/media/red.png" {
            *part = bytes.clone();
        }
    }
    let mut book = File::new(Format::Xlsx, ooxml_opc::rezip_parts(&parts).unwrap());
    book.base = wrong_checksum(std::mem::take(&mut book.base), "xl/media/red.png");
    book.sha = sha256_hex(&book.base);
    book
}

/// Security regression (Q240 review r1 M4, review r2 M1): an XLSX picture
/// past what one answer carries is read no further than that and answers
/// its alt text. Its checksum is wrong, so a read past the bound would reach
/// its end and answer it damaged instead.
#[test]
fn an_xlsx_picture_past_the_answer_bytes_is_too_large_to_send() {
    let book = workbook_with_damaged_png(png(64, 32, ANSWER_BYTES + 1));
    let answer = book.ids(&["xl/drawings/drawing1.xml#0:0"]).unwrap();
    let item = &answer["items"][0];
    assert_eq!(
        (
            &item["viewable"],
            &item["reason"],
            &item["alt"],
            item.get("data")
        ),
        (
            &json!(false),
            &json!("too large to send: a read carries at most 6 MB of pictures"),
            &json!("A red square"),
            None
        )
    );
}

/// A media part that does not read answers damaged, not too large (review
/// r2 N4).
#[test]
fn a_damaged_xlsx_picture_answers_damaged() {
    let book = workbook_with_damaged_png(png(64, 32, 100));
    let answer = book.ids(&["xl/drawings/drawing1.xml#0:0"]).unwrap();
    let item = &answer["items"][0];
    assert_eq!(
        (&item["viewable"], &item["reason"], item.get("data")),
        (
            &json!(false),
            &json!("damaged or unknown picture data"),
            None
        )
    );
}

/// A chart whose categories name a two-dimensional range (a multi-level
/// axis) keeps its values from the cells, its categories empty and the
/// reason on the series (review r2 L2).
#[test]
fn a_chart_keeps_its_values_when_its_categories_do_not_read() {
    let mut parts = ooxml_opc::unzip_parts(&workbook().base).unwrap();
    for (path, part) in &mut parts {
        if path == "xl/charts/chart1.xml" {
            *part = String::from_utf8(std::mem::take(part))
                .unwrap()
                .replace("Data!$A$2:$A$21", "Data!$A$2:$B$21")
                .into_bytes();
        }
    }
    let book = File::new(Format::Xlsx, ooxml_opc::rezip_parts(&parts).unwrap());
    let values: Vec<Value> = (0..20)
        .filter(|idx| *idx != 5)
        .map(|idx| {
            let mut point = point(idx);
            point[1] = Value::Null;
            point
        })
        .collect();
    let reason =
        "its categories could not be read from the cells their reference names; read those cells";
    // The eight positions past the twelve shown are counted.
    let listed = &sheets(&book)[0]["objects"][3]["series"][0];
    assert_eq!(
        (
            &listed["name"],
            &listed["points"],
            &listed["reason"],
            &listed["more"]
        ),
        (
            &json!("Revenue"),
            &json!(values[..11]),
            &json!(reason),
            &json!(8)
        )
    );
    let answer = book.ids(&["xl/drawings/drawing1.xml#2:chart"]).unwrap();
    let series = &answer["items"][0]["series"][0];
    assert_eq!(
        (&series["points"], &series["reason"]),
        (&json!(values), &json!(reason))
    );
}

/// A picture's cell in the sheet list moves up when rows above it are
/// deleted, onto the first row after them when its own row goes (review r2
/// N5).
#[test]
fn a_picture_cell_moves_up_with_deleted_rows() {
    let mut book = workbook();
    book.state = apply_commands(
        &book.base,
        book.checkpoint(),
        &[Command::DeleteRows {
            sheet: "Data".into(),
            at: 2,
            count: 2,
            expected: vec![
                vec!["M0".into(), "0".into()],
                vec!["M1".into(), "10".into()],
            ],
        }],
    )
    .unwrap()
    .state;
    let objects = &sheets(&book)[0]["objects"];
    assert_eq!(
        (&objects[0]["cell"], &objects[3]["cell"]),
        (&json!("B2"), &json!("B9"))
    );
}

/// A Yjs state whose `xlsx:sheet-order` holds `keys` and nothing else.
fn sheet_order_only(keys: &[&str]) -> Vec<u8> {
    use yrs::{Array, Transact, WriteTxn};
    let doc = yrs::Doc::new();
    let mut txn = doc.transact_mut();
    let order = txn.get_or_insert_array("xlsx:sheet-order");
    for key in keys {
        order.push_back(&mut txn, *key);
    }
    txn.encode_update_v1()
}
