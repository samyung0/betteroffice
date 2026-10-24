//! XLSX pictures, charts and drawn shapes, read from a sheet's drawing parts
//! (Epo 2026-10-24, Q240): the sheet list shows them, and the by-id read
//! takes them from the package alone, without opening the workbook, whose
//! recalculation costs seconds on a large book (the Editor's call: only
//! while the ids match the sheet read's). No workbook edit changes a
//! drawing (the state holds no drawing), so the package holds them as the
//! sheet shows them; of the state the by-id read takes only the sheets it
//! still holds, so a sheet deleted since the read answers its pictures as
//! gone, as the sheet read no longer lists them. A picture's
//! id is the workbook's own (xlsx-parse `sheet_images`):
//! `<drawing>#<anchor>:<n>`, `n` counting the anchor's embedded pictures; a
//! chart's is `<drawing>#<anchor>:chart`.

use std::borrow::Cow;
use std::collections::{HashMap, HashSet};

use docx_parse::xml::{ParseBudget, ParseLimits, XmlElement, parse_xml};
use docx_parse::{RelationshipMap, RelationshipTarget, parse_relationships};
use serde_json::{Value, json};
use yrs::updates::decoder::Decode;
use yrs::{Array, ReadTxn, Transact, Update};

use crate::common::malformed_id;
use crate::error::{EditCode, Error, Result};
use crate::objects::{self, Found, MAX_LISTED, MAX_TEXT_CHARS, Picture, Stored, crop_side};

/// Bytes a drawing, relationship or chart part may inflate to.
const MAX_XML_BYTES: u64 = 32 << 20;
/// Nesting a drawing walk follows (xlsx-parse's `MAX_DEPTH`).
const MAX_DEPTH: usize = 64;

enum Kind {
    /// `embed` names its media relationship; `None` for a linked picture.
    Picture {
        embed: Option<String>,
        alt: Option<String>,
        crop: [f64; 4],
    },
    Chart(String),
    Shape(String),
}

struct Drawn {
    id: Option<String>,
    cell: Option<String>,
    kind: Kind,
}

fn parsed(xml: &[u8], part: &str) -> Result<docx_parse::xml::XmlDocument> {
    let limits = ParseLimits::default();
    parse_xml(xml, part, &mut ParseBudget::new(&limits)).map_err(Error::engine)
}

fn rels_path(part: &str) -> String {
    match part.rsplit_once('/') {
        Some((directory, name)) => format!("{directory}/_rels/{name}.rels"),
        None => format!("_rels/{part}.rels"),
    }
}

fn relationships(xml: Option<&[u8]>, part: &str) -> Result<RelationshipMap> {
    let Some(xml) = xml else {
        return Ok(RelationshipMap::new());
    };
    let limits = ParseLimits::default();
    parse_relationships(xml, part, &mut ParseBudget::new(&limits)).map_err(Error::engine)
}

/// Relationship `id`'s internal target from `part`, `Err(())` for an
/// external one.
fn target(rels: &RelationshipMap, part: &str, id: &str) -> Option<std::result::Result<String, ()>> {
    let relationship = rels.get(id)?;
    match docx_parse::resolve_relationship_target(part, relationship).ok()? {
        RelationshipTarget::Internal(path) => Some(Ok(path)),
        RelationshipTarget::External(_) => Some(Err(())),
    }
}

/// `col`, `row` (from 0) as `A1`.
fn a1(col: u32, row: u32) -> String {
    let mut letters = String::new();
    let mut n = col + 1;
    while n > 0 {
        n -= 1;
        letters.insert(0, char::from(b'A' + (n % 26) as u8));
        n /= 26;
    }
    format!("{letters}{}", row + 1)
}

fn number(element: Option<&XmlElement>) -> Option<u32> {
    element?.text_content().trim().parse().ok()
}

/// A drawing part's objects in its anchors' order.
fn drawn(drawing: &str, xml: &[u8]) -> Result<Vec<Drawn>> {
    let document = parsed(xml, drawing)?;
    let Some(root) = document.root().filter(|root| root.local_name() == "wsDr") else {
        return Ok(Vec::new());
    };
    let anchors = root.child_elements().filter(|child| {
        matches!(
            child.local_name(),
            "twoCellAnchor" | "oneCellAnchor" | "absoluteAnchor"
        )
    });
    let mut out = Vec::new();
    for (index, anchor) in anchors.enumerate() {
        let cell = anchor.child_by_local_name("from").and_then(|from| {
            Some(a1(
                number(from.child_by_local_name("col"))?,
                number(from.child_by_local_name("row"))?,
            ))
        });
        let mut walk = Walk {
            prefix: format!("{drawing}#{index}"),
            cell,
            pictures: 0,
            chart: false,
            out: &mut out,
        };
        walk.element(anchor, None, 0);
    }
    Ok(out)
}

struct Walk<'a> {
    prefix: String,
    cell: Option<String>,
    pictures: usize,
    chart: bool,
    out: &'a mut Vec<Drawn>,
}

impl Walk<'_> {
    /// `element`'s objects depth first, `pic` the picture it sits in: each
    /// `a:blip` with an embed is the anchor's next picture, as xlsx-parse
    /// counts them; the first `c:chart` its chart.
    fn element(&mut self, element: &XmlElement, pic: Option<&XmlElement>, depth: usize) {
        if depth > MAX_DEPTH {
            return;
        }
        let local = element.local_name();
        if local == "blip" {
            let alt = pic
                .and_then(|pic| pic.child_by_local_name("nvPicPr"))
                .and_then(|nv| nv.child_by_local_name("cNvPr"))
                .and_then(|properties| properties.attribute(None, "descr"))
                .map(str::to_owned);
            let crop = pic
                .and_then(|pic| pic.child_by_local_name("blipFill"))
                .and_then(|fill| fill.child_by_local_name("srcRect"))
                .map_or([0.0; 4], |rect| {
                    ["l", "t", "r", "b"].map(|side| {
                        crop_side(
                            rect.attribute(None, side)
                                .and_then(|value| value.parse().ok())
                                .unwrap_or(0.0),
                        )
                    })
                });
            let embed = element.attribute(Some("r"), "embed").map(str::to_owned);
            let id = embed.as_ref().map(|_| {
                let id = format!("{}:{}", self.prefix, self.pictures);
                self.pictures += 1;
                id
            });
            if embed.is_some() || element.attribute(Some("r"), "link").is_some() {
                self.out.push(Drawn {
                    id,
                    cell: self.cell.clone(),
                    kind: Kind::Picture { embed, alt, crop },
                });
            }
            return;
        }
        if local == "chart"
            && !self.chart
            && let Some(rel) = element.attribute(Some("r"), "id")
        {
            self.chart = true;
            self.out.push(Drawn {
                id: Some(format!("{}:chart", self.prefix)),
                cell: self.cell.clone(),
                kind: Kind::Chart(rel.to_owned()),
            });
            return;
        }
        if local == "sp"
            && let Some(body) = element.child_by_local_name("txBody")
        {
            let text = body
                .children_by_local_name("p")
                .map(|paragraph| paragraph.text_content())
                .collect::<Vec<_>>()
                .join("\n");
            if let Some(text) = objects::shown(&text, MAX_TEXT_CHARS) {
                self.out.push(Drawn {
                    id: None,
                    cell: self.cell.clone(),
                    kind: Kind::Shape(text),
                });
            }
        }
        let pic = if local == "pic" { Some(element) } else { pic };
        for child in element.child_elements() {
            self.element(child, pic, depth + 1);
        }
    }
}

/// Where a read gets a part's bytes, at most the bytes it names.
type Read<'r, 'a> = dyn FnMut(&str, u64) -> Result<Option<Cow<'a, [u8]>>> + 'r;

/// One drawn object as a read finds it, its media or chart part followed.
fn found<'a>(
    object: &Drawn,
    drawing: &str,
    rels: &RelationshipMap,
    read: &mut Read<'_, 'a>,
    media_bytes: u64,
) -> Result<Option<Found<'a>>> {
    Ok(Some(match &object.kind {
        Kind::Picture { embed, alt, crop } => {
            let path = embed
                .as_deref()
                .and_then(|embed| target(rels, drawing, embed));
            let (stored, part) = match path {
                Some(Ok(path)) => match read(&path, media_bytes) {
                    Ok(Some(bytes)) => (Stored::Bytes(bytes), path),
                    Ok(None) => (Stored::Bytes(Cow::Borrowed(&[][..])), path),
                    Err(_) => (Stored::Oversized, path),
                },
                _ => (Stored::Linked, String::new()),
            };
            Found::Picture(Picture {
                id: object.id.clone(),
                alt: alt.clone(),
                stored,
                part,
                crop: *crop,
                cell: object.cell.clone(),
            })
        }
        Kind::Chart(rel) => {
            let chart = match target(rels, drawing, rel) {
                Some(Ok(path)) => read(&path, MAX_XML_BYTES)
                    .ok()
                    .flatten()
                    .and_then(|xml| objects::read_chart(&xml, &path)),
                _ => None,
            };
            Found::Chart(chart)
        }
        Kind::Shape(_) => return Ok(None),
    }))
}

/// A found object's read line, or a drawn shape's text.
fn line(object: &Drawn, found: Option<Found<'_>>) -> Result<Value> {
    let mut line = match found {
        Some(Found::Picture(picture)) => picture.line()?,
        Some(Found::Chart(chart)) => objects::chart_line(object.id.as_deref(), chart.as_ref()),
        None => match &object.kind {
            Kind::Shape(text) => json!({ "kind": "drawing", "text": text }),
            _ => return Err(Error::engine("a drawn object without its kind")),
        },
    };
    if let Some(cell) = &object.cell {
        line["cell"] = json!(cell);
    }
    Ok(line)
}

/// The pictures, charts and drawn shapes' text of the sheet at `sheet_path`
/// in `package`, at most [`MAX_LISTED`], and how many more it holds.
pub(crate) fn sheet_lines(
    package: &xlsx_parse::PreservedPackage,
    sheet_path: &str,
) -> Result<(Vec<Value>, usize)> {
    let mut read = |path: &str, _: u64| Ok(package.part_bytes(path).map(Cow::Borrowed));
    let sheet_rels = relationships(package.part_bytes(&rels_path(sheet_path)), sheet_path)?;
    let mut drawings: Vec<String> = Vec::new();
    for relationship in sheet_rels.values() {
        if relationship.relationship_type.rsplit('/').next() != Some("drawing") {
            continue;
        }
        if let Some(Ok(path)) = target(&sheet_rels, sheet_path, &relationship.id)
            && !drawings.contains(&path)
        {
            drawings.push(path);
        }
    }
    let (mut lines, mut more) = (Vec::new(), 0);
    for drawing in drawings {
        let Some(xml) = package.part_bytes(&drawing) else {
            continue;
        };
        let rels = relationships(package.part_bytes(&rels_path(&drawing)), &drawing)?;
        for object in drawn(&drawing, xml)? {
            if lines.len() == MAX_LISTED {
                more += 1;
                continue;
            }
            let found = found(&object, &drawing, &rels, &mut read, u64::MAX)?;
            lines.push(line(&object, found)?);
        }
    }
    Ok((lines, more))
}

const ID_FORMAT: &str = r#"XLSX picture ids read like "xl/drawings/drawing1.xml#0:0" and chart ids like "xl/drawings/drawing1.xml#1:chart", as the sheet list shows them"#;

/// An id's drawing part and the id itself, when it has the form of one.
fn drawing_of(id: &str) -> Option<&str> {
    let digits = |text: &str| !text.is_empty() && text.bytes().all(|byte| byte.is_ascii_digit());
    let (drawing, place) = id.rsplit_once('#')?;
    let (anchor, which) = place.split_once(':')?;
    (drawing.ends_with(".xml") && digits(anchor) && (which == "chart" || digits(which)))
        .then_some(drawing)
}

/// The drawing parts of the file's sheets, each with whether `state` still
/// holds its sheet: the state's sheet order (`xlsx:sheet-order`) names such
/// a sheet `sheet:<index>`, its place among the file's sheets
/// (betteroffice-xlsx's keys; a sheet added since names none). Read from the
/// state, the workbook part and the sheets' relationships alone.
fn live_drawings(state: &[u8], read: &mut Read<'_, 'static>) -> Result<HashMap<String, bool>> {
    let doc = yrs::Doc::new();
    doc.transact_mut()
        .apply_update(Update::decode_v1(state).map_err(Error::engine)?)
        .map_err(Error::engine)?;
    let txn = doc.transact();
    let order = txn
        .get_array("xlsx:sheet-order")
        .ok_or_else(|| Error::engine("the workbook state has no sheet order"))?;
    let held: HashSet<usize> = order
        .iter(&txn)
        .filter_map(|key| key.cast::<String>().ok())
        .filter_map(|key| key.strip_prefix("sheet:")?.parse().ok())
        .collect();
    const WORKBOOK: &str = "xl/workbook.xml";
    let workbook = read(WORKBOOK, MAX_XML_BYTES)?
        .ok_or_else(|| Error::engine("the package has no xl/workbook.xml"))?;
    let workbook = parsed(&workbook, WORKBOOK)?;
    let workbook_rels = relationships(
        read(&rels_path(WORKBOOK), MAX_XML_BYTES)?.as_deref(),
        WORKBOOK,
    )?;
    let sheets = workbook
        .root()
        .and_then(|root| root.child_by_local_name("sheets"))
        .map(|sheets| sheets.children_by_local_name("sheet").collect::<Vec<_>>())
        .unwrap_or_default();
    let mut drawings = HashMap::new();
    for (index, sheet) in sheets.into_iter().enumerate() {
        // xlsx-parse's `PreservedPackage::capture` places a sheet the same way.
        let path = sheet
            .attribute(Some("r"), "id")
            .and_then(|id| target(&workbook_rels, WORKBOOK, id))
            .and_then(std::result::Result::ok)
            .unwrap_or_else(|| format!("xl/worksheets/sheet{}.xml", index + 1));
        let rels = relationships(read(&rels_path(&path), MAX_XML_BYTES)?.as_deref(), &path)?;
        for relationship in rels.values() {
            if relationship.relationship_type.rsplit('/').next() == Some("drawing")
                && let Some(Ok(drawing)) = target(&rels, &path, &relationship.id)
            {
                *drawings.entry(drawing).or_default() |= held.contains(&index);
            }
        }
    }
    Ok(drawings)
}

fn unknown(id: &str) -> Error {
    Error::edit(
        EditCode::UnavailableTarget,
        format!(
            "{id:?} names no picture or chart of the workbook; read the sheet list for their ids"
        ),
    )
}

/// The by-id read of an XLSX package over `state` ([`objects::answer`]): a
/// picture or chart of a sheet `state` no longer holds is gone.
pub(crate) fn by_ids(base: &[u8], state: &[u8], ids: &[String]) -> Result<Value> {
    let (ids, past) = objects::requested(ids)?;
    let media_bytes = crate::limits::inspect_ids()?.answer_bytes as u64;
    let mut read = |path: &str, most: u64| -> Result<Option<Cow<'static, [u8]>>> {
        ooxml_opc::read_part(base, path, most)
            .map(|bytes| bytes.map(Cow::Owned))
            .map_err(Error::Engine)
    };
    let mut drawings: HashMap<String, (Vec<Drawn>, RelationshipMap)> = HashMap::new();
    let mut live: Option<HashMap<String, bool>> = None;
    let mut answered = Vec::new();
    for id in ids {
        let drawing =
            drawing_of(&id).ok_or_else(|| malformed_id(&id, "picture or chart", ID_FORMAT))?;
        let live = match &live {
            Some(live) => live,
            None => live.insert(live_drawings(state, &mut read)?),
        };
        match live.get(drawing) {
            Some(true) => {}
            Some(false) => {
                return Err(Error::edit(
                    EditCode::UnavailableTarget,
                    format!(
                        "{id} is no longer in the workbook: its sheet was deleted; read the sheet list for the ids"
                    ),
                ));
            }
            None => return Err(unknown(&id)),
        }
        if !drawings.contains_key(drawing) {
            let xml = read(drawing, MAX_XML_BYTES)?;
            let objects = xml
                .map(|xml| drawn(drawing, &xml))
                .transpose()?
                .unwrap_or_default();
            let rels = relationships(
                read(&rels_path(drawing), MAX_XML_BYTES)?.as_deref(),
                drawing,
            )?;
            drawings.insert(drawing.to_owned(), (objects, rels));
        }
        let (objects, rels) = &drawings[drawing];
        let object = objects
            .iter()
            .find(|object| object.id.as_deref() == Some(id.as_str()))
            .ok_or_else(|| unknown(&id))?;
        let found = found(object, drawing, rels, &mut read, media_bytes)?
            .ok_or_else(|| Error::engine("a drawn shape has no id"))?;
        answered.push((id, found));
    }
    objects::answer(answered, &past)
}
