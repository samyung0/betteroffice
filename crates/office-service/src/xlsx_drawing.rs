//! XLSX pictures, charts and drawn shapes, read from a sheet's drawing parts
//! (Epo 2026-10-24, Q240; 2026-10-26, Q381). The sheet list has the
//! workbook open: a picture's or shape's cell is its anchor moved along the
//! rows and columns edited since, as a save moves it; a chart's cell and
//! references are the workbook's (its `SheetChart`, which edits remap) and
//! its points are read from the cells those references name now. The by-id
//! read takes pictures from the package alone, without opening the workbook,
//! whose recalculation costs seconds on a large book (the Editor's call),
//! and leaves their cells to the sheet list; a chart's points are cells, so
//! a call naming one opens it. Of the state the package path reads only the
//! sheet order, so a sheet deleted since the read answers its objects as
//! gone, as the sheet list no longer lists them. A drawing part that does
//! not read is named and skipped, as the workbook skips its charts. A
//! picture's id is the workbook's own (xlsx-parse `sheet_images`):
//! `<drawing>#<anchor>:<n>`, `n` counting the anchor's embedded pictures; a
//! chart's is `<drawing>#<anchor>:chart`.

use std::borrow::Cow;
use std::collections::{HashMap, HashSet};
use std::rc::Rc;

use betteroffice_xlsx::{CellRef, CellValue, SheetChart, SheetId, Workbook, WorkbookModel};
use docx_parse::xml::{ParseBudget, ParseLimits, XmlElement, parse_xml};
use docx_parse::{RelationshipMap, RelationshipTarget, parse_relationships};
use serde_json::{Value, json};
use yrs::updates::decoder::Decode;
use yrs::{Array, ReadTxn, Transact, Update};

use crate::error::{Error, Result};
use crate::objects::{
    self, Chart, Found, Head, LINE_POINTS, MAX_LISTED, MAX_TEXT_CHARS, Missing, Parts, Picture,
    Stored, crop_side,
};

/// Bytes a drawing, relationship or chart part may inflate to.
const MAX_XML_BYTES: u64 = 32 << 20;
/// Nesting a drawing walk follows (xlsx-parse's `MAX_DEPTH`).
const MAX_DEPTH: usize = 64;
/// Cells a by-id read reads of one chart reference (xlsx-parse's
/// `MAX_CACHE_POINTS`, the most one cache carries).
const MAX_REFERENCE_CELLS: u64 = 65_536;

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
    /// Its anchor's place in the drawing.
    anchor: usize,
    /// Its anchor's top-left cell in the file, column and row from 0.
    from: Option<(u32, u32)>,
    /// Whether its anchor keeps its place on a grid edit (`editAs="absolute"`).
    fixed: bool,
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
    let mut n = u64::from(col) + 1;
    while n > 0 {
        n -= 1;
        letters.insert(0, char::from(b'A' + (n % 26) as u8));
        n /= 26;
    }
    format!("{letters}{}", u64::from(row) + 1)
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
        let from = anchor.child_by_local_name("from").and_then(|from| {
            Some((
                number(from.child_by_local_name("col"))?,
                number(from.child_by_local_name("row"))?,
            ))
        });
        let mut walk = Walk {
            prefix: format!("{drawing}#{index}"),
            anchor: index,
            from,
            fixed: anchor.attribute(None, "editAs") == Some("absolute"),
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
    anchor: usize,
    from: Option<(u32, u32)>,
    fixed: bool,
    pictures: usize,
    chart: bool,
    out: &'a mut Vec<Drawn>,
}

impl Walk<'_> {
    fn push(&mut self, id: Option<String>, kind: Kind) {
        self.out.push(Drawn {
            id,
            anchor: self.anchor,
            from: self.from,
            fixed: self.fixed,
            kind,
        });
    }

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
            // A side that is no finite number leaves the picture uncropped
            // there (review r1 L1).
            let crop = pic
                .and_then(|pic| pic.child_by_local_name("blipFill"))
                .and_then(|fill| fill.child_by_local_name("srcRect"))
                .map_or([0.0; 4], |rect| {
                    ["l", "t", "r", "b"].map(|side| {
                        crop_side(
                            rect.attribute(None, side)
                                .and_then(|value| value.parse::<f64>().ok())
                                .filter(|value| value.is_finite())
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
                self.push(id, Kind::Picture { embed, alt, crop });
            }
            return;
        }
        if local == "chart"
            && !self.chart
            && let Some(rel) = element.attribute(Some("r"), "id")
        {
            self.chart = true;
            let id = format!("{}:chart", self.prefix);
            self.push(Some(id), Kind::Chart(rel.to_owned()));
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
                self.push(None, Kind::Shape(text));
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

/// A drawing part's objects and relationships.
type Drawing = Rc<(Vec<Drawn>, RelationshipMap)>;

/// One read's drawing parts and the parts they name, each read once
/// (review r1 M3); a drawing that does not read is `None` (review r1 M1).
#[derive(Default)]
pub(crate) struct Cache {
    parts: Parts,
    drawings: HashMap<String, Option<Drawing>>,
    /// A chart frame's chart as the workbook holds it now.
    charts: HashMap<String, Option<Rc<Chart>>>,
}

impl Cache {
    /// The drawing part at `path`, `None` when it or its relationships do
    /// not read or the package does not hold it.
    fn drawing(&mut self, path: &str, read: &mut Read<'_, '_>) -> Option<Drawing> {
        if let Some(drawing) = self.drawings.get(path) {
            return drawing.clone();
        }
        let mut load = || -> Result<Option<Drawing>> {
            let Some(xml) = read(path, MAX_XML_BYTES)? else {
                return Ok(None);
            };
            let objects = drawn(path, &xml)?;
            let rels = relationships(read(&rels_path(path), MAX_XML_BYTES)?.as_deref(), path)?;
            Ok(Some(Rc::new((objects, rels))))
        };
        let drawing = load().ok().flatten();
        self.drawings.insert(path.to_owned(), drawing.clone());
        drawing
    }
}

/// A picture as a read finds it, its media part followed (at most
/// `media_bytes` of it).
fn picture<'a>(
    object: &Drawn,
    drawing: &str,
    rels: &RelationshipMap,
    read: &mut Read<'_, 'a>,
    media_bytes: u64,
    parts: &mut Parts,
) -> Result<Picture<'a>> {
    let Kind::Picture { embed, alt, crop } = &object.kind else {
        return Err(Error::engine("a drawn object that is no picture"));
    };
    let path = embed
        .as_deref()
        .and_then(|embed| target(rels, drawing, embed));
    let (head, stored) = match path {
        Some(Ok(path)) => match read(&path, media_bytes) {
            Ok(bytes) => {
                let bytes = bytes.unwrap_or(Cow::Borrowed(&[]));
                let head = parts.head(&path, || Head::of(&bytes, bytes.len(), &path));
                (head, Stored::Bytes(bytes))
            }
            Err(_) => (Head::oversized(media_bytes as usize), Stored::Nothing),
        },
        _ => (Head::linked(), Stored::Nothing),
    };
    Ok(Picture {
        id: object.id.clone(),
        alt: alt.clone(),
        head,
        stored,
        crop: *crop,
        cell: None,
    })
}

/// The open workbook a read takes current cells, charts and anchors from.
pub(crate) struct Book<'b> {
    model: &'b WorkbookModel,
    axes: Vec<Option<xlsx_parse::SheetAxes>>,
    package: Option<&'b xlsx_parse::PreservedPackage>,
    /// Each sheet's source part path, none for a sheet added since.
    sheets: Vec<Option<&'b str>>,
}

impl<'b> Book<'b> {
    pub(crate) fn of(workbook: &'b Workbook) -> Result<Self> {
        let model = workbook.model();
        let mut package = None;
        let sheets = (0..model.sheets.len())
            .map(|index| {
                let (source, path) = workbook.source_sheet(SheetId(index as u32))?;
                package = Some(source);
                Some(path)
            })
            .collect();
        Ok(Self {
            model,
            axes: workbook.source_axes().map_err(Error::engine)?,
            package,
            sheets,
        })
    }

    fn read(&self) -> impl FnMut(&str, u64) -> Result<Option<Cow<'b, [u8]>>> + 'b {
        let package = self.package;
        move |path: &str, _: u64| {
            Ok(package
                .and_then(|package| package.part_bytes(path))
                .map(Cow::Borrowed))
        }
    }

    /// `object`'s cell on sheet `sheet` now: its anchor moved along the
    /// rows and columns edited since, as a save moves it.
    fn cell(&self, sheet: usize, object: &Drawn) -> Option<String> {
        let (col, row) = object.from?;
        Some(match self.axes.get(sheet).and_then(Option::as_ref) {
            Some(axes) if !object.fixed => a1(axes.cols.marker(col), axes.rows.marker(row)),
            _ => a1(col, row),
        })
    }

    /// The chart a frame of sheet `sheet` shows now, its points read from
    /// the cells, at most `most` a reference: the workbook's references when
    /// it holds the frame (else the part's own, which no edit remaps), and
    /// the chart part's type, title and literals.
    fn chart(
        &self,
        sheet: usize,
        object: &Drawn,
        drawing: &str,
        rels: &RelationshipMap,
        most: u64,
        cache: &mut Cache,
    ) -> Option<Rc<Chart>> {
        let frame = format!("{drawing}#{}", object.anchor);
        if let Some(chart) = cache.charts.get(&frame) {
            return chart.clone();
        }
        let held = self.held(sheet, drawing, object);
        let Kind::Chart(rel) = &object.kind else {
            return None;
        };
        let part = match held {
            Some(held) => Some(held.part.clone()),
            None => target(rels, drawing, rel).and_then(std::result::Result::ok),
        };
        let chart = part.and_then(|part| {
            let xml = self.package?.part_bytes(&part)?;
            let mut chart = match held {
                Some(held) => {
                    objects::read_chart(&xlsx_parse::with_chart_refs(xml, &held.refs)?, &part)?
                }
                None => (*cache.parts.chart(&part, || Some(xml))?).clone(),
            };
            let owner = &self.model.sheets[sheet].name;
            for series in &mut chart.series {
                read_cells(series, self.model, owner, most);
            }
            Some(Rc::new(chart))
        });
        cache.charts.insert(frame, chart.clone());
        chart
    }

    /// The current sheet whose source names drawing part `drawing`.
    fn sheet_of(&self, drawing: &str) -> Option<usize> {
        let mut read = self.read();
        (0..self.sheets.len()).find(|&sheet| {
            self.sheets[sheet].is_some_and(|path| {
                drawings_of(path, &mut read)
                    .is_ok_and(|drawings| drawings.iter().any(|named| named == drawing))
            })
        })
    }

    /// The workbook's own chart at `object`'s frame of sheet `sheet`.
    fn held(&self, sheet: usize, drawing: &str, object: &Drawn) -> Option<&'b SheetChart> {
        self.model.sheets[sheet]
            .charts
            .iter()
            .find(|chart| chart.drawing == drawing && chart.anchor_index == object.anchor)
    }
}

/// A cell's value as a chart point gives it: a number, text, a truth value
/// or an error's text.
fn point_value(value: &CellValue) -> Option<Value> {
    match value {
        CellValue::Empty => None,
        CellValue::Number { value } if value.is_finite() => Some(json!(value)),
        CellValue::Number { .. } => None,
        CellValue::Text { value } => Some(json!(value)),
        CellValue::Bool { value } => Some(json!(value)),
        CellValue::Error { value } => Some(json!(value.as_str())),
    }
}

/// The cells read of a chart reference, each its shown text and point value,
/// and how many it names.
type Cells = (Vec<(Option<String>, Option<Value>)>, u64);

/// Series `series`'s name and points read from the cells its references
/// name now in `model` (`owner` holding an unqualified one), at most `most`
/// a reference; a literal keeps its own. When a reference does not read (a
/// union, a defined name, another book, a two-dimensional range) the series
/// shows no points and says so, never the file's cached ones (Q381).
fn read_cells(series: &mut objects::Series, model: &WorkbookModel, owner: &str, most: u64) {
    let cells = |formula: &str, limit: u64| -> Option<Cells> {
        let Some((sheet, start, end)) =
            xlsx_parse::chart_reference_area(formula, model, owner).ok()?
        else {
            return Some((Vec::new(), 0));
        };
        let sheet = &model.sheets[sheet];
        let named = u64::from(end.row - start.row + 1) * u64::from(end.col - start.col + 1);
        let read = named.min(limit);
        let cells = (0..read)
            .map(|at| {
                let at = at as u32;
                let place = if end.row > start.row {
                    CellRef::new(start.row + at, start.col)
                } else {
                    CellRef::new(start.row, start.col + at)
                };
                let cell = sheet.cell(place);
                let shown = cell
                    .map(|cell| xlsx_render::display_text(&model.styles, model.date_system, cell))
                    .filter(|text| !text.is_empty());
                (shown, cell.and_then(|cell| point_value(&cell.value)))
            })
            .collect();
        Some((cells, named))
    };
    let unread = |series: &mut objects::Series| {
        series.points.clear();
        series.count = None;
        series.reason = Some(
            "its points could not be read from the cells its references name; read those cells",
        );
    };
    if let Some(formula) = series.refs[0].clone() {
        match cells(&formula, 1) {
            Some((name, _)) => series.name = name.into_iter().next().and_then(|(shown, _)| shown),
            None => return unread(series),
        }
    }
    let [_, categories, values] = series.refs.clone();
    let categories = match categories {
        Some(formula) => match cells(&formula, most) {
            Some(read) => Some(read),
            None => return unread(series),
        },
        None => None,
    };
    let values = match values {
        Some(formula) => match cells(&formula, most) {
            Some(read) => Some(read),
            None => return unread(series),
        },
        None => None,
    };
    if categories.is_none() && values.is_none() {
        return;
    }
    // A literal keeps its points by idx.
    let literal: HashMap<u32, (Option<String>, Option<Value>)> = series
        .points
        .drain(..)
        .map(|(idx, category, value)| (idx, (category, value)))
        .collect();
    let read = |cells: &Option<Cells>| cells.as_ref().map_or(0, |(cells, _)| cells.len());
    let named = |cells: &Option<Cells>| cells.as_ref().map_or(0, |(_, named)| *named);
    let positions = read(&categories).max(read(&values)).max(
        literal
            .keys()
            .map(|idx| *idx as usize + 1)
            .max()
            .unwrap_or(0),
    );
    for idx in 0..positions {
        let category = match &categories {
            Some((cells, _)) => cells.get(idx).and_then(|(shown, _)| shown.clone()),
            None => literal
                .get(&(idx as u32))
                .and_then(|(category, _)| category.clone()),
        };
        let value = match &values {
            Some((cells, _)) => cells.get(idx).and_then(|(_, value)| value.clone()),
            None => literal
                .get(&(idx as u32))
                .and_then(|(_, value)| value.clone()),
        };
        if category.is_some() || value.is_some() {
            series.points.push((idx as u32, category, value));
        }
    }
    if values.is_some() {
        series.count = Some(named(&values));
    }
    series.unread = named(&categories)
        .max(named(&values))
        .saturating_sub(positions as u64);
}

/// The drawing parts sheet part `sheet_path` names.
fn drawings_of(sheet_path: &str, read: &mut Read<'_, '_>) -> Result<Vec<String>> {
    let sheet_rels = relationships(
        read(&rels_path(sheet_path), MAX_XML_BYTES)?.as_deref(),
        sheet_path,
    )?;
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
    Ok(drawings)
}

/// A found object's read line, or a drawn shape's text.
fn line(object: &Drawn, found: Option<Found<'_>>, cell: Option<String>) -> Result<Value> {
    let mut line = match found {
        Some(Found::Picture(picture)) => picture.line()?,
        Some(Found::Chart(chart)) => objects::chart_line(object.id.as_deref(), chart.as_deref()),
        None => match &object.kind {
            Kind::Shape(text) => json!({ "kind": "drawing", "text": text }),
            _ => return Err(Error::engine("a drawn object without its kind")),
        },
    };
    if let Some(cell) = cell {
        line["cell"] = json!(cell);
    }
    Ok(line)
}

/// What the sheet list shows of sheet `sheet`: its pictures, charts and
/// drawn shapes' text, at most [`MAX_LISTED`], how many more it holds, and
/// a notice naming the drawing parts that did not read.
pub(crate) fn sheet_lines(
    book: &Book<'_>,
    sheet: usize,
    cache: &mut Cache,
) -> Result<(Vec<Value>, usize, Option<String>)> {
    let Some(sheet_path) = book.sheets.get(sheet).copied().flatten() else {
        return Ok((Vec::new(), 0, None));
    };
    let mut read = book.read();
    let mut unread = Vec::new();
    let drawings = match drawings_of(sheet_path, &mut read) {
        Ok(drawings) => drawings,
        Err(_) => {
            unread.push(rels_path(sheet_path));
            Vec::new()
        }
    };
    let (mut lines, mut more) = (Vec::new(), 0);
    for drawing in drawings {
        let Some(parsed) = cache.drawing(&drawing, &mut read) else {
            if book
                .package
                .is_some_and(|package| package.part_bytes(&drawing).is_some())
            {
                unread.push(drawing);
            }
            continue;
        };
        let (objects, rels) = &*parsed;
        for object in objects {
            if lines.len() == MAX_LISTED {
                more += 1;
                continue;
            }
            let found = match &object.kind {
                Kind::Picture { .. } => Some(Found::Picture(picture(
                    object,
                    &drawing,
                    rels,
                    &mut read,
                    u64::MAX,
                    &mut cache.parts,
                )?)),
                Kind::Chart(_) => Some(Found::Chart(book.chart(
                    sheet,
                    object,
                    &drawing,
                    rels,
                    LINE_POINTS as u64,
                    cache,
                ))),
                Kind::Shape(_) => None,
            };
            // The workbook moves its charts' anchors itself.
            let cell = match (&object.kind, book.held(sheet, &drawing, object)) {
                (Kind::Chart(_), Some(held)) => {
                    held.anchor.from_cell().map(|from| a1(from.col, from.row))
                }
                _ => book.cell(sheet, object),
            };
            lines.push(line(object, found, cell)?);
        }
    }
    let notice = (!unread.is_empty()).then(|| {
        format!(
            "Not listed: the pictures and charts of {}, a part that could not be read.",
            unread.join(", ")
        )
    });
    Ok((lines, more, notice))
}

const ID_FORMAT: &str = r#"XLSX picture ids read like "xl/drawings/drawing1.xml#0:0" and chart ids like "xl/drawings/drawing1.xml#1:chart", as the sheet list shows them"#;

/// An id's drawing part, when the id has the form of one.
fn drawing_of(id: &str) -> Option<&str> {
    let digits = |text: &str| !text.is_empty() && text.bytes().all(|byte| byte.is_ascii_digit());
    let (drawing, place) = id.rsplit_once('#')?;
    let (anchor, which) = place.split_once(':')?;
    (drawing.ends_with(".xml") && digits(anchor) && (which == "chart" || digits(which)))
        .then_some(drawing)
}

/// Whether the ids a by-id read takes name a chart, whose points need the
/// workbook open.
pub(crate) fn names_a_chart(ids: &[String]) -> bool {
    objects::requested(ids).is_ok_and(|(ids, _)| ids.iter().any(|id| id.ends_with(":chart")))
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
        // xlsx-parse's `PreservedPackage::capture` places a sheet the same
        // way for the files Office writes; it resolves a target with its own
        // `resolve_part_path` and ignores `TargetMode`, so a crafted file can
        // place one differently, which only its author sees (review r1 N4).
        let path = sheet
            .attribute(Some("r"), "id")
            .and_then(|id| target(&workbook_rels, WORKBOOK, id))
            .and_then(std::result::Result::ok)
            .unwrap_or_else(|| format!("xl/worksheets/sheet{}.xml", index + 1));
        // A sheet whose relationships do not read names no drawing here, as
        // its drawings are skipped in the sheet list.
        for drawing in drawings_of(&path, read).unwrap_or_default() {
            *drawings.entry(drawing).or_default() |= held.contains(&index);
        }
    }
    Ok(drawings)
}

/// The by-id read of an XLSX package over `state` ([`objects::answer`]):
/// pictures from the package's parts, charts from `book`, the workbook open
/// over the state, when a chart is asked for ([`names_a_chart`]). A gone,
/// unknown or malformed id, and one of a drawing part that does not read,
/// is named in the notice.
pub(crate) fn by_ids(
    base: &[u8],
    state: &[u8],
    ids: &[String],
    book: Option<&Book<'_>>,
) -> Result<Value> {
    let (ids, past) = objects::requested(ids)?;
    let media_bytes = crate::limits::inspect_ids()?.answer_bytes as u64;
    let mut read = |path: &str, most: u64| -> Result<Option<Cow<'static, [u8]>>> {
        ooxml_opc::read_part(base, path, most)
            .map(|bytes| bytes.map(Cow::Owned))
            .map_err(Error::Engine)
    };
    let mut cache = Cache::default();
    let mut live: Option<HashMap<String, bool>> = None;
    let mut answered = Vec::new();
    let mut missing: Vec<Missing> = Vec::new();
    let gone = "names no picture or chart of the workbook; read the sheet list for their ids";
    for id in ids {
        let Some(drawing) = drawing_of(&id) else {
            missing.push((id.clone(), objects::malformed(ID_FORMAT)));
            continue;
        };
        let live = match &live {
            Some(live) => live,
            None => live.insert(live_drawings(state, &mut read)?),
        };
        match live.get(drawing) {
            Some(true) => {}
            Some(false) => {
                missing.push((
                    id.clone(),
                    "is no longer in the workbook: its sheet was deleted; read the sheet list for the ids".into(),
                ));
                continue;
            }
            None => {
                missing.push((id.clone(), gone.into()));
                continue;
            }
        }
        let Some(parsed) = cache.drawing(drawing, &mut read) else {
            missing.push((
                id.clone(),
                "is in a drawing part that could not be read".into(),
            ));
            continue;
        };
        let (objects, rels) = &*parsed;
        let Some(object) = objects
            .iter()
            .find(|object| object.id.as_deref() == Some(id.as_str()))
        else {
            missing.push((id.clone(), gone.into()));
            continue;
        };
        let found = match &object.kind {
            Kind::Picture { .. } => Found::Picture(picture(
                object,
                drawing,
                rels,
                &mut read,
                media_bytes,
                &mut cache.parts,
            )?),
            Kind::Chart(_) => {
                let book = book.ok_or_else(|| Error::engine("a chart by id needs the workbook"))?;
                let Some(sheet) = book.sheet_of(drawing) else {
                    missing.push((id.clone(), gone.into()));
                    continue;
                };
                Found::Chart(book.chart(
                    sheet,
                    object,
                    drawing,
                    rels,
                    MAX_REFERENCE_CELLS,
                    &mut cache,
                ))
            }
            Kind::Shape(_) => return Err(Error::engine("a drawn shape has no id")),
        };
        answered.push((id, found));
    }
    objects::answer(answered, &missing, &past)
}

#[cfg(test)]
mod tests {
    /// A cell past the grid prints without overflowing (review r1 N3).
    #[test]
    fn a_cell_at_the_largest_index_prints() {
        assert_eq!(super::a1(u32::MAX, u32::MAX), "MWLQKWV4294967296");
        assert_eq!(super::a1(0, 0), "A1");
        assert_eq!(super::a1(26, 9), "AA10");
    }
}
