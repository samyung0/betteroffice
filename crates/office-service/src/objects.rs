//! The pictures and charts the agent's reads list and its by-id read returns
//! (Epo 2026-10-24, Q240, Q278-Q280, Q284). A picture shows its alt text,
//! pixel size, crop and whether the model can see it (EMF, WMF, an SVG
//! without the PNG Office stores beside it and a linked picture cannot); a
//! chart its type, title and series, each with its cell references and
//! points paired by `c:pt/@idx`, read again from the chart part (the shared
//! chart reader ignores idx and drops text points). The by-id read answers a
//! picture's stored bytes and crop, or its alt text when it is not viewable,
//! and a chart's every point, within limits.json's `inspectIds`; what does
//! not fit is named for another call. No image is decoded (Q279): a header
//! gives the pixel size, and the AI process shrinks what it sends.

use std::borrow::Cow;
use std::collections::BTreeMap;

use base64::Engine as _;
use docx_parse::xml::{ParseBudget, ParseLimits, XmlElement, parse_xml};
use serde_json::{Value, json};

use crate::error::{EditCode, Error, Result};

/// Characters of alt text, and of a drawing's text, a read shows.
pub(crate) const MAX_TEXT_CHARS: usize = 1_000;
/// Objects a read lists of one XLSX sheet or DOCX paragraph.
pub(crate) const MAX_LISTED: usize = 50;
/// Series of a chart, and points of a series, a read line shows.
const LINE_SERIES: usize = 8;
const LINE_POINTS: usize = 12;

/// `text` trimmed, cut at `most` characters with `…`; `None` when blank.
pub(crate) fn shown(text: &str, most: usize) -> Option<String> {
    let text = text.trim();
    if text.is_empty() {
        return None;
    }
    Some(match text.char_indices().nth(most) {
        Some((at, _)) => format!("{}…", &text[..at]),
        None => text.to_owned(),
    })
}

/// A picture's stored bytes as far as a read holds them.
pub(crate) enum Stored<'a> {
    Bytes(Cow<'a, [u8]>),
    /// Linked to a file outside this one: nothing is stored.
    Linked,
    /// Past what one by-id answer can carry, left unread.
    Oversized,
}

pub(crate) struct Picture<'a> {
    pub id: Option<String>,
    pub alt: Option<String>,
    pub stored: Stored<'a>,
    /// Its part's name, for a format its bytes do not tell.
    pub part: String,
    /// Fractions of the image its frame leaves out: left, top, right, bottom
    /// (negative ones pad it).
    pub crop: [f64; 4],
    /// The cell an XLSX picture is anchored at.
    pub cell: Option<String>,
}

/// An OOXML `a:srcRect` side in thousandths of a percent, as a fraction.
pub(crate) fn crop_side(value: f64) -> f64 {
    value / 100_000.0
}

/// base64's length of `bytes` bytes.
fn encoded_len(bytes: usize) -> usize {
    bytes.div_ceil(3) * 4
}

/// What the model reads (PNG, JPEG, GIF, BMP, WebP, TIFF; the AI process
/// converts what it sends): the type and pixel size its header gives, else
/// why it is not viewable.
fn raster(
    bytes: &[u8],
    part: &str,
) -> std::result::Result<(&'static str, Option<[u32; 2]>), &'static str> {
    let at = |from: usize, len: usize| bytes.get(from..from + len);
    let le16 = |from| at(from, 2).map(|b| u32::from(u16::from_le_bytes([b[0], b[1]])));
    let be32 = |from| at(from, 4).map(|b| u32::from_be_bytes([b[0], b[1], b[2], b[3]]));
    let le32 = |from| at(from, 4).map(|b| u32::from_le_bytes([b[0], b[1], b[2], b[3]]));
    let le24 = |from| at(from, 3).map(|b| u32::from_le_bytes([b[0], b[1], b[2], 0]));
    let pair = |w: Option<u32>, h: Option<u32>| w.zip(h).map(|(w, h)| [w, h]);
    if bytes.starts_with(b"\x89PNG\r\n\x1a\n") {
        return Ok(("image/png", pair(be32(16), be32(20))));
    }
    if bytes.starts_with(&[0xff, 0xd8]) {
        return Ok(("image/jpeg", jpeg_size(bytes)));
    }
    if bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a") {
        return Ok(("image/gif", pair(le16(6), le16(8))));
    }
    if bytes.starts_with(b"BM") {
        let signed = |from| le32(from).map(|value| (value as i32).unsigned_abs());
        return Ok(("image/bmp", pair(signed(18), signed(22))));
    }
    if bytes.starts_with(b"RIFF") && at(8, 4) == Some(b"WEBP") {
        let size = match at(12, 4) {
            Some(b"VP8 ") => pair(le16(26).map(|w| w & 0x3fff), le16(28).map(|h| h & 0x3fff)),
            Some(b"VP8L") => {
                le32(21).map(|bits| [(bits & 0x3fff) + 1, ((bits >> 14) & 0x3fff) + 1])
            }
            Some(b"VP8X") => pair(le24(24).map(|w| w + 1), le24(27).map(|h| h + 1)),
            _ => None,
        };
        return Ok(("image/webp", size));
    }
    if bytes.starts_with(b"II*\0") || bytes.starts_with(b"MM\0*") {
        return Ok(("image/tiff", None));
    }
    let extension = part
        .rsplit('.')
        .next()
        .unwrap_or_default()
        .to_ascii_lowercase();
    let emf = at(40, 4) == Some(b" EMF");
    let wmf = bytes.starts_with(&[0xd7, 0xcd, 0xc6, 0x9a])
        || bytes.starts_with(&[1, 0, 9, 0, 0, 3])
        || bytes.starts_with(&[2, 0, 9, 0, 0, 3]);
    if emf || wmf || matches!(extension.as_str(), "emf" | "wmf" | "emz" | "wmz") {
        return Err("an EMF or WMF drawing");
    }
    let head = String::from_utf8_lossy(&bytes[..bytes.len().min(1024)]).to_ascii_lowercase();
    if head.contains("<svg") || extension == "svg" {
        return Err("an SVG without the PNG copy Office stores beside it");
    }
    Err("damaged or unknown picture data")
}

/// A JPEG's size from its first frame header (SOF0 to SOF15 but DHT, JPG
/// and DAC), walking the segments before it.
fn jpeg_size(bytes: &[u8]) -> Option<[u32; 2]> {
    let be16 = |at: usize| {
        bytes
            .get(at..at + 2)
            .map(|b| u32::from(u16::from_be_bytes([b[0], b[1]])))
    };
    let mut at = 2;
    loop {
        while bytes.get(at) == Some(&0xff) && bytes.get(at + 1) == Some(&0xff) {
            at += 1;
        }
        if *bytes.get(at)? != 0xff {
            return None;
        }
        let marker = *bytes.get(at + 1)?;
        if matches!(marker, 0x01 | 0xd0..=0xd7) {
            at += 2;
            continue;
        }
        if matches!(marker, 0xd9 | 0xda) {
            return None;
        }
        if matches!(marker, 0xc0..=0xcf) && !matches!(marker, 0xc4 | 0xc8 | 0xcc) {
            return Some([be16(at + 7)?, be16(at + 5)?]);
        }
        at += 2 + be16(at + 2)? as usize;
    }
}

/// A viewable picture: its type, pixel size when its header gives one, and
/// bytes.
type Seen<'b> = (&'static str, Option<[u32; 2]>, &'b [u8]);

impl Picture<'_> {
    /// What the model gets of it, else why it cannot see it.
    fn seen(&self, answer_bytes: usize) -> std::result::Result<Seen<'_>, String> {
        let bytes = match &self.stored {
            Stored::Bytes(bytes) => bytes.as_ref(),
            Stored::Linked => return Err("linked to a file outside this one".into()),
            Stored::Oversized => return Err(too_large(answer_bytes)),
        };
        let (mime, pixels) = raster(bytes, &self.part).map_err(str::to_owned)?;
        if encoded_len(bytes.len()) + ITEM_RESERVE > answer_bytes {
            return Err(too_large(answer_bytes));
        }
        Ok((mime, pixels, bytes))
    }

    /// Its read line: id, alt text, pixel size, crop, anchor cell, and
    /// whether it is viewable (the reason when not).
    pub(crate) fn line(&self) -> Result<Value> {
        let limits = crate::limits::inspect_ids()?;
        let mut line = json!({ "kind": "picture" });
        if let Some(id) = &self.id {
            line["id"] = json!(id);
        }
        self.describe(&mut line, limits.answer_bytes);
        Ok(line)
    }

    fn describe(&self, line: &mut Value, answer_bytes: usize) {
        if let Some(alt) = self
            .alt
            .as_deref()
            .and_then(|alt| shown(alt, MAX_TEXT_CHARS))
        {
            line["alt"] = json!(alt);
        }
        if let Some(cell) = &self.cell {
            line["cell"] = json!(cell);
        }
        match self.seen(answer_bytes) {
            Ok((_, pixels, _)) => {
                line["viewable"] = json!(true);
                if let Some(pixels) = pixels {
                    line["pixels"] = json!(pixels);
                }
                if self.crop.iter().any(|side| *side != 0.0) {
                    let side = |value: f64| (value * 10_000.0).round() / 10_000.0;
                    line["crop"] = json!({
                        "left": side(self.crop[0]),
                        "top": side(self.crop[1]),
                        "right": side(self.crop[2]),
                        "bottom": side(self.crop[3]),
                    });
                }
            }
            Err(reason) => {
                line["viewable"] = json!(false);
                line["reason"] = json!(reason);
            }
        }
    }
}

fn too_large(answer_bytes: usize) -> String {
    format!(
        "too large to send: a read carries at most {} MB",
        encoded_capacity(answer_bytes) / 1_000_000
    )
}

/// The stored bytes the base64 of one answer can carry.
fn encoded_capacity(answer_bytes: usize) -> usize {
    (answer_bytes - ITEM_RESERVE) / 4 * 3
}

/// JSON an answer holds besides one picture's base64: its frame, notice and
/// the picture's other fields (alt text of up to [`MAX_TEXT_CHARS`]).
const ITEM_RESERVE: usize = 16_384;

/// A chart part as the reads show it.
#[derive(Clone, Debug, Default, PartialEq)]
pub(crate) struct Chart {
    /// Its plot types (`column`, `line`, `pie` …), joined for a combination.
    pub kind: String,
    pub title: Option<String>,
    pub series: Vec<Series>,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub(crate) struct Series {
    pub name: Option<String>,
    /// The cell references of its name, categories (or x values) and values.
    pub refs: [Option<String>; 3],
    /// Its values cache's declared point count.
    pub count: Option<u64>,
    /// `(idx, category, value)` in idx order: each idx either cache holds.
    pub points: Vec<(u32, Option<String>, Option<Value>)>,
}

fn text_of(element: &XmlElement) -> String {
    element.text_content().trim().to_owned()
}

/// A chart part's plot types, title and series; `None` when it is no
/// `c:chartSpace` with a plot this reads.
pub(crate) fn read_chart(xml: &[u8], part: &str) -> Option<Chart> {
    let limits = ParseLimits::default();
    let document = parse_xml(xml, part, &mut ParseBudget::new(&limits)).ok()?;
    let root = document
        .root()
        .filter(|root| root.local_name() == "chartSpace")?;
    let chart = root.child_by_local_name("chart")?;
    let plot = chart.child_by_local_name("plotArea")?;
    let mut kinds: Vec<String> = Vec::new();
    let mut series = Vec::new();
    for group in plot.child_elements() {
        let Some(kind) = group.local_name().strip_suffix("Chart") else {
            continue;
        };
        let column = group
            .child_by_local_name("barDir")
            .and_then(|dir| dir.attribute(None, "val"))
            != Some("bar");
        let kind = match kind.strip_prefix("bar") {
            Some(rest) if column => format!("column{rest}"),
            _ => kind.to_owned(),
        };
        if !kinds.contains(&kind) {
            kinds.push(kind);
        }
        series.extend(group.children_by_local_name("ser").map(read_series));
    }
    if kinds.is_empty() {
        return None;
    }
    let title = chart.child_by_local_name("title").and_then(|title| {
        let tx = title.child_by_local_name("tx")?;
        let text = match tx.child_by_local_name("rich") {
            Some(rich) => rich
                .children_by_local_name("p")
                .map(text_of)
                .filter(|line| !line.is_empty())
                .collect::<Vec<_>>()
                .join(" "),
            None => data(Some(tx)).1.into_values().next().unwrap_or_default(),
        };
        (!text.is_empty()).then_some(text)
    });
    Some(Chart {
        kind: kinds.join(", "),
        title,
        series,
    })
}

fn read_series(ser: &XmlElement) -> Series {
    let field = |names: &[&str]| names.iter().find_map(|name| ser.child_by_local_name(name));
    let tx = ser.child_by_local_name("tx");
    let (name_ref, names) = data(tx);
    let name = names
        .into_values()
        .next()
        .or_else(|| tx.and_then(|tx| tx.child_by_local_name("v")).map(text_of));
    let (categories_ref, categories) = data(field(&["cat", "xVal"]));
    let values_element = field(&["val", "yVal"]);
    let (values_ref, values) = data(values_element);
    let count = values_element
        .and_then(|values| values.child_elements().next())
        .and_then(cache)
        .and_then(|cache| cache.child_by_local_name("ptCount"))
        .and_then(|count| count.attribute(None, "val"))
        .and_then(|count| count.trim().parse().ok());
    let mut paired: BTreeMap<u32, (Option<String>, Option<Value>)> = BTreeMap::new();
    for (idx, category) in categories {
        paired.entry(idx).or_default().0 = Some(category);
    }
    for (idx, value) in values {
        let number = value
            .parse::<f64>()
            .ok()
            .filter(|number| number.is_finite());
        paired.entry(idx).or_default().1 = Some(number.map_or_else(|| json!(value), |n| json!(n)));
    }
    Series {
        name: name.filter(|name| !name.is_empty()),
        refs: [name_ref, categories_ref, values_ref],
        count,
        points: paired
            .into_iter()
            .map(|(idx, (category, value))| (idx, category, value))
            .collect(),
    }
}

/// The cache of a `c:strRef`, `c:numRef` or `c:multiLvlStrRef`, or a
/// literal's own points.
fn cache(source: &XmlElement) -> Option<&XmlElement> {
    match source.local_name() {
        "strLit" | "numLit" => Some(source),
        _ => source
            .child_elements()
            .find(|child| child.local_name().ends_with("Cache")),
    }
}

/// A `c:tx`, `c:cat` or `c:val`'s cell reference and cached points by idx;
/// a multi-level category joins its levels, outermost first.
fn data(element: Option<&XmlElement>) -> (Option<String>, BTreeMap<u32, String>) {
    let mut points = BTreeMap::new();
    let Some(source) = element.and_then(|element| {
        element.child_elements().find(|child| {
            child.local_name().ends_with("Ref") || child.local_name().ends_with("Lit")
        })
    }) else {
        return (None, points);
    };
    let reference = source
        .child_by_local_name("f")
        .map(text_of)
        .filter(|f| !f.is_empty());
    let Some(cache) = cache(source) else {
        return (reference, points);
    };
    let read = |parent: &XmlElement, points: &mut BTreeMap<u32, String>| {
        for point in parent.children_by_local_name("pt") {
            let Some(idx) = point
                .attribute(None, "idx")
                .filter(|idx| !idx.is_empty() && idx.bytes().all(|byte| byte.is_ascii_digit()))
                .and_then(|idx| idx.parse::<u32>().ok())
            else {
                continue;
            };
            let value = point
                .child_by_local_name("v")
                .map(text_of)
                .unwrap_or_default();
            points.insert(idx, value);
        }
    };
    if cache.local_name() == "multiLvlStrCache" {
        let mut levels: Vec<BTreeMap<u32, String>> = Vec::new();
        for level in cache.children_by_local_name("lvl") {
            let mut level_points = BTreeMap::new();
            read(level, &mut level_points);
            levels.push(level_points);
        }
        let idxs: std::collections::BTreeSet<u32> = levels
            .iter()
            .flat_map(|level| level.keys().copied())
            .collect();
        for idx in idxs {
            let joined: Vec<&str> = levels
                .iter()
                .rev()
                .filter_map(|level| level.get(&idx).map(String::as_str))
                .filter(|text| !text.is_empty())
                .collect();
            points.insert(idx, joined.join(" / "));
        }
    } else {
        read(cache, &mut points);
    }
    (reference, points)
}

fn point_json((idx, category, value): &(u32, Option<String>, Option<Value>)) -> Value {
    json!([idx, category, value])
}

impl Series {
    fn head(&self) -> Value {
        let mut series = json!({});
        if let Some(name) = &self.name {
            series["name"] = json!(name);
        }
        let mut refs = json!({});
        for (key, reference) in ["name", "categories", "values"].iter().zip(&self.refs) {
            if let Some(reference) = reference {
                refs[*key] = json!(reference);
            }
        }
        if refs.as_object().is_some_and(|refs| !refs.is_empty()) {
            series["refs"] = refs;
        }
        if let Some(count) = self.count {
            series["count"] = json!(count);
        }
        series
    }
}

impl Chart {
    fn head(&self, line: &mut Value) {
        line["chartType"] = json!(self.kind);
        if let Some(title) = &self.title {
            line["title"] = json!(title);
        }
    }

    /// Its read line's fields: type, title, the first [`LINE_SERIES`]
    /// series with their first [`LINE_POINTS`] points each, and how many of
    /// both it leaves out.
    pub(crate) fn describe(&self, line: &mut Value) {
        self.head(line);
        line["series"] = self
            .series
            .iter()
            .take(LINE_SERIES)
            .map(|series| {
                let mut entry = series.head();
                entry["points"] = series
                    .points
                    .iter()
                    .take(LINE_POINTS)
                    .map(point_json)
                    .collect();
                if series.points.len() > LINE_POINTS {
                    entry["more"] = json!(series.points.len() - LINE_POINTS);
                }
                entry
            })
            .collect();
        if self.series.len() > LINE_SERIES {
            line["moreSeries"] = json!(self.series.len() - LINE_SERIES);
        }
    }
}

/// A chart's read line, or what it shows when its part does not read.
pub(crate) fn chart_line(id: Option<&str>, chart: Option<&Chart>) -> Value {
    let mut line = json!({ "kind": "chart" });
    if let Some(id) = id {
        line["id"] = json!(id);
    }
    match chart {
        Some(chart) => chart.describe(&mut line),
        None => line["reason"] = json!(UNREAD_CHART),
    }
    line
}

const UNREAD_CHART: &str = "its chart data could not be read";

/// What a by-id read found for one id.
pub(crate) enum Found<'a> {
    Picture(Picture<'a>),
    /// `None` when the chart part does not read.
    Chart(Option<Chart>),
}

/// The by-id read's ids: each trimmed of spaces, quotes and the brackets a
/// read shows around ids (a clear repair, Q277), repeats dropped; then those
/// limits.json's per-call count takes, and the rest.
pub(crate) fn requested(ids: &[String]) -> Result<(Vec<String>, Vec<String>)> {
    let mut clean: Vec<String> = Vec::new();
    for id in ids {
        let id = id
            .trim()
            .trim_matches(|c: char| matches!(c, '[' | ']' | '"' | '\'' | '`'))
            .trim();
        if !id.is_empty() && !clean.iter().any(|known| known == id) {
            clean.push(id.to_owned());
        }
    }
    if clean.is_empty() {
        return Err(Error::edit(
            EditCode::InvalidInput,
            "ids names the pictures and charts to read, each id copied from an earlier read",
        ));
    }
    let most = crate::limits::inspect_ids()?.per_call;
    let rest = clean.split_off(clean.len().min(most));
    Ok((clean, rest))
}

/// The by-id answer: `items` in the order asked, each picture's bytes
/// (`data`, base64) or, when not viewable, its alt text and why, each
/// chart's every point; within limits.json's `answerBytes`, a picture that
/// does not fit and the points of a series past it are left for another
/// call, and `notice` names them with the ids past the per-call count.
pub(crate) fn answer(found: Vec<(String, Found<'_>)>, past_limit: &[String]) -> Result<Value> {
    let limits = crate::limits::inspect_ids()?;
    let names: usize = found
        .iter()
        .map(|(id, _)| id.len())
        .chain(past_limit.iter().map(String::len))
        .map(|len| len + 4)
        .sum();
    // The answer's frame and its notice, whatever it names.
    let mut used = 400 + names;
    let budget = limits.answer_bytes;
    let size = |value: &Value| serde_json::to_vec(value).map(|bytes| bytes.len() + 1);
    let mut items = Vec::new();
    let mut later = Vec::new();
    let mut cut_charts = false;
    for (id, found) in found {
        let mut item = json!({ "id": id });
        match found {
            Found::Picture(picture) => {
                item["kind"] = json!("picture");
                picture.describe(&mut item, budget);
                let mut cost = size(&item).map_err(Error::engine)?;
                let seen = picture.seen(budget).ok();
                if let Some((mime, _, bytes)) = seen {
                    item["mimeType"] = json!(mime);
                    cost +=
                        r#","mimeType":"","data":"""#.len() + mime.len() + encoded_len(bytes.len());
                }
                if used + cost > budget {
                    later.push(id);
                    continue;
                }
                used += cost;
                if let Some((_, _, bytes)) = seen {
                    item["data"] = json!(base64::engine::general_purpose::STANDARD.encode(bytes));
                }
            }
            Found::Chart(chart) => {
                item["kind"] = json!("chart");
                let Some(chart) = chart else {
                    item["reason"] = json!(UNREAD_CHART);
                    used += size(&item).map_err(Error::engine)?;
                    items.push(item);
                    continue;
                };
                chart.head(&mut item);
                let head = size(&item).map_err(Error::engine)?
                    + r#","series":[],"seriesCut":4294967295"#.len();
                if used + head > budget {
                    later.push(id);
                    continue;
                }
                used += head;
                let mut listed = Vec::new();
                let mut series_cut = 0;
                for series in &chart.series {
                    let mut entry = series.head();
                    let head = size(&entry).map_err(Error::engine)?
                        + r#","points":[],"cut":4294967295"#.len();
                    if used + head > budget {
                        series_cut += 1;
                        continue;
                    }
                    used += head;
                    let mut points = Vec::new();
                    for point in &series.points {
                        let point = point_json(point);
                        let cost = size(&point).map_err(Error::engine)?;
                        if used + cost > budget {
                            break;
                        }
                        used += cost;
                        points.push(point);
                    }
                    if points.len() < series.points.len() {
                        entry["cut"] = json!(series.points.len() - points.len());
                    }
                    entry["points"] = Value::Array(points);
                    listed.push(entry);
                }
                cut_charts |=
                    series_cut > 0 || listed.iter().any(|entry| entry.get("cut").is_some());
                item["series"] = Value::Array(listed);
                if series_cut > 0 {
                    item["seriesCut"] = json!(series_cut);
                }
            }
        }
        items.push(item);
    }
    let mut notices = Vec::new();
    if !later.is_empty() {
        notices.push(format!(
            "Not sent: {}, past the {} MB one read carries; read them in another call.",
            later.join(", "),
            budget / 1_000_000
        ));
    }
    if cut_charts {
        notices.push(format!(
            "The points a series' cut counts, and the series seriesCut counts, did not fit the {} MB one read carries; the cells their references name hold them.",
            budget / 1_000_000
        ));
    }
    if !past_limit.is_empty() {
        notices.push(format!(
            "Not read: {}, past the limit of {} ids a call; read them in another call.",
            past_limit.join(", "),
            limits.per_call
        ));
    }
    let mut answer = json!({ "items": items });
    if !notices.is_empty() {
        answer["notice"] = json!(notices.join(" "));
    }
    Ok(answer)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn png(width: u32, height: u32) -> Vec<u8> {
        let mut bytes = b"\x89PNG\r\n\x1a\n\0\0\0\x0dIHDR".to_vec();
        bytes.extend(width.to_be_bytes());
        bytes.extend(height.to_be_bytes());
        bytes.extend([8, 6, 0, 0, 0]);
        bytes
    }

    #[test]
    fn headers_give_type_and_pixel_size() {
        assert_eq!(
            raster(&png(640, 480), "a.png"),
            Ok(("image/png", Some([640, 480])))
        );
        // SOI, an APP0 segment, then SOF0: height 300, width 400.
        let jpeg = [
            0xff, 0xd8, 0xff, 0xe0, 0, 4, 0, 0, 0xff, 0xc0, 0, 11, 8, 1, 44, 1, 144, 3, 1, 0x11, 0,
        ];
        assert_eq!(raster(&jpeg, "a.jpg"), Ok(("image/jpeg", Some([400, 300]))));
        assert_eq!(
            raster(b"GIF89a\x20\x00\x10\x00", "a.gif"),
            Ok(("image/gif", Some([32, 16])))
        );
        let mut emf = vec![1, 0, 0, 0];
        emf.resize(40, 0);
        emf.extend(b" EMF");
        assert_eq!(raster(&emf, "image1.bin"), Err("an EMF or WMF drawing"));
        assert_eq!(
            raster(&[0xd7, 0xcd, 0xc6, 0x9a, 0, 0], "x"),
            Err("an EMF or WMF drawing")
        );
        assert_eq!(
            raster(
                b"<?xml version=\"1.0\"?><svg xmlns=\"http://www.w3.org/2000/svg\"/>",
                "a.bin"
            ),
            Err("an SVG without the PNG copy Office stores beside it")
        );
        assert_eq!(
            raster(b"garbage", "a.png"),
            Err("damaged or unknown picture data")
        );
        // A truncated header names no size and never reads past the bytes.
        assert_eq!(raster(&png(1, 1)[..18], "a.png"), Ok(("image/png", None)));
        assert_eq!(raster(&jpeg[..12], "a.jpg"), Ok(("image/jpeg", None)));
    }

    fn picture(bytes: Vec<u8>) -> Found<'static> {
        Found::Picture(Picture {
            id: Some("p".into()),
            alt: Some("A photo".into()),
            stored: Stored::Bytes(Cow::Owned(bytes)),
            part: "a.png".into(),
            crop: [0.0; 4],
            cell: None,
        })
    }

    /// Security regression (limits.json `inspectIds.answerBytes`): a
    /// picture past what one answer carries answers its alt text, and a
    /// chart after a picture that fills most of the answer is cut with the
    /// count of the points left out; the answer stays within the bytes.
    #[test]
    fn the_answer_bytes_cut_a_series_and_refuse_a_picture_too_large() {
        let mut big = png(4000, 3000);
        big.resize(6_200_000, 0);
        let mut too_large = png(8000, 6000);
        too_large.resize(6_300_000, 0);
        let chart = Chart {
            kind: "line".into(),
            title: None,
            series: vec![Series {
                name: Some("Daily".into()),
                refs: [None, None, Some("Sheet1!$B$2:$B$10001".into())],
                count: Some(10_000),
                points: (0..10_000)
                    .map(|idx| (idx, Some(format!("D{idx}")), Some(json!(idx))))
                    .collect(),
            }],
        };
        let answer = answer(
            vec![
                ("p".into(), picture(big)),
                ("q".into(), picture(too_large)),
                ("c".into(), Found::Chart(Some(chart))),
            ],
            &[],
        )
        .unwrap();
        let items = answer["items"].as_array().unwrap();
        assert!(items[0]["data"].is_string());
        assert_eq!(
            (&items[1]["viewable"], &items[1]["reason"], &items[1]["alt"]),
            (
                &json!(false),
                &json!("too large to send: a read carries at most 6 MB"),
                &json!("A photo")
            )
        );
        let series = &items[2]["series"][0];
        let sent = series["points"].as_array().unwrap().len();
        assert!(sent > 0 && sent < 10_000, "{sent}");
        assert_eq!(series["cut"], json!(10_000 - sent));
        assert!(answer["notice"].as_str().unwrap().contains("cut"));
        let limit = crate::limits::inspect_ids().unwrap().answer_bytes;
        assert!(serde_json::to_vec(&answer).unwrap().len() <= limit);
    }

    const SPARSE: &str = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx><c:cat><c:strRef><c:f>Sheet1!$A$2:$A$6</c:f><c:strCache><c:ptCount val="5"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="3"><c:v>Q4</c:v></c:pt><c:pt idx="4"><c:v>Q5</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$B$2:$B$6</c:f><c:numCache><c:ptCount val="5"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="2"><c:v>30.5</c:v></c:pt><c:pt idx="4"><c:v>#N/A</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#;

    /// A sparse cache pairs categories and values by `c:pt/@idx` and keeps
    /// a point that is not a number (the shared reader packs the values and
    /// drops it).
    #[test]
    fn a_sparse_cache_pairs_points_by_idx() {
        let chart = read_chart(SPARSE.as_bytes(), "chart1.xml").unwrap();
        assert_eq!(
            (chart.kind.as_str(), chart.title.as_deref()),
            ("column", Some("Sales"))
        );
        let series = &chart.series[0];
        assert_eq!(series.name.as_deref(), Some("Revenue"));
        assert_eq!(series.count, Some(5));
        assert_eq!(
            series.refs,
            [
                Some("Sheet1!$B$1".into()),
                Some("Sheet1!$A$2:$A$6".into()),
                Some("Sheet1!$B$2:$B$6".into())
            ]
        );
        let points: Vec<Value> = series.points.iter().map(point_json).collect();
        assert_eq!(
            points,
            [
                json!([0, "Q1", 10.0]),
                json!([1, "Q2", null]),
                json!([2, null, 30.5]),
                json!([3, "Q4", null]),
                json!([4, "Q5", "#N/A"]),
            ]
        );
    }
}
