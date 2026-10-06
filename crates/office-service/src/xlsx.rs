//! XLSX sessions over `betteroffice_xlsx::Workbook`, as office-checkpoint.ts
//! drives `XlsxDocument` and `XlsxEffectsReader`.

use betteroffice_xlsx::{CalculationOptions, CellRef, SheetId, Workbook};
use serde::Serialize;

use crate::common::{assert_checkpoint, image_mime_type, sha256_hex};
use crate::env;
use crate::error::{EditCode, Error, Result};
use crate::js::{J, canonical};
use crate::session::{Applied, Item};
use crate::types::{
    Asset, Checkpoint, Command, EffectKind, Entry, Format, NetEffect, ObjectKind, ObjectRef, Target,
};

/// Days from the 1900 epoch to 1970-01-01, phantom leap day included.
const UNIX_EPOCH_SERIAL: f64 = 25569.0;
const MS_PER_DAY: f64 = 86_400_000.0;

pub(crate) fn now_options() -> CalculationOptions {
    CalculationOptions {
        now_serial: Some(env::now_ms() / MS_PER_DAY + UNIX_EPOCH_SERIAL),
    }
}

pub(crate) struct XlsxSession {
    workbook: Workbook,
}

impl XlsxSession {
    pub fn open(base: &[u8], state: Option<&[u8]>) -> Result<Self> {
        let mut workbook =
            Workbook::open_collaborative_recalculated(base, env::next_client(), now_options())
                .map_err(Error::engine)?;
        if let Some(state) = state {
            workbook
                .apply_update_v1(state, now_options())
                .map_err(Error::engine)?;
            workbook.sheet_info().map_err(Error::engine)?;
        }
        Ok(Self { workbook })
    }

    pub fn state(&self) -> Vec<u8> {
        self.workbook.encode_state_as_update_v1()
    }

    /// `checkpointProjectionJson` (`full`) or `checkpointCellsJson`.
    fn projection(&self, full: bool) -> Result<J> {
        J::parse(&projection_json(&self.workbook, full)?).map_err(Error::Engine)
    }

    pub fn entries(&self) -> Result<Vec<Item>> {
        let projection = self.projection(true)?;
        let mut entries = Vec::new();
        for (index, sheet) in array(projection.get("sheets")).iter().enumerate() {
            let object = sheet.as_obj().cloned().unwrap_or_default();
            let id = text(sheet.get("id"));
            let name = text(sheet.get("name"));
            entries.push(Item::text(
                &id,
                format!("Sheet {name}"),
                &name,
                index.to_string(),
            ));
            let mut layout = object.clone();
            for key in ["cells", "images", "id", "name"] {
                layout.remove(key);
            }
            entries.push(Item::visual(
                format!("{id}:layout"),
                format!("{name} layout"),
                &J::Obj(layout.clone()),
                "",
            ));
            if let Some(J::Arr(links)) = layout.get("hyperlinks") {
                for (link_index, link) in links.iter().enumerate() {
                    entries.push(Item::text(
                        format!("{id}:link:{link_index}"),
                        format!("{name}, hyperlink"),
                        canonical(link),
                        "",
                    ));
                }
            }
            for image in array(sheet.get("images")) {
                let bytes: Vec<u8> = array(image.get("bytes"))
                    .iter()
                    .map(|byte| byte.as_f64().unwrap_or_default() as u8)
                    .collect();
                let part = text(image.get("part"));
                let image_id = text(image.get("id"));
                let asset = Asset {
                    sha256: sha256_hex(&bytes),
                    mime_type: image_mime_type(&part)?.to_owned(),
                    bytes,
                };
                entries.push(Item {
                    id: format!("{id}:image:{image_id}"),
                    kind: EffectKind::Image,
                    label: format!("{name}, image"),
                    value: asset.sha256.clone(),
                    position: canonical(image.get("anchor").unwrap_or(&J::Null)),
                    asset_ref: Some(ObjectRef {
                        format: Format::Xlsx,
                        kind: ObjectKind::Image,
                        id: image_id,
                        story_id: None,
                        sheet_id: Some(id.clone()),
                        slide_id: None,
                    }),
                    asset: Some(asset),
                });
            }
            for cell in array(sheet.get("cells")) {
                let cell_id = text(cell.get("id"));
                let address = text(cell.get("address"));
                let value = match cell.get("formula") {
                    Some(J::Null) | None => {
                        let value = cell.get("value").unwrap_or(&J::Null);
                        if value.get("kind").and_then(J::as_str) == Some("empty") {
                            String::new()
                        } else {
                            canonical(value)
                        }
                    }
                    Some(formula) => format!("={}", js_string(formula)),
                };
                if !value.is_empty() {
                    entries.push(Item::text(
                        &cell_id,
                        format!("{name}!{address}"),
                        value,
                        format!("{id}:{address}"),
                    ));
                }
                entries.push(Item::visual(
                    format!("{cell_id}:format"),
                    format!("{name}!{address} formatting"),
                    cell.get("format").unwrap_or(&J::Null),
                    "",
                ));
            }
        }
        for item in array(projection.get("definedNames")) {
            let name = text(item.get("name"));
            let local = item.get("local_sheet").unwrap_or(&J::Null);
            entries.push(Item::text(
                format!("name:{}:{name}", js_string(local)),
                format!("Defined name {name}"),
                text(item.get("formula")),
                "",
            ));
        }
        Ok(entries)
    }

    pub fn editable(&self) -> Result<Vec<Entry>> {
        let cells = self.projection(false)?;
        let mut entries = Vec::new();
        for sheet in array(cells.get("sheets")) {
            let (sheet_id, name) = (text(sheet.get("id")), text(sheet.get("name")));
            for cell in array(sheet.get("cells")) {
                let address = text(cell.get("address"));
                entries.push(Entry {
                    id: text(cell.get("id")),
                    label: format!("{name}!{address}"),
                    value: cell_value(cell),
                    position: format!("{sheet_id}:{address}"),
                });
            }
        }
        Ok(entries)
    }

    pub fn apply(&mut self, command: &Command) -> Result<Applied> {
        let Command::SetCell {
            sheet,
            cell,
            expected_value,
            value,
        } = command
        else {
            return Err(Error::edit(
                EditCode::UnsupportedOperation,
                "XLSX sources support set_cell only",
            ));
        };
        let cells = self.projection(false)?;
        let (index, found) = find_sheet(&cells, sheet)?;
        let sheet_id = text(found.get("id"));
        let (row, col, address) = a1(cell)?;
        let at = CellRef::new(row, col);
        let read_input = |workbook: &Workbook| {
            workbook
                .cell(SheetId(index as u32), at)
                .map(|cell| cell.input)
                .map_err(Error::engine)
        };
        let before = cell_at(found, &address).cloned();
        let input = read_input(&self.workbook)?;
        if *expected_value != input
            && *expected_value != before.as_ref().map(cell_value).unwrap_or_default()
        {
            return Err(Error::edit(
                EditCode::StaleTarget,
                "the cell differs from expected_value",
            ));
        }
        if before.is_none() && value.is_empty() {
            return Err(Error::edit(
                EditCode::InvalidInput,
                "the cell is already empty",
            ));
        }
        let result = self
            .workbook
            .edit_cell(SheetId(index as u32), at, value, now_options())
            .map_err(Error::engine)?;
        self.workbook.sheet_info().map_err(Error::engine)?;
        if !result.applied {
            return Err(Error::edit(
                EditCode::InvalidInput,
                "the workbook rejected this cell value",
            ));
        }
        let cell = match before {
            Some(cell) => cell,
            None => {
                let cells = self.projection(false)?;
                let (_, sheet) = find_sheet(&cells, &sheet_id)?;
                cell_at(sheet, &address)
                    .cloned()
                    .ok_or_else(|| Error::engine("edited cell is missing from the projection"))?
            }
        };
        Ok(Applied {
            id: text(cell.get("id")),
            inverse: Command::SetCell {
                sheet: sheet_id,
                cell: address,
                expected_value: read_input(&self.workbook)?,
                value: input,
            },
        })
    }

    pub fn locate(&self, id: &str) -> Result<Target> {
        let marker = id.find(":[").filter(|marker| *marker > 0).ok_or_else(|| {
            Error::edit(EditCode::UnavailableTarget, "target_id is not an XLSX cell")
        })?;
        let sheet_id = &id[..marker];
        let info = self.workbook.sheet_info().map_err(Error::engine)?;
        if !info.sheet_ids.iter().any(|known| known == sheet_id) {
            return Err(Error::edit(
                EditCode::UnavailableTarget,
                "the sheet is no longer in the workbook",
            ));
        }
        Ok(Target {
            id: id.to_owned(),
            path: vec![
                "xlsx:sheets".into(),
                sheet_id.into(),
                "contents".into(),
                id[marker + 1..].into(),
            ],
            range: None,
        })
    }

    /// `saveBytesAt(Date.parse(now) / 86_400_000 + 25569)`.
    pub fn export(&mut self, now: &str) -> Result<Vec<u8>> {
        let now_serial = parse_iso_ms(now) / MS_PER_DAY + UNIX_EPOCH_SERIAL;
        if !now_serial.is_finite() {
            return Err(Error::engine("export time must be finite"));
        }
        self.workbook.recalculate_all(CalculationOptions {
            now_serial: Some(now_serial),
        });
        self.workbook.save().map_err(Error::engine)
    }

    pub fn effects(&self) -> Result<Vec<NetEffect>> {
        let json = self
            .workbook
            .pending_effects_json()
            .map_err(Error::engine)?;
        serde_json::from_str(&json).map_err(Error::engine)
    }
}

/// `seedOffice` for XLSX: the state every replica starts from, built
/// without opening one.
pub(crate) fn seed(base: &[u8]) -> Result<Vec<u8>> {
    Workbook::seed_collaborative(base).map_err(Error::engine)
}

/// A source opened only to read pending effects (`XlsxEffectsReader`).
pub(crate) struct EffectsReader {
    workbook: Workbook,
}

impl EffectsReader {
    pub fn open(base: &[u8]) -> Result<Self> {
        Workbook::open_collaborative_for_effects(base, env::next_client())
            .map(|workbook| Self { workbook })
            .map_err(Error::engine)
    }

    /// The effects of `state`, or none when the reader would not adopt it whole.
    pub fn effects(&self, state: &[u8]) -> Result<Option<Vec<NetEffect>>> {
        let json = self
            .workbook
            .pending_effects_of_state_json(state, now_options())
            .map_err(Error::engine)?;
        json.map(|json| serde_json::from_str(&json).map_err(Error::engine))
            .transpose()
    }
}

/// office-checkpoint.ts `readXlsxEffects` without a room.
pub(crate) fn pending_effects(
    base: &[u8],
    checkpoint: Checkpoint,
    reader: Option<&EffectsReader>,
) -> Result<Vec<NetEffect>> {
    let opened;
    let reader = match reader {
        Some(reader) => reader,
        None => {
            opened = EffectsReader::open(base)?;
            &opened
        }
    };
    if let Some(effects) = reader.effects(checkpoint.state)? {
        return Ok(effects);
    }
    XlsxSession::open(base, Some(checkpoint.state))?.effects()
}

pub(crate) fn check(base: &[u8], checkpoint: Checkpoint) -> Result<()> {
    assert_checkpoint(Format::Xlsx, base, Some(checkpoint))
}

/// `XlsxDocument.rebaseCheckpoint`; an engine refusal is a rebase refusal.
pub(crate) fn rebase(
    base: &[u8],
    captured: &[u8],
    latest: &[u8],
    exported: &[u8],
) -> Result<Vec<u8>> {
    Workbook::rebase_checkpoint(base, captured, latest, exported, env::next_client())
        .map_err(|error| Error::rebase(error.to_string()))
}

fn array(value: Option<&J>) -> &[J] {
    value.and_then(J::as_arr).map(Vec::as_slice).unwrap_or(&[])
}

fn text(value: Option<&J>) -> String {
    value.map(js_string).unwrap_or_else(|| "undefined".into())
}

/// `${value}` in a template literal.
pub(crate) fn js_string(value: &J) -> String {
    match value {
        J::Null => "null".into(),
        J::Bool(value) => value.to_string(),
        J::Num(value) => crate::js::number_to_string(*value),
        J::Str(text) => text.clone(),
        J::Arr(items) => items
            .iter()
            .map(|item| match item {
                J::Null => String::new(),
                item => js_string(item),
            })
            .collect::<Vec<_>>()
            .join(","),
        J::Obj(_) => "[object Object]".into(),
    }
}

/// office-checkpoint.ts `xlsxCellValue`.
fn cell_value(cell: &J) -> String {
    match cell.get("formula") {
        Some(J::Null) | None => {}
        Some(formula) => return format!("={}", js_string(formula)),
    }
    let value = cell.get("value");
    if value
        .and_then(|value| value.get("kind"))
        .and_then(J::as_str)
        == Some("empty")
    {
        return String::new();
    }
    match value.and_then(|value| value.get("value")) {
        None => String::new(),
        Some(J::Str(text)) => text.clone(),
        Some(other) => canonical(other),
    }
}

/// office-checkpoint.ts `xlsxSheet`.
fn find_sheet<'a>(cells: &'a J, sheet: &str) -> Result<(usize, &'a J)> {
    let wanted = js_trim(sheet).to_lowercase();
    array(cells.get("sheets"))
        .iter()
        .enumerate()
        .find(|(position, item)| {
            item.get("id").and_then(J::as_str) == Some(sheet)
                || item.get("name").and_then(J::as_str).map(str::to_lowercase)
                    == Some(wanted.clone())
                || position.to_string() == wanted
        })
        .ok_or_else(|| Error::edit(EditCode::UnavailableTarget, "unknown sheet"))
}

fn cell_at<'a>(sheet: &'a J, address: &str) -> Option<&'a J> {
    array(sheet.get("cells"))
        .iter()
        .find(|cell| cell.get("address").and_then(J::as_str) == Some(address))
}

/// `String.prototype.trim`: White_Space and the byte order mark.
fn js_trim(text: &str) -> &str {
    text.trim_matches(|ch: char| ch.is_whitespace() || ch == '\u{feff}')
}

/// office-checkpoint.ts `a1`: zero-based row and column and the address.
fn a1(cell: &str) -> Result<(u32, u32, String)> {
    let invalid = || Error::edit(EditCode::InvalidInput, "cell must be an A1 address");
    let text = js_trim(cell);
    let text = text.strip_prefix('$').unwrap_or(text);
    let letters_end = text
        .find(|ch: char| !ch.is_ascii_alphabetic())
        .unwrap_or(text.len());
    let letters = &text[..letters_end];
    let rest = &text[letters_end..];
    let digits = rest.strip_prefix('$').unwrap_or(rest);
    if !(1..=3).contains(&letters.len())
        || !(1..=7).contains(&digits.len())
        || !digits.bytes().all(|byte| byte.is_ascii_digit())
    {
        return Err(invalid());
    }
    let letters = letters.to_ascii_uppercase();
    let col = letters
        .bytes()
        .fold(0_u32, |col, letter| col * 26 + (letter - b'A' + 1) as u32);
    let row: u32 = digits.parse().map_err(|_| invalid())?;
    if row < 1 {
        return Err(invalid());
    }
    Ok((row - 1, col - 1, format!("{letters}{row}")))
}

/// `Date.parse` of a `YYYY-MM-DDTHH:MM:SS.sssZ` string; NaN when invalid.
pub(crate) fn parse_iso_ms(text: &str) -> f64 {
    let bytes = text.as_bytes();
    let number = |range: std::ops::Range<usize>| -> Option<i64> {
        let part = text.get(range)?;
        part.bytes()
            .all(|byte| byte.is_ascii_digit())
            .then(|| part.parse().ok())
            .flatten()
    };
    let shaped = bytes.len() == 24
        && bytes[4] == b'-'
        && bytes[7] == b'-'
        && bytes[10] == b'T'
        && bytes[13] == b':'
        && bytes[16] == b':'
        && bytes[19] == b'.'
        && bytes[23] == b'Z';
    let parts = (
        number(0..4),
        number(5..7),
        number(8..10),
        number(11..13),
        number(14..16),
        number(17..19),
        number(20..23),
    );
    let (Some(year), Some(month), Some(day), Some(hour), Some(minute), Some(second), Some(millis)) =
        parts
    else {
        return f64::NAN;
    };
    let leap = (year % 4 == 0 && year % 100 != 0) || year % 400 == 0;
    let days_in_month = match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if leap => 29,
        2 => 28,
        _ => 0,
    };
    let midnight = hour == 24 && minute == 0 && second == 0 && millis == 0;
    if !shaped
        || day < 1
        || day > days_in_month
        || !(hour < 24 || midnight)
        || minute > 59
        || second > 59
    {
        return f64::NAN;
    }
    // Days from 1970-01-01 to the civil date (Howard Hinnant's algorithm).
    let y = if month <= 2 { year - 1 } else { year };
    let era = y.div_euclid(400);
    let yoe = y - era * 400;
    let mp = (month + 9) % 12;
    let doy = (153 * mp + 2) / 5 + day - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    let days = era * 146_097 + doe - 719_468;
    ((days * 86_400 + hour * 3600 + minute * 60 + second) * 1000 + millis) as f64
}

// The projection `crates/xlsx-wasm/src/core.rs` serializes for
// `checkpointProjectionJson` and `checkpointCellsJson`, repeated here because
// that crate is the wasm-bindgen boundary.

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct Projection<'a> {
    sheets: Vec<ProjectedSheet<'a>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    defined_names: Option<&'a Vec<betteroffice_xlsx::DefinedName>>,
}

#[derive(Serialize)]
struct ProjectedSheet<'a> {
    id: &'a str,
    name: &'a str,
    cells: ProjectedCells<'a>,
    #[serde(flatten)]
    layout: Option<ProjectedLayout<'a>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    images: Option<Vec<ProjectedImage>>,
}

#[derive(Serialize)]
struct ProjectedImage {
    id: String,
    part: String,
    anchor: betteroffice_xlsx::ChartAnchor,
    bytes: Vec<u8>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct ProjectedLayout<'a> {
    freeze_pane: &'a Option<betteroffice_xlsx::FreezePane>,
    hyperlinks: &'a Vec<betteroffice_xlsx::Hyperlink>,
    merges: &'a Vec<betteroffice_xlsx::CellRange>,
    col_widths: &'a std::collections::BTreeMap<betteroffice_xlsx::ColId, f64>,
    row_heights: &'a std::collections::BTreeMap<betteroffice_xlsx::RowId, f64>,
    charts: &'a Vec<betteroffice_xlsx::SheetChart>,
}

struct ProjectedCells<'a> {
    sheet: &'a betteroffice_xlsx::Sheet,
    ids: &'a [String],
    styles: Option<&'a betteroffice_xlsx::Stylesheet>,
}

impl Serialize for ProjectedCells<'_> {
    fn serialize<S: serde::Serializer>(
        &self,
        serializer: S,
    ) -> std::result::Result<S::Ok, S::Error> {
        #[derive(Serialize)]
        struct ProjectedCell<'a, F> {
            id: &'a str,
            address: String,
            value: &'a betteroffice_xlsx::CellValue,
            formula: &'a Option<String>,
            #[serde(skip_serializing_if = "Option::is_none")]
            format: Option<F>,
        }
        serializer.collect_seq(
            self.sheet
                .iter_cells()
                .zip(self.ids)
                .map(|((at, cell), id)| ProjectedCell {
                    id,
                    address: at.to_a1(),
                    value: &cell.value,
                    formula: &cell.formula,
                    format: self.styles.map(|styles| styles.cell_format(cell.style)),
                }),
        )
    }
}

fn projection_json(workbook: &Workbook, full: bool) -> Result<String> {
    let sheet_ids = workbook.sheet_info().map_err(Error::engine)?.sheet_ids;
    let model = workbook.model();
    let identities = workbook
        .cell_identities(model.sheets.iter().enumerate().flat_map(|(index, sheet)| {
            sheet
                .iter_cells()
                .map(move |(at, _)| (SheetId(index as u32), at))
        }))
        .map_err(Error::engine)?;
    let mut offset = 0;
    let mut sheets = Vec::with_capacity(model.sheets.len());
    for (index, sheet) in model.sheets.iter().enumerate() {
        let count = sheet.iter_cells().count();
        let images = if full {
            let images = workbook
                .embedded_images(SheetId(index as u32))
                .map_err(Error::engine)?;
            Some(
                images
                    .into_iter()
                    .map(|image| ProjectedImage {
                        id: image.id,
                        part: image.part,
                        anchor: image.anchor,
                        bytes: image.bytes,
                    })
                    .collect(),
            )
        } else {
            None
        };
        sheets.push(ProjectedSheet {
            id: &sheet_ids[index],
            name: &sheet.name,
            cells: ProjectedCells {
                sheet,
                ids: &identities[offset..offset + count],
                styles: full.then_some(&model.styles),
            },
            layout: full.then_some(ProjectedLayout {
                freeze_pane: &sheet.freeze_pane,
                hyperlinks: &sheet.hyperlinks,
                merges: &sheet.merges,
                col_widths: &sheet.col_widths,
                row_heights: &sheet.row_heights,
                charts: &sheet.charts,
            }),
            images,
        });
        offset += count;
    }
    serde_json::to_string(&Projection {
        sheets,
        defined_names: full.then_some(&model.defined_names),
    })
    .map_err(Error::engine)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a1_addresses_parse_as_the_ts_regex_does() {
        assert_eq!(a1(" $b$12 ").unwrap(), (11, 1, "B12".into()));
        assert_eq!(a1("AA1").unwrap(), (0, 26, "AA1".into()));
        assert!(a1("A0").is_err());
        assert!(a1("ABCD1").is_err());
        assert!(a1("A12345678").is_err());
    }

    #[test]
    fn iso_times_parse_as_date_parse_does() {
        assert_eq!(parse_iso_ms("2000-01-01T00:00:00.000Z"), 946_684_800_000.0);
        assert_eq!(
            parse_iso_ms("2026-09-29T12:34:56.789Z"),
            1_790_685_296_789.0
        );
        assert!(parse_iso_ms("2026-02-30T00:00:00.000Z").is_nan());
    }
}
