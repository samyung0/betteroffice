//! XLSX sessions over `betteroffice_xlsx::Workbook`, as office-checkpoint.ts
//! drives `XlsxDocument` and `XlsxEffectsReader`.

use std::collections::BTreeMap;

use betteroffice_xlsx::{
    CalculationOptions, CellInput, CellRange, CellRef, MAX_COLS, MAX_ROWS, NumberFormatMutation,
    Op, SheetId, StylePatch, Workbook,
};
use serde::Serialize;

use crate::common::{assert_checkpoint, image_mime_type, sha256_hex, storable_text};
use crate::env;
use crate::error::{EditCode, Error, Result};
use crate::formula::Formulas;
use crate::js::{J, canonical};
use crate::session::{Applied, Item};
use crate::types::{
    Asset, CellStyle, Checkpoint, Command, EffectKind, Entry, Format, InspectQuery, NetEffect,
    ObjectKind, ObjectRef, Target,
};

/// Days from the 1900 epoch to 1970-01-01, phantom leap day included.
const UNIX_EPOCH_SERIAL: f64 = 25569.0;
const MS_PER_DAY: f64 = 86_400_000.0;
/// Cells one `set_range` sets or `format_cells` formats, or one row or column
/// delete removes; also the cells one call's commands name across them (a
/// range's area, a delete's expected cells; Epo 2026-10-20, review finding 3).
const MAX_EDIT_CELLS: u64 = 10_000;
/// Cells one scoped read covers.
const MAX_READ_CELLS: u64 = 2_000;
/// Characters of a refused range's TSV the refusal carries.
const MAX_SHOWN_CHARS: usize = 2_000;
/// Cells of a refused band's grid the refusal carries, an empty row counting
/// as one.
const MAX_SHOWN_CELLS: usize = 2_000;

pub(crate) fn now_options() -> CalculationOptions {
    CalculationOptions {
        now_serial: Some(env::now_ms() / MS_PER_DAY + UNIX_EPOCH_SERIAL),
    }
}

/// One call's workbook: `apply_commands` and `create` open one per call.
pub(crate) struct XlsxSession {
    workbook: Workbook,
    /// Cells the call's commands wrote so far, toward [`MAX_EDIT_CELLS`].
    written: u64,
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
        Ok(Self {
            workbook,
            written: 0,
        })
    }

    pub fn state(&self) -> Vec<u8> {
        self.workbook.encode_state_as_update_v1()
    }

    /// Counts a command's `cells` toward its call, refused once the call
    /// passes [`MAX_EDIT_CELLS`]: a range's area, a delete's expected cells.
    /// The pipeline sends a call's commands only up to the limit, so the
    /// refusal guards other callers.
    fn write(&mut self, cells: u64) -> Result<()> {
        self.written += cells;
        if self.written > MAX_EDIT_CELLS {
            return Err(invalid(format!(
                "this call's commands write at least {} cells; a call writes at most {MAX_EDIT_CELLS} cells across its commands, so split the edit into several calls",
                self.written
            )));
        }
        Ok(())
    }

    /// `checkpointProjectionJson` (`full`) or `checkpointCellsJson` without
    /// the sheets' cells, which [`Self::each_cell`] reads one at a time: the
    /// whole projection as one value peaked at 2 GiB for an 8 MiB workbook.
    fn projection(&self, full: bool) -> Result<J> {
        J::parse(&projection_json(&self.workbook, full)?).map_err(Error::Engine)
    }

    /// Sheet `index`'s projected cells in order.
    fn each_cell(
        &self,
        index: usize,
        full: bool,
        mut visit: impl FnMut(J) -> Result<()>,
    ) -> Result<()> {
        let model = self.workbook.model();
        let sheet = &model.sheets[index];
        let ids = self
            .workbook
            .cell_identities(
                sheet
                    .iter_cells()
                    .map(|(at, _)| (SheetId(index as u32), at)),
            )
            .map_err(Error::engine)?;
        for ((at, cell), id) in sheet.iter_cells().zip(&ids) {
            let json = serde_json::to_string(&ProjectedCell {
                id,
                address: at.to_a1(),
                value: &cell.value,
                formula: &cell.formula,
                format: full.then(|| model.styles.cell_format(cell.style)),
            })
            .map_err(Error::engine)?;
            visit(J::parse(&json).map_err(Error::Engine)?)?;
        }
        Ok(())
    }

    /// The projected cell at `address` on sheet `index`.
    fn cell_at(&self, index: usize, address: &str) -> Result<Option<J>> {
        let mut found = None;
        self.each_cell(index, false, |cell| {
            if found.is_none() && cell.get("address").and_then(J::as_str) == Some(address) {
                found = Some(cell);
            }
            Ok(())
        })?;
        Ok(found)
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
                format!("workbook:{index}"),
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
            self.each_cell(index, true, |cell| {
                let cell = &cell;
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
                Ok(())
            })?;
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
        let sheets = self.projection(false)?;
        let mut entries = Vec::new();
        for (index, sheet) in array(sheets.get("sheets")).iter().enumerate() {
            let (sheet_id, name) = (text(sheet.get("id")), text(sheet.get("name")));
            self.each_cell(index, false, |cell| {
                let address = text(cell.get("address"));
                entries.push(Entry {
                    id: text(cell.get("id")),
                    label: format!("{name}!{address}"),
                    value: cell_value(&cell),
                    position: format!("{sheet_id}:{address}"),
                });
                Ok(())
            })?;
        }
        Ok(entries)
    }

    pub fn apply(&mut self, command: &Command) -> Result<Applied> {
        match command {
            Command::SetCell {
                sheet,
                cell,
                expected_value,
                value,
            } => self.set_cell(sheet, cell, expected_value, value),
            Command::SetRange {
                sheet,
                range,
                expected,
                values,
            } => self.set_range(sheet, range, expected, values),
            Command::FormatCells {
                sheet,
                range,
                style,
            } => self.format_cells(sheet, range, style),
            Command::AddSheet { name, after } => self.add_sheet(name, after.as_deref()),
            Command::RenameSheet {
                sheet,
                expected_name,
                name,
            } => self.rename_sheet(sheet, expected_name, name),
            Command::DeleteSheet {
                sheet,
                expected_name,
            } => self.delete_sheet(sheet, expected_name),
            Command::InsertRows { sheet, at, count } => {
                self.insert_band(sheet, Band::rows(*at, *count)?)
            }
            Command::DeleteRows {
                sheet,
                at,
                count,
                expected,
            } => self.delete_band(sheet, Band::rows(*at, *count)?, expected),
            Command::InsertColumns { sheet, at, count } => {
                self.insert_band(sheet, Band::columns(at, *count)?)
            }
            Command::DeleteColumns {
                sheet,
                at,
                count,
                expected,
            } => self.delete_band(sheet, Band::columns(at, *count)?, expected),
            Command::ReplaceText { .. } => Err(Error::edit(
                EditCode::UnsupportedOperation,
                "XLSX sources support set_cell only",
            )),
            Command::InsertParagraph { .. } | Command::DeleteParagraph { .. } => Err(Error::edit(
                EditCode::UnsupportedOperation,
                "this command edits DOCX and PPTX files",
            )),
            Command::SetList { .. }
            | Command::InsertTable { .. }
            | Command::DeleteTable { .. }
            | Command::FormatText { .. }
            | Command::SetAlignment { .. } => Err(Error::edit(
                EditCode::UnsupportedOperation,
                "this command edits DOCX files",
            )),
        }
    }

    fn set_cell(
        &mut self,
        sheet: &str,
        cell: &str,
        expected_value: &str,
        value: &str,
    ) -> Result<Applied> {
        self.write(1)?;
        let cells = self.projection(false)?;
        let (index, found) = find_sheet(&cells, sheet)?;
        let sheet_id = text(found.get("id"));
        let (row, col, address) = a1(cell)?;
        self.formulas().check(&address, value)?;
        let at = CellRef::new(row, col);
        let read_input = |workbook: &Workbook| {
            workbook
                .cell(SheetId(index as u32), at)
                .map(|cell| cell.input)
                .map_err(refused)
        };
        let before = self.cell_at(index, &address)?;
        let input = read_input(&self.workbook)?;
        if expected_value != input
            && expected_value != before.as_ref().map(cell_value).unwrap_or_default()
        {
            // The cell as inspect lists it; an empty cell is not listed.
            return Err(Error::Edit {
                code: EditCode::StaleTarget,
                message: "the cell differs from expected_value".into(),
                details: before.as_ref().map(|cell| {
                    serde_json::json!({
                        "block": { "id": text(cell.get("id")), "type": "cell", "text": cell_value(cell) }
                    })
                }),
            });
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
            .map_err(refused)?;
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
                let sheets = self.projection(false)?;
                let (index, _) = find_sheet(&sheets, &sheet_id)?;
                self.cell_at(index, &address)?
                    .ok_or_else(|| Error::engine("edited cell is missing from the projection"))?
            }
        };
        Ok(Applied {
            ids: vec![text(cell.get("id"))],
            inverse: Some(Command::SetCell {
                sheet: sheet_id,
                cell: address,
                expected_value: read_input(&self.workbook)?,
                value: storable_text(&input),
            }),
        })
    }

    /// One `edit_cells` for the whole range, refused whole when a cell
    /// differs from `expected`. The inverse swaps the two grids.
    fn set_range(
        &mut self,
        sheet: &str,
        range: &str,
        expected: &[Vec<String>],
        values: &[Vec<String>],
    ) -> Result<Applied> {
        let sheet = self.sheet_ref(sheet)?;
        let area = cell_range(range)?;
        let (rows, columns) = shape(area);
        if rows * columns > MAX_EDIT_CELLS {
            return Err(invalid(format!(
                "{} has {} cells; set at most {MAX_EDIT_CELLS} cells per command",
                area.to_a1(),
                rows * columns
            )));
        }
        for (label, grid) in [("expected", expected), ("values", values)] {
            if grid.len() as u64 != rows || grid.iter().any(|row| row.len() as u64 != columns) {
                return Err(invalid(format!(
                    "{label} must be {rows} rows of {columns} cells, the shape of {}",
                    area.to_a1()
                )));
            }
        }
        self.write(rows * columns)?;
        let formulas = self.formulas();
        for (at, value) in cells(area).zip(values.iter().flatten()) {
            formulas.check(&at.to_a1(), value)?;
        }
        let before = cells(area)
            .map(|at| self.read(sheet.index, at))
            .collect::<Result<Vec<_>>>()?;
        let stale = before
            .iter()
            .zip(expected.iter().flatten())
            .filter(|((input, shown), want)| *want != input && *want != shown)
            .count();
        if stale > 0 {
            let inputs: Vec<&str> = before.iter().map(|(input, _)| input.as_str()).collect();
            return Err(Error::Edit {
                code: EditCode::StaleTarget,
                message: format!(
                    "{stale} of the range's cells differ from expected; read the range again"
                ),
                details: Some(serde_json::json!({ "block": {
                    "id": format!("{}!{}", sheet.name, area.to_a1()),
                    "type": "range",
                    "text": capped(tsv(inputs.chunks(columns as usize))),
                }})),
            });
        }
        let at = SheetId(sheet.index as u32);
        let edits: Vec<CellInput> = cells(area)
            .zip(values.iter().flatten())
            .map(|(cell, input)| CellInput {
                cell,
                input: input.clone(),
            })
            .collect();
        self.workbook
            .edit_cells(at, &edits, now_options())
            .map_err(refused)?;
        let after = cells(area)
            .map(|cell| self.workbook.cell(at, cell).map(|cell| cell.input))
            .collect::<std::result::Result<Vec<_>, _>>()
            .map_err(refused)?;
        let changed: Vec<(SheetId, CellRef)> = cells(area)
            .zip(before.iter().zip(&after))
            .filter(|(_, ((old, _), new))| old != *new)
            .map(|(cell, _)| (at, cell))
            .collect();
        let grid = |inputs: Vec<String>| -> Vec<Vec<String>> {
            inputs
                .chunks(columns as usize)
                .map(<[String]>::to_vec)
                .collect()
        };
        Ok(Applied {
            ids: self
                .workbook
                .cell_identities(changed)
                .map_err(Error::engine)?,
            inverse: Some(Command::SetRange {
                sheet: sheet.id,
                range: area.to_a1(),
                expected: grid(after),
                values: grid(
                    before
                        .iter()
                        .map(|(input, _)| storable_text(input))
                        .collect(),
                ),
            }),
        })
    }

    /// One `apply_ops` for the style and the number format. No inverse.
    fn format_cells(&mut self, sheet: &str, range: &str, style: &CellStyle) -> Result<Applied> {
        let sheet = SheetId(self.sheet_ref(sheet)?.index as u32);
        let range = cell_range(range)?;
        let (rows, columns) = shape(range);
        if rows * columns > MAX_EDIT_CELLS {
            return Err(invalid(format!(
                "{} has {} cells; format at most {MAX_EDIT_CELLS} cells per command",
                range.to_a1(),
                rows * columns
            )));
        }
        self.write(rows * columns)?;
        let patch = &style.patch;
        if let Some(font) = &patch.font_family
            && !(1..=31).contains(&font.chars().count())
        {
            return Err(invalid("fontFamily must be 1 to 31 characters"));
        }
        for (field, color) in [
            ("textColor", patch.text_color.as_ref()),
            ("fillColor", patch.fill_color.as_ref()),
            (
                "border color",
                patch
                    .border
                    .as_ref()
                    .and_then(|border| border.color.as_ref()),
            ),
        ] {
            if let Some(color) = color
                && !(color.len() == 7
                    && color.starts_with('#')
                    && color[1..].bytes().all(|byte| byte.is_ascii_hexdigit()))
            {
                return Err(invalid(format!(
                    "{field} {color:?} must be a #RRGGBB colour"
                )));
            }
        }
        if let Some(NumberFormatMutation::Custom { pattern }) = &style.number_format
            && !(1..=255).contains(&pattern.chars().count())
        {
            return Err(invalid(
                "a custom numberFormat pattern must be 1 to 255 characters",
            ));
        }
        let mut ops = Vec::new();
        if *patch != StylePatch::default() {
            ops.push(Op::PatchRangeStyle {
                sheet,
                range,
                patch: patch.clone(),
            });
        }
        if let Some(format) = &style.number_format {
            ops.push(Op::SetRangeNumberFormat {
                sheet,
                range,
                format: format.clone(),
            });
        }
        if ops.is_empty() {
            return Err(invalid(
                "style sets nothing; set bold, italic, strikethrough, fontFamily, fontSize, textColor, fillColor, border, horizontalAlignment, verticalAlignment, textWrapping, clear or numberFormat",
            ));
        }
        self.apply_ops(ops, None)
    }

    fn add_sheet(&mut self, name: &str, after: Option<&str>) -> Result<Applied> {
        let index = match after {
            Some(after) => self.sheet_ref(after)?.index + 1,
            None => self.workbook.sheet_count(),
        };
        self.apply_ops(
            vec![Op::AddSheet {
                index,
                name: name.to_owned(),
            }],
            None,
        )
    }

    /// The inverse renames it back.
    fn rename_sheet(&mut self, sheet: &str, expected_name: &str, name: &str) -> Result<Applied> {
        let sheet = self.named_sheet(sheet, expected_name)?;
        self.apply_ops(
            vec![Op::RenameSheet {
                sheet: SheetId(sheet.index as u32),
                name: name.to_owned(),
            }],
            Some(Command::RenameSheet {
                sheet: sheet.id,
                expected_name: name.to_owned(),
                name: sheet.name,
            }),
        )
    }

    fn delete_sheet(&mut self, sheet: &str, expected_name: &str) -> Result<Applied> {
        let sheet = self.named_sheet(sheet, expected_name)?;
        if self.workbook.sheet_count() == 1 {
            return Err(invalid(format!(
                "{:?} is the workbook's only sheet; a workbook keeps at least one",
                sheet.name
            )));
        }
        self.apply_ops(vec![Op::RemoveSheet { index: sheet.index }], None)
    }

    /// The inverse deletes the new band, expecting it still empty.
    fn insert_band(&mut self, sheet: &str, band: Band) -> Result<Applied> {
        let sheet = self.sheet_ref(sheet)?;
        let (index, at, count) = (SheetId(sheet.index as u32), band.at, band.count);
        let (op, inverse) = if band.rows {
            (
                Op::InsertRows {
                    sheet: index,
                    at,
                    count,
                },
                Command::DeleteRows {
                    sheet: sheet.id,
                    at: at + 1,
                    count,
                    expected: Vec::new(),
                },
            )
        } else {
            (
                Op::InsertCols {
                    sheet: index,
                    at,
                    count,
                },
                Command::DeleteColumns {
                    sheet: sheet.id,
                    at: column_letters(at),
                    count,
                    expected: Vec::new(),
                },
            )
        };
        self.apply_ops(vec![op], Some(inverse))
    }

    /// Refused when the band holds more than [`MAX_EDIT_CELLS`] stored cells
    /// or its inputs differ from `expected`, a grid of rows whose trailing
    /// empty cells and rows may be left out (Epo 2026-10-20, review finding
    /// 6). No inverse.
    fn delete_band(
        &mut self,
        sheet: &str,
        band: Band,
        expected: &[Vec<String>],
    ) -> Result<Applied> {
        let sheet = self.sheet_ref(sheet)?;
        let model = self.workbook.model();
        let stored: Vec<CellRef> = model.sheets[sheet.index]
            .iter_cells()
            .map(|(at, _)| at)
            .filter(|at| band.holds(*at))
            .collect();
        if stored.len() as u64 > MAX_EDIT_CELLS {
            return Err(invalid(format!(
                "{} holds {} cells; delete at most {MAX_EDIT_CELLS} cells per command",
                band.label(),
                stored.len()
            )));
        }
        // The call's count takes the cells the delete names, as the pipeline
        // counts them before it sends a call (Epo 2026-10-21).
        self.write(expected.iter().map(|row| row.len() as u64).sum())?;
        let mut current: BTreeMap<(u32, u32), (String, String)> = BTreeMap::new();
        for at in stored {
            let cell = self.read(sheet.index, at)?;
            if !cell.0.is_empty() || !cell.1.is_empty() {
                current.insert(band.place(at), cell);
            }
        }
        let empty = (String::new(), String::new());
        let wanted_matches = expected.iter().enumerate().all(|(line, cells)| {
            cells.iter().enumerate().all(|(column, want)| {
                let (input, shown) = current.get(&(line as u32, column as u32)).unwrap_or(&empty);
                want == input || want == shown
            })
        });
        let current_covered = current.keys().all(|(line, column)| {
            expected
                .get(*line as usize)
                .is_some_and(|cells| (*column as usize) < cells.len())
        });
        if !wanted_matches || !current_covered {
            // The exact grid to copy into `expected`, and its TSV to read.
            let (grid, cut) = band_grid(&current);
            let mut block = serde_json::json!({
                "id": format!("{}!{}", sheet.name, band.label()),
                "type": "range",
                "text": capped(tsv(grid.iter().map(Vec::as_slice))),
                "grid": grid,
            });
            if cut {
                block["gridCut"] = serde_json::Value::Bool(true);
            }
            return Err(Error::Edit {
                code: EditCode::StaleTarget,
                message: format!("{} differs from expected; read it again", band.label()),
                details: Some(serde_json::json!({ "block": block })),
            });
        }
        let index = SheetId(sheet.index as u32);
        let op = if band.rows {
            Op::DeleteRows {
                sheet: index,
                at: band.at,
                count: band.count,
            }
        } else {
            Op::DeleteCols {
                sheet: index,
                at: band.at,
                count: band.count,
            }
        };
        self.apply_ops(vec![op], None)
    }

    /// `ops` as one engine edit; touched ids are left to the cells' commands.
    fn apply_ops(&mut self, ops: Vec<Op>, inverse: Option<Command>) -> Result<Applied> {
        self.workbook
            .apply_ops(ops, now_options())
            .map_err(refused)?;
        Ok(Applied {
            ids: Vec::new(),
            inverse,
        })
    }

    /// Scoped reads without projecting the workbook: the sheet list, or the
    /// non-empty cells of a range of at most [`MAX_READ_CELLS`].
    pub fn inspect_scoped(&self, query: &InspectQuery) -> Result<serde_json::Value> {
        let Some(range) = &query.range else {
            let info = self.workbook.sheet_info().map_err(Error::engine)?;
            let sheets: Vec<_> = info
                .sheet_ids
                .iter()
                .zip(&info.sheet_names)
                .enumerate()
                .map(|(index, (id, name))| self.summary(index, id, name))
                .collect();
            return Ok(serde_json::json!({ "sheets": sheets }));
        };
        let sheet = query
            .sheet
            .as_deref()
            .ok_or_else(|| invalid("a range read names its sheet"))?;
        let sheet = self.sheet_ref(sheet)?;
        let area = cell_range(range)?;
        let (rows, columns) = shape(area);
        if rows * columns > MAX_READ_CELLS {
            return Err(invalid(format!(
                "{} has {} cells; read at most {MAX_READ_CELLS} cells at a time, as a smaller range",
                area.to_a1(),
                rows * columns
            )));
        }
        let model = self.workbook.model();
        let mut cells = Vec::new();
        for (at, _) in model.sheets[sheet.index].iter_cells_in_rect(
            area.start.row..area.end.row + 1,
            area.start.col..area.end.col + 1,
        ) {
            let (input, value) = self.read(sheet.index, at)?;
            if !input.is_empty() {
                cells.push(serde_json::json!({
                    "address": at.to_a1(),
                    "input": input,
                    "value": value,
                }));
            }
        }
        Ok(serde_json::json!({
            "sheet": self.summary(sheet.index, &sheet.id, &sheet.name),
            "range": area.to_a1(),
            "cells": cells,
        }))
    }

    /// `rows` and `columns` reach the used range's last row and column.
    fn summary(&self, index: usize, id: &str, name: &str) -> serde_json::Value {
        let used = self.workbook.model().sheets[index].used_range();
        serde_json::json!({
            "id": id,
            "name": name,
            "usedRange": used.map(|used| used.to_a1()),
            "rows": used.map_or(0, |used| used.end.row + 1),
            "columns": used.map_or(0, |used| used.end.col + 1),
        })
    }

    /// A cell's input as typed and its text as the grid shows it.
    fn read(&self, index: usize, at: CellRef) -> Result<(String, String)> {
        let input = self
            .workbook
            .cell(SheetId(index as u32), at)
            .map_err(refused)?
            .input;
        let model = self.workbook.model();
        let shown = model.sheets[index]
            .cell(at)
            .map(|cell| xlsx_render::display_text(&model.styles, model.date_system, cell))
            .unwrap_or_default();
        Ok((input, shown))
    }

    fn formulas(&self) -> Formulas<'_> {
        let model = self.workbook.model();
        Formulas {
            sheets: model
                .sheets
                .iter()
                .map(|sheet| sheet.name.as_str())
                .collect(),
            names: &model.defined_names,
        }
    }

    /// A sheet by id, name or index, as `xlsxSheet` finds one; an unknown one
    /// is refused with the workbook's sheets.
    fn sheet_ref(&self, sheet: &str) -> Result<SheetRef> {
        let info = self.workbook.sheet_info().map_err(Error::engine)?;
        let wanted = js_trim(sheet).to_lowercase();
        let found = info
            .sheet_ids
            .iter()
            .zip(&info.sheet_names)
            .enumerate()
            .find(|(index, (id, name))| {
                *id == sheet || name.to_lowercase() == wanted || index.to_string() == wanted
            });
        match found {
            Some((index, (id, name))) => Ok(SheetRef {
                index,
                id: id.clone(),
                name: name.clone(),
            }),
            None => Err(Error::Edit {
                code: EditCode::UnavailableTarget,
                message: format!("the workbook has no sheet {sheet:?}"),
                details: Some(serde_json::json!({
                    "blocks": info
                        .sheet_ids
                        .iter()
                        .zip(&info.sheet_names)
                        .map(|(id, name)| serde_json::json!({ "id": id, "type": "sheet", "text": name }))
                        .collect::<Vec<_>>(),
                })),
            }),
        }
    }

    /// [`Self::sheet_ref`] refused as stale unless it is named `expected`.
    fn named_sheet(&self, sheet: &str, expected: &str) -> Result<SheetRef> {
        let sheet = self.sheet_ref(sheet)?;
        if sheet.name != expected {
            return Err(Error::Edit {
                code: EditCode::StaleTarget,
                message: format!("the sheet is named {:?}, not {expected:?}", sheet.name),
                details: Some(serde_json::json!({
                    "block": { "id": sheet.id, "type": "sheet", "text": sheet.name }
                })),
            });
        }
        Ok(sheet)
    }

    /// Each cell id's Yrs location, against one sheet list.
    pub fn locate_all(&self, ids: &[String]) -> Result<Vec<Target>> {
        let info = self.workbook.sheet_info().map_err(Error::engine)?;
        ids.iter()
            .map(|id| cell_target(&info.sheet_ids, id))
            .collect()
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

/// The workbook's refusal of a cell, value, style or structure change the
/// agent chose (outside the grid, over Excel's limits, not a worksheet, a
/// reference it cannot rewrite) is an edit refusal with its sentence; any
/// other failure stays the engine's.
/// An agent edit the workbook refused. Faults valid agent input cannot reach
/// once [`XlsxSession::sheet_ref`] found the sheet (no sheets, a missing
/// sheet or chart frame) stay engine errors, so they reach Sentry.
fn refused(error: betteroffice_xlsx::Error) -> Error {
    use betteroffice_xlsx::Error as Xlsx;
    use xlsx_ops::OpError as Op;
    match error {
        Xlsx::CellOutOfRange(_)
        | Xlsx::InvalidOperation(_)
        | Xlsx::Operation(Op::InvalidStyle(_)) => invalid(error.to_string()),
        Xlsx::RangeTooLarge { rows, cols, max } => invalid(format!(
            "the range is {rows} by {cols} cells; an edit covers at most {max} cells"
        )),
        Xlsx::Operation(
            Op::FormulaNotRewritable { .. }
            | Op::DefinedNameNotRewritable { .. }
            | Op::ChartRefNotRewritable { .. }
            | Op::ChartAnchorNotMovable { .. }
            | Op::NumFmtTableFull,
        ) => Error::edit(EditCode::UnsupportedOperation, error.to_string()),
        error => Error::engine(error),
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::edit(EditCode::InvalidInput, message)
}

/// A sheet an agent command names.
struct SheetRef {
    index: usize,
    id: String,
    name: String,
}

/// Rows or columns `at..at + count`, 0-based.
struct Band {
    rows: bool,
    at: u32,
    count: u32,
}

impl Band {
    /// `at` is 1-based, as rows are numbered.
    fn rows(at: u32, count: u32) -> Result<Self> {
        if !(1..=MAX_ROWS).contains(&at) {
            return Err(invalid(format!(
                "row {at} is outside the sheet (rows 1 to {MAX_ROWS})"
            )));
        }
        Self::counted(true, at - 1, count)
    }

    fn columns(at: &str, count: u32) -> Result<Self> {
        Self::counted(false, column(at)?, count)
    }

    fn counted(rows: bool, at: u32, count: u32) -> Result<Self> {
        if count == 0 {
            return Err(invalid("count must be at least 1"));
        }
        Ok(Self { rows, at, count })
    }

    fn holds(&self, cell: CellRef) -> bool {
        let index = if self.rows { cell.row } else { cell.col };
        index >= self.at && index - self.at < self.count
    }

    /// The cell's line and column in the band's grid: rows list their cells
    /// from column A, columns list theirs from row 1.
    fn place(&self, cell: CellRef) -> (u32, u32) {
        if self.rows {
            (cell.row - self.at, cell.col)
        } else {
            (cell.row, cell.col - self.at)
        }
    }

    /// `5:6` or `C:D`, as Excel names whole rows and columns.
    fn label(&self) -> String {
        let last = self.at.saturating_add(self.count - 1);
        if self.rows {
            format!("{}:{}", self.at + 1, u64::from(last) + 1)
        } else {
            format!("{}:{}", column_letters(self.at), column_letters(last))
        }
    }
}

/// `A1` or `A1:C3` (any corner order, `$` anchors ignored) inside the grid.
fn cell_range(range: &str) -> Result<CellRange> {
    let parsed = CellRange::parse_a1(&js_trim(range).to_ascii_uppercase()).map_err(|_| {
        invalid(format!(
            "range {range:?} must be a cell such as B2 or a range such as A1:C3 inside the sheet"
        ))
    })?;
    Ok(CellRange::new(
        CellRef::new(parsed.start.row, parsed.start.col),
        CellRef::new(parsed.end.row, parsed.end.col),
    ))
}

/// Rows and columns.
fn shape(range: CellRange) -> (u64, u64) {
    (
        u64::from(range.end.row - range.start.row) + 1,
        u64::from(range.end.col - range.start.col) + 1,
    )
}

/// The range's cells, row by row.
fn cells(range: CellRange) -> impl Iterator<Item = CellRef> {
    (range.start.row..=range.end.row).flat_map(move |row| {
        (range.start.col..=range.end.col).map(move |col| CellRef::new(row, col))
    })
}

/// A band's non-empty inputs by (line, cell) as rows of cells, each to its
/// last non-empty cell, cut after [`MAX_SHOWN_CELLS`]; `true` when cut.
fn band_grid(current: &BTreeMap<(u32, u32), (String, String)>) -> (Vec<Vec<&str>>, bool) {
    let lines = current.keys().next_back().map_or(0, |(line, _)| line + 1);
    let mut left = MAX_SHOWN_CELLS;
    let mut grid = Vec::new();
    for line in 0..lines {
        if left == 0 {
            return (grid, true);
        }
        let width = current
            .range((line, 0)..(line + 1, 0))
            .next_back()
            .map_or(0, |((_, column), _)| *column as usize + 1);
        let shown = width.min(left);
        grid.push(
            (0..shown as u32)
                .map(|column| {
                    current
                        .get(&(line, column))
                        .map_or("", |(input, _)| input.as_str())
                })
                .collect(),
        );
        left -= shown.max(1);
        if shown < width {
            return (grid, true);
        }
    }
    (grid, false)
}

fn tsv<'a>(rows: impl Iterator<Item = &'a [&'a str]>) -> String {
    rows.map(|row| row.join("\t"))
        .collect::<Vec<_>>()
        .join("\n")
}

/// At most [`MAX_SHOWN_CHARS`] characters, a cut marked with `…`.
fn capped(text: String) -> String {
    match text.char_indices().nth(MAX_SHOWN_CHARS) {
        Some((end, _)) => format!("{}…", &text[..end]),
        None => text,
    }
}

/// A column letter (`C`, `AA`) as a 0-based index.
fn column(letters: &str) -> Result<u32> {
    let text = js_trim(letters).to_ascii_uppercase();
    let refused = || {
        invalid(format!(
            "column {letters:?} must be a column letter from A to {}",
            column_letters(MAX_COLS - 1)
        ))
    };
    if !(1..=3).contains(&text.len()) || !text.bytes().all(|byte| byte.is_ascii_uppercase()) {
        return Err(refused());
    }
    let col = text
        .bytes()
        .fold(0_u32, |col, letter| col * 26 + u32::from(letter - b'A' + 1));
    if col > MAX_COLS {
        return Err(refused());
    }
    Ok(col - 1)
}

fn column_letters(col: u32) -> String {
    let address = CellRef::new(0, col).to_a1();
    address.trim_end_matches('1').to_owned()
}

/// The Yrs location of cell id `<sheet key>:<cell key>` (its contents entry).
fn cell_target(sheet_ids: &[String], id: &str) -> Result<Target> {
    let marker = id
        .find(":[")
        .filter(|marker| *marker > 0)
        .ok_or_else(|| Error::edit(EditCode::UnavailableTarget, "target_id is not an XLSX cell"))?;
    let sheet_id = &id[..marker];
    if !sheet_ids.iter().any(|known| known == sheet_id) {
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

#[derive(Serialize)]
struct ProjectedCell<'a, F> {
    id: &'a str,
    address: String,
    value: &'a betteroffice_xlsx::CellValue,
    formula: &'a Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    format: Option<F>,
}

fn projection_json(workbook: &Workbook, full: bool) -> Result<String> {
    let sheet_ids = workbook.sheet_info().map_err(Error::engine)?.sheet_ids;
    let model = workbook.model();
    let mut sheets = Vec::with_capacity(model.sheets.len());
    for (index, sheet) in model.sheets.iter().enumerate() {
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

    /// Which workbook errors an agent edit gets back as a refusal (409) and
    /// which stay engine faults (503, Sentry).
    #[test]
    fn workbook_errors_map_to_refusals_or_engine_faults() {
        use betteroffice_xlsx::Error as Xlsx;
        use xlsx_ops::OpError as Op;
        let cell = CellRef::new(0, 0);
        let part = || "xl/charts/chart1.xml".to_owned();
        let frame = || "Chart 1".to_owned();
        let cases = [
            (Xlsx::CellOutOfRange(cell), "invalid_input"),
            (Xlsx::InvalidOperation("bad".into()), "invalid_input"),
            (
                Xlsx::Operation(Op::InvalidStyle("bad".into())),
                "invalid_input",
            ),
            (
                Xlsx::RangeTooLarge {
                    rows: 200,
                    cols: 200,
                    max: 10_000,
                },
                "invalid_input",
            ),
            (
                Xlsx::Operation(Op::FormulaNotRewritable {
                    sheet: SheetId(0),
                    cell,
                }),
                "unsupported_operation",
            ),
            (
                Xlsx::Operation(Op::DefinedNameNotRewritable {
                    name: "Rate".into(),
                }),
                "unsupported_operation",
            ),
            (
                Xlsx::Operation(Op::ChartRefNotRewritable { part: part() }),
                "unsupported_operation",
            ),
            (
                Xlsx::Operation(Op::ChartAnchorNotMovable { part: part() }),
                "unsupported_operation",
            ),
            (
                Xlsx::Operation(Op::NumFmtTableFull),
                "unsupported_operation",
            ),
            (Xlsx::NoSheets, "engine"),
            (Xlsx::Operation(Op::SheetNotFound(SheetId(3))), "engine"),
            (Xlsx::Operation(Op::SheetIndexOutOfRange(3)), "engine"),
            (
                Xlsx::Operation(Op::ChartNotFound { frame: frame() }),
                "engine",
            ),
            (
                Xlsx::Operation(Op::ChartFrameShifted { frame: frame() }),
                "engine",
            ),
            (Xlsx::SheetOutOfRange(SheetId(3)), "engine"),
        ];
        for (error, expected) in cases {
            let shown = error.to_string();
            let kind = match refused(error) {
                Error::Edit { code, .. } => code.as_str(),
                Error::Engine(_) => "engine",
                other => panic!("{shown}: {other:?}"),
            };
            assert_eq!(kind, expected, "{shown}");
        }
    }

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

    /// Sheets sit in the workbook container (`workbook:<n>`, as slides in
    /// `deck`): a sheet inserted first shifts the others without moving
    /// them; a sheet the order puts elsewhere still moves.
    #[test]
    fn sheets_an_insert_shifts_do_not_move_a_reordered_sheet_does() {
        use yrs::types::ToJson;
        use yrs::{Array, ReadTxn, Transact, Update, updates::decoder::Decode};

        const SHOWCASE: &[u8] = include_bytes!("../../../apps/demo/public/showcase.xlsx");
        let sha = sha256_hex(SHOWCASE);
        let checkpoint = |state| Checkpoint {
            format: Format::Xlsx,
            schema_version: 1,
            base_sha256: &sha,
            state,
        };
        let seeded = crate::seed(Format::Xlsx, SHOWCASE).expect("seeds");
        let mut session = XlsxSession::open(SHOWCASE, Some(&seeded)).expect("opens");
        session
            .workbook
            .apply_ops(
                vec![betteroffice_xlsx::Op::AddSheet {
                    index: 0,
                    name: "First".into(),
                }],
                now_options(),
            )
            .expect("adds a sheet");
        let inserted = session.state();
        let before = crate::baseline(SHOWCASE, checkpoint(&seeded)).expect("baseline");
        let after = crate::baseline(SHOWCASE, checkpoint(&inserted)).expect("baseline");
        let moves = |effects: &[NetEffect]| {
            effects
                .iter()
                .filter(|effect| effect.operation == crate::types::Operation::Move)
                .count()
        };
        assert!(moves(&crate::compare_baselines(&before, &after)) > 0);
        let saved = crate::unshifted_effects(&before, &after);
        assert_eq!(moves(&saved), 0);
        assert!(saved.iter().any(|effect| effect.label == "Sheet First"));

        // The last sheet moved first in the order.
        let doc = yrs::Doc::new();
        let mut txn = doc.transact_mut();
        txn.apply_update(Update::decode_v1(&seeded).expect("decodes"))
            .expect("applies");
        let order = txn.get_array("xlsx:sheet-order").expect("a sheet order");
        let last = order.len(&txn) - 1;
        assert!(last >= 1, "two sheets at least");
        let moved = order.get(&txn, last).expect("the last sheet").to_json(&txn);
        order.remove(&mut txn, last);
        order.insert(&mut txn, 0, moved);
        drop(txn);
        let reordered = doc
            .transact()
            .encode_state_as_update_v1(&yrs::StateVector::default());
        let after = crate::baseline(SHOWCASE, checkpoint(&reordered)).expect("baseline");
        assert!(moves(&crate::unshifted_effects(&before, &after)) > 0);
    }
}
