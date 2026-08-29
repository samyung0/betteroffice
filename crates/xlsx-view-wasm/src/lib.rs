use serde::{Deserialize, Serialize};
use wasm_bindgen::prelude::*;
use xlsx_model::{CellRange, CellRef, MAX_COLS, MAX_ROWS, Sheet, SheetChart, SheetId, Workbook};
use xlsx_parse::PreservedPackage;
use xlsx_render::{
    ChartRegion, GridGeometry, RenderError, Viewport, build_display_list_with_charts,
    chart_at_point, chart_regions, display_text,
};

const MAX_DISPLAY_CELLS: u64 = 250_000;

#[wasm_bindgen]
pub struct XlsxViewDocument {
    workbook: Workbook,
    package: PreservedPackage,
    active_sheet: SheetId,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct SheetInfo {
    sheet_ids: Vec<String>,
    sheet_names: Vec<String>,
    active_sheet: u32,
    content_width: f32,
    content_height: f32,
    frozen_rows: u32,
    frozen_cols: u32,
    initial_scroll_x: f32,
    initial_scroll_y: f32,
}

#[derive(Deserialize)]
struct CellArgs {
    sheet: u32,
    row: u32,
    col: u32,
}

#[derive(Deserialize)]
struct RangeArgs {
    sheet: u32,
    range: String,
}

#[derive(Deserialize)]
struct ChartHitArgs {
    viewport: Viewport,
    x: f32,
    y: f32,
}

#[derive(Serialize)]
struct CellPosition {
    x: f32,
    y: f32,
}

#[derive(Serialize)]
struct MergedRanges {
    ranges: Vec<CellRange>,
}

#[wasm_bindgen]
impl XlsxViewDocument {
    pub fn open(bytes: &[u8]) -> Result<XlsxViewDocument, JsValue> {
        let parts = ooxml_opc::unzip_parts(bytes).map_err(js_error)?;
        let parsed = xlsx_parse::parse_workbook_with_package(&parts).map_err(js_error)?;
        let active_sheet = if (parsed.active_sheet.0 as usize) < parsed.workbook.sheets.len() {
            parsed.active_sheet
        } else {
            SheetId(0)
        };
        Ok(Self {
            workbook: parsed.workbook,
            package: parsed.package,
            active_sheet,
        })
    }

    #[wasm_bindgen(js_name = sheetInfoJson)]
    pub fn sheet_info_json(&self) -> Result<String, JsValue> {
        serde_json::to_string(&self.sheet_info()?).map_err(js_error)
    }

    #[wasm_bindgen(js_name = setActiveSheet)]
    pub fn set_active_sheet(&mut self, index: u32) -> Result<(), JsValue> {
        self.sheet(SheetId(index))?;
        self.active_sheet = SheetId(index);
        Ok(())
    }

    #[wasm_bindgen(js_name = displayListJson)]
    pub fn display_list_json(&self, viewport_json: &str) -> Result<String, JsValue> {
        let viewport: Viewport = serde_json::from_str(viewport_json)
            .map_err(|error| js_error(format!("bad viewport: {error}")))?;
        let sheet = self.sheet(self.active_sheet)?;
        validate_display_region(sheet, &viewport)?;
        let theme = &self.workbook.styles.theme;
        let owner = sheet.name.clone();
        let display_list =
            build_display_list_with_charts(&self.workbook, self.active_sheet, &viewport, |chart| {
                self.resolve_chart(theme, &owner, chart)
            })
            .map_err(js_error)?;
        serde_json::to_string(&display_list).map_err(js_error)
    }

    #[wasm_bindgen(js_name = chartAtPointJson)]
    pub fn chart_at_point_json(&self, args: &str) -> Result<String, JsValue> {
        let args: ChartHitArgs = serde_json::from_str(args)
            .map_err(|error| js_error(format!("bad chart hit args: {error}")))?;
        let sheet = self.sheet(self.active_sheet)?;
        validate_display_region(sheet, &args.viewport)?;
        let regions = chart_regions(sheet, &args.viewport).map_err(js_error)?;
        let hit: Option<ChartRegion> = chart_at_point(&regions, args.x, args.y).cloned();
        serde_json::to_string(&hit).map_err(js_error)
    }

    #[wasm_bindgen(js_name = cellPositionJson)]
    pub fn cell_position_json(&self, args: &str) -> Result<String, JsValue> {
        let args: CellArgs = serde_json::from_str(args)
            .map_err(|error| js_error(format!("bad cell args: {error}")))?;
        let cell = CellRef::new(args.row, args.col);
        validate_cell(cell)?;
        let sheet = self.sheet(SheetId(args.sheet))?;
        let geometry = GridGeometry::new(sheet);
        let (frozen_rows, frozen_cols) = sheet
            .freeze_pane
            .map_or((0, 0), |pane| (pane.rows, pane.cols));
        serde_json::to_string(&CellPosition {
            x: (geometry.col_x(cell.col) - geometry.col_x(frozen_cols)).max(0.0),
            y: (geometry.row_y(cell.row) - geometry.row_y(frozen_rows)).max(0.0),
        })
        .map_err(js_error)
    }

    #[wasm_bindgen(js_name = cellText)]
    pub fn cell_text(&self, sheet: u32, row: u32, col: u32) -> Result<String, JsValue> {
        let at = CellRef::new(row, col);
        validate_cell(at)?;
        let cell = self
            .sheet(SheetId(sheet))?
            .cell(at)
            .cloned()
            .unwrap_or_default();
        Ok(display_text(
            &self.workbook.styles,
            self.workbook.date_system,
            &cell,
        ))
    }

    #[wasm_bindgen(js_name = mergedRangesJson)]
    pub fn merged_ranges_json(&self, args: &str) -> Result<String, JsValue> {
        let args: RangeArgs = serde_json::from_str(args)
            .map_err(|error| js_error(format!("bad range args: {error}")))?;
        let range = CellRange::parse_a1(&args.range)
            .map_err(|error| js_error(format!("bad range: {error}")))?;
        validate_range(range)?;
        let ranges = self
            .sheet(SheetId(args.sheet))?
            .merges
            .iter()
            .copied()
            .filter(|merged| ranges_intersect(*merged, range))
            .collect();
        serde_json::to_string(&MergedRanges { ranges }).map_err(js_error)
    }

    pub fn version() -> String {
        env!("CARGO_PKG_VERSION").to_owned()
    }
}

impl XlsxViewDocument {
    fn sheet(&self, id: SheetId) -> Result<&Sheet, JsValue> {
        self.workbook
            .sheet(id)
            .ok_or_else(|| js_error(format!("sheet {} is out of range", id.0)))
    }

    fn sheet_info(&self) -> Result<SheetInfo, JsValue> {
        let sheet = self.sheet(self.active_sheet)?;
        let geometry = GridGeometry::new(sheet);
        let used_range = sheet.used_range();
        let mut content_col = used_range
            .map_or(26, |range| range.end.col.saturating_add(2))
            .min(MAX_COLS);
        let mut content_row = used_range
            .map_or(50, |range| range.end.row.saturating_add(2))
            .min(MAX_ROWS);
        let (frozen_rows, frozen_cols, initial_scroll_x, initial_scroll_y) = match sheet.freeze_pane
        {
            Some(pane) => {
                content_col = content_col
                    .max(pane.cols.saturating_add(1))
                    .max(pane.top_left.col.saturating_add(2))
                    .min(MAX_COLS);
                content_row = content_row
                    .max(pane.rows.saturating_add(1))
                    .max(pane.top_left.row.saturating_add(2))
                    .min(MAX_ROWS);
                (
                    pane.rows,
                    pane.cols,
                    (geometry.col_x(pane.top_left.col) - geometry.col_x(pane.cols)).max(0.0),
                    (geometry.row_y(pane.top_left.row) - geometry.row_y(pane.rows)).max(0.0),
                )
            }
            None => (0, 0, 0.0, 0.0),
        };
        Ok(SheetInfo {
            sheet_ids: (0..self.workbook.sheets.len())
                .map(|index| format!("sheet:{index}"))
                .collect(),
            sheet_names: self
                .workbook
                .sheets
                .iter()
                .map(|sheet| sheet.name.clone())
                .collect(),
            active_sheet: self.active_sheet.0,
            content_width: geometry.col_x(content_col),
            content_height: geometry.row_y(content_row),
            frozen_rows,
            frozen_cols,
            initial_scroll_x,
            initial_scroll_y,
        })
    }

    fn resolve_chart(
        &self,
        theme: &xlsx_model::Theme,
        owner: &str,
        chart: &SheetChart,
    ) -> Result<ooxml_drawingml::chart::ChartSpace, RenderError> {
        let bytes =
            self.package
                .part_bytes(&chart.part)
                .ok_or_else(|| RenderError::ChartPartMissing {
                    part: chart.part.clone(),
                })?;
        xlsx_parse::preserved_chart_space(bytes, &self.workbook, owner, theme).ok_or_else(|| {
            RenderError::ChartParseFailed {
                part: chart.part.clone(),
            }
        })
    }
}

fn validate_display_region(sheet: &Sheet, viewport: &Viewport) -> Result<(), JsValue> {
    if !viewport.x.is_finite()
        || !viewport.y.is_finite()
        || !viewport.width.is_finite()
        || !viewport.height.is_finite()
        || viewport.width <= 0.0
        || viewport.height <= 0.0
        || !(viewport.x + viewport.width).is_finite()
        || !(viewport.y + viewport.height).is_finite()
    {
        return Err(js_error("invalid viewport"));
    }
    let geometry = GridGeometry::new(sheet);
    if viewport.x + viewport.width > geometry.col_x(MAX_COLS)
        || viewport.y + viewport.height > geometry.row_y(MAX_ROWS)
    {
        return Err(js_error("viewport is outside worksheet bounds"));
    }
    let (rows, columns) = geometry.viewport_range(viewport);
    let (frozen_rows, frozen_cols) = sheet
        .freeze_pane
        .map_or((0, 0), |pane| (pane.rows, pane.cols));
    let cells = u64::from(rows.end - rows.start)
        .saturating_add(u64::from(frozen_rows))
        .saturating_mul(
            u64::from(columns.end - columns.start).saturating_add(u64::from(frozen_cols)),
        );
    if cells > MAX_DISPLAY_CELLS {
        return Err(js_error(format!(
            "display covers {cells} cells, exceeds {MAX_DISPLAY_CELLS}"
        )));
    }
    Ok(())
}

fn validate_cell(cell: CellRef) -> Result<(), JsValue> {
    if cell.row >= MAX_ROWS || cell.col >= MAX_COLS {
        return Err(js_error("cell is outside worksheet bounds"));
    }
    Ok(())
}

fn validate_range(range: CellRange) -> Result<(), JsValue> {
    validate_cell(range.start)?;
    validate_cell(range.end)
}

fn ranges_intersect(left: CellRange, right: CellRange) -> bool {
    left.start.row <= right.end.row
        && right.start.row <= left.end.row
        && left.start.col <= right.end.col
        && right.start.col <= left.end.col
}

fn js_error(error: impl std::fmt::Display) -> JsValue {
    JsValue::from_str(&error.to_string())
}
