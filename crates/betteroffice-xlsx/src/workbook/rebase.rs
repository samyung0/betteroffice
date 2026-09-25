use super::*;

/// What a checkpoint shows of one sheet, formula caches and source-only sheet
/// data (column styles, tables, array ranges) aside.
#[derive(Debug, PartialEq)]
struct PublishedSheet {
    name: String,
    cells: BTreeMap<(u32, u32), (Option<String>, CellValue, CellFormat)>,
    merges: BTreeSet<(u32, u32, u32, u32)>,
    hyperlinks: Vec<Hyperlink>,
    freeze_pane: Option<FreezePane>,
    row_heights: BTreeMap<u32, f64>,
    col_widths: BTreeMap<u32, f64>,
    charts: BTreeMap<String, ChartAnchor>,
}

fn published(model: &WorkbookModel) -> (Vec<PublishedSheet>, Vec<xlsx_model::DefinedName>) {
    let corners = |range: &CellRange| {
        (
            range.start.row,
            range.start.col,
            range.end.row,
            range.end.col,
        )
    };
    let sheets = model
        .sheets
        .iter()
        .map(|sheet| {
            let mut hyperlinks = sheet.hyperlinks.clone();
            hyperlinks.sort_by_key(|link| corners(&link.range));
            PublishedSheet {
                name: sheet.name.clone(),
                cells: sheet
                    .iter_cells()
                    .map(|(at, cell)| {
                        let value = match cell.formula {
                            Some(_) => CellValue::Empty,
                            None => cell.value.clone(),
                        };
                        let format = model.styles.cell_format(cell.style);
                        ((at.row, at.col), (cell.formula.clone(), value, format))
                    })
                    .filter(|(_, cell)| *cell != (None, CellValue::Empty, CellFormat::default()))
                    .collect(),
                merges: sheet.merges.iter().map(corners).collect(),
                hyperlinks,
                freeze_pane: sheet.freeze_pane,
                row_heights: sheet.row_heights.clone(),
                col_widths: sheet.col_widths.clone(),
                charts: sheet
                    .charts
                    .iter()
                    .map(|chart| (chart.frame_id(), chart.anchor))
                    .collect(),
            }
        })
        .collect();
    (sheets, model.defined_names.clone())
}

impl Workbook {
    /// Re-expresses `latest` as overrides over `new_source`, the publication of
    /// `captured`: sheet, row and column changes made since the capture replay
    /// as structural edits and everything else as overrides. Fails when the
    /// result does not reproduce `latest`.
    pub fn rebase_checkpoint(
        old_source: &[u8],
        captured: &[u8],
        latest: &[u8],
        new_source: &[u8],
        client_id: u64,
    ) -> Result<Vec<u8>> {
        let options = CalculationOptions::default();
        let mut before = Self::open_collaborative(old_source, client_id)?;
        before.apply_update_v1(captured, options)?;
        let mut current = Self::open_collaborative(old_source, client_id)?;
        current.apply_update_v1(latest, options)?;
        let rebased = Self::open_collaborative(new_source, client_id)?;
        if before.model.sheets.len() != rebased.model.sheets.len()
            || before
                .model
                .sheets
                .iter()
                .zip(&rebased.model.sheets)
                .any(|(a, b)| a.name != b.name)
        {
            return Err(Error::CollaborativeState(
                "published workbook does not match captured sheet order".into(),
            ));
        }
        if before.model == current.model {
            return Ok(rebased.encode_state_as_update_v1());
        }
        rebased
            .authority
            .rebase(&before.authority, &current.authority, &current.model)
            .map_err(authority_error)?;
        let state = rebased.encode_state_as_update_v1();
        validate_collaboration_size(&state)?;
        let checker = if client_id == 1 { 2 } else { client_id - 1 };
        let mut check = Self::open_collaborative(new_source, checker)?;
        check.apply_update_v1(&state, options)?;
        if published(&check.model) != published(&current.model) {
            return Err(Error::CollaborativeState(
                "rebased checkpoint does not reproduce the latest workbook".into(),
            ));
        }
        Ok(state)
    }
}
