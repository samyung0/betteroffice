use super::topology::{Axis, Change, Point, Span};
use super::*;
use serde::{Deserialize, Serialize};
use xlsx_model::{AnchorCell, CellRange, CellRef};
use xlsx_ops::ReferenceAddress;

pub(super) const VERSION: i64 = 7;
pub(super) const CATALOG: &str = "xlsx:axis-catalog";
pub(super) const DEFINED_NAMES: &str = "xlsx:defined-names";
const ROWS: &str = "rows";
const COLS: &str = "cols";

#[derive(Clone)]
struct Context {
    keys: Vec<String>,
    names: Vec<String>,
    axes: BTreeMap<String, (Axis, Axis)>,
}
impl Context {
    fn index(&self, key: &str) -> Option<usize> {
        self.keys.iter().position(|item| item == key)
    }
    fn sheet(&self, sheet: SheetId) -> Result<&str, String> {
        self.keys
            .get(sheet.0 as usize)
            .map(String::as_str)
            .ok_or("sheet index out of bounds".into())
    }
    fn axes(&self, key: &str) -> Result<&(Axis, Axis), String> {
        self.axes.get(key).ok_or("missing sheet axes".into())
    }
    fn key(&self, sheet: &str, at: CellRef) -> Result<String, String> {
        let (rows, cols) = self.axes(sheet)?;
        serde_json::to_string(&(rows.at(at.row)?, cols.at(at.col)?))
            .map_err(|error| error.to_string())
    }
    fn at(&self, sheet: &str, key: &str) -> Result<Option<CellRef>, String> {
        let (row, col): (Point, Point) = serde_json::from_str(key)
            .map_err(|error| format!("invalid stable cell key: {error}"))?;
        let (rows, cols) = self.axes(sheet)?;
        rows.validate_point(&row)?;
        cols.validate_point(&col)?;
        Ok(rows
            .index(&row)
            .zip(cols.index(&col))
            .map(|(row, col)| CellRef::new(row, col)))
    }
}

fn context_fingerprint(context: &Context) -> Result<String, String> {
    let topology = context
        .keys
        .iter()
        .map(|key| {
            context
                .axes
                .get(key)
                .map(|(rows, cols)| (key, &rows.spans, &cols.spans))
                .ok_or("missing axes")
        })
        .collect::<Result<Vec<_>, _>>()?;
    Ok(format!(
        "{:x}",
        Sha256::digest(json(&(&context.names, topology))?)
    ))
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(deny_unknown_fields)]
struct BoundRange {
    rows: Vec<Span>,
    cols: Vec<Span>,
    flags: [bool; 4],
}
impl BoundRange {
    fn validate(&self, context: &Context, sheet: &str) -> Result<(), String> {
        let (rows, cols) = context.axes(sheet)?;
        if self.rows.is_empty() || self.cols.is_empty() {
            return Err("empty bound range".into());
        }
        rows.validate_spans(&self.rows)?;
        cols.validate_spans(&self.cols)
    }
    fn bind(context: &Context, sheet: &str, range: CellRange) -> Result<Self, String> {
        let (rows, cols) = context.axes(sheet)?;
        Ok(Self {
            rows: rows.range(range.start.row, range.end.row)?,
            cols: cols.range(range.start.col, range.end.col)?,
            flags: [
                range.start.abs_row,
                range.start.abs_col,
                range.end.abs_row,
                range.end.abs_col,
            ],
        })
    }
    fn clip_start(&self, context: &Context, sheet: &str) -> Option<CellRef> {
        let (rows, cols) = context.axes.get(sheet)?;
        let row = self.rows.first()?;
        let col = self.cols.first()?;
        Some(CellRef::new(
            rows.clip(&Point {
                run: row.run.clone(),
                offset: row.start,
            })?,
            cols.clip(&Point {
                run: col.run.clone(),
                offset: col.start,
            })?,
        ))
    }
    fn resolve(&self, context: &Context, sheet: &str) -> Option<CellRange> {
        let (rows, cols) = context.axes.get(sheet)?;
        let (row, end_row) = rows.resolve_range(&self.rows)?;
        let (col, end_col) = cols.resolve_range(&self.cols)?;
        let mut start = CellRef::new(row, col);
        let mut end = CellRef::new(end_row, end_col);
        start.abs_row = self.flags[0];
        start.abs_col = self.flags[1];
        end.abs_row = self.flags[2];
        end.abs_col = self.flags[3];
        Some(CellRange::new(start, end))
    }
}
#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(deny_unknown_fields)]
struct Binding {
    start: usize,
    end: usize,
    sheet: String,
    qualified: bool,
    span: Option<Vec<(String, BoundRange)>>,
    range: BoundRange,
    original: CellRange,
    original_sheet_name: Option<String>,
    axis: Option<bool>,
    prefix: String,
    suffix: String,
}
#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(deny_unknown_fields)]
struct Formula {
    text: String,
    bindings: Vec<Binding>,
    opaque_context: Option<String>,
}
impl Formula {
    fn bind(text: &str, owner: Option<&str>, context: &Context) -> Result<Self, String> {
        let mut bindings = Vec::new();
        let references = match xlsx_ops::formula_references(text) {
            Ok(references) => references,
            Err(_) => {
                return Ok(Self {
                    text: text.to_owned(),
                    bindings,
                    opaque_context: Some(context_fingerprint(context)?),
                });
            }
        };
        for reference in references {
            let span_keys = reference
                .sheet
                .as_ref()
                .and_then(|name| name.split_once(':'))
                .and_then(|(first, last)| {
                    let first = context
                        .names
                        .iter()
                        .position(|item| item.eq_ignore_ascii_case(first))?;
                    let last = context
                        .names
                        .iter()
                        .position(|item| item.eq_ignore_ascii_case(last))?;
                    Some(context.keys[first.min(last)..=first.max(last)].to_vec())
                });
            let sheet = match &reference.sheet {
                Some(name) => span_keys
                    .as_ref()
                    .and_then(|keys| keys.first().map(String::as_str))
                    .or_else(|| {
                        context
                            .names
                            .iter()
                            .position(|item| item.eq_ignore_ascii_case(name))
                            .map(|index| context.keys[index].as_str())
                    }),
                None => {
                    owner.or_else(|| (context.keys.len() == 1).then(|| context.keys[0].as_str()))
                }
            };
            let Some(sheet) = sheet else {
                continue;
            };
            let (range, axis) = match reference.address {
                ReferenceAddress::Cells(range) => (range, None),
                ReferenceAddress::Rows {
                    start,
                    end,
                    start_absolute,
                    end_absolute,
                } => {
                    let mut range =
                        CellRange::new(CellRef::new(start, 0), CellRef::new(end, MAX_COLS - 1));
                    range.start.abs_row = start_absolute;
                    range.end.abs_row = end_absolute;
                    (range, Some(true))
                }
                ReferenceAddress::Cols {
                    start,
                    end,
                    start_absolute,
                    end_absolute,
                } => {
                    let mut range =
                        CellRange::new(CellRef::new(0, start), CellRef::new(MAX_ROWS - 1, end));
                    range.start.abs_col = start_absolute;
                    range.end.abs_col = end_absolute;
                    (range, Some(false))
                }
            };
            let span = span_keys
                .as_ref()
                .map(|keys| {
                    keys.iter()
                        .map(|key| Ok((key.clone(), BoundRange::bind(context, key, range)?)))
                        .collect::<Result<Vec<_>, String>>()
                })
                .transpose()?;
            bindings.push(Binding {
                span,
                start: reference.start,
                end: reference.end,
                sheet: sheet.to_owned(),
                qualified: reference.sheet.is_some(),
                range: BoundRange::bind(context, sheet, range)?,
                original: range,
                original_sheet_name: reference.sheet,
                axis,
                prefix: reference.prefix,
                suffix: reference.suffix,
            });
        }
        Ok(Self {
            text: text.to_owned(),
            bindings,
            opaque_context: None,
        })
    }
    fn resolve(&self, context: &Context) -> Result<String, String> {
        if let Some(fingerprint) = &self.opaque_context {
            if *fingerprint != context_fingerprint(context)? {
                return Err("formula uses reference syntax that cannot follow collaborative structural edits".into());
            }
        }
        let mut out = String::new();
        let mut cursor = 0;
        for binding in &self.bindings {
            binding.range.validate(context, &binding.sheet)?;
            if binding.start < cursor
                || binding.end > self.text.len()
                || !self.text.is_char_boundary(binding.start)
                || !self.text.is_char_boundary(binding.end)
            {
                return Err("invalid formula binding span".into());
            }
            out.push_str(&self.text[cursor..binding.start]);
            let mut sheet = binding.sheet.as_str();
            let mut last_sheet = None;
            let range = if let Some(span) = &binding.span {
                let mut surviving = span
                    .iter()
                    .filter(|(key, _)| context.index(key).is_some())
                    .collect::<Vec<_>>();
                surviving.sort_by_key(|(key, _)| context.index(key));
                if let Some((key, range)) = surviving.first() {
                    sheet = key;
                    last_sheet = surviving.last().map(|(key, _)| key.as_str());
                    let resolved = range.resolve(context, key);
                    for (other, range) in &surviving {
                        if range.resolve(context, other) != resolved {
                            return Err("a 3-D reference cannot represent different row or column topology on its sheets".into());
                        }
                    }
                    resolved
                } else {
                    None
                }
            } else {
                context
                    .index(sheet)
                    .and_then(|_| binding.range.resolve(context, sheet))
            };
            if let Some(range) = range {
                let current_name = context.index(sheet).map(|index| {
                    let first = context.names[index].clone();
                    match last_sheet
                        .filter(|last| *last != sheet)
                        .and_then(|last| context.index(last))
                    {
                        Some(last) => format!("{first}:{}", context.names[last]),
                        None => first,
                    }
                });
                let same_sheet = !binding.qualified
                    || binding.original_sheet_name.as_deref() == current_name.as_deref();
                if same_sheet && range == binding.original {
                    out.push_str(&self.text[binding.start..binding.end]);
                    cursor = binding.end;
                    continue;
                }
                out.push_str(&binding.prefix);
                if binding.qualified {
                    let name = current_name.as_deref().ok_or("missing reference sheet")?;
                    out.push('\'');
                    out.push_str(&name.replace('\'', "''"));
                    out.push_str("'!");
                }
                match binding.axis {
                    None => out.push_str(&range.to_a1()),
                    Some(true) => out.push_str(&format!(
                        "{}{}:{}{}",
                        if range.start.abs_row { "$" } else { "" },
                        range.start.row + 1,
                        if range.end.abs_row { "$" } else { "" },
                        range.end.row + 1
                    )),
                    Some(false) => out.push_str(&format!(
                        "{}{}:{}{}",
                        if range.start.abs_col { "$" } else { "" },
                        xlsx_model::addr::col_to_letters(range.start.col),
                        if range.end.abs_col { "$" } else { "" },
                        xlsx_model::addr::col_to_letters(range.end.col)
                    )),
                }
                out.push_str(&binding.suffix);
            } else {
                out.push_str("#REF!");
            }
            cursor = binding.end;
        }
        out.push_str(&self.text[cursor..]);
        Ok(out)
    }
}
#[derive(Clone, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct Content {
    value: CellValue,
    formula: Option<Formula>,
}
#[derive(Clone, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct Link {
    value: Hyperlink,
    range: BoundRange,
    location: Option<Formula>,
}
#[derive(Clone, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct Name {
    value: DefinedName,
    sheet: Option<String>,
    formula: Formula,
}
#[derive(Clone, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct Chart {
    value: SheetChart,
    refs: Vec<Formula>,
    from: Option<BoundRange>,
    to: Option<BoundRange>,
}
#[derive(Clone, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct Freeze {
    value: FreezePane,
    rows: Vec<Span>,
    cols: Vec<Span>,
    top_left: BoundRange,
}

fn json<T: Serialize>(value: &T) -> Result<String, String> {
    serde_json::to_string(value).map_err(|error| error.to_string())
}
fn decode<T: for<'a> Deserialize<'a>>(value: Out) -> Result<T, String> {
    let value = value
        .cast::<String>()
        .map_err(|_| "stable record must be JSON text")?;
    serde_json::from_str(&value).map_err(|error| format!("invalid stable record: {error}"))
}
fn set_json<T: Serialize>(
    map: &MapRef,
    txn: &mut TransactionMut<'_>,
    key: &str,
    value: &T,
) -> Result<(), String> {
    map.try_update(txn, key, json(value)?);
    Ok(())
}
fn map<T: ReadTxn>(txn: &T, name: &str) -> Result<MapRef, String> {
    txn.get_map(name).ok_or_else(|| format!("missing {name}"))
}
fn sheet_map<T: ReadTxn>(txn: &T, key: &str) -> Result<MapRef, String> {
    nested_map(&map(txn, SHEETS)?, txn, key)
}
fn axis<T: ReadTxn>(
    txn: &T,
    sheet: &MapRef,
    key: &str,
    axis: &str,
    limit: u32,
) -> Result<Axis, String> {
    let catalog = map(txn, CATALOG)?;
    let records = nested_map(&nested_map(&catalog, txn, key)?, txn, axis)?;
    let changes = records
        .iter(txn)
        .map(|(id, value)| Ok((id.to_owned(), decode::<Change>(value)?)))
        .collect::<Result<BTreeMap<_, _>, String>>()?;
    let live = nested_map(sheet, txn, axis)?;
    let active = live
        .iter(txn)
        .map(|(id, value)| {
            if value != Out::Any(Any::Bool(true)) || !changes.contains_key(id) {
                return Err("invalid active axis change".into());
            }
            Ok(id.to_owned())
        })
        .collect::<Result<BTreeSet<_>, String>>()?;
    Axis::project(limit, &changes, &active)
}
fn context<T: ReadTxn>(txn: &T) -> Result<Context, String> {
    let order = txn.get_array(SHEET_ORDER).ok_or("missing sheet order")?;
    let sheets = map(txn, SHEETS)?;
    let mut keys = sheet_keys(&order, txn)?;
    let mut seen = BTreeSet::new();
    keys.retain(|key| seen.insert(key.clone()));
    if keys.is_empty() {
        keys.push(
            sheets
                .keys(txn)
                .min()
                .ok_or("missing retained worksheet")?
                .to_owned(),
        );
    }
    let originals = keys
        .iter()
        .map(|key| {
            let sheet = nested_map(&sheets, txn, key)?;
            let name = sheet
                .get(txn, NAME)
                .and_then(|value| value.cast::<String>().ok())
                .ok_or("missing sheet name")?;
            Ok((key.clone(), name))
        })
        .collect::<Result<BTreeMap<_, _>, String>>()?;
    let mut reserved = originals
        .values()
        .map(|name| name.to_lowercase())
        .collect::<BTreeSet<_>>();
    let mut assigned = BTreeSet::new();
    let mut resolved = BTreeMap::new();
    for (key, original) in &originals {
        let mut name = original.clone();
        if !assigned.insert(original.to_lowercase()) {
            let digest = format!("{:x}", Sha256::digest(key.as_bytes()));
            let mut attempt = 0;
            loop {
                let suffix = if attempt == 0 {
                    format!(" ({})", &digest[..8])
                } else {
                    format!(" ({}-{attempt})", &digest[..8])
                };
                name = format!(
                    "{}{}",
                    original
                        .chars()
                        .take(31usize.saturating_sub(suffix.len()))
                        .collect::<String>(),
                    suffix
                );
                if reserved.insert(name.to_lowercase()) {
                    break;
                }
                attempt += 1;
            }
        }
        resolved.insert(key.clone(), name);
    }
    let mut names = Vec::new();
    let mut axes = BTreeMap::new();
    for key in &keys {
        let sheet = nested_map(&sheets, txn, key)?;
        names.push(resolved[key].clone());
        axes.insert(
            key.clone(),
            (
                axis(txn, &sheet, key, ROWS, MAX_ROWS)?,
                axis(txn, &sheet, key, COLS, MAX_COLS)?,
            ),
        );
    }
    for (key, _) in sheets.iter(txn) {
        if !axes.contains_key(key) {
            let sheet = nested_map(&sheets, txn, key)?;
            axes.insert(
                key.to_owned(),
                (
                    axis(txn, &sheet, key, ROWS, MAX_ROWS)?,
                    axis(txn, &sheet, key, COLS, MAX_COLS)?,
                ),
            );
        }
    }
    Ok(Context { keys, names, axes })
}
fn blank_sheet(txn: &mut TransactionMut<'_>, key: &str, name: &str) -> Result<MapRef, String> {
    let sheets = txn.get_or_insert_map(SHEETS);
    let sheet = sheets.insert(txn, key, MapPrelim::default());
    sheet.insert(txn, NAME, name);
    for name in [
        CONTENTS,
        STYLES,
        ROW_HEIGHTS,
        COL_WIDTHS,
        ROWS,
        COLS,
        MERGES,
        HYPERLINKS,
        CHARTS,
    ] {
        sheet.insert(txn, name, MapPrelim::default());
    }
    let catalog = txn.get_or_insert_map(CATALOG);
    let catalog_sheet = catalog.insert(txn, key, MapPrelim::default());
    catalog_sheet.insert(txn, ROWS, MapPrelim::default());
    catalog_sheet.insert(txn, COLS, MapPrelim::default());
    catalog_sheet.insert(txn, CHARTS, MapPrelim::default());
    Ok(sheet)
}
fn chart_bind(value: &SheetChart, key: &str, context: &Context) -> Result<Chart, String> {
    let bind = |cell: AnchorCell| {
        BoundRange::bind(
            context,
            key,
            CellRange::new(
                CellRef::new(cell.row, cell.col),
                CellRef::new(cell.row, cell.col),
            ),
        )
    };
    let (from, to) = match value.anchor {
        ChartAnchor::TwoCell { from, to, .. } => (Some(bind(from)?), Some(bind(to)?)),
        ChartAnchor::OneCell { from, .. } => (Some(bind(from)?), None),
        ChartAnchor::Absolute { .. } => (None, None),
    };
    Ok(Chart {
        value: value.clone(),
        refs: value
            .refs
            .iter()
            .map(|reference| Formula::bind(&reference.formula, Some(key), context))
            .collect::<Result<_, _>>()?,
        from,
        to,
    })
}
fn write_cell(
    txn: &mut TransactionMut<'_>,
    sheet: &MapRef,
    key: &str,
    at: CellRef,
    cell: &Cell,
    context: &Context,
    styles: &Stylesheet,
    content: bool,
    format: bool,
) -> Result<(), String> {
    let cell_key = context.key(key, at)?;
    if content {
        let contents = nested_map(sheet, txn, CONTENTS)?;
        if cell.formula.is_none() && matches!(cell.value, CellValue::Empty) {
            contents.remove(txn, &cell_key);
        } else {
            set_json(
                &contents,
                txn,
                &cell_key,
                &Content {
                    value: cell.value.clone(),
                    formula: cell
                        .formula
                        .as_ref()
                        .map(|formula| Formula::bind(formula, Some(key), context))
                        .transpose()?,
                },
            )?;
        }
    }
    if format {
        let formats = nested_map(sheet, txn, STYLES)?;
        match cell.style {
            Some(style) => {
                formats.try_update(txn, cell_key, style_key(styles, style)?);
            }
            None => {
                formats.remove(txn, &cell_key);
            }
        }
    }
    Ok(())
}
fn write_names(
    txn: &mut TransactionMut<'_>,
    names: &[DefinedName],
    context: &Context,
) -> Result<(), String> {
    let map = txn.get_or_insert_map(DEFINED_NAMES);
    let mut desired = BTreeMap::from([("$schema".to_owned(), Any::BigInt(VERSION))]);
    for name in names {
        let sheet = name
            .local_sheet
            .map(|sheet| context.sheet(sheet).map(str::to_owned))
            .transpose()?;
        let key = json(&(sheet.as_deref(), name.name.to_lowercase()))?;
        desired.insert(
            key,
            Any::from(json(&Name {
                value: name.clone(),
                formula: Formula::bind(&name.formula, sheet.as_deref(), context)?,
                sheet,
            })?),
        );
    }
    sync_map(&map, txn, desired);
    Ok(())
}
pub(super) fn seed(
    doc: &Doc,
    base: &WorkbookBase,
    model: &WorkbookModel,
    keys: &[String],
) -> Result<(), String> {
    let mut txn = doc.transact_mut_with(BOOTSTRAP_ORIGIN);
    let formats = txn.get_or_insert_map(CELL_FORMATS);
    sync_cell_formats(&formats, &mut txn, &model.styles)?;
    let meta = txn.get_or_insert_map(META);
    meta.insert(&mut txn, BASE_FINGERPRINT, base.fingerprint.as_str());
    meta.insert(&mut txn, "schemaVersion", VERSION);
    meta.insert(&mut txn, STRUCTURE_GENERATION, 0_i64);
    let order = txn.get_or_insert_array(SHEET_ORDER);
    order.insert_range(&mut txn, 0, keys.iter().cloned());
    txn.get_or_insert_map(DEFINED_NAMES);
    for (key, sheet) in keys.iter().zip(&model.sheets) {
        blank_sheet(&mut txn, key, &sheet.name)?;
    }
    let context = context(&txn)?;
    for (key, value) in keys.iter().zip(&model.sheets) {
        let sheet = sheet_map(&txn, key)?;
        for (at, cell) in value.iter_cells() {
            write_cell(
                &mut txn,
                &sheet,
                key,
                at,
                cell,
                &context,
                &model.styles,
                true,
                true,
            )?;
        }
        let (rows, cols) = context.axes(key)?;
        let heights = nested_map(&sheet, &txn, ROW_HEIGHTS)?;
        for (row, value) in &value.row_heights {
            heights.insert(&mut txn, json(&rows.at(*row)?)?, *value);
        }
        let widths = nested_map(&sheet, &txn, COL_WIDTHS)?;
        for (col, value) in &value.col_widths {
            widths.insert(&mut txn, json(&cols.at(*col)?)?, *value);
        }
        write_freeze(&mut txn, &sheet, key, value.freeze_pane, &context)?;
        let merges = nested_map(&sheet, &txn, MERGES)?;
        for (index, range) in value.merges.iter().enumerate() {
            set_json(
                &merges,
                &mut txn,
                &format!("base:{index}"),
                &BoundRange::bind(&context, key, *range)?,
            )?;
        }
        write_links(&mut txn, &sheet, key, &value.hyperlinks, &context)?;
        let charts = nested_map(&nested_map(&map(&txn, CATALOG)?, &txn, key)?, &txn, CHARTS)?;
        for chart in &value.charts {
            set_json(
                &charts,
                &mut txn,
                &chart.frame_id(),
                &chart_bind(chart, key, &context)?,
            )?;
        }
    }
    write_names(&mut txn, &model.defined_names, &context)
}
fn write_freeze(
    txn: &mut TransactionMut<'_>,
    sheet: &MapRef,
    key: &str,
    pane: Option<FreezePane>,
    context: &Context,
) -> Result<(), String> {
    if let Some(value) = pane {
        let (rows, cols) = context.axes(key)?;
        let rows = if value.rows == 0 {
            Vec::new()
        } else {
            rows.range(0, value.rows - 1)?
        };
        let cols = if value.cols == 0 {
            Vec::new()
        } else {
            cols.range(0, value.cols - 1)?
        };
        set_json(
            sheet,
            txn,
            FREEZE_PANE,
            &Freeze {
                value,
                rows,
                cols,
                top_left: BoundRange::bind(
                    context,
                    key,
                    CellRange::new(value.top_left, value.top_left),
                )?,
            },
        )
    } else {
        sheet.remove(txn, FREEZE_PANE);
        Ok(())
    }
}
fn write_links(
    txn: &mut TransactionMut<'_>,
    sheet: &MapRef,
    key: &str,
    links: &[Hyperlink],
    context: &Context,
) -> Result<(), String> {
    let map = nested_map(sheet, txn, HYPERLINKS)?;
    let mut desired = BTreeMap::new();
    for value in links {
        let range = BoundRange::bind(context, key, value.range)?;
        let location = value
            .location
            .as_ref()
            .map(|location| Formula::bind(location, Some(key), context))
            .transpose()?;
        desired.insert(
            json(&range)?,
            Any::from(json(&Link {
                value: value.clone(),
                range,
                location,
            })?),
        );
    }
    sync_map(&map, txn, desired);
    Ok(())
}

pub(super) fn materialize<T: ReadTxn>(
    txn: &T,
    base: &WorkbookBase,
) -> Result<(WorkbookModel, WorkbookStructure), String> {
    let mut roots = vec![
        CELL_FORMATS,
        META,
        SHEET_ORDER,
        SHEETS,
        CATALOG,
        DEFINED_NAMES,
    ];
    if let Some(data) = &base.rebase {
        roots.push(crate::workbook::rebase::ROOT);
        data.validate(txn)?;
    }
    require_root_keys(txn, &roots)?;
    let meta = map(txn, META)?;
    let fingerprint = meta
        .get(txn, BASE_FINGERPRINT)
        .and_then(|value| value.cast::<String>().ok())
        .ok_or("missing base fingerprint")?;
    if !base.accepts_fingerprint(VERSION, &fingerprint) {
        return Err("workbook base fingerprint does not match shared state".into());
    }
    require_map_keys(
        &meta,
        txn,
        &[BASE_FINGERPRINT, "schemaVersion", STRUCTURE_GENERATION],
        "workbook metadata",
    )?;
    let context = context(txn)?;
    let sheets = map(txn, SHEETS)?;
    for (key, _) in sheets.iter(txn) {
        let sheet = nested_map(&sheets, txn, key)?;
        let catalog = nested_map(&map(txn, CATALOG)?, txn, key)?;
        require_map_keys(
            &catalog,
            txn,
            &[ROWS, COLS, CHARTS],
            "sheet topology catalog",
        )?;
        let BranchID::Nested(start) = sheet.as_ref().id() else {
            return Err("invalid sheet container identity".into());
        };
        let containers = [
            (nested_map(&sheet, txn, CONTENTS)?, 2),
            (nested_map(&sheet, txn, STYLES)?, 3),
            (nested_map(&sheet, txn, ROW_HEIGHTS)?, 4),
            (nested_map(&sheet, txn, COL_WIDTHS)?, 5),
            (nested_map(&sheet, txn, ROWS)?, 6),
            (nested_map(&sheet, txn, COLS)?, 7),
            (nested_map(&sheet, txn, MERGES)?, 8),
            (nested_map(&sheet, txn, HYPERLINKS)?, 9),
            (nested_map(&sheet, txn, CHARTS)?, 10),
            (catalog.clone(), 11),
            (nested_map(&catalog, txn, ROWS)?, 12),
            (nested_map(&catalog, txn, COLS)?, 13),
            (nested_map(&catalog, txn, CHARTS)?, 14),
        ];
        for (container, offset) in containers {
            if container.as_ref().id()
                != BranchID::Nested(yrs::ID::new(start.client, start.clock + offset))
            {
                return Err("sheet container identities were replaced".into());
            }
        }
        let source_index = key
            .strip_prefix("sheet:")
            .and_then(|index| index.parse::<usize>().ok());
        if source_index.is_some_and(|index| {
            index >= base.charts.len() || start.client.get() != base.bootstrap_client_id
        }) {
            return Err("source sheet identity does not match the exact package".into());
        }
        if source_index.is_none() && key != format!("replica:{}:{}", start.client, start.clock) {
            return Err("new sheet identity does not match its creation".into());
        }
        let source_charts = source_index
            .map(|index| base.charts[index].as_slice())
            .unwrap_or(&[]);
        for (baseline, records) in [
            (true, nested_map(&catalog, txn, CHARTS)?),
            (false, nested_map(&sheet, txn, CHARTS)?),
        ] {
            if baseline && records.len(txn) != source_charts.len() as u32 {
                return Err("source chart catalog changed".into());
            }
            for (id, value) in records.iter(txn) {
                let chart: Chart = decode(value)?;
                let authored = source_charts
                    .iter()
                    .find(|original| original.frame_id() == id)
                    .ok_or("unknown chart frame")?;
                if baseline && &chart.value != authored {
                    return Err("source chart catalog changed".into());
                }
                if chart.value.part != authored.part
                    || chart.value.drawing != authored.drawing
                    || chart.value.anchor_index != authored.anchor_index
                    || AnchorShape::of(&chart.value.anchor) != AnchorShape::of(&authored.anchor)
                    || chart.value.refs.len() != authored.refs.len()
                    || chart.refs.len() != authored.refs.len()
                {
                    return Err("source chart identity changed".into());
                }
                for range in [&chart.from, &chart.to].into_iter().flatten() {
                    range.validate(&context, key)?;
                }
                for formula in &chart.refs {
                    formula.resolve(&context)?;
                }
            }
        }
        require_map_keys_with_optional(
            &sheet,
            txn,
            &[
                NAME,
                CONTENTS,
                STYLES,
                ROW_HEIGHTS,
                COL_WIDTHS,
                ROWS,
                COLS,
                MERGES,
                HYPERLINKS,
                CHARTS,
            ],
            &[FREEZE_PANE],
            "stable sheet",
        )?;
        for (cell, value) in nested_map(&sheet, txn, CONTENTS)?.iter(txn) {
            context.at(key, cell)?;
            let content: Content = decode(value)?;
            if let Some(formula) = content.formula {
                formula.resolve(&context)?;
            }
        }
        for (_, value) in nested_map(&sheet, txn, MERGES)?.iter(txn) {
            decode::<BoundRange>(value)?.validate(&context, key)?;
        }
        for (_, value) in nested_map(&sheet, txn, HYPERLINKS)?.iter(txn) {
            let link: Link = decode(value)?;
            link.range.validate(&context, key)?;
            if let Some(location) = link.location {
                location.resolve(&context)?;
            }
        }
    }
    let (styles, style_indices) =
        materialize_cell_formats(&map(txn, CELL_FORMATS)?, txn, &base.styles)?;
    let mut model = base.workbook();
    model.styles = styles;
    model.defined_names.clear();
    for (index, key) in context.keys.iter().enumerate() {
        let source = sheet_map(txn, key)?;
        let mut sheet = Sheet::new(context.names[index].clone());
        for (cell_key, value) in nested_map(&source, txn, CONTENTS)?.iter(txn) {
            let content: Content = decode(value)?;
            let Some(at) = context.at(key, cell_key)? else {
                continue;
            };
            let formula = content
                .formula
                .map(|formula| formula.resolve(&context))
                .transpose()?;
            sheet.set_cell(
                at,
                Cell {
                    value: content.value,
                    formula,
                    style: None,
                },
            );
        }
        for (cell_key, value) in nested_map(&source, txn, STYLES)?.iter(txn) {
            let style = value
                .cast::<String>()
                .map_err(|_| "invalid cell style key")?;
            let style = *style_indices.get(&style).ok_or("unknown cell style")?;
            let Some(at) = context.at(key, cell_key)? else {
                continue;
            };
            let mut cell = sheet.cell(at).cloned().unwrap_or_default();
            cell.style = style;
            sheet.set_cell(at, cell);
        }
        let (rows, cols) = context.axes(key)?;
        for (name, axis, values) in [
            (ROW_HEIGHTS, rows, &mut sheet.row_heights),
            (COL_WIDTHS, cols, &mut sheet.col_widths),
        ] {
            for (point, value) in nested_map(&source, txn, name)?.iter(txn) {
                let point: Point = serde_json::from_str(point)
                    .map_err(|error| format!("invalid dimension identity: {error}"))?;
                axis.validate_point(&point)?;
                let value = value.cast::<f64>().map_err(|_| "invalid axis dimension")?;
                if !value.is_finite() || value < 0.0 {
                    return Err("invalid axis dimension".into());
                }
                if let Some(index) = axis.index(&point) {
                    values.insert(index, value);
                }
            }
        }
        if let Some(value) = source.get(txn, FREEZE_PANE) {
            let freeze: Freeze = decode(value)?;
            rows.validate_spans(&freeze.rows)?;
            cols.validate_spans(&freeze.cols)?;
            freeze.top_left.validate(&context, key)?;
            let mut value = freeze.value;
            value.rows = rows
                .resolve_range(&freeze.rows)
                .map_or(0, |(_, end)| end + 1);
            value.cols = cols
                .resolve_range(&freeze.cols)
                .map_or(0, |(_, end)| end + 1);
            value.top_left = freeze
                .top_left
                .clip_start(&context, key)
                .ok_or("freeze pane anchor has no topology identity")?;
            sheet.freeze_pane = Some(value);
        }
        let mut merges = nested_map(&source, txn, MERGES)?
            .iter(txn)
            .map(|(id, value)| Ok((id.to_owned(), decode::<BoundRange>(value)?)))
            .collect::<Result<Vec<_>, String>>()?;
        merges.sort_by(|left, right| left.0.cmp(&right.0));
        for (_, merge) in merges {
            if let Some(range) = merge.resolve(&context, key) {
                if sheet.merges.iter().any(|other| intersects(*other, range)) {
                    continue;
                }
                sheet.merges.push(range);
            }
        }
        let mut links = nested_map(&source, txn, HYPERLINKS)?
            .iter(txn)
            .map(|(id, value)| Ok((id.to_owned(), decode::<Link>(value)?)))
            .collect::<Result<Vec<_>, String>>()?;
        links.sort_by(|left, right| left.0.cmp(&right.0));
        for (_, mut link) in links {
            if let Some(range) = link.range.resolve(&context, key) {
                link.value.range = range;
                link.value.location = link
                    .location
                    .map(|value| value.resolve(&context))
                    .transpose()?;
                sheet.hyperlinks.push(link.value);
            }
        }
        let baseline = nested_map(&nested_map(&map(txn, CATALOG)?, txn, key)?, txn, CHARTS)?;
        let mut chart_records = baseline
            .iter(txn)
            .map(|(id, value)| Ok((id.to_owned(), decode::<Chart>(value)?)))
            .collect::<Result<BTreeMap<_, _>, String>>()?;
        for (id, value) in nested_map(&source, txn, CHARTS)?.iter(txn) {
            chart_records.insert(id.to_owned(), decode::<Chart>(value)?);
        }
        let mut charts = chart_records.into_iter().collect::<Vec<_>>();
        charts.sort_by(|left, right| left.0.cmp(&right.0));
        for (_, mut chart) in charts {
            if chart.refs.len() != chart.value.refs.len() {
                return Err("chart reference binding count mismatch".into());
            }
            for (target, formula) in chart.value.refs.iter_mut().zip(chart.refs) {
                target.formula = formula.resolve(&context)?;
            }
            let resolve = |cell: &mut AnchorCell, range: &Option<BoundRange>| {
                if let Some(point) = range
                    .as_ref()
                    .and_then(|range| range.clip_start(&context, key))
                {
                    cell.row = point.row;
                    cell.col = point.col;
                }
            };
            match &mut chart.value.anchor {
                ChartAnchor::TwoCell { from, to, .. } => {
                    resolve(from, &chart.from);
                    resolve(to, &chart.to);
                }
                ChartAnchor::OneCell { from, .. } => resolve(from, &chart.from),
                ChartAnchor::Absolute { .. } => {}
            }
            sheet.charts.push(chart.value);
        }
        model.sheets.push(sheet);
    }
    for (id, value) in map(txn, DEFINED_NAMES)?.iter(txn) {
        if id == "$schema" {
            if value != Out::Any(Any::BigInt(VERSION)) {
                return Err("invalid defined names schema".into());
            }
            continue;
        }
        let mut name: Name = decode(value)?;
        if let Some(sheet) = &name.sheet {
            let Some(index) = context.index(sheet) else {
                continue;
            };
            name.value.local_sheet = Some(SheetId(index as u32));
        }
        name.value.formula = name.formula.resolve(&context)?;
        model.defined_names.push(name.value);
    }
    model.defined_names.sort_by(|left, right| {
        (left.local_sheet.map(|id| id.0), &left.name)
            .cmp(&(right.local_sheet.map(|id| id.0), &right.name))
    });
    project_shared_frame_anchors(&mut model.sheets);
    let axis_changes = context
        .axes
        .iter()
        .map(|(key, (rows, cols))| (key.clone(), (rows.first_changed(), cols.first_changed())))
        .collect();
    let structure = WorkbookStructure {
        generation: 0,
        axis_changes,
        sheet_keys: context.keys,
        sheet_names: context.names,
        freeze_panes: model.sheets.iter().map(|sheet| sheet.freeze_pane).collect(),
        hyperlinks: model
            .sheets
            .iter()
            .map(|sheet| sheet.hyperlinks.clone())
            .collect(),
        charts: model
            .sheets
            .iter()
            .map(|sheet| sheet.charts.iter().map(ChartIdentity::of).collect())
            .collect(),
        merges: model
            .sheets
            .iter()
            .map(|sheet| sheet.merges.clone())
            .collect(),
        shared_types: BTreeMap::new(),
    };
    Ok((model, structure))
}
fn intersects(left: CellRange, right: CellRange) -> bool {
    left.start.row <= right.end.row
        && right.start.row <= left.end.row
        && left.start.col <= right.end.col
        && right.start.col <= left.end.col
}

pub(super) fn apply(
    doc: &Doc,
    base: &WorkbookBase,
    ops: &[Op],
    origin: SyncOrigin,
) -> Result<(), String> {
    for op in ops {
        let before = materialize(&doc.transact(), base)?.0;
        if let Op::SetChartAnchor {
            sheet, frame, to, ..
        } = op
        {
            if before
                .sheet(*sheet)
                .and_then(|sheet| sheet.charts.iter().find(|chart| chart.frame_id() == *frame))
                .is_some_and(|chart| chart.anchor == *to)
            {
                continue;
            }
        }
        let mut after = before.clone();
        xlsx_ops::apply(&mut after, op).map_err(|error| error.to_string())?;
        let mut txn = doc.transact_mut_with(origin.as_str());
        let context = context(&txn)?;
        let new_id = format!(
            "{}:{}",
            doc.client_id(),
            txn.state_vector().get(&doc.client_id())
        );
        let formats = map(&txn, CELL_FORMATS)?;
        sync_cell_formats(&formats, &mut txn, &after.styles)?;
        match op {
            Op::AddSheet { index, name } => {
                let key = format!("replica:{new_id}");
                blank_sheet(&mut txn, &key, name)?;
                let order = txn.get_array(SHEET_ORDER).ok_or("missing sheet order")?;
                if order.len(&txn) == 0 {
                    order.insert(&mut txn, 0, context.keys[0].clone());
                }
                order.insert(&mut txn, *index as u32, key);
            }
            Op::RemoveSheet { index } => {
                let order = txn.get_array(SHEET_ORDER).ok_or("missing sheet order")?;
                order.remove(&mut txn, *index as u32);
            }
            Op::SetDefinedNames { defined_names } => {
                write_names(&mut txn, defined_names, &context)?
            }
            Op::RestoreSheet { .. } => {
                return Err("internal restore is not a collaborative command".into());
            }
            _ => {
                let target = match op {
                    Op::SetCell { sheet, .. }
                    | Op::InsertRows { sheet, .. }
                    | Op::DeleteRows { sheet, .. }
                    | Op::InsertCols { sheet, .. }
                    | Op::DeleteCols { sheet, .. }
                    | Op::SetColWidth { sheet, .. }
                    | Op::SetRowHeight { sheet, .. }
                    | Op::SetFreezePane { sheet, .. }
                    | Op::SetHyperlinks { sheet, .. }
                    | Op::SetCharts { sheet, .. }
                    | Op::SetChartAnchor { sheet, .. }
                    | Op::MergeCells { sheet, .. }
                    | Op::UnmergeCells { sheet, .. }
                    | Op::PatchRangeStyle { sheet, .. }
                    | Op::SetRangeNumberFormat { sheet, .. }
                    | Op::ApplyRangeFormat { sheet, .. }
                    | Op::RenameSheet { sheet, .. } => *sheet,
                    _ => return Err("unhandled collaborative command".into()),
                };
                let key = context.sheet(target)?;
                let sheet = sheet_map(&txn, key)?;
                let output = after.sheet(target).ok_or("missing projected sheet")?;
                match op {
                    Op::InsertRows { at, count, .. }
                    | Op::InsertCols { at, count, .. }
                    | Op::DeleteRows { at, count, .. }
                    | Op::DeleteCols { at, count, .. } => {
                        let rows = matches!(op, Op::InsertRows { .. } | Op::DeleteRows { .. });
                        let axis_name = if rows { ROWS } else { COLS };
                        let axes = context.axes(key)?;
                        let axis = if rows { &axes.0 } else { &axes.1 };
                        let change = if matches!(op, Op::InsertRows { .. } | Op::InsertCols { .. })
                        {
                            Change::Insert {
                                before: axis.at(*at)?,
                                after: at.checked_sub(1).map(|index| axis.at(index)).transpose()?,
                                count: *count,
                            }
                        } else {
                            Change::Delete {
                                spans: axis.range(
                                    *at,
                                    at.checked_add(*count)
                                        .and_then(|end| end.checked_sub(1))
                                        .ok_or("axis deletion overflow")?,
                                )?,
                            }
                        };
                        let catalog = nested_map(
                            &nested_map(&map(&txn, CATALOG)?, &txn, key)?,
                            &txn,
                            axis_name,
                        )?;
                        set_json(&catalog, &mut txn, &new_id, &change)?;
                        let active = nested_map(&sheet, &txn, axis_name)?;
                        active.insert(&mut txn, new_id, true);
                    }
                    Op::SetCell { at, cell, .. } => {
                        let current = before.sheet(target).and_then(|sheet| sheet.cell(*at));
                        let format = current.and_then(|cell| cell.style) != cell.style;
                        let content = current.is_none_or(|old| {
                            old.formula != cell.formula
                                || (cell.formula.is_none() && old.value != cell.value)
                        });
                        write_cell(
                            &mut txn,
                            &sheet,
                            key,
                            *at,
                            &cell.clone().into(),
                            &context,
                            &after.styles,
                            content,
                            format,
                        )?;
                    }
                    Op::PatchRangeStyle { range, .. }
                    | Op::SetRangeNumberFormat { range, .. }
                    | Op::ApplyRangeFormat { range, .. } => {
                        for row in range.start.row..=range.end.row {
                            for col in range.start.col..=range.end.col {
                                let at = CellRef::new(row, col);
                                let cell = output.cell(at).cloned().unwrap_or_default();
                                write_cell(
                                    &mut txn,
                                    &sheet,
                                    key,
                                    at,
                                    &cell,
                                    &context,
                                    &after.styles,
                                    false,
                                    true,
                                )?;
                            }
                        }
                    }
                    Op::SetRowHeight { row, height, .. } => {
                        let map = nested_map(&sheet, &txn, ROW_HEIGHTS)?;
                        let id = json(&context.axes(key)?.0.at(*row)?)?;
                        sync_optional(&map, &mut txn, &id, height.map(Any::Number));
                    }
                    Op::SetColWidth { col, width, .. } => {
                        let map = nested_map(&sheet, &txn, COL_WIDTHS)?;
                        let id = json(&context.axes(key)?.1.at(*col)?)?;
                        sync_optional(&map, &mut txn, &id, width.map(Any::Number));
                    }
                    Op::RenameSheet { name, .. } => {
                        sheet.try_update(&mut txn, NAME, name.as_str());
                    }
                    Op::SetFreezePane { pane, .. } => {
                        write_freeze(&mut txn, &sheet, key, *pane, &context)?
                    }
                    Op::SetHyperlinks { hyperlinks, .. } => {
                        write_links(&mut txn, &sheet, key, hyperlinks, &context)?
                    }
                    Op::MergeCells { range, .. } | Op::UnmergeCells { range, .. } => {
                        let map = nested_map(&sheet, &txn, MERGES)?;
                        let remove = map
                            .iter(&txn)
                            .filter_map(|(id, value)| {
                                match decode::<BoundRange>(value)
                                    .and_then(|bound| Ok(bound.resolve(&context, key)))
                                {
                                    Ok(Some(old))
                                        if (matches!(op, Op::MergeCells { .. })
                                            && intersects(old, *range))
                                            || old == *range =>
                                    {
                                        Some(Ok(id.to_owned()))
                                    }
                                    Err(error) => Some(Err(error)),
                                    _ => None,
                                }
                            })
                            .collect::<Result<Vec<_>, String>>()?;
                        for id in remove {
                            map.remove(&mut txn, &id);
                        }
                        if matches!(op, Op::MergeCells { .. }) {
                            set_json(
                                &map,
                                &mut txn,
                                &new_id,
                                &BoundRange::bind(&context, key, *range)?,
                            )?;
                        }
                    }
                    Op::SetChartAnchor { frame, .. } => {
                        let chart = output
                            .charts
                            .iter()
                            .find(|chart| chart.frame_id() == *frame)
                            .ok_or("missing chart frame")?;
                        for shared_key in &context.keys {
                            let shared_sheet = sheet_map(&txn, shared_key)?;
                            let charts = nested_map(&shared_sheet, &txn, CHARTS)?;
                            if let Some(source) = before
                                .sheet(SheetId(
                                    context.index(shared_key).ok_or("missing chart sheet")? as u32,
                                ))
                                .and_then(|sheet| {
                                    sheet.charts.iter().find(|item| item.frame_id() == *frame)
                                })
                            {
                                let mut shared = source.clone();
                                shared.anchor = chart.anchor;
                                set_json(
                                    &charts,
                                    &mut txn,
                                    frame,
                                    &chart_bind(&shared, shared_key, &context)?,
                                )?;
                            }
                        }
                    }
                    Op::SetCharts { charts, .. } => {
                        let map = nested_map(&sheet, &txn, CHARTS)?;
                        let desired = charts
                            .iter()
                            .map(|chart| {
                                Ok((
                                    chart.frame_id(),
                                    Any::from(json(&chart_bind(chart, key, &context)?)?),
                                ))
                            })
                            .collect::<Result<_, String>>()?;
                        sync_map(&map, &mut txn, desired);
                    }
                    _ => return Err("unhandled sheet command".into()),
                }
            }
        }
    }
    Ok(())
}

/// Sheet containers are durable identities. Adding a sheet records its order entry
/// in Undo, but its original map containers must survive so remote edits and redo
/// can still address them after the sheet is hidden.
pub(super) fn immutable_sheet_insertions(doc: &Doc) -> Result<yrs::IdSet, String> {
    let txn = doc.transact();
    let mut immutable = yrs::IdSet::new();
    for (key, value) in map(&txn, SHEETS)?.iter(&txn) {
        let sheet = value.cast::<MapRef>().map_err(|_| "invalid sheet record")?;
        let catalog = nested_map(&map(&txn, CATALOG)?, &txn, key)?;
        let last = nested_map(&catalog, &txn, CHARTS)?;
        match (sheet.as_ref().id(), last.as_ref().id()) {
            (BranchID::Nested(start), BranchID::Nested(end))
                if start.client == end.client && end.clock == start.clock + 14 =>
            {
                immutable.insert(start, 15)
            }
            _ => return Err("sheet container identities were replaced".into()),
        }
    }
    Ok(immutable)
}

pub(super) fn undo_scopes(doc: &Doc, undo: &mut UndoManager<()>) -> Result<(), String> {
    let order = doc.get_or_insert_array(SHEET_ORDER);
    let names = doc.get_or_insert_map(DEFINED_NAMES);
    let txn = doc.transact();
    let sheets = map(&txn, SHEETS)?
        .iter(&txn)
        .map(|(_, value)| {
            value
                .cast::<MapRef>()
                .map_err(|_| "invalid sheet record".to_owned())
        })
        .collect::<Result<Vec<_>, _>>()?;
    drop(txn);
    for sheet in sheets {
        undo.expand_scope(doc, &sheet);
    }
    undo.expand_scope(doc, &order);
    undo.expand_scope(doc, &names);
    Ok(())
}

pub(super) fn cell_identity(doc: &Doc, sheet: SheetId, at: CellRef) -> Result<String, String> {
    let context = context(&doc.transact())?;
    let key = context.sheet(sheet)?;
    Ok(format!("{key}:{}", context.key(key, at)?))
}

pub(super) fn cell_identities(
    doc: &Doc,
    cells: impl IntoIterator<Item = (SheetId, CellRef)>,
) -> Result<Vec<String>, String> {
    let context = context(&doc.transact())?;
    cells
        .into_iter()
        .map(|(sheet, at)| {
            let key = context.sheet(sheet)?;
            Ok(format!("{key}:{}", context.key(key, at)?))
        })
        .collect()
}

pub(super) fn rebase_aliases(
    before: &Doc,
    latest: &Doc,
) -> Result<Vec<crate::workbook::rebase::SheetAlias>, String> {
    use crate::workbook::rebase::{AliasSpan, SheetAlias};
    let before = context(&before.transact())?;
    let latest = context(&latest.transact())?;
    fn spans(old: &Axis, current: Option<&Axis>) -> Vec<AliasSpan> {
        let mut targets = BTreeMap::<&str, Vec<(u64, u64, u64)>>::new();
        let mut position = 0;
        for span in current.into_iter().flat_map(|axis| &axis.spans) {
            targets.entry(&span.run).or_default().push((
                span.start,
                span.start + span.len,
                position,
            ));
            position += span.len;
        }
        for spans in targets.values_mut() {
            spans.sort_by_key(|span| span.0);
        }
        let mut result: Vec<AliasSpan> = Vec::new();
        let mut append = |start: u64, len: u64, target: Option<u64>| {
            if let Some(last) = result.last_mut() {
                if last.start + last.len == start
                    && match (last.target, target) {
                        (Some(previous), Some(current)) => previous + last.len == current,
                        (None, None) => true,
                        _ => false,
                    }
                {
                    last.len += len;
                    return;
                }
            }
            result.push(AliasSpan { start, len, target });
        };
        let mut old_position = 0;
        for span in &old.spans {
            let mut cursor = span.start;
            let end = span.start + span.len;
            if let Some(targets) = targets.get(span.run.as_str()) {
                let first = targets.partition_point(|(_, end, _)| *end <= cursor);
                for &(start, target_end, position) in &targets[first..] {
                    if start >= end {
                        break;
                    }
                    let low = cursor.max(start);
                    let high = end.min(target_end);
                    if low > cursor {
                        append(old_position + cursor - span.start, low - cursor, None);
                    }
                    append(
                        old_position + low - span.start,
                        high - low,
                        Some(position + low - start),
                    );
                    cursor = high;
                }
            }
            if cursor < end {
                append(old_position + cursor - span.start, end - cursor, None);
            }
            old_position += span.len;
        }
        result
    }

    before
        .keys
        .iter()
        .enumerate()
        .map(|(index, key)| {
            let (rows, cols) = before.axes(key)?;
            let target = latest.index(key);
            let axes = target.map(|_| latest.axes(key)).transpose()?;
            Ok(SheetAlias {
                target: target.map_or_else(
                    || format!("indexed-deleted-sheet:{index}"),
                    |index| format!("sheet:{index}"),
                ),
                rows: spans(rows, axes.map(|axes| &axes.0)),
                cols: spans(cols, axes.map(|axes| &axes.1)),
            })
        })
        .collect()
}
