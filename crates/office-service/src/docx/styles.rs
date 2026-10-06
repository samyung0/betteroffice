//! The paragraph-property side of styles/styleResolver.ts, numbering's list
//! levels (docx/numberingParser.ts, yrs/paragraphSeed.ts) and
//! yrs/tableParagraphFormatting.ts, as the save projection reads them.

use std::collections::HashMap;

use crate::jsv::{Obj, V};
use crate::obj;

use super::item::{is, o};

/// utils/fontFamilyMerge.ts `mergeFontFamily`.
fn merge_font_family(target: &V, source: &Obj) -> V {
    let result = o(target).map(|target| target.spread()).unwrap_or_default();
    let pairs = [
        ("ascii", "asciiTheme"),
        ("hAnsi", "hAnsiTheme"),
        ("eastAsia", "eastAsiaTheme"),
        ("cs", "csTheme"),
    ];
    for (explicit, theme) in pairs {
        let (name, themed) = (source.get(explicit), source.get(theme));
        if !name.is_undef() || !themed.is_undef() {
            result.delete(explicit);
            result.delete(theme);
            if !name.is_undef() {
                result.set(explicit, name);
            }
            if !themed.is_undef() {
                result.set(theme, themed);
            }
        }
    }
    for key in source.keys() {
        let paired = pairs
            .iter()
            .any(|(explicit, theme)| *explicit == &*key || *theme == &*key);
        if !paired && !source.get(&key).is_undef() {
            result.set(&key, source.get(&key));
        }
    }
    V::Obj(result)
}

/// utils/textFormattingMerge.ts `mergeTextFormatting`.
fn merge_text_formatting(target: &V, source: &V) -> V {
    if !source.truthy() && !target.truthy() {
        return V::Undef;
    }
    let Some(source_object) = o(source).filter(|_| source.truthy()) else {
        return target.clone();
    };
    if !target.truthy() {
        return V::Obj(source_object.spread());
    }
    let result = o(target).map(|target| target.spread()).unwrap_or_default();
    for key in source_object.keys() {
        let value = source_object.get(&key);
        if value.is_undef() {
            continue;
        }
        if &*key == "fontFamily"
            && let Some(family) = o(&value)
        {
            result.set(
                "fontFamily",
                merge_font_family(&target.get("fontFamily"), &family),
            );
            continue;
        }
        if &*key == "color" && o(&value).is_some() {
            let explicit = ["rgb", "themeColor", "themeTint", "themeShade"]
                .iter()
                .any(|field| value.get(field).truthy());
            if !value.get("auto").truthy() || explicit {
                result.set("color", value);
            }
            continue;
        }
        if let Some(value) = o(&value) {
            let merged = o(&target.get(&key))
                .map(|base| base.spread())
                .unwrap_or_default();
            merged.assign(&value);
            result.set(&key, V::Obj(merged));
            continue;
        }
        result.set(&key, value);
    }
    V::Obj(result)
}

/// utils/paragraphFormattingMerge.ts `mergeParagraphFormatting`.
pub(crate) fn merge_paragraph_formatting(target: &V, source: &V) -> V {
    if !source.truthy() {
        return target.clone();
    }
    let Some(source_object) = o(source) else {
        return target.clone();
    };
    if !target.truthy() {
        return V::Obj(source_object.spread());
    }
    let result = o(target).map(|target| target.spread()).unwrap_or_default();
    for key in source_object.keys() {
        let value = source_object.get(&key);
        if value.is_undef() {
            continue;
        }
        match &*key {
            "borders" | "numPr" | "frame" => {
                let merged = o(&result.get(&key))
                    .map(|base| base.spread())
                    .unwrap_or_default();
                if let Some(value) = o(&value) {
                    merged.assign(&value);
                }
                result.set(&key, V::Obj(merged));
            }
            "runProperties" => {
                let merged = merge_text_formatting(&result.get("runProperties"), &value);
                result.set("runProperties", merged);
            }
            "tabs" if value.arr().is_some() => {
                result.set(
                    "tabs",
                    V::Arr(value.arr().map(|tabs| tabs.copy()).unwrap_or_default()),
                );
            }
            _ => result.set(&key, value),
        }
    }
    V::Obj(result)
}

/// styles/styleResolver.ts `StyleResolver`, paragraph properties only.
pub(crate) struct StyleResolver {
    styles: HashMap<String, V>,
    /// Map insertion order of the style ids (first insertion of each id).
    order: Vec<String>,
    doc_defaults: V,
    default_paragraph: V,
    default_table: V,
    builtin_normal: V,
}

impl StyleResolver {
    pub fn new(definitions: &V) -> Self {
        let mut styles = HashMap::new();
        let mut order = Vec::new();
        for style in definitions.get("styles").items() {
            let id = style.get("styleId");
            if id.truthy() {
                let id = id.to_js_string();
                if !styles.contains_key(&id) {
                    order.push(id.clone());
                }
                styles.insert(id, style);
            }
        }
        let doc_defaults = definitions.get("docDefaults");
        let builtin_normal = obj! {
            "styleId": "Normal",
            "type": "paragraph",
            "name": "Normal",
            "default": true,
            "pPr": obj! { "spaceAfter": 160.0, "lineSpacing": 259.0, "lineSpacingRule": "auto" },
        };
        let mut resolver = Self {
            styles,
            order,
            doc_defaults,
            default_paragraph: V::Undef,
            default_table: V::Undef,
            builtin_normal,
        };
        resolver.default_paragraph = resolver.find_default("paragraph");
        resolver.default_table = resolver.find_default("table");
        resolver
    }

    fn find_default(&self, kind: &str) -> V {
        for id in &self.order {
            let style = &self.styles[id];
            if style.get("type").as_str().as_deref() == Some(kind) && style.get("default").truthy()
            {
                return style.clone();
            }
        }
        if kind == "paragraph" {
            return self.styles.get("Normal").cloned().unwrap_or_else(|| {
                if self.doc_defaults.truthy() {
                    V::Undef
                } else {
                    self.builtin_normal.clone()
                }
            });
        }
        V::Undef
    }

    pub fn style(&self, id: &str) -> V {
        self.styles.get(id).cloned().unwrap_or_default()
    }

    pub fn default_table_style(&self) -> V {
        self.default_table.clone()
    }

    /// `resolveParagraphStyle(styleId, table).paragraphFormatting`.
    pub fn resolve_paragraph_style(&self, style_id: &V, table: &V) -> V {
        let pdefault = self.doc_defaults.get("pPr");
        let mut result = if pdefault.truthy() {
            V::Obj(o(&pdefault).unwrap_or_default().spread())
        } else {
            V::Undef
        };
        let style = if style_id.truthy() {
            self.style(&style_id.to_js_string())
        } else {
            V::Undef
        }
        .or_else(|| self.default_paragraph.clone());
        let builtin = style.same(&self.builtin_normal);
        if builtin && style.get("pPr").truthy() {
            result = merge_paragraph_formatting(&result, &style.get("pPr"));
        }
        if table.truthy() {
            result = merge_paragraph_formatting(&result, table);
        }
        if style.truthy() && !builtin && style.get("pPr").truthy() {
            result = merge_paragraph_formatting(&result, &style.get("pPr"));
        }
        result
    }
}

/// docx/numberingParser.ts `createNumberingMap(definitions).getLevel`.
pub(crate) struct Numbering {
    abstracts: HashMap<String, V>,
    abstract_order: Vec<String>,
    nums: HashMap<String, V>,
}

fn key(value: &V) -> String {
    value.to_js_string()
}

impl Numbering {
    pub fn new(definitions: &V) -> Self {
        let mut abstracts = HashMap::new();
        let mut abstract_order = Vec::new();
        for item in definitions.get("abstractNums").items() {
            let id = key(&item.get("abstractNumId"));
            if !abstracts.contains_key(&id) {
                abstract_order.push(id.clone());
            }
            abstracts.insert(id, item);
        }
        let mut nums = HashMap::new();
        for item in definitions.get("nums").items() {
            nums.insert(key(&item.get("numId")), item);
        }
        Self {
            abstracts,
            abstract_order,
            nums,
        }
    }

    pub fn level(&self, num_id: &V, ilvl: &V) -> V {
        let Some(num) = self.nums.get(&key(num_id)) else {
            return V::Null;
        };
        let overrides = num.get("levelOverrides");
        if overrides.truthy() {
            let found = overrides
                .items()
                .into_iter()
                .find(|item| item.get("ilvl").same(ilvl));
            if let Some(found) = found {
                if found.get("lvl").truthy() {
                    return found.get("lvl");
                }
                if let Some(abstract_num) = self.abstracts.get(&key(&num.get("abstractNumId"))) {
                    let base = abstract_num
                        .get("levels")
                        .items()
                        .into_iter()
                        .find(|level| level.get("ilvl").same(ilvl));
                    if let Some(base) = base
                        && !found.get("startOverride").is_undef()
                    {
                        return V::Obj(
                            o(&base)
                                .unwrap_or_default()
                                .with(&[("start", found.get("startOverride"))]),
                        );
                    }
                }
            }
        }
        let Some(mut abstract_num) = self.abstracts.get(&key(&num.get("abstractNumId"))).cloned()
        else {
            return V::Null;
        };
        let link = abstract_num.get("numStyleLink");
        if link.truthy() && abstract_num.get("levels").items().is_empty() {
            for id in &self.abstract_order {
                let candidate = &self.abstracts[id];
                if candidate.get("styleLink").same(&link)
                    && !candidate.get("levels").items().is_empty()
                {
                    abstract_num = candidate.clone();
                    break;
                }
            }
        }
        abstract_num
            .get("levels")
            .items()
            .into_iter()
            .find(|level| level.get("ilvl").same(ilvl))
            .unwrap_or(V::Null)
    }
}

/// paragraphSeed.ts `styleListRendering`: the indents of a style's list
/// level, the only fields the save reads off it; none without a level.
pub(crate) fn style_list_rendering(style_ppr: &V, numbering: &Numbering) -> Option<V> {
    let num_pr = style_ppr.get("numPr");
    if !num_pr.truthy() {
        return None;
    }
    let num_id = num_pr.get("numId");
    // computeListRendering: `const { numId, ilvl = 0 } = numPr`.
    let ilvl = match num_pr.get("ilvl") {
        V::Undef => V::Num(0.0),
        other => other,
    };
    if num_id.is_undef() || num_id.same(&V::Num(0.0)) || numbering.level(&num_id, &ilvl).nullish() {
        return None;
    }
    let level = numbering
        .level(
            &num_id.clone().or_else(|| V::Num(0.0)),
            &num_pr.get("ilvl").or_else(|| V::Num(0.0)),
        )
        .get("pPr");
    let rendering = Obj::new();
    if style_ppr.get("indentLeft").nullish() {
        rendering.set("indentLeft", level.get("indentLeft"));
    }
    if style_ppr.get("indentFirstLine").nullish() && style_ppr.get("hangingIndent").nullish() {
        rendering.set("indentFirstLine", level.get("indentFirstLine"));
        rendering.set("hangingIndent", level.get("hangingIndent"));
    }
    Some(V::Obj(rendering))
}

/// tableParagraphFormatting.ts `tableCellParagraphFormatting`.
pub(crate) fn table_cell_paragraph_formatting(
    table: &V,
    style: &V,
    row_index: usize,
    start_column: f64,
    end_column: f64,
    columns: f64,
) -> V {
    let mut result = style.get("pPr");
    let parts = style.get("tblStylePr");
    if !parts.items().iter().any(|part| part.get("pPr").truthy()) {
        return result;
    }
    let look = table
        .get("formatting")
        .get("look")
        .or_else(|| style.get("tblPr").get("look"));
    let flag = |key: &str, mask: i64| -> bool {
        let own = look.get(key);
        if !own.nullish() {
            return own.truthy();
        }
        let value = look
            .get("value")
            .or_else(|| V::str(if look.truthy() { "0" } else { "04A0" }));
        let digits: String = value
            .to_js_string()
            .trim_start()
            .chars()
            .take_while(char::is_ascii_hexdigit)
            .collect();
        i64::from_str_radix(&digits, 16)
            .map(|bits| bits & mask != 0)
            .unwrap_or(false)
    };
    let first_row = flag("firstRow", 0x20);
    let last_row = flag("lastRow", 0x40);
    let first_column = flag("firstColumn", 0x80);
    let last_column = flag("lastColumn", 0x100);
    let rows = table.get("rows").items();
    let grid_before = rows
        .get(row_index)
        .map(|row| {
            row.get("formatting")
                .get("gridBefore")
                .or_else(|| V::Num(0.0))
                .to_number()
        })
        .unwrap_or(0.0);
    let start_column = start_column + grid_before;
    let end_column = end_column + grid_before;
    let at_first_row = first_row && row_index == 0;
    let at_last_row = last_row && row_index + 1 == rows.len();
    let at_first_column = first_column && start_column == 0.0;
    let at_last_column = last_column && end_column == columns;
    let band = |key: &str| {
        table
            .get("formatting")
            .get(key)
            .or_else(|| style.get("tblPr").get(key))
            .or_else(|| V::Num(1.0))
            .to_number()
    };
    let row_band = band("styleRowBandSize");
    let column_band = band("styleColBandSize");
    let mut regions = Vec::new();
    if !flag("noHBand", 0x200) && !at_first_row && !at_last_row && row_band > 0.0 {
        let index = ((row_index as f64 - f64::from(u8::from(first_row))) / row_band).floor();
        regions.push(if index % 2.0 == 0.0 {
            "band1Horz"
        } else {
            "band2Horz"
        });
    }
    if !flag("noVBand", 0x400) && !at_first_column && !at_last_column && column_band > 0.0 {
        let index = ((start_column - f64::from(u8::from(first_column))) / column_band).floor();
        regions.push(if index % 2.0 == 0.0 {
            "band1Vert"
        } else {
            "band2Vert"
        });
    }
    for (region, active) in [
        ("firstCol", at_first_column),
        ("lastCol", at_last_column),
        ("firstRow", at_first_row),
        ("lastRow", at_last_row),
        ("nwCell", at_first_row && at_first_column),
        ("neCell", at_first_row && at_last_column),
        ("swCell", at_last_row && at_first_column),
        ("seCell", at_last_row && at_last_column),
    ] {
        if active {
            regions.push(region);
        }
    }
    for region in regions {
        let part = parts
            .items()
            .into_iter()
            .find(|part| part.get("type").as_str().as_deref() == Some(region))
            .map(|part| part.get("pPr"))
            .unwrap_or_default();
        result = merge_paragraph_formatting(&result, &part);
    }
    result
}

/// tableParagraphFormatting.ts `tableColumnCount`.
pub(crate) fn table_column_count(table: &V) -> f64 {
    let mut count = table
        .get("columnWidths")
        .arr()
        .map(|widths| widths.len() as f64)
        .unwrap_or(0.0);
    for row in table.get("rows").items() {
        let formatting = row.get("formatting");
        let mut sum = formatting
            .get("gridBefore")
            .or_else(|| V::Num(0.0))
            .to_number()
            + formatting
                .get("gridAfter")
                .or_else(|| V::Num(0.0))
                .to_number();
        for cell in row.get("cells").items() {
            sum += cell
                .get("formatting")
                .get("gridSpan")
                .or_else(|| V::Num(1.0))
                .to_number();
        }
        count = count.max(sum);
    }
    count
}

/// Occupied grid slots of a table embed's payload rows (rowspan and colspan).
struct Occupied(Vec<Vec<bool>>);

impl Occupied {
    fn taken(&self, row: usize, column: usize) -> bool {
        self.0
            .get(row)
            .and_then(|row| row.get(column))
            .copied()
            .unwrap_or(false)
    }

    fn take(&mut self, row: usize, column: usize, rowspan: usize, colspan: usize) {
        for r in row..row + rowspan {
            if self.0.len() <= r {
                self.0.resize(r + 1, Vec::new());
            }
            for c in column..column + colspan {
                if self.0[r].len() <= c {
                    self.0[r].resize(c + 1, false);
                }
                self.0[r][c] = true;
            }
        }
    }
}

/// `Number(value) || 1` as a count.
pub(crate) fn span(value: &V) -> usize {
    let number = value.to_number();
    if number.is_nan() || number == 0.0 {
        1
    } else {
        number.max(0.0) as usize
    }
}

/// tableParagraphFormatting.ts `tablePayloadCellFormatting`.
pub(crate) fn table_payload_cell_formatting(
    payload: &V,
    styles: Option<&StyleResolver>,
) -> Vec<(String, V)> {
    let mut cells = Vec::new();
    let Some(styles) = styles else { return cells };
    let tbl_pr = payload.get("tblPr").or_else(|| obj! {});
    let original = tbl_pr.get("_originalFormatting").or_else(|| obj! {});
    let default_style = styles.default_table_style();
    let style_id = tbl_pr
        .get("styleId")
        .as_str()
        .map(V::Str)
        .unwrap_or_else(|| default_style.get("styleId"));
    let style = if style_id.truthy() {
        styles.style(&style_id.to_js_string())
    } else {
        V::Undef
    }
    .or_else(|| default_style.clone());
    let rows = payload.get("rows").items();
    let formatting = o(&original)
        .unwrap_or_default()
        .with(&[("look", tbl_pr.get("look").or_else(|| original.get("look")))]);
    let table = obj! {
        "formatting": V::Obj(formatting),
        "columnWidths": payload.get("grid").arr().map(|grid| V::Arr(grid.clone())).unwrap_or(V::Undef),
        "rows": V::array(rows.iter().map(|row| obj! {
            "formatting": row.get("trPr").get("_originalFormatting"),
            "cells": V::array(row.get("cells").items().iter().map(|cell| obj! {
                "formatting": obj! { "gridSpan": span(&cell.get("tcPr").get("colspan")) as f64 },
            }).collect()),
        }).collect()),
    };
    let columns = table_column_count(&table);
    let mut occupied = Occupied(Vec::new());
    for (row_index, row) in rows.iter().enumerate() {
        let mut column = 0;
        for cell in row.get("cells").items() {
            while occupied.taken(row_index, column) {
                column += 1;
            }
            let colspan = span(&cell.get("tcPr").get("colspan"));
            let rowspan = span(&cell.get("tcPr").get("rowspan"));
            occupied.take(row_index, column, rowspan, colspan);
            let formatting = table_cell_paragraph_formatting(
                &table,
                &style,
                row_index,
                column as f64,
                (column + colspan) as f64,
                columns,
            );
            let story = cell.get("story");
            if story.truthy() && formatting.truthy() {
                cells.push((story.to_js_string(), formatting));
            }
            column += colspan;
        }
    }
    cells
}

/// The cell story whose paragraphs `story` shares the table formatting of.
pub(crate) fn enclosing_cell_story(story: &str) -> Option<String> {
    let mut id = story.to_owned();
    while let Some(stripped) = strip_control(&id) {
        id = stripped;
    }
    is_cell_story(&id).then_some(id)
}

/// `id` without a trailing `:sdtN`.
fn strip_control(id: &str) -> Option<String> {
    let at = id.rfind(":sdt")?;
    let digits = &id[at + 4..];
    (!digits.is_empty() && digits.bytes().all(|byte| byte.is_ascii_digit()))
        .then(|| id[..at].to_owned())
}

/// Ends in `:tN:rNcN`.
fn is_cell_story(id: &str) -> bool {
    let Some(at) = id.rfind(":t") else {
        return false;
    };
    let rest = &id[at + 2..];
    let Some((table, cell)) = rest.split_once(":r") else {
        return false;
    };
    let Some((row, column)) = cell.split_once('c') else {
        return false;
    };
    [table, row, column]
        .iter()
        .all(|part| !part.is_empty() && part.bytes().all(|byte| byte.is_ascii_digit()))
}

/// A seeded cell story's source (tableParagraphFormatting.ts `SourceCell`).
#[derive(Clone)]
pub(crate) struct SourceCell {
    pub table: V,
    pub row: usize,
    pub column: f64,
    pub end: f64,
    pub columns: f64,
}

/// A source document's blocks by the story id the seed gives each, and its
/// cell stories' source cells (`sourceStories`).
pub(crate) struct SourceStories {
    pub stories: HashMap<String, Vec<V>>,
    /// Story ids in the order the seed visits them.
    pub order: Vec<String>,
    pub cells: HashMap<String, SourceCell>,
}

/// Each cell's rowspan by `row-gridColumn`; a covered vMerge continuation is
/// `skip` (`calculateRowSpans`).
fn row_span_skips(table: &V) -> HashMap<String, bool> {
    let mut result: HashMap<String, (usize, bool)> = HashMap::new();
    let mut active: HashMap<i64, usize> = HashMap::new();
    for (row_index, row) in table.get("rows").items().iter().enumerate() {
        let mut column = 0_i64;
        let cells: Vec<(i64, V, String)> = row
            .get("cells")
            .items()
            .iter()
            .map(|cell| {
                let current = column;
                column += cell
                    .get("formatting")
                    .get("gridSpan")
                    .or_else(|| V::Num(1.0))
                    .to_number() as i64;
                (
                    current,
                    cell.get("formatting").get("vMerge"),
                    format!("{row_index}-{current}"),
                )
            })
            .collect();
        let empty = !cells.is_empty()
            && cells.iter().all(|(column, merge, _)| {
                merge.as_str().as_deref() == Some("continue") && active.contains_key(column)
            });
        if empty {
            for (column, _, key) in &cells {
                active.remove(column);
                result.insert(key.clone(), (1, false));
            }
            continue;
        }
        for (column, merge, key) in cells {
            match merge.as_str().as_deref() {
                Some("restart") => {
                    active.insert(column, row_index);
                    result.insert(key, (1, false));
                }
                Some("continue") => match active.get(&column) {
                    None => {
                        result.insert(key, (1, false));
                    }
                    Some(start) => {
                        if let Some(owner) = result.get_mut(&format!("{start}-{column}")) {
                            owner.0 += 1;
                        }
                        result.insert(key, (1, true));
                    }
                },
                _ => {
                    active.remove(&column);
                    result.insert(key, (1, false));
                }
            }
        }
    }
    result
        .into_iter()
        .map(|(key, (_, skip))| (key, skip))
        .collect()
}

pub(crate) fn source_stories(document: &V) -> SourceStories {
    let mut read = SourceStories {
        stories: HashMap::new(),
        order: Vec::new(),
        cells: HashMap::new(),
    };
    fn visit(read: &mut SourceStories, story: String, blocks: Vec<V>) {
        if read.stories.insert(story.clone(), blocks.clone()).is_none() {
            read.order.push(story.clone());
        }
        let mut table_index = 0;
        let mut sdt_index = 0;
        for block in &blocks {
            if is(block, "blockSdt") {
                visit(
                    read,
                    format!("{story}:sdt{sdt_index}"),
                    block.get("content").items(),
                );
                sdt_index += 1;
                continue;
            }
            if !is(block, "table") {
                continue;
            }
            let current = table_index;
            table_index += 1;
            let skips = row_span_skips(block);
            let columns = table_column_count(block);
            for (row_index, row) in block.get("rows").items().iter().enumerate() {
                let mut column = 0.0;
                let mut cell_index = 0;
                for cell in row.get("cells").items() {
                    let start = column;
                    column += cell
                        .get("formatting")
                        .get("gridSpan")
                        .or_else(|| V::Num(1.0))
                        .to_number();
                    if skips
                        .get(&format!("{row_index}-{start}"))
                        .copied()
                        .unwrap_or(false)
                    {
                        continue;
                    }
                    let id = format!("{story}:t{current}:r{row_index}c{cell_index}");
                    cell_index += 1;
                    read.cells.insert(
                        id.clone(),
                        SourceCell {
                            table: block.clone(),
                            row: row_index,
                            column: start,
                            end: column,
                            columns,
                        },
                    );
                    visit(read, id, cell.get("content").items());
                }
            }
        }
    }
    let package = document.get("package");
    visit(
        &mut read,
        "body".into(),
        package.get("document").get("content").items(),
    );
    for parts in [package.get("headers"), package.get("footers")] {
        for (id, part) in map_entries(&parts) {
            visit(&mut read, format!("hf:{id}"), part.get("content").items());
        }
    }
    for (prefix, notes) in [
        ("fn", package.get("footnotes")),
        ("en", package.get("endnotes")),
    ] {
        for note in notes.items() {
            visit(
                &mut read,
                format!("{prefix}:{}", note.get("id").to_js_string()),
                note.get("content").items(),
            );
        }
    }
    read
}

/// The entries of a decoded `Map` (an array of `[key, value]` pairs here).
pub(crate) fn map_entries(map: &V) -> Vec<(String, V)> {
    map.items()
        .iter()
        .filter_map(|entry| {
            let pair = entry.arr()?;
            Some((pair.get(0).to_js_string(), pair.get(1)))
        })
        .collect()
}

/// The table-style paragraph formatting the seed gave a source cell.
pub(crate) fn source_cell_formatting(cell: &SourceCell, styles: &StyleResolver) -> V {
    let fallback = styles.default_table_style();
    let style_id = cell
        .table
        .get("formatting")
        .get("styleId")
        .or_else(|| fallback.get("styleId"));
    let style = if style_id.truthy() {
        styles.style(&style_id.to_js_string())
    } else {
        V::Undef
    }
    .or_else(|| fallback.clone());
    table_cell_paragraph_formatting(
        &cell.table,
        &style,
        cell.row,
        cell.column,
        cell.end,
        cell.columns,
    )
}

/// yrsToDocument.ts `sameJson`.
pub(crate) fn same_json(a: &V, b: &V) -> bool {
    if a.same(b) {
        return true;
    }
    if let (Some(left), Some(right)) = (a.arr(), b.arr()) {
        let (left, right) = (left.items(), right.items());
        return left.len() == right.len() && left.iter().zip(&right).all(|(x, y)| same_json(x, y));
    }
    let (Some(left), Some(right)) = (a.obj(), b.obj()) else {
        return false;
    };
    let keys = left.keys();
    keys.len() == right.len()
        && keys
            .iter()
            .all(|key| right.has(key) && same_json(&left.get(key), &right.get(key)))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn story_ids_parse_as_the_ts_patterns_do() {
        assert_eq!(
            enclosing_cell_story("body:t0:r1c2:sdt3:sdt0").as_deref(),
            Some("body:t0:r1c2")
        );
        assert_eq!(enclosing_cell_story("body:sdt1"), None);
        assert!(is_cell_story("hf:rId7:t2:r10c0"));
        assert!(!is_cell_story("body:t0:r1"));
    }
}
