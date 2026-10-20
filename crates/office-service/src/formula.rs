//! What a value the agent writes into a cell must pass (Epo 2026-10-20):
//! agent commands only, so people keep Excel parity. A formula uses only
//! functions the engine implements and never `HYPERLINK`, references only
//! this workbook's sheets and defined names (no `[...]` workbooks, URLs,
//! paths or DDE), parses, and stays within 2,000 characters; text stays
//! within Excel's 32,767. A defined name the formula uses passes the same
//! function and reference rules.

use std::collections::BTreeSet;

use betteroffice_xlsx::DefinedName;
use xlsx_calc::functions::Func;
use xlsx_calc::{Expr, parse_formula};
use xlsx_ops::{ParsedInput, parse_input};

use crate::error::{EditCode, Error, Result};

const MAX_FORMULA_CHARS: usize = 2_000;
const MAX_TEXT_CHARS: usize = 32_767;

/// The workbook a value is checked against.
pub(crate) struct Formulas<'a> {
    pub sheets: Vec<&'a str>,
    pub names: &'a [DefinedName],
}

impl Formulas<'_> {
    /// Refuses `value` for cell `at` as `invalid_input` naming the rule.
    pub fn check(&self, at: &str, value: &str) -> Result<()> {
        let refusal = match parse_input(value) {
            ParsedInput::Formula(body) => self.formula(&body).err(),
            ParsedInput::Text(text) if text.chars().count() > MAX_TEXT_CHARS => Some(format!(
                "the text is {} characters long; a cell holds at most {MAX_TEXT_CHARS}",
                text.chars().count()
            )),
            _ => None,
        };
        match refusal {
            Some(rule) => Err(Error::edit(EditCode::InvalidInput, format!("{at}: {rule}"))),
            None => Ok(()),
        }
    }

    fn formula(&self, body: &str) -> std::result::Result<(), String> {
        let length = body.chars().count();
        if length > MAX_FORMULA_CHARS {
            return Err(format!(
                "the formula is {length} characters long; a formula may hold at most {MAX_FORMULA_CHARS}"
            ));
        }
        let expr = parse_formula(body).map_err(|error| match unquoted(body) {
            Some('|') => "the formula uses '|', a DDE link, which is not allowed".to_owned(),
            Some('[') => {
                "the formula references another workbook ([...]), which is not allowed".to_owned()
            }
            _ => format!(
                "the formula does not parse ({}); write it as it would be typed in Excel",
                error.message
            ),
        })?;
        self.walk(&expr, &mut BTreeSet::new())
    }

    /// `seen` stops names that refer to each other.
    fn walk(&self, expr: &Expr, seen: &mut BTreeSet<String>) -> std::result::Result<(), String> {
        match expr {
            Expr::FuncCall { name, func, args } => {
                match func {
                    None => {
                        return Err(format!(
                            "the function {name} is not available; use a function the spreadsheet engine supports"
                        ));
                    }
                    Some(Func::Hyperlink) => {
                        return Err("the HYPERLINK function is not allowed".to_owned());
                    }
                    Some(_) => {}
                }
                args.iter().try_for_each(|arg| self.walk(arg, seen))
            }
            Expr::Ref { sheet, .. }
            | Expr::Range { sheet, .. }
            | Expr::ColumnRange { sheet, .. }
            | Expr::RowRange { sheet, .. } => sheet.as_deref().map_or(Ok(()), |s| self.sheet(s)),
            Expr::Name { scope, name } => {
                if let Some(scope) = scope {
                    self.sheet(scope)?;
                }
                self.name(name, seen)
            }
            Expr::Unary { expr, .. } | Expr::Percent(expr) => self.walk(expr, seen),
            Expr::Binary { lhs, rhs, .. } => {
                self.walk(lhs, seen)?;
                self.walk(rhs, seen)
            }
            Expr::RangeJoin { start, end } => {
                self.walk(start, seen)?;
                self.walk(end, seen)
            }
            Expr::ArrayLiteral { values, .. } => {
                values.iter().try_for_each(|value| self.walk(value, seen))
            }
            Expr::Number(_)
            | Expr::Literal(_)
            | Expr::Text(_)
            | Expr::Bool(_)
            | Expr::Error(_)
            | Expr::TableRef { .. } => Ok(()),
        }
    }

    fn sheet(&self, sheet: &str) -> std::result::Result<(), String> {
        if sheet.contains(['[', ']']) {
            return Err(format!(
                "the reference to '{sheet}' names another workbook, which is not allowed"
            ));
        }
        if sheet.contains("://") || sheet.contains(['\\', '/']) {
            return Err(format!(
                "the reference to '{sheet}' names a URL or a file path, which is not allowed"
            ));
        }
        let wanted = sheet.to_lowercase();
        if self.sheets.iter().any(|name| name.to_lowercase() == wanted) {
            return Ok(());
        }
        Err(format!(
            "the sheet '{sheet}' is not in this workbook; its sheets are {}",
            self.sheets
                .iter()
                .map(|name| format!("'{name}'"))
                .collect::<Vec<_>>()
                .join(", ")
        ))
    }

    fn name(&self, name: &str, seen: &mut BTreeSet<String>) -> std::result::Result<(), String> {
        let folded = name.to_lowercase();
        if !seen.insert(folded.clone()) {
            return Ok(());
        }
        let defined: Vec<&DefinedName> = self
            .names
            .iter()
            .filter(|defined| defined.name.to_lowercase() == folded)
            .collect();
        if defined.is_empty() {
            return Err(format!("the name {name} is not defined in this workbook"));
        }
        for defined in defined {
            let formula = &defined.formula;
            let expr = parse_formula(formula).map_err(|_| {
                format!(
                    "the name {name} refers outside this workbook ({formula}), which is not allowed"
                )
            })?;
            self.walk(&expr, seen)
                .map_err(|rule| format!("the name {name} ({formula}) is not allowed: {rule}"))?;
        }
        Ok(())
    }
}

/// The first `|`, or `[` that opens a reference rather than a table's
/// columns, outside string literals and quoted sheet names.
fn unquoted(body: &str) -> Option<char> {
    let (mut string, mut quoted, mut table, mut previous) = (false, false, 0_u32, ' ');
    for ch in body.chars() {
        match ch {
            _ if (string && ch != '"') || (quoted && ch != '\'') => {}
            '"' => string = !string,
            '\'' => quoted = !quoted,
            '|' => return Some(ch),
            '[' if table > 0 || previous.is_alphanumeric() || matches!(previous, '_' | '.') => {
                table += 1
            }
            '[' => return Some(ch),
            ']' => table = table.saturating_sub(1),
            _ => {}
        }
        if !ch.is_whitespace() {
            previous = ch;
        }
    }
    None
}
