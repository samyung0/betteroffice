//! turning raw editor input into a `CellState`, classified the way excel does:
//! leading `=` is a formula, leading `'` forces text, numbers/booleans coerce.

use xlsx_calc::{EvalContext, Expr, evaluate, parse_formula};
use xlsx_model::{CellProvider, CellValue, SheetId};

use crate::op::CellState;

/// the classification of a raw editor string, before it becomes a `CellState`.
#[derive(Debug, Clone, PartialEq)]
pub enum ParsedInput {
    /// empty input — clears the cell back to default.
    Clear,
    /// a formula, carrying the text with the leading `=` stripped.
    Formula(String),
    /// literal text (either forced with a leading `'`, or a plain string).
    Text(String),
    Number(f64),
    Bool(bool),
}

/// classify a raw editor string the way excel does on cell entry.
pub fn parse_input(input: &str) -> ParsedInput {
    if input.is_empty() {
        return ParsedInput::Clear;
    }
    if let Some(rest) = input.strip_prefix('=') {
        return ParsedInput::Formula(rest.to_string());
    }
    if let Some(rest) = input.strip_prefix('\'') {
        return ParsedInput::Text(rest.to_string());
    }
    if input.eq_ignore_ascii_case("true") {
        return ParsedInput::Bool(true);
    }
    if input.eq_ignore_ascii_case("false") {
        return ParsedInput::Bool(false);
    }
    if let Some(n) = parse_number(input) {
        return ParsedInput::Number(n);
    }
    ParsedInput::Text(input.to_string())
}

/// parse a bare numeric literal; rejects the `inf`/`nan` tokens that
/// `f64::from_str` would accept, so those fall through to text.
fn parse_number(input: &str) -> Option<f64> {
    let trimmed = input.trim();
    if trimmed.is_empty() {
        return None;
    }
    let n: f64 = trimmed.parse().ok()?;
    n.is_finite().then_some(n)
}

/// build the `CellState` for a raw editor string, evaluating formulas
/// immediately against `provider`; a formula left with open parentheses is
/// closed, and one that still does not parse degrades to raw text.
pub fn cell_state_for_input(input: &str, provider: &dyn CellProvider, sheet: SheetId) -> CellState {
    match parse_input(input) {
        ParsedInput::Clear => CellState::default(),
        ParsedInput::Number(value) => CellState {
            value: CellValue::Number { value },
            ..Default::default()
        },
        ParsedInput::Bool(value) => CellState {
            value: CellValue::Bool { value },
            ..Default::default()
        },
        ParsedInput::Text(value) => CellState {
            value: CellValue::Text { value },
            ..Default::default()
        },
        ParsedInput::Formula(text) => match typed_formula(text) {
            Some((text, expr)) => {
                let ctx = EvalContext::new(provider, sheet);
                CellState {
                    value: evaluate(&expr, &ctx),
                    formula: Some(text),
                    ..Default::default()
                }
            }
            None => CellState {
                value: CellValue::Text {
                    value: input.to_string(),
                },
                ..Default::default()
            },
        },
    }
}

/// build the `CellState` without evaluating: formula text is stored with an
/// `Empty` value for the dependency graph's recalc to fill in afterwards.
pub fn cell_state_for_input_no_eval(input: &str) -> CellState {
    match parse_input(input) {
        ParsedInput::Clear => CellState::default(),
        ParsedInput::Number(value) => CellState {
            value: CellValue::Number { value },
            ..Default::default()
        },
        ParsedInput::Bool(value) => CellState {
            value: CellValue::Bool { value },
            ..Default::default()
        },
        ParsedInput::Text(value) => CellState {
            value: CellValue::Text { value },
            ..Default::default()
        },
        ParsedInput::Formula(text) => match typed_formula(text) {
            Some((text, _)) => CellState {
                value: CellValue::Empty,
                formula: Some(text),
                ..Default::default()
            },
            None => CellState {
                value: CellValue::Text {
                    value: input.to_string(),
                },
                ..Default::default()
            },
        },
    }
}

/// a typed formula's text and parse: as typed, or with the parentheses it
/// leaves open closed at its end, as Excel and Sheets complete them.
fn typed_formula(text: String) -> Option<(String, Expr)> {
    match parse_formula(&text) {
        Ok(expr) => Some((text, expr)),
        Err(_) => {
            let closed = close_open_parens(&text)?;
            let expr = parse_formula(&closed).ok()?;
            Some((closed, expr))
        }
    }
}

/// `text` with a `)` for each `(` it leaves open, or `None` when none is open,
/// a `)` has no `(`, a string, quoted sheet name or `[` is left open, or the
/// text ends in `(` (a call with no arguments yet is not a finished formula).
/// Parentheses inside strings, quoted names and structured references do not count.
fn close_open_parens(text: &str) -> Option<String> {
    if text.trim_end().ends_with('(') {
        return None;
    }
    let mut depth = 0usize;
    let mut brackets = 0usize;
    let mut quote: Option<char> = None;
    let mut chars = text.chars().peekable();
    while let Some(c) = chars.next() {
        if let Some(open) = quote {
            if c == open {
                if chars.peek() == Some(&open) {
                    chars.next();
                } else {
                    quote = None;
                }
            }
            continue;
        }
        match c {
            '"' | '\'' if brackets == 0 => quote = Some(c),
            '[' => brackets += 1,
            ']' => brackets = brackets.checked_sub(1)?,
            '(' if brackets == 0 => depth += 1,
            ')' if brackets == 0 => depth = depth.checked_sub(1)?,
            _ => {}
        }
    }
    (quote.is_none() && brackets == 0 && depth > 0).then(|| format!("{text}{}", ")".repeat(depth)))
}

#[cfg(test)]
mod tests {
    use super::*;
    use xlsx_model::{Cell, CellRef, Sheet, Workbook};

    #[test]
    fn classifies_edge_cases() {
        assert_eq!(parse_input(""), ParsedInput::Clear);
        assert_eq!(parse_input("00"), ParsedInput::Number(0.0));
        assert_eq!(parse_input("1e3"), ParsedInput::Number(1000.0));
        assert_eq!(parse_input(" 42 "), ParsedInput::Number(42.0));
        assert_eq!(parse_input("="), ParsedInput::Formula(String::new()));
        assert_eq!(parse_input("'=danger"), ParsedInput::Text("=danger".into()));
        assert_eq!(parse_input("-"), ParsedInput::Text("-".into()));
        assert_eq!(parse_input("3.14.15"), ParsedInput::Text("3.14.15".into()));
    }

    #[test]
    fn booleans_are_case_insensitive() {
        assert_eq!(parse_input("true"), ParsedInput::Bool(true));
        assert_eq!(parse_input("TRUE"), ParsedInput::Bool(true));
        assert_eq!(parse_input("False"), ParsedInput::Bool(false));
    }

    #[test]
    fn rejects_non_finite_number_tokens() {
        assert_eq!(parse_input("inf"), ParsedInput::Text("inf".into()));
        assert_eq!(parse_input("nan"), ParsedInput::Text("nan".into()));
        assert_eq!(
            parse_input("infinity"),
            ParsedInput::Text("infinity".into())
        );
    }

    fn wb_with(cells: &[(&str, f64)]) -> Workbook {
        let mut sheet = Sheet::new("Sheet1");
        for (a1, v) in cells {
            sheet.set_cell(
                CellRef::parse_a1(a1).unwrap(),
                Cell {
                    value: CellValue::Number { value: *v },
                    ..Cell::default()
                },
            );
        }
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);
        wb
    }

    #[test]
    fn formula_evaluates_against_provider_on_entry() {
        let wb = wb_with(&[("A1", 2.0), ("A2", 3.0)]);
        let state = cell_state_for_input("=SUM(A1:A2)", &wb, SheetId(0));
        assert_eq!(state.value, CellValue::Number { value: 5.0 });
        assert_eq!(state.formula.as_deref(), Some("SUM(A1:A2)"));
    }

    #[test]
    fn unparseable_formula_degrades_to_raw_text() {
        let wb = wb_with(&[]);
        let state = cell_state_for_input("=SUM(", &wb, SheetId(0));
        assert_eq!(
            state.value,
            CellValue::Text {
                value: "=SUM(".into()
            }
        );
        assert!(state.formula.is_none());
    }

    #[test]
    fn empty_input_clears() {
        let wb = wb_with(&[]);
        assert_eq!(
            cell_state_for_input("", &wb, SheetId(0)),
            CellState::default()
        );
    }

    #[test]
    fn no_eval_stores_formula_without_value() {
        let state = cell_state_for_input_no_eval("=SUM(A1:A2)");
        assert_eq!(state.value, CellValue::Empty);
        assert_eq!(state.formula.as_deref(), Some("SUM(A1:A2)"));
    }

    #[test]
    fn no_eval_coerces_literals_and_degrades_bad_formula() {
        assert_eq!(
            cell_state_for_input_no_eval("42").value,
            CellValue::Number { value: 42.0 }
        );
        let bad = cell_state_for_input_no_eval("=SUM(");
        assert_eq!(
            bad.value,
            CellValue::Text {
                value: "=SUM(".into()
            }
        );
        assert!(bad.formula.is_none());
    }

    #[test]
    fn closes_the_parentheses_a_typed_formula_leaves_open() {
        for (typed, stored) in [
            ("=SUM(A1:A2", "SUM(A1:A2)"),
            ("=ROUND(SUM(A1:A2", "ROUND(SUM(A1:A2))"),
            ("=IF(A1>1,\"(\",\"b\"", "IF(A1>1,\"(\",\"b\")"),
            ("=SUM('Q (1)'!A1:A2", "SUM('Q (1)'!A1:A2)"),
        ] {
            let state = cell_state_for_input_no_eval(typed);
            assert_eq!(state.formula.as_deref(), Some(stored), "{typed}");
        }
        let wb = wb_with(&[("A1", 2.0), ("A2", 3.0)]);
        let state = cell_state_for_input("=SUM(A1:A2", &wb, SheetId(0));
        assert_eq!(state.value, CellValue::Number { value: 5.0 });
        assert_eq!(state.formula.as_deref(), Some("SUM(A1:A2)"));
    }

    #[test]
    fn leaves_a_formula_closing_cannot_fix_as_text() {
        for typed in ["=", "=A1+", "=SUM(", "=A1)", "=IF(A1=\"x", "=SUM(A1:A2;B1"] {
            let state = cell_state_for_input_no_eval(typed);
            assert!(state.formula.is_none(), "{typed}");
            assert_eq!(
                state.value,
                CellValue::Text {
                    value: typed.into()
                },
                "{typed}"
            );
        }
    }
}
