//! Cell-anchored parts a save moves along a preserved sheet's rows and
//! columns: notes and threaded comments with their VML shapes, drawing anchors
//! other than charts (which move with the model), and protected ranges. A note
//! whose cell was deleted goes with it, as in Excel.

use xlsx_model::addr::{CellRange, CellRef, col_to_letters};

use crate::ParseError;
use crate::axis::SheetAxes;
use crate::chart::{directory_of, parse_relationships};
use crate::tree::{Edit, Element, Part, Replacement};
use crate::xml::resolve_part_path;

#[derive(Clone, Copy)]
pub(crate) enum Anchored {
    Comments,
    ThreadedComments,
    Vml,
    Drawing,
}

/// The parts a sheet relates to that hold cell anchors.
pub(crate) fn anchored_parts(
    sheet_path: &str,
    rels: Option<&[u8]>,
) -> Result<Vec<(Anchored, String)>, ParseError> {
    let Some(rels) = rels else {
        return Ok(Vec::new());
    };
    let directory = directory_of(sheet_path);
    Ok(parse_relationships(rels)?
        .into_iter()
        .filter_map(|(_, kind, target)| {
            let anchored = match kind.rsplit('/').next()? {
                "comments" => Anchored::Comments,
                "threadedComment" => Anchored::ThreadedComments,
                "vmlDrawing" => Anchored::Vml,
                "drawing" => Anchored::Drawing,
                _ => return None,
            };
            Some((anchored, resolve_part_path(directory, &target)))
        })
        .collect())
}

/// The part with its anchors moved along `axes`, or `None` when none moved.
pub(crate) fn move_anchored(
    anchored: Anchored,
    part: &[u8],
    axes: &SheetAxes,
) -> Result<Option<Vec<u8>>, ParseError> {
    let local = match anchored {
        Anchored::Drawing => return crate::chart::move_drawing_anchors(part, axes),
        Anchored::Comments => Some("comment"),
        Anchored::ThreadedComments => Some("threadedComment"),
        Anchored::Vml => None,
    };
    let parsed = Part::decode(part).and_then(|source| {
        let root = source.tree()?;
        Ok((source, root))
    });
    let (source, root) = match (parsed, local) {
        (Ok(parsed), _) => parsed,
        // Note VML is often not well-formed XML (an unclosed `<br>` in a
        // textbox); it is written back as it was (Epo 2026-10-13).
        (Err(_), None) => return Ok(None),
        (Err(error), Some(_)) => return Err(error),
    };
    let mut edits = Vec::new();
    match local {
        Some(local) => move_refs(&source, &root, local, axes, &mut edits),
        None => move_notes(&root, axes, &mut edits),
    }
    if edits.is_empty() {
        return Ok(None);
    }
    edits.sort_by_key(|(span, _)| span.start);
    source.splice(&edits).map(Some)
}

/// A comment follows its cell, or goes when its row or column was deleted.
fn move_refs(
    source: &Part,
    element: &Element,
    local: &str,
    axes: &SheetAxes,
    edits: &mut Vec<Edit>,
) {
    if element.local_name() == local {
        let Some(cell) = element
            .attribute_local("ref")
            .and_then(|value| CellRef::parse_a1(value).ok())
        else {
            return;
        };
        match (axes.rows.current(cell.row), axes.cols.current(cell.col)) {
            (Some(row), Some(col)) if (row, col) == (cell.row, cell.col) => {}
            (Some(row), Some(col)) => {
                let tag = source.text(element.start_tag());
                if let Some(tag) = with_attribute(tag, "ref", &CellRef::new(row, col).to_a1()) {
                    edits.push((element.start_tag(), Replacement::Markup(tag)));
                }
            }
            _ => edits.push((element.span.clone(), Replacement::Markup(String::new()))),
        }
        return;
    }
    for child in element.child_elements() {
        move_refs(source, child, local, axes, edits);
    }
}

/// A note's VML shape follows its cell, its box keeping its size, or goes
/// with the cell.
fn move_notes(element: &Element, axes: &SheetAxes, edits: &mut Vec<Edit>) {
    let data = element.child("ClientData").filter(|data| {
        element.local_name() == "shape" && data.attribute_local("ObjectType") == Some("Note")
    });
    let Some(data) = data else {
        for child in element.child_elements() {
            move_notes(child, axes, edits);
        }
        return;
    };
    let index = |local: &str| {
        let child = data.child(local)?;
        Some((child, child.text_content().trim().parse::<u32>().ok()?))
    };
    let (Some((row_element, row)), Some((col_element, col))) = (index("Row"), index("Column"))
    else {
        return;
    };
    let (Some(new_row), Some(new_col)) = (axes.rows.current(row), axes.cols.current(col)) else {
        edits.push((element.span.clone(), Replacement::Markup(String::new())));
        return;
    };
    let mut set = |child: &Element, text: String| {
        if let Some(span) = child.splice_target() {
            edits.push((span, Replacement::Text(text)));
        }
    };
    if new_row != row {
        set(row_element, new_row.to_string());
    }
    if new_col != col {
        set(col_element, new_col.to_string());
    }
    // LeftColumn, LeftOffset, TopRow, TopOffset, RightColumn, RightOffset,
    // BottomRow, BottomOffset.
    if let Some(anchor) = data.child("Anchor") {
        let values: Option<Vec<i64>> = anchor
            .text_content()
            .split(',')
            .map(|value| value.trim().parse().ok())
            .collect();
        if let Some(mut values) = values.filter(|values| values.len() == 8) {
            let (rows, cols) = (
                i64::from(new_row) - i64::from(row),
                i64::from(new_col) - i64::from(col),
            );
            for (index, delta) in [(0, cols), (2, rows), (4, cols), (6, rows)] {
                values[index] = (values[index] + delta).max(0);
            }
            if rows != 0 || cols != 0 {
                let text = values.iter().map(i64::to_string).collect::<Vec<_>>();
                set(anchor, text.join(", "));
            }
        }
    }
}

/// A `<protectedRanges>` fragment with each range moved along `axes`: a
/// range keeps its surviving rows and columns (and what was inserted between
/// them), one wholly deleted leaves its `sqref`, and a `<protectedRange>` with
/// none left goes. `None` when no range is left at all.
pub(crate) fn move_protected_ranges(
    fragment: &[u8],
    axes: &SheetAxes,
) -> Result<Option<Vec<u8>>, ParseError> {
    let source = Part::decode(fragment)?;
    let root = source.tree()?;
    let mut edits = Vec::new();
    let mut kept = 0;
    for range in root
        .child_elements()
        .filter(|child| child.local_name() == "protectedRange")
    {
        let Some(sqref) = range.attribute_local("sqref") else {
            kept += 1;
            continue;
        };
        let moved = sqref
            .split_whitespace()
            .filter_map(|reference| move_reference(reference, axes))
            .collect::<Vec<_>>()
            .join(" ");
        if moved.is_empty() {
            edits.push((range.span.clone(), Replacement::Markup(String::new())));
            continue;
        }
        kept += 1;
        if moved != sqref
            && let Some(tag) = with_attribute(source.text(range.start_tag()), "sqref", &moved)
        {
            edits.push((range.start_tag(), Replacement::Markup(tag)));
        }
    }
    if kept == 0 {
        return Ok(None);
    }
    source.splice(&edits).map(Some)
}

/// One `sqref` reference moved along `axes`: a cell or range (`A1`, `A1:B2`),
/// whole columns (`B:D`) or whole rows (`2:4`) keep their surviving part;
/// `None` when it was wholly deleted. Anything else is not a reference
/// `ST_Sqref` allows and stays as authored.
fn move_reference(reference: &str, axes: &SheetAxes) -> Option<String> {
    if let Ok(cells) = CellRange::parse_a1(reference) {
        let (top, bottom) = axes.rows.span(cells.start.row, cells.end.row)?;
        let (left, right) = axes.cols.span(cells.start.col, cells.end.col)?;
        return Some(CellRange::new(CellRef::new(top, left), CellRef::new(bottom, right)).to_a1());
    }
    let Some((first, last)) = reference.split_once(':') else {
        return Some(reference.to_owned());
    };
    let column = |letters: &str| {
        letters
            .bytes()
            .all(|byte| byte.is_ascii_alphabetic())
            .then(|| {
                CellRef::parse_a1(&format!("{letters}1"))
                    .ok()
                    .map(|cell| cell.col)
            })
            .flatten()
    };
    let row = |digits: &str| {
        digits
            .bytes()
            .all(|byte| byte.is_ascii_digit())
            .then(|| digits.parse::<u32>().ok()?.checked_sub(1))
            .flatten()
    };
    if let (Some(first), Some(last)) = (column(first), column(last)) {
        let (left, right) = axes.cols.span(first.min(last), first.max(last))?;
        return Some(format!(
            "{}:{}",
            col_to_letters(left),
            col_to_letters(right)
        ));
    }
    if let (Some(first), Some(last)) = (row(first), row(last)) {
        let (top, bottom) = axes.rows.span(first.min(last), first.max(last))?;
        return Some(format!("{}:{}", top + 1, bottom + 1));
    }
    Some(reference.to_owned())
}

/// A start tag with one attribute's value replaced, its quoting kept. Values
/// written here are cell references, which need no escaping.
fn with_attribute(tag: &str, name: &str, value: &str) -> Option<String> {
    let mut from = 0;
    while let Some(found) = tag[from..].find(name) {
        let start = from + found;
        from = start + name.len();
        if !tag[..start].ends_with(|c: char| c.is_ascii_whitespace()) {
            continue;
        }
        let Some(rest) = tag[from..].trim_start().strip_prefix('=') else {
            continue;
        };
        let rest = rest.trim_start();
        let Some(quote) = rest.chars().next().filter(|c| matches!(c, '"' | '\'')) else {
            continue;
        };
        let value_start = tag.len() - rest.len() + 1;
        let value_end = value_start + tag[value_start..].find(quote)?;
        return Some(format!(
            "{}{value}{}",
            &tag[..value_start],
            &tag[value_end..]
        ));
    }
    None
}
