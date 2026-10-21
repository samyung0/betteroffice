//! The chat agent's per-call limits (Epo 2026-10-21, 2026-10-22): the engine
//! alone counts them, from one table, `limits.json`, which Capy's gen:openapi
//! also writes into edit_office_document's description. A call applies its
//! commands up to the first that would pass a limit and names the rest in
//! its notice, unapplied; people's own edits are not counted. A format adds
//! its counts to the table.

use std::collections::BTreeMap;
use std::sync::OnceLock;

use serde::Deserialize;

use crate::error::{EditCode, Error, Result};
use crate::types::{Command, Format};

#[derive(Deserialize)]
struct Table {
    /// Commands a call holds, in every format.
    commands: u64,
    /// Each format's counts.
    formats: BTreeMap<String, Vec<Limit>>,
}

/// One count: the commands that add to it and the most a call holds.
#[derive(Deserialize)]
struct Limit {
    count: String,
    limit: u64,
    commands: Vec<String>,
    /// Counts the text a command writes instead of [`counted`]'s amount.
    #[serde(default)]
    text: Option<Text>,
}

/// What a count takes of the text a PPTX shape or text box gets.
#[derive(Clone, Copy, Deserialize)]
#[serde(rename_all = "lowercase")]
enum Text {
    Paragraphs,
    Characters,
}

/// The paragraphs or characters of text `command` writes into a shape.
fn written(command: &Command, text: Text) -> u64 {
    let paragraphs = match command {
        Command::SetShapeText { paragraphs, .. } | Command::AddTextBox { paragraphs, .. } => {
            paragraphs.as_slice()
        }
        _ => &[],
    };
    match text {
        Text::Paragraphs => paragraphs.len() as u64,
        Text::Characters => paragraphs
            .iter()
            .map(|paragraph| paragraph.chars().count() as u64)
            .sum(),
    }
}

fn table() -> Result<&'static Table> {
    static TABLE: OnceLock<std::result::Result<Table, String>> = OnceLock::new();
    TABLE
        .get_or_init(|| {
            serde_json::from_str(include_str!("../limits.json")).map_err(|error| error.to_string())
        })
        .as_ref()
        .map_err(|error| Error::engine(format!("limits.json: {error}")))
}

/// A command's tool name and what it adds to the counts naming it: cells by
/// range area or grid, rows and columns by count, a list by its paragraphs,
/// a table by its cells, 1 otherwise.
fn counted(command: &Command) -> (&'static str, u64) {
    let cells = |grid: &[Vec<String>]| grid.iter().map(|row| row.len() as u64).sum();
    match command {
        Command::SetCell { .. } => ("set_cell", 1),
        Command::SetRange { range, .. } => ("set_range", crate::xlsx::range_cells(range)),
        Command::FormatCells { range, .. } => ("format_cells", crate::xlsx::range_cells(range)),
        Command::AddSheet { .. } => ("add_sheet", 1),
        Command::RenameSheet { .. } => ("rename_sheet", 1),
        Command::DeleteSheet { .. } => ("delete_sheet", 1),
        Command::InsertRows { count, .. } => ("insert_rows", u64::from(*count)),
        Command::InsertColumns { count, .. } => ("insert_columns", u64::from(*count)),
        Command::DeleteRows { expected, .. } => ("delete_rows", cells(expected)),
        Command::DeleteColumns { expected, .. } => ("delete_columns", cells(expected)),
        Command::ReplaceText { .. } => ("replace_text", 1),
        Command::InsertParagraph { .. } => ("insert_paragraph", 1),
        Command::DeleteParagraph { .. } => ("delete_paragraph", 1),
        Command::SetList { ids, .. } => ("set_list", ids.len() as u64),
        Command::InsertTable { rows, .. } => ("insert_table", cells(rows)),
        Command::DeleteTable { expected, .. } => ("delete_table", cells(expected)),
        Command::FormatText { .. } => ("format_text", 1),
        Command::SetAlignment { .. } => ("set_alignment", 1),
        Command::SetShapeText { .. } => ("set_shape_text", 1),
        Command::AddSlide { .. } => ("add_slide", 1),
        Command::DeleteSlide { .. } => ("delete_slide", 1),
        Command::MoveSlide { .. } => ("move_slide", 1),
        Command::AddTextBox { .. } => ("add_text_box", 1),
    }
}

/// `10000` as `10,000`.
fn number(value: u64) -> String {
    let digits = value.to_string();
    let mut out = String::new();
    for (index, digit) in digits.chars().enumerate() {
        if index > 0 && (digits.len() - index).is_multiple_of(3) {
            out.push(',');
        }
        out.push(digit);
    }
    out
}

/// How many leading `commands` a call applies, and the notice naming the
/// rest when a limit stops it. `first` is the index the agent's call gives
/// `commands[0]` (1 after a create command, which counts toward the
/// commands). Refused when the first command alone passes a limit.
pub(crate) fn fitting(
    format: Format,
    commands: &[Command],
    first: usize,
) -> Result<(usize, Option<String>)> {
    let table = table()?;
    let limits = table
        .formats
        .get(format.as_str())
        .map_or(&[][..], Vec::as_slice);
    let mut counts = vec![0u64; limits.len()];
    for (index, command) in commands.iter().enumerate() {
        let mut passed = ((first + index) as u64 >= table.commands)
            .then(|| format!("{} commands", number(table.commands)));
        let (name, amount) = counted(command);
        for (limit, count) in limits.iter().zip(counts.iter_mut()) {
            if limit.commands.iter().any(|command| command == name) {
                *count += limit.text.map_or(amount, |text| written(command, text));
                if *count > limit.limit && passed.is_none() {
                    passed = Some(format!("{} {}", number(limit.limit), limit.count));
                }
            }
        }
        let Some(passed) = passed else {
            continue;
        };
        if index == 0 {
            return Err(Error::edit(
                EditCode::InvalidInput,
                format!(
                    "the command at index {first} alone passes the limit of {passed} a call; split it"
                ),
            ));
        }
        let (from, to) = (first + index, first + commands.len() - 1);
        let which = if from == to {
            format!("the command at index {from}")
        } else {
            format!("the commands at indexes {from} to {to}")
        };
        return Ok((
            index,
            Some(format!(
                "Not applied: {which} (counting from 0), past the limit of {passed} a call; send them in another call."
            )),
        ));
    }
    Ok((commands.len(), None))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn format_text() -> Command {
        Command::FormatText {
            target_id: "body:paragraph:1".into(),
            expected_text: "x".into(),
            format: Default::default(),
        }
    }

    /// The table parses, and every command it names is one `counted` names.
    #[test]
    fn the_table_names_known_commands() {
        let table = table().unwrap();
        assert_eq!(table.commands, 50);
        for limits in table.formats.values() {
            for limit in limits {
                for name in &limit.commands {
                    assert!(
                        [
                            "set_cell",
                            "set_range",
                            "format_cells",
                            "add_sheet",
                            "rename_sheet",
                            "delete_sheet",
                            "insert_rows",
                            "insert_columns",
                            "delete_rows",
                            "delete_columns",
                            "insert_paragraph",
                            "delete_paragraph",
                            "set_list",
                            "insert_table",
                            "delete_table",
                            "format_text",
                            "set_alignment",
                            "set_shape_text",
                            "add_slide",
                            "delete_slide",
                            "move_slide",
                            "add_text_box",
                        ]
                        .contains(&name.as_str()),
                        "{name}"
                    );
                }
            }
        }
    }

    /// The commands before the first past a limit apply; the notice names
    /// the rest by index, after a create command one later; a first command
    /// alone past a limit is refused.
    #[test]
    fn a_call_applies_up_to_its_first_command_past_a_limit() {
        let spans = vec![format_text(); 33];
        assert_eq!(
            fitting(Format::Docx, &spans, 0).unwrap(),
            (
                30,
                Some(
                    "Not applied: the commands at indexes 30 to 32 (counting from 0), past the limit of 30 formatted spans a call; send them in another call."
                        .into()
                )
            )
        );
        // After a create command the 50 commands stop a call one earlier.
        let many = vec![
            Command::AddSlide {
                after_slide: None,
                before_slide: None,
                layout: None,
            };
            51
        ];
        assert_eq!(
            fitting(Format::Xlsx, &many, 1).unwrap(),
            (
                49,
                Some(
                    "Not applied: the commands at indexes 50 to 51 (counting from 0), past the limit of 50 commands a call; send them in another call."
                        .into()
                )
            )
        );
        let rows = |count| Command::InsertRows {
            sheet: "Sheet1".into(),
            at: 1,
            count,
        };
        assert_eq!(
            fitting(Format::Xlsx, &[rows(600), rows(401)], 0)
                .unwrap()
                .1
                .unwrap(),
            "Not applied: the command at index 1 (counting from 0), past the limit of 1,000 inserted rows and columns a call; send them in another call."
        );
        assert_eq!(
            fitting(Format::Xlsx, &[rows(1_001)], 0).unwrap_err(),
            Error::edit(
                EditCode::InvalidInput,
                "the command at index 0 alone passes the limit of 1,000 inserted rows and columns a call; split it"
            )
        );
        assert_eq!(fitting(Format::Xlsx, &[rows(1_000)], 0).unwrap(), (1, None));
    }
}
