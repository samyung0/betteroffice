//! The chat agent's XLSX commands, scoped reads and `create` (Epo
//! 2026-10-20): ranges, formats, sheets, rows and columns, the formula rules
//! and the starter workbook.

use betteroffice_xlsx::{
    CellRange, CellRef, DefinedName, NumberFormatMutation, Sheet, SheetId, StylePatch, Workbook,
    WorkbookModel,
};
use office_service::{
    CellStyle, Checkpoint, Command, CommandResult, Determinism, EditCode, Error, Format,
    InspectQuery, apply_commands, create, export, inspect_scoped, seed, sha256_hex, starter,
};
use serde_json::{Value, json};

const DETERMINISM: Determinism = Determinism {
    seed: "0000000000000000000000000000000000000000000000000000000000000000",
    now: "2026-10-20T00:00:00.000Z",
};

/// A seeded source and its latest state.
struct Book {
    base: Vec<u8>,
    sha: String,
    state: Vec<u8>,
}

impl Book {
    fn new(base: Vec<u8>) -> Self {
        let state = seed(Format::Xlsx, &base).unwrap();
        Self {
            sha: sha256_hex(&base),
            base,
            state,
        }
    }

    /// The starter with `commands` applied.
    fn created(commands: &[Command]) -> Self {
        Self::new(blank_with(commands))
    }

    fn checkpoint(&self) -> Checkpoint<'_> {
        Checkpoint {
            format: Format::Xlsx,
            schema_version: 1,
            base_sha256: &self.sha,
            state: &self.state,
        }
    }

    /// Applies `commands`, keeping the state only when they land.
    fn apply(&mut self, commands: &[Command]) -> Result<CommandResult, Error> {
        let result = apply_commands(&self.base, self.checkpoint(), commands)?;
        self.state = result.state.clone();
        Ok(result)
    }

    fn read(&self, sheet: Option<&str>, range: Option<&str>) -> Result<Value, Error> {
        inspect_scoped(
            &self.base,
            self.checkpoint(),
            &InspectQuery {
                sheet: sheet.map(str::to_owned),
                range: range.map(str::to_owned),
            },
        )
    }

    /// `address=input→value` for each non-empty cell of `range`.
    fn cells(&self, sheet: &str, range: &str) -> Vec<String> {
        self.read(Some(sheet), Some(range)).unwrap()["cells"]
            .as_array()
            .unwrap()
            .iter()
            .map(|cell| {
                format!(
                    "{}={}→{}",
                    cell["address"].as_str().unwrap(),
                    cell["input"].as_str().unwrap(),
                    cell["value"].as_str().unwrap()
                )
            })
            .collect()
    }

    fn sheet_names(&self) -> Vec<String> {
        self.read(None, None).unwrap()["sheets"]
            .as_array()
            .unwrap()
            .iter()
            .map(|sheet| sheet["name"].as_str().unwrap().to_owned())
            .collect()
    }

    fn sheet_id(&self, name: &str) -> String {
        self.read(None, None).unwrap()["sheets"]
            .as_array()
            .unwrap()
            .iter()
            .find(|sheet| sheet["name"] == name)
            .unwrap()["id"]
            .as_str()
            .unwrap()
            .to_owned()
    }
}

fn blank_with(commands: &[Command]) -> Vec<u8> {
    create(
        Format::Xlsx,
        starter(Format::Xlsx).unwrap(),
        commands,
        DETERMINISM,
    )
    .unwrap()
    .bytes
}

fn grid(rows: &[&[&str]]) -> Vec<Vec<String>> {
    rows.iter()
        .map(|row| row.iter().map(|cell| (*cell).to_owned()).collect())
        .collect()
}

fn set_range(sheet: &str, range: &str, expected: &[&[&str]], values: &[&[&str]]) -> Command {
    Command::SetRange {
        sheet: sheet.into(),
        range: range.into(),
        expected: grid(expected),
        values: grid(values),
    }
}

fn refusal(code: EditCode, message: &str, details: Option<Value>) -> Error {
    Error::Edit {
        code,
        message: message.into(),
        details,
    }
}

fn invalid(message: &str) -> Error {
    refusal(EditCode::InvalidInput, message, None)
}

#[test]
fn set_range_writes_values_formulas_and_empties_and_its_inverse_restores_them() {
    let mut book = Book::created(&[set_range(
        "Sheet1",
        "A1:B2",
        &[&["", ""], &["", ""]],
        &[&["1", "2"], &["3", ""]],
    )]);
    let before = book.cells("Sheet1", "A1:C2");
    assert_eq!(before, ["A1=1→1", "B1=2→2", "A2=3→3"]);

    let edit = book
        .apply(&[set_range(
            "sheet:0",
            "C2:A1",
            &[&["1", "2", ""], &["3", "", ""]],
            &[&["10", "=A1*2", ""], &["", "x", "'007"]],
        )])
        .unwrap();
    let after = book.cells("Sheet1", "A1:C2");
    assert_eq!(after, ["A1=10→10", "B1==A1*2→20", "B2=x→x", "C2='007→007"]);
    assert_eq!(edit.targets.len(), 5, "every changed cell is a target");
    assert!(edit.targets.iter().all(|target| {
        target.path[..3] == ["xlsx:sheets", "sheet:0", "contents"]
            && target.id.starts_with("sheet:0:[")
    }));
    assert_eq!(
        edit.inverse,
        Some(vec![set_range(
            "sheet:0",
            "A1:C2",
            &[&["10", "=A1*2", ""], &["", "x", "'007"]],
            &[&["1", "2", ""], &["3", "", ""]],
        )])
    );

    let undo = book.apply(&edit.inverse.unwrap()).unwrap();
    assert_eq!(book.cells("Sheet1", "A1:C2"), before);
    book.apply(&undo.inverse.unwrap()).unwrap();
    assert_eq!(book.cells("Sheet1", "A1:C2"), after);
}

#[test]
fn a_stale_set_range_carries_the_current_inputs_as_tsv() {
    let mut book = Book::created(&[set_range(
        "Sheet1",
        "A1:B2",
        &[&["", ""], &["", ""]],
        &[&["1", "=A1*2"], &["3", ""]],
    )]);
    // The shown value matches as well as the input.
    let shown = set_range("Sheet1", "B1", &[&["2"]], &[&["7"]]);
    assert!(book.clone_apply(&shown).is_ok());

    assert_eq!(
        book.apply(&[set_range(
            "Sheet1",
            "A1:B2",
            &[&["1", "=A1*2"], &["4", ""]],
            &[&["", ""], &["", ""]],
        )])
        .unwrap_err(),
        refusal(
            EditCode::StaleTarget,
            "1 of the range's cells differ from expected; read the range again",
            Some(json!({ "block": {
                "id": "Sheet1!A1:B2",
                "type": "range",
                "text": "1\t=A1*2\n3\t",
            }})),
        )
    );
}

impl Book {
    fn clone_apply(&self, command: &Command) -> Result<CommandResult, Error> {
        apply_commands(&self.base, self.checkpoint(), std::slice::from_ref(command))
    }
}

#[test]
fn set_range_refuses_a_wrong_shape_or_too_many_cells() {
    let mut book = Book::created(&[]);
    assert_eq!(
        book.apply(&[set_range("Sheet1", "A1:B2", &[&["", ""]], &[&["1", "2"]])])
            .unwrap_err(),
        invalid("expected must be 2 rows of 2 cells, the shape of A1:B2")
    );
    assert_eq!(
        book.apply(&[set_range("Sheet1", "A1:B5001", &[], &[])])
            .unwrap_err(),
        invalid("A1:B5001 has 10002 cells; set at most 10000 cells per command")
    );
    let Error::Edit { code, details, .. } = book
        .apply(&[set_range("Nope", "A1", &[&[""]], &[&["1"]])])
        .unwrap_err()
    else {
        panic!("an edit refusal");
    };
    assert_eq!(code, EditCode::UnavailableTarget);
    assert_eq!(
        details,
        Some(json!({ "blocks": [{ "id": "sheet:0", "type": "sheet", "text": "Sheet1" }] }))
    );
}

#[test]
fn format_cells_patches_the_style_and_number_format_without_an_inverse() {
    let mut book = Book::created(&[set_range("Sheet1", "A1", &[&[""]], &[&["3.5"]])]);
    let edit = book
        .apply(&[Command::FormatCells {
            sheet: "Sheet1".into(),
            range: "A1:B2".into(),
            style: CellStyle {
                patch: StylePatch {
                    bold: Some(true),
                    fill_color: Some("#ffcc00".into()),
                    ..StylePatch::default()
                },
                number_format: Some(NumberFormatMutation::Custom {
                    pattern: "0.00".into(),
                }),
            },
        }])
        .unwrap();
    assert_eq!(edit.inverse, None);
    assert_eq!(book.cells("Sheet1", "A1"), ["A1=3.5→3.50"]);
    let saved = export(&book.base, book.checkpoint(), DETERMINISM).unwrap();
    let formatting = Workbook::open(&saved)
        .unwrap()
        .selection_formatting(
            SheetId(0),
            CellRange::new(CellRef::new(0, 0), CellRef::new(1, 1)),
        )
        .unwrap();
    assert_eq!(formatting.bold, Some(true));
    assert_eq!(formatting.fill_color.as_deref(), Some("#ffcc00"));

    let format = |style: CellStyle| {
        book.clone_apply(&Command::FormatCells {
            sheet: "Sheet1".into(),
            range: "A1".into(),
            style,
        })
    };
    assert_eq!(
        format(CellStyle {
            patch: StylePatch {
                text_color: Some("red".into()),
                ..StylePatch::default()
            },
            number_format: None,
        })
        .unwrap_err(),
        invalid("textColor \"red\" must be a #RRGGBB colour")
    );
    assert!(matches!(
        format(CellStyle::default()).unwrap_err(),
        Error::Edit { code: EditCode::InvalidInput, message, .. } if message.starts_with("style sets nothing")
    ));
    assert_eq!(
        book.clone_apply(&Command::FormatCells {
            sheet: "Sheet1".into(),
            range: "A1:B5001".into(),
            style: CellStyle {
                patch: StylePatch {
                    bold: Some(true),
                    ..StylePatch::default()
                },
                number_format: None,
            },
        })
        .unwrap_err(),
        invalid("A1:B5001 has 10002 cells; format at most 10000 cells per command")
    );
}

#[test]
fn sheets_are_added_renamed_with_an_inverse_and_deleted() {
    let mut book = Book::created(&[]);
    book.apply(&[
        Command::AddSheet {
            name: "Data".into(),
            after: None,
        },
        set_range("Data", "A1", &[&[""]], &[&["5"]]),
        set_range("Sheet1", "A1", &[&[""]], &[&["=Data!A1*2"]]),
        Command::AddSheet {
            name: "Notes".into(),
            after: Some("Sheet1".into()),
        },
    ])
    .unwrap();
    assert_eq!(book.sheet_names(), ["Sheet1", "Notes", "Data"]);
    assert_eq!(book.cells("Sheet1", "A1"), ["A1==Data!A1*2→10"]);

    let data = book.sheet_id("Data");
    let rename = book
        .apply(&[Command::RenameSheet {
            sheet: data.clone(),
            expected_name: "Data".into(),
            name: "Grades".into(),
        }])
        .unwrap();
    assert_eq!(book.sheet_names(), ["Sheet1", "Notes", "Grades"]);
    assert_eq!(book.cells("Sheet1", "A1"), ["A1=='Grades'!A1*2→10"]);
    assert_eq!(
        book.clone_apply(&Command::DeleteSheet {
            sheet: data.clone(),
            expected_name: "Data".into(),
        })
        .unwrap_err(),
        refusal(
            EditCode::StaleTarget,
            "the sheet is named \"Grades\", not \"Data\"",
            Some(json!({ "block": { "id": data, "type": "sheet", "text": "Grades" } })),
        )
    );

    book.apply(&rename.inverse.unwrap()).unwrap();
    assert_eq!(book.sheet_names(), ["Sheet1", "Notes", "Data"]);
    assert_eq!(book.cells("Sheet1", "A1"), ["A1==Data!A1*2→10"]);

    let delete = book
        .apply(&[Command::DeleteSheet {
            sheet: "Notes".into(),
            expected_name: "Notes".into(),
        }])
        .unwrap();
    assert_eq!(delete.inverse, None);
    assert_eq!(book.sheet_names(), ["Sheet1", "Data"]);

    let mut single = Book::created(&[]);
    assert_eq!(
        single
            .apply(&[Command::DeleteSheet {
                sheet: "Sheet1".into(),
                expected_name: "Sheet1".into(),
            }])
            .unwrap_err(),
        invalid("\"Sheet1\" is the workbook's only sheet; a workbook keeps at least one")
    );
}

#[test]
fn rows_and_columns_are_inserted_and_deleted_when_their_cells_are_as_expected() {
    let mut book = Book::created(&[set_range(
        "Sheet1",
        "A1:B3",
        &[&["", ""], &["", ""], &["", ""]],
        &[&["1", "x"], &["2", "y"], &["3", ""]],
    )]);
    let start = book.cells("Sheet1", "A1:D4");

    let insert = book
        .apply(&[Command::InsertRows {
            sheet: "Sheet1".into(),
            at: 2,
            count: 1,
        }])
        .unwrap();
    assert_eq!(
        book.cells("Sheet1", "A1:A4"),
        ["A1=1→1", "A3=2→2", "A4=3→3"]
    );
    assert_eq!(
        insert.inverse,
        Some(vec![Command::DeleteRows {
            sheet: "sheet:0".into(),
            at: 2,
            count: 1,
            expected: String::new(),
        }])
    );
    book.apply(&insert.inverse.unwrap()).unwrap();
    assert_eq!(book.cells("Sheet1", "A1:D4"), start);

    assert_eq!(
        book.clone_apply(&Command::DeleteRows {
            sheet: "Sheet1".into(),
            at: 1,
            count: 1,
            expected: "1".into(),
        })
        .unwrap_err(),
        refusal(
            EditCode::StaleTarget,
            "1:1 differs from expected; read it again",
            Some(json!({ "block": { "id": "Sheet1!1:1", "type": "range", "text": "1\tx" } })),
        )
    );
    let delete = book
        .apply(&[Command::DeleteRows {
            sheet: "Sheet1".into(),
            at: 2,
            count: 1,
            expected: "2\ty".into(),
        }])
        .unwrap();
    assert_eq!(delete.inverse, None);
    assert_eq!(
        book.cells("Sheet1", "A1:B3"),
        ["A1=1→1", "B1=x→x", "A2=3→3"]
    );

    book.apply(&[
        Command::DeleteColumns {
            sheet: "Sheet1".into(),
            at: "b".into(),
            count: 1,
            expected: "x".into(),
        },
        Command::InsertColumns {
            sheet: "Sheet1".into(),
            at: "A".into(),
            count: 2,
        },
    ])
    .unwrap();
    assert_eq!(book.cells("Sheet1", "A1:D3"), ["C1=1→1", "C2=3→3"]);
    assert_eq!(
        book.clone_apply(&Command::DeleteColumns {
            sheet: "Sheet1".into(),
            at: "C".into(),
            count: 1,
            expected: String::new(),
        })
        .unwrap_err(),
        refusal(
            EditCode::StaleTarget,
            "C:C differs from expected; read it again",
            Some(json!({ "block": { "id": "Sheet1!C:C", "type": "range", "text": "1\n3" } })),
        )
    );
}

#[test]
fn a_row_or_column_delete_removes_at_most_ten_thousand_cells() {
    let column: Vec<Vec<String>> = (1..=10_000).map(|row| vec![row.to_string()]).collect();
    let mut book = Book::created(&[
        Command::SetRange {
            sheet: "Sheet1".into(),
            range: "A1:A10000".into(),
            expected: vec![vec![String::new()]; 10_000],
            values: column,
        },
        set_range("Sheet1", "A10001", &[&[""]], &[&["last"]]),
    ]);
    assert_eq!(
        book.apply(&[Command::DeleteColumns {
            sheet: "Sheet1".into(),
            at: "A".into(),
            count: 1,
            expected: String::new(),
        }])
        .unwrap_err(),
        invalid("A:A holds 10001 cells; delete at most 10000 cells per command")
    );
}

/// The rules for formulas the agent writes (Epo 2026-10-20), one case each,
/// next to formulas that pass.
#[test]
fn agent_formulas_pass_the_rules_or_are_refused_naming_one() {
    let mut model = WorkbookModel::default();
    model.sheets.push(Sheet::new("Sheet1"));
    model.sheets.push(Sheet::new("Data"));
    for (name, formula) in [("Rate", "Data!$A$1"), ("Ext", "'[1]Sheet1'!$A$1")] {
        model.defined_names.push(DefinedName {
            name: name.into(),
            formula: formula.into(),
            local_sheet: None,
            hidden: false,
        });
    }
    let base = Workbook::from_model(model).unwrap().save().unwrap();
    let book = Book::new(base);
    let long_formula = format!("={}1", "1+".repeat(1000));
    let long_text = "x".repeat(32_768);
    let cases: [(&str, Option<&str>); 20] = [
        ("=SUM(Data!A1:A3)+VLOOKUP(1,Data!A1:B3,2,FALSE)", None),
        ("=IF(Rate>1,\"big\",\"small\")", None),
        ("='Data'!A1*Rate+_xlfn.XLOOKUP(1,Data!A:A,Data!B:B)", None),
        ("plain text", None),
        (
            "=HYPERLINK(\"http://example.com\")",
            Some("the HYPERLINK function is not allowed"),
        ),
        (
            "=WEBSERVICE(\"http://example.com\")",
            Some(
                "the function WEBSERVICE is not available; use a function the spreadsheet engine supports",
            ),
        ),
        (
            "=_xll.EVIL()",
            Some(
                "the function _xll.EVIL is not available; use a function the spreadsheet engine supports",
            ),
        ),
        (
            "=PY(1)",
            Some(
                "the function PY is not available; use a function the spreadsheet engine supports",
            ),
        ),
        (
            "=CELL(\"filename\")",
            Some(
                "the function CELL is not available; use a function the spreadsheet engine supports",
            ),
        ),
        (
            "='[1]Sheet1'!A1",
            Some("the reference to '[1]Sheet1' names another workbook, which is not allowed"),
        ),
        (
            "=[Book.xlsx]Sheet1!A1",
            Some("the formula references another workbook ([...]), which is not allowed"),
        ),
        (
            "='C:\\files\\Sheet1'!A1",
            Some(
                "the reference to 'C:\\files\\Sheet1' names a URL or a file path, which is not allowed",
            ),
        ),
        (
            "='http://example.com/x'!A1",
            Some(
                "the reference to 'http://example.com/x' names a URL or a file path, which is not allowed",
            ),
        ),
        (
            "=cmd|' /c calc'!A0",
            Some("the formula uses '|', a DDE link, which is not allowed"),
        ),
        (
            "=Nope!A1",
            Some("the sheet 'Nope' is not in this workbook; its sheets are 'Sheet1', 'Data'"),
        ),
        (
            "=Missing+1",
            Some("the name Missing is not defined in this workbook"),
        ),
        (
            "=Ext+1",
            Some(
                "the name Ext refers outside this workbook ('[1]Sheet1'!$A$1), which is not allowed",
            ),
        ),
        ("=1+", Some("the formula does not parse")),
        (
            &long_formula,
            Some("the formula is 2001 characters long; a formula may hold at most 2000"),
        ),
        (
            &long_text,
            Some("the text is 32768 characters long; a cell holds at most 32767"),
        ),
    ];
    for (value, refused) in cases {
        let result = book.clone_apply(&set_range("Sheet1", "B2", &[&[""]], &[&[value]]));
        match refused {
            None => assert!(result.is_ok(), "{value}: {result:?}"),
            Some(rule) => match result.unwrap_err() {
                Error::Edit {
                    code: EditCode::InvalidInput,
                    message,
                    ..
                } => assert!(
                    message.starts_with(&format!("B2: {rule}")),
                    "{value}: {message}"
                ),
                other => panic!("{value}: {other:?}"),
            },
        }
    }
    // set_cell, the other way an agent writes a cell, passes the same rules.
    assert_eq!(
        book.clone_apply(&Command::SetCell {
            sheet: "Sheet1".into(),
            cell: "C3".into(),
            expected_value: String::new(),
            value: "=HYPERLINK(\"x\")".into(),
        })
        .unwrap_err(),
        invalid("C3: the HYPERLINK function is not allowed")
    );
}

#[test]
fn scoped_reads_list_sheets_and_read_a_bounded_range() {
    let column: Vec<Vec<String>> = (1..=1500)
        .map(|row| vec![row.to_string(), format!("=A{row}*2")])
        .collect();
    let book = Book::created(&[
        Command::SetRange {
            sheet: "Sheet1".into(),
            range: "A1:B1500".into(),
            expected: vec![vec![String::new(); 2]; 1500],
            values: column,
        },
        Command::AddSheet {
            name: "Empty".into(),
            after: None,
        },
    ]);
    assert_eq!(
        book.read(None, None).unwrap(),
        json!({ "sheets": [
            { "id": "sheet:0", "name": "Sheet1", "usedRange": "A1:B1500", "rows": 1500, "columns": 2 },
            { "id": "sheet:1", "name": "Empty", "usedRange": null, "rows": 0, "columns": 0 },
        ]})
    );
    assert_eq!(
        book.read(Some("Sheet1"), Some("B1499:C1501")).unwrap(),
        json!({
            "sheet": { "id": "sheet:0", "name": "Sheet1", "usedRange": "A1:B1500", "rows": 1500, "columns": 2 },
            "range": "B1499:C1501",
            "cells": [
                { "address": "B1499", "input": "=A1499*2", "value": "2998" },
                { "address": "B1500", "input": "=A1500*2", "value": "3000" },
            ],
        })
    );
    assert_eq!(
        book.read(Some("Sheet1"), Some("A1:B1001")).unwrap_err(),
        invalid("A1:B1001 has 2002 cells; read at most 2000 cells at a time, as a smaller range")
    );
}

#[test]
fn create_applies_commands_to_the_starter_and_exports_a_workbook() {
    let created = create(
        Format::Xlsx,
        starter(Format::Xlsx).unwrap(),
        &[
            set_range(
                "Sheet1",
                "A1:B2",
                &[&["", ""], &["", ""]],
                &[&["Name", "Score"], &["Ada\u{1}", "=40+2"]],
            ),
            Command::AddSheet {
                name: "Notes".into(),
                after: None,
            },
        ],
        DETERMINISM,
    )
    .unwrap();
    assert_eq!(
        created.notice.as_deref(),
        Some("Removed 1 character Office files cannot store (U+0001).")
    );
    let workbook = Workbook::open(&created.bytes).unwrap();
    assert_eq!(
        workbook
            .model()
            .sheets
            .iter()
            .map(|sheet| sheet.name.as_str())
            .collect::<Vec<_>>(),
        ["Sheet1", "Notes"]
    );
    let input = |a1: &str| {
        workbook
            .cell(SheetId(0), CellRef::parse_a1(a1).unwrap())
            .unwrap()
            .input
    };
    assert_eq!(
        [input("A1"), input("B1"), input("A2"), input("B2")],
        ["Name", "Score", "Ada", "=40+2"]
    );
}

/// The committed starter (the blank workbook Excel saved, starters/NOTICE.md)
/// seeds, exports and reopens as one empty sheet "Sheet1", and `create`
/// builds on it.
#[test]
fn the_xlsx_starter_seeds_exports_and_reopens_as_one_empty_sheet() {
    let book = Book::new(starter(Format::Xlsx).unwrap().to_vec());
    let exported =
        Workbook::open(&export(&book.base, book.checkpoint(), DETERMINISM).unwrap()).unwrap();
    assert_eq!(exported.sheet_count(), 1);
    assert_eq!(exported.model().sheets[0].name, "Sheet1");
    assert_eq!(exported.model().sheets[0].iter_cells().count(), 0);

    let created = Book::created(&[set_range("Sheet1", "A1", &[&[""]], &[&["=1+1"]])]);
    assert_eq!(created.cells("Sheet1", "A1"), ["A1==1+1→2"]);
}

/// The wire shapes Capy sends: snake_case `type`, camelCase fields.
#[test]
fn commands_parse_from_their_wire_json() {
    let commands: Vec<Command> = serde_json::from_value(json!([
        { "type": "set_range", "sheet": "sheet:0", "range": "A1:B1", "expected": [["", ""]], "values": [["1", "=A1+1"]] },
        { "type": "format_cells", "sheet": "Sheet1", "range": "A1:B1", "style": {
            "bold": true, "fillColor": "#ffcc00", "horizontalAlignment": "center",
            "numberFormat": { "type": "custom", "pattern": "0.00" } } },
        { "type": "add_sheet", "name": "Data", "after": "Sheet1" },
        { "type": "rename_sheet", "sheet": "sheet:1", "expectedName": "Data", "name": "Grades" },
        { "type": "delete_sheet", "sheet": "sheet:1", "expectedName": "Grades" },
        { "type": "insert_rows", "sheet": "Sheet1", "at": 2, "count": 3 },
        { "type": "delete_rows", "sheet": "Sheet1", "at": 2, "count": 1, "expected": "2\ty" },
        { "type": "insert_columns", "sheet": "Sheet1", "at": "B", "count": 1 },
        { "type": "delete_columns", "sheet": "Sheet1", "at": "B", "count": 1, "expected": "" },
    ]))
    .unwrap();
    assert_eq!(
        commands[1],
        Command::FormatCells {
            sheet: "Sheet1".into(),
            range: "A1:B1".into(),
            style: CellStyle {
                patch: StylePatch {
                    bold: Some(true),
                    fill_color: Some("#ffcc00".into()),
                    horizontal_alignment: Some(betteroffice_xlsx::HorizontalAlignment::Center),
                    ..StylePatch::default()
                },
                number_format: Some(NumberFormatMutation::Custom {
                    pattern: "0.00".into()
                }),
            },
        }
    );
    assert_eq!(
        commands[8],
        Command::DeleteColumns {
            sheet: "Sheet1".into(),
            at: "B".into(),
            count: 1,
            expected: String::new(),
        }
    );
    // A delete names the band it expects (Epo 2026-10-20); the collaboration
    // server answers a command that does not parse as invalid_input.
    let missing = serde_json::from_value::<Command>(
        json!({ "type": "delete_rows", "sheet": "Sheet1", "at": 2, "count": 1 }),
    )
    .unwrap_err();
    assert!(
        missing.to_string().contains("missing field `expected`"),
        "{missing}"
    );
}

#[test]
fn create_office_over_the_wire_answers_the_new_file() {
    let request = json!({
        "method": "createOffice",
        "args": [
            "xlsx",
            [{ "type": "set_range", "sheet": "Sheet1", "range": "A1:B1",
               "expected": [["", ""]], "values": [["2", "=A1*3"]] }],
            { "seed": DETERMINISM.seed, "now": DETERMINISM.now },
        ],
    });
    let answer: Value =
        serde_json::from_slice(&office_service::wire::call(request.to_string().as_bytes()))
            .unwrap();
    assert_eq!(answer["value"]["notice"], Value::Null, "{answer}");
    let bytes = office_service::wire::bytes(&answer["value"]["bytes"]).unwrap();
    assert_eq!(
        Book::new(bytes).cells("Sheet1", "A1:B1"),
        ["A1=2→2", "B1==A1*3→6"]
    );
}
