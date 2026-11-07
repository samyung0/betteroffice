//! Typed dates, times, percentages and accounting negatives are stored as
//! numbers, and a General cell takes the format they imply (upstream #1491,
//! ported onto the schema-8 engine).

use betteroffice_xlsx::{
    CalculationOptions, Cell, CellInput, CellRef, CellValue, DateSystem, ProposalEditInput,
    ProposalRequest, Sheet, SheetId, Workbook, WorkbookModel,
};
use xlsx_model::{ColStyle, NumberFormat, Xf};

fn at() -> CellRef {
    CellRef::new(0, 0)
}

fn model(format: u16, date_system: DateSystem) -> WorkbookModel {
    let mut model = WorkbookModel {
        date_system,
        ..Default::default()
    };
    model.styles.cell_xfs.push(Xf {
        num_fmt_id: Some(format),
        ..Default::default()
    });
    let mut sheet = Sheet::new("Sheet1");
    sheet.set_cell(
        at(),
        Cell {
            value: CellValue::Number { value: 45292.0 },
            style: Some(0),
            ..Default::default()
        },
    );
    model.sheets.push(sheet);
    model
}

fn fixture(format: u16, date_system: DateSystem) -> Vec<u8> {
    let parts = xlsx_parse::serialize_workbook(&model(format, date_system)).unwrap();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn assert_cell(workbook: &Workbook, value: f64, format: u16) {
    let sheet = &workbook.model().sheets[0];
    assert_eq!(sheet.cell(at()).unwrap().value, CellValue::Number { value });
    assert_eq!(
        workbook
            .model()
            .styles
            .cell_format(sheet.effective_style(at()))
            .number_format,
        NumberFormat::Builtin { id: format }
    );
}

fn assert_saved(workbook: &Workbook, value: f64, format: u16) {
    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_cell(&reopened, value, format);
}

#[test]
fn date_entry_in_date_cell_saves_numeric_and_preserves_format() {
    let mut workbook = Workbook::open(&fixture(14, DateSystem::V1900)).unwrap();
    let options = CalculationOptions::default();
    let before = workbook.model().sheets[0].cell(at()).unwrap().clone();
    workbook
        .edit_cell(SheetId(0), at(), "1/15/2024", options)
        .unwrap();
    assert_cell(&workbook, 45306.0, 14);
    assert_eq!(
        workbook.model().sheets[0].cell(at()).unwrap().style,
        before.style
    );
    assert_saved(&workbook, 45306.0, 14);
    workbook.undo(options).unwrap();
    assert_eq!(workbook.model().sheets[0].cell(at()), Some(&before));
    workbook.redo(options).unwrap();
    assert_saved(&workbook, 45306.0, 14);
}

#[test]
fn general_entry_applies_date_and_percent_formats_in_one_history_step() {
    for (input, value, format) in [("1/15/2024", 45306.0, 14), ("15%", 0.15, 9)] {
        let mut workbook = Workbook::open(&fixture(0, DateSystem::V1900)).unwrap();
        let options = CalculationOptions::default();
        let before = workbook.model().sheets[0].cell(at()).unwrap().clone();
        workbook
            .edit_cell(SheetId(0), at(), input, options)
            .unwrap();
        assert_cell(&workbook, value, format);
        assert_saved(&workbook, value, format);
        workbook.undo(options).unwrap();
        assert_cell(&workbook, 45292.0, 0);
        assert_eq!(workbook.model().sheets[0].cell(at()), Some(&before));
        workbook.redo(options).unwrap();
        assert_saved(&workbook, value, format);
    }
}

#[test]
fn batch_entry_uses_clock_and_workbook_epoch_and_undoes_format() {
    let mut workbook = Workbook::open(&fixture(0, DateSystem::V1904)).unwrap();
    let options = CalculationOptions {
        now_serial: Some(45306.0),
    };
    workbook
        .edit_cells(
            SheetId(0),
            &[CellInput {
                cell: at(),
                input: "1/15 12:00".into(),
            }],
            options,
        )
        .unwrap();
    assert_saved(&workbook, 43844.5, 22);
    workbook.undo(options).unwrap();
    assert_cell(&workbook, 45292.0, 0);
    workbook.redo(options).unwrap();
    assert_saved(&workbook, 43844.5, 22);
}

#[test]
fn entry_respects_inherited_date_order_and_existing_formats() {
    let mut model = model(0, DateSystem::V1900);
    model.styles.num_fmts.push((164, "dd.mm.yyyy".into()));
    model.styles.cell_xfs.push(Xf {
        num_fmt_id: Some(164),
        ..Default::default()
    });
    model.sheets[0].set_cell(at(), Cell::default());
    model.sheets[0].col_styles.push(ColStyle {
        first: 0,
        last: 0,
        xf: 1,
    });
    let mut workbook = Workbook::from_model(model).unwrap();
    workbook
        .edit_cell(SheetId(0), at(), "15/1/2024", CalculationOptions::default())
        .unwrap();
    let sheet = &workbook.model().sheets[0];
    assert_eq!(
        sheet.cell(at()).unwrap().value,
        CellValue::Number { value: 45306.0 }
    );
    assert_eq!(
        workbook
            .model()
            .styles
            .cell_format(sheet.effective_style(at()))
            .number_format,
        NumberFormat::Custom {
            pattern: "dd.mm.yyyy".into()
        }
    );
}

#[test]
fn collaborative_entry_and_history_replicate_value_and_format() {
    let sync = |from: &mut Workbook, to: &mut Workbook| {
        let update = from.encode_state_as_update_v1();
        to.apply_update_v1(&update, CalculationOptions::default())
            .unwrap();
    };
    for (input, value, format) in [("1/15/2024", 45306.0, 14), ("15%", 0.15, 9)] {
        let mut local =
            Workbook::from_model_collaborative(model(0, DateSystem::V1900), 11).unwrap();
        let mut peer = Workbook::from_model_collaborative(model(0, DateSystem::V1900), 12).unwrap();
        let options = CalculationOptions::default();
        local.edit_cell(SheetId(0), at(), input, options).unwrap();
        sync(&mut local, &mut peer);
        assert_cell(&peer, value, format);
        local.undo(options).unwrap();
        sync(&mut local, &mut peer);
        assert_cell(&peer, 45292.0, 0);
        local.redo(options).unwrap();
        sync(&mut local, &mut peer);
        assert_cell(&peer, value, format);
    }
}

#[test]
fn a_proposal_applies_its_inferred_format() {
    let mut workbook = Workbook::open(&fixture(0, DateSystem::V1900)).unwrap();
    let proposal = workbook
        .propose(
            ProposalRequest {
                agent_id: "test".into(),
                note: None,
                edits: vec![ProposalEditInput {
                    sheet: SheetId(0),
                    cell: at(),
                    input: "1/15/2024".into(),
                    number_format: None,
                }],
            },
            CalculationOptions::default(),
        )
        .unwrap();
    workbook
        .accept_proposal(&proposal.id, false, CalculationOptions::default())
        .unwrap();
    assert_saved(&workbook, 45306.0, 14);
    workbook.undo(CalculationOptions::default()).unwrap();
    assert_cell(&workbook, 45292.0, 0);
}

#[test]
fn text_read_back_from_a_cell_stays_text_when_written_unchanged() {
    let mut model = model(0, DateSystem::V1900);
    model.styles.num_fmts.push((164, "dd.mm.yyyy".into()));
    model.styles.cell_xfs.push(Xf {
        num_fmt_id: Some(164),
        ..Default::default()
    });
    let text = |value: &str| Cell {
        value: CellValue::Text {
            value: value.into(),
        },
        ..Default::default()
    };
    model.sheets[0].set_cell(CellRef::new(0, 0), text("1/15"));
    model.sheets[0].set_cell(
        CellRef::new(0, 1),
        Cell {
            style: Some(1),
            ..text("15.01.2024")
        },
    );
    let mut workbook = Workbook::from_model(model).unwrap();
    let options = CalculationOptions {
        now_serial: Some(45_000.25),
    };
    for (cell, value) in [
        (CellRef::new(0, 0), "1/15"),
        (CellRef::new(0, 1), "15.01.2024"),
    ] {
        let input = workbook.cell(SheetId(0), cell).unwrap().input;
        workbook
            .edit_cell(SheetId(0), cell, &input, options)
            .unwrap();
        assert_eq!(
            workbook.model().sheets[0].cell(cell).unwrap().value,
            CellValue::Text {
                value: value.into()
            }
        );
    }
}

#[test]
fn text_format_cells_copy_as_text_into_general_cells() {
    let mut model = model(0, DateSystem::V1900);
    model.styles.cell_xfs.push(Xf {
        num_fmt_id: Some(49),
        ..Default::default()
    });
    model.sheets[0].set_cell(
        at(),
        Cell {
            value: CellValue::Text {
                value: "00123".into(),
            },
            style: Some(1),
            ..Default::default()
        },
    );
    let mut workbook = Workbook::from_model(model).unwrap();
    let input = workbook.cell(SheetId(0), at()).unwrap().input;
    for cell in [at(), CellRef::new(0, 1)] {
        workbook
            .edit_cell(SheetId(0), cell, &input, CalculationOptions::default())
            .unwrap();
        assert_eq!(
            workbook.model().sheets[0].cell(cell).unwrap().value,
            CellValue::Text {
                value: "00123".into()
            }
        );
    }
}

/// S1's rule, kept: an unstyled cell reads xf 0 (Excel's default cell format),
/// so in a workbook whose xf 0 is Text a typed `=` formula is stored as text
/// by both the edit and the typed-formula check, not refused as unparsed.
#[test]
fn an_unstyled_cell_in_a_text_default_workbook_keeps_a_typed_formula_as_text() {
    let mut model = model(49, DateSystem::V1900);
    model.sheets[0].set_cell(at(), Cell::default());
    let mut workbook = Workbook::from_model(model).unwrap();
    workbook
        .edit_typed_cell(SheetId(0), at(), "=1+1", CalculationOptions::default())
        .unwrap();
    assert_eq!(
        workbook.model().sheets[0].cell(at()).unwrap().value,
        CellValue::Text {
            value: "=1+1".into()
        }
    );
}
