//! A value change in a wrapped cell refits its row while editing, as Excel and
//! Google Sheets do; a row with a custom height keeps it (Q759 A).

use betteroffice_xlsx::{
    CalculationOptions, CellInput, CellRef, GridGeometry, SheetId, Workbook, WorkbookModel,
};
use xlsx_model::CellValue;
use xlsx_model::styles::{Alignment, Font, Xf};
use xlsx_model::workbook::{Cell, Sheet};

const NOTE: &str = "Note";
const PARAGRAPH: &str = "A long paragraph that wraps over many lines in a narrow column, \
    so the row it sits in has to grow to show all of it";

fn source() -> Vec<u8> {
    let mut model = WorkbookModel::default();
    model.styles.fonts = vec![Font::default()];
    model.styles.cell_xfs = vec![
        Xf::default(),
        Xf {
            alignment: Some(Alignment {
                wrap_text: true,
                ..Default::default()
            }),
            ..Default::default()
        },
    ];
    let mut sheet = Sheet::new("Notes");
    sheet.col_widths.insert(0, 12.0);
    for row in [0, 2] {
        sheet.set_cell(
            CellRef::new(row, 0),
            Cell {
                value: CellValue::Text { value: NOTE.into() },
                style: Some(1),
                ..Default::default()
            },
        );
    }
    sheet.row_heights.insert(2, 30.0);
    model.sheets.push(sheet);
    ooxml_opc::rezip_parts(&xlsx_parse::serialize_workbook(&model).unwrap()).unwrap()
}

/// The row's height in the geometry the workbook keeps for its active sheet.
fn kept_height(workbook: &Workbook, row: u32) -> f32 {
    workbook
        .cell_scroll_bounds(SheetId(0), CellRef::new(row, 0))
        .unwrap()
        .3
}

fn fresh(workbook: &Workbook) -> GridGeometry {
    GridGeometry::new(
        workbook.sheet(SheetId(0)).unwrap(),
        &workbook.model().styles,
    )
}

fn fresh_height(workbook: &Workbook, row: u32) -> f32 {
    let geometry = fresh(workbook);
    geometry.row_y(row + 1) - geometry.row_y(row)
}

fn edit(workbook: &mut Workbook, row: u32, text: &str) {
    workbook
        .edit_cell(
            SheetId(0),
            CellRef::new(row, 0),
            text,
            CalculationOptions::default(),
        )
        .unwrap();
}

/// Equal up to the float error of offsets kept by increments.
fn assert_near(left: f32, right: f32) {
    assert!((left - right).abs() < 1e-3, "{left} != {right}");
}

/// The kept row heights and content height equal a fresh geometry's.
fn assert_fitted(workbook: &Workbook) {
    for row in 0..3 {
        assert_near(kept_height(workbook, row), fresh_height(workbook, row));
    }
    let info = workbook.sheet_info().unwrap();
    assert_near(
        info.content_height,
        fresh(workbook).row_y(info.used_range.unwrap().end.row + 2),
    );
}

#[test]
fn a_wrapped_cells_value_change_refits_its_row() {
    let mut workbook = Workbook::open(&source()).unwrap();
    workbook.sheet_info().unwrap();
    let note = kept_height(&workbook, 0);
    let custom = kept_height(&workbook, 2);

    edit(&mut workbook, 0, PARAGRAPH);
    let paragraph = kept_height(&workbook, 0);
    assert!(paragraph > note * 2.0);
    assert_fitted(&workbook);

    edit(&mut workbook, 0, NOTE);
    assert_near(kept_height(&workbook, 0), note);
    assert_fitted(&workbook);

    workbook.undo(CalculationOptions::default()).unwrap();
    assert_near(kept_height(&workbook, 0), paragraph);
    workbook.redo(CalculationOptions::default()).unwrap();
    assert_near(kept_height(&workbook, 0), note);

    // a paste over both rows; the custom-height row keeps its height
    workbook
        .edit_cells(
            SheetId(0),
            &[0, 2].map(|row| CellInput {
                cell: CellRef::new(row, 0),
                input: PARAGRAPH.into(),
            }),
            CalculationOptions::default(),
        )
        .unwrap();
    assert_near(kept_height(&workbook, 0), paragraph);
    assert_near(kept_height(&workbook, 2), custom);
    assert_fitted(&workbook);
}

#[test]
fn a_remote_value_change_in_a_wrapped_cell_refits_its_row() {
    let bytes = source();
    let mut local = Workbook::open_collaborative(&bytes, 7).unwrap();
    let mut peer = Workbook::open_collaborative(&bytes, 8).unwrap();
    local.sheet_info().unwrap();
    let note = kept_height(&local, 0);
    edit(&mut peer, 0, PARAGRAPH);
    local
        .apply_update_v1(
            &peer.encode_state_as_update_v1(),
            CalculationOptions::default(),
        )
        .unwrap();
    assert!(kept_height(&local, 0) > note * 2.0);
    assert_fitted(&local);
}
