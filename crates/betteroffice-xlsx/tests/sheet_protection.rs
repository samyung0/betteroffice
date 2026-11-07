//! Protected sheets (#1994 on Capy's engine): the sheet info says whether the
//! active sheet is protected and each cell read says whether its style locks
//! it; the editor refuses edits to locked cells of protected sheets.

use std::time::{Duration, Instant};

use betteroffice_xlsx::{
    CalculationOptions, Cell, CellRange, CellRef, CellValue, ColStyle, Op, Sheet, SheetId,
    Workbook, WorkbookModel,
};
use xlsx_model::Xf;
use xlsx_model::styles::Protection;

fn r(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

/// "Open" without protection, then "Locked" with `sheetProtection`, whose A1
/// keeps the default (locked) style and B1 a style with `locked="0"`.
fn package() -> Vec<u8> {
    let mut sheets = [Sheet::new("Open"), Sheet::new("Locked")];
    for (sheet, address, style) in [(0, "A1", None), (1, "A1", None), (1, "B1", Some(1))] {
        sheets[sheet].set_cell(
            r(address),
            Cell {
                value: CellValue::Number { value: 1.0 },
                formula: None,
                style,
            },
        );
    }
    let mut model = WorkbookModel {
        sheets: sheets.into(),
        ..WorkbookModel::default()
    };
    model.styles.cell_xfs = vec![
        Xf::default(),
        Xf {
            protection: Some(Protection {
                locked: false,
                hidden: false,
            }),
            ..Xf::default()
        },
    ];
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    let sheet = &mut parts
        .iter_mut()
        .find(|(name, _)| name == "xl/worksheets/sheet2.xml")
        .unwrap()
        .1;
    *sheet = String::from_utf8(sheet.clone())
        .unwrap()
        .replace(
            "</sheetData>",
            r#"</sheetData><sheetProtection sheet="1" objects="1" scenarios="1"/>"#,
        )
        .into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn open(bytes: &[u8], collaborative: bool) -> Workbook {
    if collaborative {
        Workbook::open_collaborative(bytes, 7).unwrap()
    } else {
        Workbook::open(bytes).unwrap()
    }
}

fn protected(workbook: &mut Workbook, sheet: u32) -> bool {
    workbook.set_active_sheet(SheetId(sheet)).unwrap();
    workbook.sheet_info().unwrap().protected
}

#[test]
fn the_sheet_info_follows_the_protected_sheet_through_sheet_removals_and_adds() {
    let bytes = package();
    for collaborative in [false, true] {
        let mut workbook = open(&bytes, collaborative);
        assert!(!protected(&mut workbook, 0), "{collaborative}");
        assert!(protected(&mut workbook, 1), "{collaborative}");

        workbook
            .apply_ops(
                vec![Op::RemoveSheet { index: 0 }],
                CalculationOptions::default(),
            )
            .unwrap();
        assert!(protected(&mut workbook, 0), "{collaborative}");
        workbook
            .apply_ops(
                vec![Op::AddSheet {
                    index: 1,
                    name: "Added".into(),
                }],
                CalculationOptions::default(),
            )
            .unwrap();
        assert!(protected(&mut workbook, 0), "{collaborative}");
        assert!(!protected(&mut workbook, 1), "{collaborative}");
    }

    // the standalone engine also adds a sheet in front of it
    let mut workbook = open(&bytes, false);
    workbook
        .apply_ops(
            vec![Op::AddSheet {
                index: 0,
                name: "First".into(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert!(!protected(&mut workbook, 0));
    assert!(!protected(&mut workbook, 1));
    assert!(protected(&mut workbook, 2));
}

#[test]
fn cells_are_locked_unless_their_style_unlocks_them() {
    let bytes = package();
    for collaborative in [false, true] {
        let workbook = open(&bytes, collaborative);
        let locked = |address| workbook.cell(SheetId(1), r(address)).unwrap().locked;
        assert!(locked("A1"), "{collaborative}: default style");
        assert!(!locked("B1"), "{collaborative}: locked=\"0\"");
        assert!(locked("C1"), "{collaborative}: empty cell");
        assert!(workbook.cell(SheetId(0), r("A1")).unwrap().locked);
    }
}

/// Rows of the unlocked column A [`column_package`] stores.
const COLUMN_ROWS: u32 = 20_000;

/// "Open" unprotected and "Locked" protected, both with column A unlocked
/// (style and stored cells) and row 5 unlocked by its row style; "Locked"
/// also stores B2 in the default (locked) style.
fn column_package() -> Vec<u8> {
    let mut model = WorkbookModel::default();
    for name in ["Open", "Locked"] {
        let mut sheet = Sheet::new(name);
        sheet.col_styles.push(ColStyle {
            first: 0,
            last: 0,
            xf: 1,
        });
        sheet.row_styles.insert(4, 1);
        for row in 0..COLUMN_ROWS {
            sheet.set_cell(
                CellRef::new(row, 0),
                Cell {
                    value: CellValue::Number { value: 1.0 },
                    formula: None,
                    style: Some(1),
                },
            );
        }
        sheet.set_cell(
            r("B2"),
            Cell {
                value: CellValue::Number { value: 2.0 },
                formula: None,
                style: None,
            },
        );
        model.sheets.push(sheet);
    }
    model.styles.cell_xfs = vec![
        Xf::default(),
        Xf {
            protection: Some(Protection {
                locked: false,
                hidden: false,
            }),
            ..Xf::default()
        },
    ];
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    let sheet = &mut parts
        .iter_mut()
        .find(|(name, _)| name == "xl/worksheets/sheet2.xml")
        .unwrap()
        .1;
    *sheet = String::from_utf8(sheet.clone())
        .unwrap()
        .replace(
            "</sheetData>",
            r#"</sheetData><sheetProtection sheet="1"/>"#,
        )
        .into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn ranges(ranges: &[&str]) -> Vec<CellRange> {
    ranges
        .iter()
        .map(|range| CellRange::parse_a1(range).unwrap())
        .collect()
}

#[test]
fn locked_in_ranges_reads_the_formats_of_a_protected_sheet() {
    let bytes = column_package();
    for collaborative in [false, true] {
        let workbook = open(&bytes, collaborative);
        let locked = |sheet, wanted: &[&str]| {
            workbook
                .locked_in_ranges(SheetId(sheet), &ranges(wanted))
                .unwrap()
        };
        // an unlocked whole column and an unlocked whole row
        assert!(!locked(1, &["A1:A1048576"]), "{collaborative}");
        assert!(!locked(1, &["A5:XFD5"]), "{collaborative}");
        // B2 is the range's one locked cell
        assert!(locked(1, &["A2:B2"]), "{collaborative}");
        assert!(locked(1, &["A1:A3", "B2"]), "{collaborative}");
        // a missing cell outside column A and row 5 takes xf 0, locked
        assert!(locked(1, &["C7"]), "{collaborative}");
        // an unprotected sheet locks nothing
        assert!(!locked(0, &["A1:XFD1048576"]), "{collaborative}");
    }
}

#[test]
fn locked_in_ranges_over_a_whole_column_does_not_walk_its_rows() {
    let workbook = Workbook::open(&column_package()).unwrap();
    let wanted = ranges(&["A1:A1048576", "A1:A1048576", "A5:XFD5"]);
    let started = Instant::now();
    assert!(!workbook.locked_in_ranges(SheetId(1), &wanted).unwrap());
    let elapsed = started.elapsed();
    // the stored cells and the defaults, never the 1,048,576 rows: well under
    // a second even in a debug build
    assert!(elapsed < Duration::from_secs(1), "{elapsed:?}");
}
