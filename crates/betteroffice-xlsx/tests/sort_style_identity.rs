//! upstream #1894 on Capy's engine: sorting and format copies give every moved
//! cell exactly its source cell's style, and interning never reuses an xf that
//! carries XML the model does not.

use betteroffice_xlsx::{
    CalculationOptions, CapturedFormat, CellRange, CellRef, Op, Sheet, SheetId, Workbook,
    WorkbookModel,
};
use serde_json::json;

fn r(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn options() -> CalculationOptions {
    CalculationOptions::default()
}

fn open(bytes: &[u8], collaborative: bool) -> Workbook {
    if collaborative {
        Workbook::open_collaborative(bytes, 7).unwrap()
    } else {
        Workbook::open(bytes).unwrap()
    }
}

fn package(cell_xfs: &str, cells: &str) -> Vec<u8> {
    let count = cell_xfs.matches("<xf").count();
    let styles = format!(
        r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><name val="Calibri"/><sz val="11"/></font></fonts>
<fills count="1"><fill><patternFill patternType="none"/></fill></fills>
<borders count="1"><border/></borders>
<cellStyleXfs count="2"><xf/><xf numFmtId="2"/></cellStyleXfs>
<cellXfs count="{count}">{cell_xfs}</cellXfs>
<cellStyles count="2"><cellStyle name="Normal" xfId="0" builtinId="0"/>
<cellStyle name="Currency" xfId="1" builtinId="4"/></cellStyles></styleSheet>"#
    );
    let sheet = format!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1">{cells}</row></sheetData>
</worksheet>"#
    );
    let mut model = WorkbookModel {
        sheets: vec![Sheet::new("Sheet1")],
        ..WorkbookModel::default()
    };
    // a stylesheet, so the package has the styles part replaced below
    model.styles.cell_xfs.push(xlsx_model::Xf::default());
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    for (name, bytes) in &mut parts {
        match name.as_str() {
            "xl/styles.xml" => *bytes = styles.as_bytes().to_vec(),
            "xl/worksheets/sheet1.xml" => *bytes = sheet.as_bytes().to_vec(),
            _ => {}
        }
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn part(bytes: &[u8], name: &str) -> String {
    let parts = ooxml_opc::unzip_parts(bytes).unwrap();
    String::from_utf8(parts.into_iter().find(|(path, _)| path == name).unwrap().1).unwrap()
}

fn style(workbook: &Workbook, address: &str) -> Option<u32> {
    workbook.model().sheets[0].cell(r(address)).unwrap().style
}

#[test]
fn captured_sort_keeps_default_and_extended_source_xfs() {
    let default_xf = r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#;
    let extended_xf = r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><extLst><ext uri="{6E1A2B3C-4D5E-4F60-8A9B-0C1D2E3F4A5B}"><x:keep xmlns:x="urn:protected-sort" value="kept"/></ext></extLst></xf>"#;
    let bytes = package(
        &format!("{default_xf}<xf numFmtId=\"2\"/><xf numFmtId=\"14\"/>{extended_xf}"),
        r#"<c r="A1" s="3"><v>7</v></c><c r="B1"><v>8</v></c><c r="C1" s="0"><v>9</v></c>"#,
    );
    for collaborative in [false, true] {
        for plain_source in ["B1", "C1"] {
            for move_values in [false, true] {
                let mut workbook = open(&bytes, collaborative);
                let capture = |address| {
                    serde_json::to_value(
                        workbook
                            .capture_format(SheetId(0), CellRange::new(r(address), r(address)))
                            .unwrap(),
                    )
                    .unwrap()
                };
                let plain = capture(plain_source);
                let extended = capture("A1");
                let mut payload = Vec::new();
                if move_values {
                    payload.extend([
                        json!({"type": "setCell", "sheet": 0, "at": {"row": 0, "col": 0},
                            "cell": {"value": {"kind": "number", "value": if plain_source == "B1" { 8 } else { 9 }}, "formula": null}}),
                        json!({"type": "setCell", "sheet": 0, "at": {"row": 0, "col": 1},
                            "cell": {"value": {"kind": "number", "value": 7}, "formula": null}}),
                    ]);
                }
                payload.push(json!({"type": "applyRangeFormat", "sheet": 0,
                    "range": {"start": {"row": 0, "col": 0}, "end": {"row": 0, "col": 1}},
                    "format": {"rows": 1, "columns": 2,
                        "formats": [plain["formats"][0], extended["formats"][0]],
                        "sourceStyles": [plain["sourceStyles"][0], extended["sourceStyles"][0]],
                        "sourceFormats": [plain["sourceFormats"][0], extended["sourceFormats"][0]]}}));
                let ops: Vec<Op> = serde_json::from_value(json!(payload)).unwrap();
                assert!(workbook.apply_ops(ops, options()).unwrap().applied);
                let expected_plain = (plain_source == "C1").then_some(0);
                let case = format!("{collaborative}: {plain_source}: {move_values}");
                // a replica keys a written format by its facets and saves a cell
                // whose facets did not change with its source xf, so there each
                // cell keeps its source's format rather than its index
                let kept = |workbook: &Workbook| {
                    let styles = &workbook.model().styles;
                    for (address, expected) in [("A1", expected_plain), ("B1", Some(3))] {
                        if collaborative {
                            assert_eq!(
                                styles.cell_format(style(workbook, address)),
                                styles.cell_format(expected),
                                "{case}: {address}"
                            );
                        } else {
                            assert_eq!(style(workbook, address), expected, "{case}: {address}");
                        }
                    }
                };
                kept(&workbook);
                let saved = workbook.save().unwrap();
                assert_eq!(part(&saved, "xl/styles.xml"), part(&bytes, "xl/styles.xml"));
                kept(&Workbook::open(&saved).unwrap());
            }
        }
    }
}

#[test]
fn painting_an_unstyled_cell_keeps_the_destination_cell_xml() {
    let bytes = package(
        r#"<xf/><xf numFmtId="4"/>"#,
        r#"<c r="A1"><v>7</v></c><c r="B1"><v>8</v></c>"#,
    );
    for collaborative in [false, true] {
        let mut workbook = open(&bytes, collaborative);
        let captured = workbook
            .capture_format(SheetId(0), CellRange::new(r("A1"), r("A1")))
            .unwrap();
        assert_eq!(captured.source_styles, vec![None]);
        assert!(
            !workbook
                .apply_format(
                    SheetId(0),
                    CellRange::new(r("B1"), r("B1")),
                    captured,
                    options(),
                )
                .unwrap()
                .applied
        );
        assert_eq!(style(&workbook, "B1"), None);
        assert_eq!(
            part(&workbook.save().unwrap(), "xl/worksheets/sheet1.xml"),
            part(&bytes, "xl/worksheets/sheet1.xml")
        );
    }
}

#[test]
fn painting_a_missing_cell_in_a_styled_row_keeps_the_row_xf() {
    let row_xf = r#"<xf numFmtId="2" fontId="0"><extLst><ext uri="row-source"/></extLst></xf>"#;
    let source = package(&format!("<xf/>{row_xf}"), r#"<c r="B1"><v>8</v></c>"#);
    let mut parts = ooxml_opc::unzip_parts(&source).unwrap();
    for (name, bytes) in &mut parts {
        if name == "xl/worksheets/sheet1.xml" {
            *bytes = String::from_utf8(bytes.clone())
                .unwrap()
                .replace(
                    "</sheetData>",
                    r#"<row r="2" s="1" customFormat="1"/></sheetData>"#,
                )
                .into_bytes();
        }
    }
    let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
    for collaborative in [false, true] {
        let mut workbook = open(&bytes, collaborative);
        let captured = workbook
            .capture_format(SheetId(0), CellRange::new(r("A2"), r("A2")))
            .unwrap();
        assert_eq!(captured.source_styles, vec![Some(1)]);
        workbook
            .apply_format(
                SheetId(0),
                CellRange::new(r("B1"), r("B1")),
                captured,
                options(),
            )
            .unwrap();
        assert_eq!(style(&workbook, "B1"), Some(1));
        let saved = workbook.save().unwrap();
        assert_eq!(part(&saved, "xl/styles.xml"), part(&bytes, "xl/styles.xml"));
        assert_eq!(style(&Workbook::open(&saved).unwrap(), "B1"), Some(1));
    }
}

#[test]
fn facet_interning_avoids_xfs_with_extra_xml() {
    for (attributes, children) in [
        ("", "<extLst><ext uri=\"extra-style\"/></extLst>"),
        ("", "<unknown value=\"extra-style\"/>"),
        ("", "<!--extra-style-->"),
        (" custom=\"extra-style\"", ""),
        ("", "<alignment custom=\"extra-style\"/>"),
        ("", "<protection custom=\"extra-style\"/>"),
    ] {
        for size in [2, 128] {
            for has_clean in [false, true] {
                let extra = format!(r#"<xf numFmtId="2"{attributes}>{children}</xf>"#);
                let fillers = (2..size)
                    .map(|index| format!(r#"<xf numFmtId="{}"/>"#, index + 10))
                    .collect::<String>();
                let clean = if !has_clean {
                    ""
                } else if children.starts_with("<protection") {
                    "<xf numFmtId=\"2\"><protection locked=\"1\" hidden=\"0\"/></xf>"
                } else {
                    "<xf numFmtId=\"2\"/>"
                };
                let bytes = package(
                    &format!("<xf/>{extra}{fillers}{clean}"),
                    r#"<c r="A1" s="1"><v>7</v></c><c r="B1"><v>8</v></c>"#,
                );
                let mut workbook = Workbook::open(&bytes).unwrap();
                let mut styles = workbook.model().styles.clone();
                let format = styles.cell_format(Some(1));
                let index = styles.intern_cell_format(&format).unwrap().unwrap();
                let case = format!("{size}: {has_clean}: {attributes}: {children}");
                assert_ne!(index, 1, "{case}");
                assert_eq!(index, size, "{case}");
                assert_eq!(styles.intern_cell_format(&format).unwrap(), Some(index));
                workbook
                    .apply_ops(
                        vec![Op::ApplyRangeFormat {
                            sheet: SheetId(0),
                            range: CellRange::new(r("B1"), r("B1")),
                            format: CapturedFormat {
                                rows: 1,
                                columns: 1,
                                formats: vec![format],
                                kept: Vec::new(),
                                source_styles: Vec::new(),
                                source_formats: Vec::new(),
                            },
                        }],
                        options(),
                    )
                    .unwrap();
                let saved = workbook.save().unwrap();
                assert_eq!(
                    part(&saved, "xl/styles.xml").matches("extra-style").count(),
                    1,
                    "{case}"
                );
                assert_eq!(
                    style(&Workbook::open(&saved).unwrap(), "B1"),
                    Some(index),
                    "{case}"
                );
            }
        }
    }
}
