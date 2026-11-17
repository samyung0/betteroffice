//! The viewer opens a workbook without the editor's save capture (Q691 C1b,
//! Q749 A): it reads only what it draws, so a part only a save needs cannot
//! stop it, while the edit engine still refuses that part.

use xlsx_view_wasm::XlsxViewDocument;

const FIXTURE: &[u8] = include_bytes!("../../../packages/xlsx/test-fixtures/charts.xlsx");
const CONTENT_TYPES: &str = "[Content_Types].xml";

fn rewrite(parts: &mut [(String, Vec<u8>)], path: &str, edit: impl FnOnce(&str) -> String) {
    let slot = parts
        .iter_mut()
        .find(|(name, _)| name == path)
        .unwrap_or_else(|| panic!("{path} is not in the fixture"));
    slot.1 = edit(std::str::from_utf8(&slot.1).unwrap()).into_bytes();
}

/// `charts.xlsx` (one sheet, four charts) with a second sheet, `Locked`, whose
/// source names `sheetProtection`.
fn with_protected_sheet() -> Vec<(String, Vec<u8>)> {
    let mut parts = ooxml_opc::unzip_parts(FIXTURE).unwrap();
    parts.push((
        "xl/worksheets/sheet2.xml".into(),
        br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Locked</t></is></c></row></sheetData><sheetProtection sheet="1" objects="1"/></worksheet>"#.to_vec(),
    ));
    rewrite(&mut parts, "xl/workbook.xml", |text| {
        text.replace(
            "</sheets>",
            r#"<sheet name="Locked" sheetId="2" r:id="rId2"/></sheets>"#,
        )
    });
    rewrite(&mut parts, "xl/_rels/workbook.xml.rels", |text| {
        text.replace(
            "</Relationships>",
            r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>"#,
        )
    });
    rewrite(&mut parts, CONTENT_TYPES, |text| {
        text.replace(
            "</Types>",
            r#"<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#,
        )
    });
    parts
}

fn open_view(parts: &[(String, Vec<u8>)]) -> XlsxViewDocument {
    let bytes = ooxml_opc::rezip_parts(parts).unwrap();
    XlsxViewDocument::open(&bytes).unwrap_or_else(|_| panic!("the view refused the package"))
}

fn active_sheet_protected(document: &XlsxViewDocument) -> bool {
    let info = document
        .sheet_info_json()
        .unwrap_or_else(|_| panic!("no sheet info"));
    serde_json::from_str::<serde_json::Value>(&info).unwrap()["protected"]
        .as_bool()
        .unwrap()
}

fn cell_text(document: &XlsxViewDocument, sheet: u32, row: u32, col: u32) -> String {
    document
        .cell_text(sheet, row, col)
        .unwrap_or_else(|_| panic!("no cell {sheet}:{row}:{col}"))
}

#[cfg(feature = "test-counters")]
#[test]
fn a_view_open_captures_only_the_sheet_that_names_sheet_protection() {
    let parts = with_protected_sheet();
    let before = xlsx_parse::template_captures();
    let mut document = open_view(&parts);
    assert_eq!(xlsx_parse::template_captures() - before, 1);
    assert!(!active_sheet_protected(&document));
    document
        .set_active_sheet(1)
        .unwrap_or_else(|_| panic!("no second sheet"));
    assert!(active_sheet_protected(&document));
}

#[cfg(feature = "test-counters")]
#[test]
fn a_view_builds_each_sheets_grid_geometry_once() {
    use xlsx_render::geometry::{geometry_counters, reset_geometry_counters};
    let mut document = open_view(&with_protected_sheet());
    reset_geometry_counters();
    let viewport = r#"{"x":0,"y":0,"width":1600,"height":1000}"#;
    let scrolled = r#"{"x":0,"y":400,"width":1600,"height":1000}"#;
    for sheet in 0..2 {
        document
            .set_active_sheet(sheet)
            .unwrap_or_else(|_| panic!("no sheet {sheet}"));
        for viewport in [viewport, scrolled] {
            document
                .sheet_info_json()
                .unwrap_or_else(|_| panic!("no sheet info"));
            document
                .display_list_json(viewport)
                .unwrap_or_else(|_| panic!("no display list"));
            document
                .chart_at_point_json(&format!(r#"{{"viewport":{viewport},"x":200,"y":100}}"#))
                .unwrap_or_else(|_| panic!("no chart hit"));
            document
                .cell_position_json(r#"{"sheet":0,"row":5,"col":3}"#)
                .unwrap_or_else(|_| panic!("no cell position"));
        }
        assert_eq!(geometry_counters().0, u64::from(sheet) + 1);
    }
}

#[test]
fn a_package_broken_only_in_a_part_view_never_reads_opens_in_view_not_in_edit() {
    let whole = with_protected_sheet();
    let edit = |parts: &[(String, Vec<u8>)]| {
        betteroffice_xlsx::Workbook::open(&ooxml_opc::rezip_parts(parts).unwrap())
    };
    assert!(edit(&whole).is_ok());
    let mut broken = whole;
    rewrite(&mut broken, CONTENT_TYPES, |text| {
        text.replace("</Types>", "<Override PartName=")
    });

    let mut document = open_view(&broken);
    assert_eq!(cell_text(&document, 0, 0, 0), "Quarter");
    assert_eq!(cell_text(&document, 0, 1, 1), "12");
    assert!(!active_sheet_protected(&document));
    document
        .set_active_sheet(1)
        .unwrap_or_else(|_| panic!("no second sheet"));
    assert_eq!(cell_text(&document, 1, 0, 0), "Locked");
    assert!(active_sheet_protected(&document));

    assert!(edit(&broken).is_err());
}

#[test]
fn a_view_serves_each_chart_part_as_the_full_package_does() {
    let parts = with_protected_sheet();
    let viewed = xlsx_parse::parse_workbook_for_view(parts.clone()).unwrap();
    let preserved = xlsx_parse::parse_workbook_with_owned_package(parts).unwrap();
    let charts = &viewed.workbook.sheets[0].charts;
    assert_eq!(charts.len(), 4);
    for chart in charts {
        let bytes = viewed.package.part_bytes(&chart.part);
        assert!(bytes.is_some_and(|bytes| !bytes.is_empty()));
        assert_eq!(bytes, preserved.package.part_bytes(&chart.part));
    }
}

/// The geometry a view keeps for its open draws and hit-tests exactly what a
/// freshly built one does (Q691 R2), on a sheet with charts, merges, a frozen
/// pane and hidden tracks.
#[test]
fn a_views_kept_geometry_draws_and_hits_as_a_fresh_one() {
    use xlsx_model::{SheetChart, SheetId};
    use xlsx_render::{RenderError, build_display_list_with_charts, chart_at_point, chart_regions};

    let mut parts = ooxml_opc::unzip_parts(FIXTURE).unwrap();
    rewrite(&mut parts, "xl/worksheets/sheet1.xml", |text| {
        text.replace(
            "</sheetData>",
            r#"</sheetData><mergeCells count="2"><mergeCell ref="D2:E4"/><mergeCell ref="A6:C6"/></mergeCells>"#,
        )
    });
    let document = open_view(&parts);
    let viewed = xlsx_parse::parse_workbook_for_view(parts).unwrap();
    let workbook = &viewed.workbook;
    let sheet = &workbook.sheets[0];
    assert!(sheet.freeze_pane.is_some());
    assert_eq!(sheet.merges.len(), 2);
    assert_eq!(sheet.charts.len(), 4);
    let resolve = |chart: &SheetChart| {
        let bytes = viewed.package.part_bytes(&chart.part).ok_or_else(|| {
            RenderError::ChartPartMissing {
                part: chart.part.clone(),
            }
        })?;
        xlsx_parse::preserved_chart_space(bytes, workbook, &sheet.name, &workbook.styles.theme)
            .ok_or_else(|| RenderError::ChartParseFailed {
                part: chart.part.clone(),
            })
    };
    let mut hits = 0;
    for (x, y) in [(0.0, 0.0), (0.0, 140.0), (90.0, 420.0)] {
        let viewport = xlsx_render::Viewport {
            x,
            y,
            width: 1200.0,
            height: 800.0,
        };
        let json = format!(r#"{{"x":{x},"y":{y},"width":1200,"height":800}}"#);
        let fresh =
            build_display_list_with_charts(workbook, SheetId(0), &viewport, resolve).unwrap();
        assert!(!fresh.charts.is_empty());
        assert_eq!(
            document
                .display_list_json(&json)
                .unwrap_or_else(|_| panic!("no display list")),
            serde_json::to_string(&fresh).unwrap()
        );
        let regions = chart_regions(sheet, &workbook.styles, &viewport).unwrap();
        for px in (0..12).map(|step| step as f32 * 100.0 + 5.0) {
            for py in (0..8).map(|step| step as f32 * 100.0 + 5.0) {
                let fresh = serde_json::to_string(&chart_at_point(&regions, px, py)).unwrap();
                let kept = document
                    .chart_at_point_json(&format!(r#"{{"viewport":{json},"x":{px},"y":{py}}}"#))
                    .unwrap_or_else(|_| panic!("no chart hit"));
                assert_eq!(kept, fresh, "hit at {px},{py} in {json}");
                hits += usize::from(fresh != "null");
            }
        }
    }
    assert!(hits > 0);
}
