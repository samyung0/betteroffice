use betteroffice_xlsx::Workbook;

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STYLED_ROWS: u32 = 100_000;

fn package(styled_rows: u32) -> Vec<u8> {
    let mut rows =
        String::from(r#"<row r="1"><c r="A1" t="inlineStr"><is><t>Year</t></is></c></row>"#);
    for row in 2..=styled_rows {
        rows.push_str(&format!(r#"<row r="{row}" s="1" customFormat="1"/>"#));
    }
    let parts = [
        (
            "[Content_Types].xml",
            concat!(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
                r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#,
                r#"<Default Extension="xml" ContentType="application/xml"/>"#,
                r#"<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>"#,
                r#"<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
                r#"<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"#,
                r#"</Types>"#,
            )
            .to_owned(),
        ),
        (
            "_rels/.rels",
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
            ),
        ),
        (
            "xl/workbook.xml",
            format!(
                r#"<workbook xmlns="{MAIN}" xmlns:r="{REL}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{REL}/styles" Target="styles.xml"/></Relationships>"#
            ),
        ),
        (
            "xl/styles.xml",
            format!(
                concat!(
                    r#"<styleSheet xmlns="{main}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>"#,
                    r#"<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills>"#,
                    r#"<borders count="1"><border/></borders>"#,
                    r#"<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#,
                    r#"<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
                    r#"<xf numFmtId="0" fontId="0" fillId="1" borderId="0" xfId="0" applyFill="1"/></cellXfs>"#,
                    r#"</styleSheet>"#,
                ),
                main = MAIN,
            ),
        ),
        (
            "xl/worksheets/sheet1.xml",
            format!(
                r#"<worksheet xmlns="{MAIN}"><sheetData>{rows}</sheetData></worksheet>"#
            ),
        ),
    ]
    .map(|(path, xml)| (path.to_owned(), xml.into_bytes()));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn assert_styled_rows(workbook: &Workbook) {
    let row_styles = &workbook.model().sheets[0].row_styles;
    assert_eq!(row_styles.len(), STYLED_ROWS as usize - 1);
    assert_eq!(row_styles.get(&1), Some(&1));
    assert_eq!(row_styles.get(&(STYLED_ROWS - 1)), Some(&1));
    assert_eq!(row_styles.get(&0), None);
}

#[test]
fn a_sheet_styled_past_sixty_five_thousand_rows_opens() {
    let bytes = package(STYLED_ROWS);
    assert_styled_rows(&Workbook::open(&bytes).unwrap());
    assert_styled_rows(&Workbook::open_for_read(&bytes).unwrap());
    assert_styled_rows(&Workbook::open_collaborative(&bytes, 7).unwrap());
}

#[test]
fn a_sheet_styled_past_sixty_five_thousand_rows_saves_and_reopens() {
    let workbook = Workbook::open(&package(STYLED_ROWS)).unwrap();
    let saved = workbook.save().unwrap();
    assert_styled_rows(&Workbook::open(&saved).unwrap());
}

fn edit_heading(workbook: &mut Workbook, input: &str) {
    use betteroffice_xlsx::{CalculationOptions, CellInput, CellRef, CellValue, SheetId};

    assert!(
        workbook
            .edit_cells(
                SheetId(0),
                &[CellInput {
                    cell: CellRef::new(0, 0),
                    input: input.into(),
                }],
                CalculationOptions::default(),
            )
            .unwrap()
            .applied
    );
    assert_eq!(
        workbook
            .sheet(SheetId(0))
            .unwrap()
            .cell(CellRef::new(0, 0))
            .unwrap()
            .value,
        CellValue::Text {
            value: input.into(),
        }
    );
}

#[test]
fn seventy_thousand_rows_sync() {
    use betteroffice_xlsx::CalculationOptions;

    let bytes = package(70_000);
    let mut peer_a = Workbook::open_collaborative(&bytes, 104).unwrap();
    let mut peer_b = Workbook::open_collaborative(&bytes, 105).unwrap();
    edit_heading(&mut peer_a, "Synced");
    peer_b
        .apply_update_v1(
            &peer_a.encode_state_as_update_v1(),
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(peer_b.model(), peer_a.model());
    assert_eq!(peer_b.model().sheets[0].row_styles.len(), 69_999);
}
