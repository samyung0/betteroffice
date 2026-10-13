//! Parts the engine does not model (notes and threaded comments, pictures,
//! sheet and workbook protection) survive a cell edit and the collaborative
//! save, spelled as Excel writes them.

use std::collections::BTreeMap;

use betteroffice_xlsx::{
    CalculationOptions, Cell, CellRef, CellValue, Op, Sheet, SheetId, Workbook, WorkbookModel,
};

const PNG: &[u8] = &[0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A, 1, 2, 3, 4];

const WORKBOOK: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr/><workbookProtection workbookAlgorithmName="SHA-512" workbookHashValue="bm90IGEgcmVhbCBoYXNo" workbookSaltValue="c2FsdA==" workbookSpinCount="100000" lockStructure="1"/><bookViews><workbookView activeTab="0"/></bookViews><sheets><sheet name="Grades" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029"/></workbook>"#;

const SHEET_PROTECTION: &str = r#"<sheetProtection algorithmName="SHA-512" hashValue="bm90IGEgcmVhbCBoYXNo" saltValue="c2FsdA==" spinCount="100000" sheet="1" objects="1" scenarios="1" formatColumns="0"/>"#;
const PROTECTED_RANGES: &str =
    r#"<protectedRanges><protectedRange name="Scores" sqref="B2:B3"/></protectedRanges>"#;

fn worksheet() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:B3"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Score</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Ada</t></is></c><c r="B2"><v>90</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>Lin</t></is></c><c r="B3"><v>85</v></c></row></sheetData>{SHEET_PROTECTION}{PROTECTED_RANGES}<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><drawing r:id="rId1"/><legacyDrawing r:id="rId2"/></worksheet>"#
    )
}

const SHEET_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/><Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/></Relationships>"#;

/// The legacy comments Excel writes beside threaded ones: a placeholder for
/// the thread on B2 and a plain note on A3.
const COMMENTS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr"><authors><author>tc={6A1C2F0B-3D4E-4F50-8A61-7B8C9D0E1F20}</author><author>Ms Tan</author></authors><commentList><comment ref="B2" authorId="0" shapeId="0" xr:uid="{6A1C2F0B-3D4E-4F50-8A61-7B8C9D0E1F20}"><text><t>[Threaded comment]

Your version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924

Comment:
    Check this score
Reply:
    Checked, it is right</t></text></comment><comment ref="A3" authorId="1" shapeId="0"><text><r><rPr><b/><sz val="9"/><rFont val="Tahoma"/></rPr><t>Ms Tan:</t></r><r><rPr><sz val="9"/><rFont val="Tahoma"/></rPr><t>
Joined in week 3</t></r></text></comment></commentList></comments>"#;

const THREADED: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><threadedComment ref="B2" dT="2026-10-01T09:00:00.00" personId="{0B8E4D2A-1C3F-4A5B-9C6D-7E8F9A0B1C2D}" id="{6A1C2F0B-3D4E-4F50-8A61-7B8C9D0E1F20}"><text>Check this score</text></threadedComment><threadedComment ref="B2" dT="2026-10-01T09:05:00.00" personId="{0B8E4D2A-1C3F-4A5B-9C6D-7E8F9A0B1C2D}" id="{7B2D3E1C-4E5F-4061-9B72-8C9D0E1F2031}" parentId="{6A1C2F0B-3D4E-4F50-8A61-7B8C9D0E1F20}"><text>Checked, it is right</text></threadedComment></ThreadedComments>"#;

const PERSONS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><person displayName="Ms Tan" id="{0B8E4D2A-1C3F-4A5B-9C6D-7E8F9A0B1C2D}" userId="tan@example.invalid" providerId="None"/></personList>"#;

/// One note shape per comment: the thread's placeholder on B2 and the note on
/// A3, each with its popup box anchor (column, offset, row, offset, twice).
const VML: &str = r##"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>{NOTE_B2}{NOTE_A3}</xml>"##;

const DRAWING: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>5</xdr:col><xdr:colOff>304800</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>95250</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 1" descr="School logo"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="1828800" y="190500"/><a:ext cx="1524000" cy="1047750"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>{SIZED}</xdr:wsDr>"#;

/// A picture that moves and sizes with its cells (`editAs` left at twoCell).
const SIZED: &str = r#"<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="3" name="Picture 2" descr="Class photo"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>"#;

fn drawing() -> String {
    DRAWING.replace("{SIZED}", SIZED)
}

const DRAWING_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>"#;

const PERSON_REL: &str = r#"<Relationship Id="rId9" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="persons/person.xml"/>"#;

const NOTE_B2: &str = r##"<v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;margin-left:155pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>2, 15, 0, 2, 4, 15, 4, 4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>1</x:Row><x:Column>1</x:Column></x:ClientData></v:shape>"##;
const NOTE_A3: &str = r##"<v:shape id="_x0000_s1026" type="#_x0000_t202" style="position:absolute;margin-left:107pt;margin-top:16.5pt;width:108pt;height:59.25pt;z-index:2;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>1, 15, 1, 2, 3, 15, 5, 4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>2</x:Row><x:Column>0</x:Column></x:ClientData></v:shape>"##;

fn vml() -> String {
    VML.replace("{NOTE_B2}", NOTE_B2)
        .replace("{NOTE_A3}", NOTE_A3)
}

fn at(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn set(parts: &mut Vec<(String, Vec<u8>)>, path: &str, body: &[u8]) {
    match parts.iter_mut().find(|(name, _)| name == path) {
        Some(part) => part.1 = body.to_vec(),
        None => parts.push((path.to_owned(), body.to_vec())),
    }
}

fn edit_text(parts: &mut [(String, Vec<u8>)], path: &str, from: &str, to: &str) {
    let part = parts.iter_mut().find(|(name, _)| name == path).unwrap();
    let text = String::from_utf8(part.1.clone()).unwrap();
    assert!(text.contains(from), "{path} has no {from}");
    part.1 = text.replacen(from, to, 1).into_bytes();
}

fn source() -> Vec<u8> {
    let mut sheet = Sheet::new("Grades");
    sheet.set_cell(
        at("B2"),
        Cell {
            value: CellValue::Number { value: 90.0 },
            ..Cell::default()
        },
    );
    let model = WorkbookModel {
        sheets: vec![sheet],
        ..WorkbookModel::default()
    };
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    set(&mut parts, "xl/workbook.xml", WORKBOOK.as_bytes());
    set(
        &mut parts,
        "xl/worksheets/sheet1.xml",
        worksheet().as_bytes(),
    );
    set(
        &mut parts,
        "xl/worksheets/_rels/sheet1.xml.rels",
        SHEET_RELS.as_bytes(),
    );
    set(&mut parts, "xl/comments1.xml", COMMENTS.as_bytes());
    set(
        &mut parts,
        "xl/threadedComments/threadedComment1.xml",
        THREADED.as_bytes(),
    );
    set(&mut parts, "xl/persons/person.xml", PERSONS.as_bytes());
    set(&mut parts, "xl/drawings/vmlDrawing1.vml", vml().as_bytes());
    set(&mut parts, "xl/drawings/drawing1.xml", drawing().as_bytes());
    set(
        &mut parts,
        "xl/drawings/_rels/drawing1.xml.rels",
        DRAWING_RELS.as_bytes(),
    );
    set(&mut parts, "xl/media/image1.png", PNG);
    edit_text(
        &mut parts,
        "xl/_rels/workbook.xml.rels",
        "</Relationships>",
        &format!("{PERSON_REL}</Relationships>"),
    );
    edit_text(
        &mut parts,
        "[Content_Types].xml",
        "<Override",
        r#"<Default Extension="png" ContentType="image/png"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/><Override PartName="/xl/threadedComments/threadedComment1.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/><Override PartName="/xl/persons/person.xml" ContentType="application/vnd.ms-excel.person+xml"/><Override"#,
    );
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn parts(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    ooxml_opc::unzip_parts(bytes).unwrap().into_iter().collect()
}

fn text(parts: &BTreeMap<String, Vec<u8>>, path: &str) -> String {
    String::from_utf8(parts[path].clone()).unwrap()
}

/// Edits a cell on the sheet that holds everything, saves as a collaborative
/// replica does, and returns the source and saved packages.
fn edit_and_save() -> (BTreeMap<String, Vec<u8>>, BTreeMap<String, Vec<u8>>) {
    let source = source();
    let mut workbook = Workbook::open_collaborative(&source, 5101).unwrap();
    workbook
        .edit_cell(SheetId(0), at("B3"), "88", CalculationOptions::default())
        .unwrap();
    let saved = workbook.save().unwrap();
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(
        reopened.model().sheets[0].cell(at("B3")).unwrap().value,
        CellValue::Number { value: 88.0 }
    );
    (parts(&source), parts(&saved))
}

fn assert_parts_kept(
    source: &BTreeMap<String, Vec<u8>>,
    saved: &BTreeMap<String, Vec<u8>>,
    paths: &[&str],
) {
    for path in paths {
        assert_eq!(saved.get(*path), Some(&source[*path]), "{path} changed");
    }
}

#[test]
fn notes_and_threaded_comments_survive_an_edit_and_a_save() {
    let (source, saved) = edit_and_save();
    assert_parts_kept(
        &source,
        &saved,
        &[
            "xl/comments1.xml",
            "xl/threadedComments/threadedComment1.xml",
            "xl/persons/person.xml",
            "xl/drawings/vmlDrawing1.vml",
            "xl/worksheets/_rels/sheet1.xml.rels",
        ],
    );
    assert!(text(&saved, "xl/worksheets/sheet1.xml").contains(r#"<legacyDrawing r:id="rId2"/>"#));
    assert!(text(&saved, "xl/_rels/workbook.xml.rels").contains(PERSON_REL));
    let content_types = text(&saved, "[Content_Types].xml");
    for part in [
        "/xl/comments1.xml",
        "/xl/threadedComments/threadedComment1.xml",
        "/xl/persons/person.xml",
    ] {
        assert!(content_types.contains(part), "no content type for {part}");
    }
    assert!(content_types.contains(r#"Extension="vml""#));
}

#[test]
fn a_placed_picture_survives_an_edit_and_a_save() {
    let (source, saved) = edit_and_save();
    assert_parts_kept(
        &source,
        &saved,
        &[
            "xl/drawings/drawing1.xml",
            "xl/drawings/_rels/drawing1.xml.rels",
            "xl/media/image1.png",
        ],
    );
    assert!(text(&saved, "xl/worksheets/sheet1.xml").contains(r#"<drawing r:id="rId1"/>"#));
    let bytes = ooxml_opc::rezip_parts(&saved.into_iter().collect::<Vec<_>>()).unwrap();
    let images = Workbook::open(&bytes)
        .unwrap()
        .embedded_images(SheetId(0))
        .unwrap();
    assert_eq!(images.len(), 2);
    assert_eq!(images[0].part, "xl/media/image1.png");
    assert_eq!(images[0].bytes, PNG);
}

#[test]
fn sheet_and_workbook_protection_survive_an_edit_and_a_save() {
    let (_, saved) = edit_and_save();
    let worksheet = text(&saved, "xl/worksheets/sheet1.xml");
    let order = [
        "</sheetData>",
        SHEET_PROTECTION,
        PROTECTED_RANGES,
        "<pageMargins",
    ]
    .map(|fragment| {
        worksheet
            .find(fragment)
            .unwrap_or_else(|| panic!("no {fragment}: {worksheet}"))
    });
    assert!(
        order.windows(2).all(|pair| pair[0] < pair[1]),
        "{worksheet}"
    );
    let workbook = text(&saved, "xl/workbook.xml");
    let protection = r#"<workbookProtection workbookAlgorithmName="SHA-512" workbookHashValue="bm90IGEgcmVhbCBoYXNo" workbookSaltValue="c2FsdA==" workbookSpinCount="100000" lockStructure="1"/>"#;
    let order = ["<workbookPr", protection, "<bookViews>"].map(|fragment| {
        workbook
            .find(fragment)
            .unwrap_or_else(|| panic!("no {fragment}: {workbook}"))
    });
    assert!(order.windows(2).all(|pair| pair[0] < pair[1]), "{workbook}");
}

/// Applies structural edits on a collaborative replica and returns the source
/// and saved packages.
fn restructure_and_save(ops: Vec<Op>) -> (BTreeMap<String, Vec<u8>>, BTreeMap<String, Vec<u8>>) {
    let source = source();
    let mut workbook = Workbook::open_collaborative(&source, 5103).unwrap();
    workbook
        .apply_ops(ops, CalculationOptions::default())
        .unwrap();
    let saved = workbook.save().unwrap();
    Workbook::open(&saved).expect("the saved workbook reopens");
    (parts(&source), parts(&saved))
}

#[test]
fn inserted_rows_and_columns_move_notes_pictures_and_protected_ranges() {
    let (source, saved) = restructure_and_save(vec![
        Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
        Op::InsertCols {
            sheet: SheetId(0),
            at: 0,
            count: 2,
        },
    ]);
    // Everything moves one row down and two columns right, and nothing else
    // in these parts changes.
    let moved = |path: &str, pairs: &[(&str, &str)]| {
        let mut expected = text(&source, path);
        for (from, to) in pairs {
            assert!(expected.contains(from), "{path} has no {from}");
            expected = expected.replace(from, to);
        }
        assert_eq!(text(&saved, path), expected, "{path}");
    };
    moved(
        "xl/comments1.xml",
        &[
            (r#"ref="B2""#, r#"ref="D3""#),
            (r#"ref="A3""#, r#"ref="C4""#),
        ],
    );
    moved(
        "xl/threadedComments/threadedComment1.xml",
        &[(r#"ref="B2""#, r#"ref="D3""#)],
    );
    moved(
        "xl/drawings/vmlDrawing1.vml",
        &[
            (
                "<x:Anchor>2, 15, 0, 2, 4, 15, 4, 4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>1</x:Row><x:Column>1</x:Column>",
                "<x:Anchor>4, 15, 1, 2, 6, 15, 5, 4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>2</x:Row><x:Column>3</x:Column>",
            ),
            (
                "<x:Anchor>1, 15, 1, 2, 3, 15, 5, 4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>2</x:Row><x:Column>0</x:Column>",
                "<x:Anchor>3, 15, 2, 2, 5, 15, 6, 4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>3</x:Row><x:Column>2</x:Column>",
            ),
        ],
    );
    moved(
        "xl/drawings/drawing1.xml",
        &[
            (
                "<xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row>",
                "<xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row>",
            ),
            (
                "<xdr:to><xdr:col>5</xdr:col><xdr:colOff>304800</xdr:colOff><xdr:row>6</xdr:row>",
                "<xdr:to><xdr:col>7</xdr:col><xdr:colOff>304800</xdr:colOff><xdr:row>7</xdr:row>",
            ),
            (
                "<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row>",
                "<xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row>",
            ),
            (
                "<xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row>",
                "<xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>9</xdr:row>",
            ),
        ],
    );
    let sheet = text(&saved, "xl/worksheets/sheet1.xml");
    assert!(
        sheet.contains(
            r#"<protectedRanges><protectedRange name="Scores" sqref="D3:D4"/></protectedRanges>"#
        ),
        "{sheet}"
    );
}

#[test]
fn deleting_noted_cells_drops_their_notes_and_pulls_pictures_to_the_deletion() {
    // Row 2 holds the thread on B2; column A holds the note on A3 (A2 once
    // row 2 is gone). Excel drops a comment with its cell.
    let (source, saved) = restructure_and_save(vec![
        Op::DeleteRows {
            sheet: SheetId(0),
            at: 1,
            count: 1,
        },
        Op::DeleteCols {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
    ]);
    let count = |path: &str, needle: &str| text(&saved, path).matches(needle).count();
    assert_eq!(count("xl/comments1.xml", "<comment "), 0);
    assert!(text(&saved, "xl/comments1.xml").contains("<authors>"));
    assert_eq!(
        count(
            "xl/threadedComments/threadedComment1.xml",
            "<threadedComment "
        ),
        0
    );
    assert_eq!(count("xl/drawings/vmlDrawing1.vml", "<v:shape "), 0);
    assert_eq!(count("xl/drawings/vmlDrawing1.vml", "<v:shapetype "), 1);

    // The logo moves without resizing: its top row was deleted, so it stays at
    // the deletion point, and it moves one column left. The photo's left
    // column was deleted: its edge stays at the deletion point and it narrows.
    let mut expected = text(&source, "xl/drawings/drawing1.xml");
    for (from, to) in [
        (
            "<xdr:from><xdr:col>3</xdr:col>",
            "<xdr:from><xdr:col>2</xdr:col>",
        ),
        (
            "<xdr:to><xdr:col>5</xdr:col>",
            "<xdr:to><xdr:col>4</xdr:col>",
        ),
        (
            "<xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>",
            "<xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>",
        ),
        (
            "<xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row>",
            "<xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>7</xdr:row>",
        ),
    ] {
        assert_eq!(expected.matches(from).count(), 1, "{from}");
        expected = expected.replace(from, to);
    }
    assert_eq!(text(&saved, "xl/drawings/drawing1.xml"), expected);

    // B2:B3 loses its deleted row and moves into column A.
    let sheet = text(&saved, "xl/worksheets/sheet1.xml");
    assert!(
        sheet.contains(r#"<protectedRange name="Scores" sqref="A2"/>"#),
        "{sheet}"
    );
}
