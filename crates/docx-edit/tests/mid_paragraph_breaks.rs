//! A page or column break inside a paragraph's text, as Word lays it out: the
//! text after it continues at the top of the next page or column, still in
//! the same paragraph (no space-before, no first-line indent, no list
//! number), and nothing before the break is painted again.

use docx_edit::{EditCtx, EngineSession, FormatPolicy, Position, seed_from_docx};
use serde_json::json;

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const ONE_COLUMN: &str = "";
const TWO_COLUMNS: &str = r#"<w:cols w:num="2" w:space="720"/>"#;
/// One inch of page margin.
const LEFT: f64 = 96.0;
/// `w:ind w:left="720" w:firstLine="720"`, half an inch each.
const INDENT: &str = r#"<w:ind w:left="720" w:firstLine="720"/>"#;
const HALF_INCH: f64 = 48.0;
const NUMBERING: &str = r#"<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>"#;

fn document(body: &str, columns: &str) -> Vec<u8> {
    let styles = format!(
        r#"<w:styles xmlns:w="{W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>"#
    );
    let main = "application/vnd.openxmlformats-officedocument.wordprocessingml";
    let relationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let parts = [
        (
            "[Content_Types].xml",
            format!(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{main}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="{main}.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="{main}.numbering+xml"/></Types>"#
            ),
        ),
        (
            "_rels/.rels",
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="{relationship}/officeDocument" Target="word/document.xml"/></Relationships>"#
            ),
        ),
        (
            "word/_rels/document.xml.rels",
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="{relationship}/styles" Target="styles.xml"/><Relationship Id="numbering" Type="{relationship}/numbering" Target="numbering.xml"/></Relationships>"#
            ),
        ),
        (
            "word/document.xml",
            format!(
                r#"<w:document xmlns:w="{W}"><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>{columns}</w:sectPr></w:body></w:document>"#
            ),
        ),
        ("word/styles.xml", styles),
        (
            "word/numbering.xml",
            format!(r#"<w:numbering xmlns:w="{W}">{NUMBERING}</w:numbering>"#),
        ),
    ];
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name.to_owned(), value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

fn p(properties: &str, runs: &str) -> String {
    format!("<w:p><w:pPr>{properties}</w:pPr>{runs}</w:p>")
}

fn t(text: &str) -> String {
    format!("<w:r><w:t>{text}</w:t></w:r>")
}

fn br(kind: &str) -> String {
    format!(r#"<w:r><w:br w:type="{kind}"/></w:r>"#)
}

struct Laid {
    engine: EngineSession,
    extras: String,
}

fn lay_out(bytes: &[u8]) -> Laid {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let engine = EngineSession::new(76401);
    seed_from_docx(engine.doc(), bytes).unwrap();
    engine
        .layout_document_with_regions_json(
            &json!({
                "bodyStory": "body", "renderEnv": {},
                "regions": {"sections": [{"properties": package.document.final_section_properties}], "settings": package.settings},
                "measurement": {"fontChains": {"arial|0|0": [font]}, "defaults": {"fontFamily": "Arial", "fontSize": 12}}
            })
            .to_string(),
        )
        .unwrap();
    let extras = json!({"fontChains": {"arial|0|0": [font]}}).to_string();
    engine.build_display_list_frame(&extras, 0).unwrap();
    Laid { engine, extras }
}

/// Each page's painted lines in paint order: (text, left x, baseline).
fn lines(laid: &Laid) -> Vec<Vec<(String, f64, f64)>> {
    let display = laid
        .engine
        .with_display_list(|display| serde_json::to_value(display).unwrap())
        .unwrap();
    display["pages"]
        .as_array()
        .unwrap()
        .iter()
        .map(|page| {
            let mut lines: Vec<(String, f64, f64)> = Vec::new();
            for primitive in page["primitives"].as_array().unwrap() {
                let origin = match primitive["kind"].as_str() {
                    Some("text") => primitive,
                    Some("glyphRun") => &primitive["glyphs"][0],
                    _ => continue,
                };
                let text = primitive["text"].as_str().unwrap();
                let x = origin["x"].as_f64().unwrap();
                let baseline = origin["baselineY"]
                    .as_f64()
                    .or(origin["y"].as_f64())
                    .unwrap();
                match lines.last_mut() {
                    Some(line) if line.2 == baseline && x >= line.1 => line.0.push_str(text),
                    _ => lines.push((text.to_owned(), x, baseline)),
                }
            }
            lines
        })
        .collect()
}

fn texts(pages: &[Vec<(String, f64, f64)>]) -> Vec<Vec<&str>> {
    pages
        .iter()
        .map(|page| page.iter().map(|line| line.0.as_str()).collect())
        .collect()
}

fn line<'a>(pages: &'a [Vec<(String, f64, f64)>], text: &str) -> &'a (String, f64, f64) {
    pages
        .iter()
        .flatten()
        .find(|line| line.0 == text)
        .unwrap_or_else(|| panic!("no line {text:?} in {pages:?}"))
}

#[test]
fn text_after_a_mid_paragraph_page_break_starts_the_next_page() {
    let runs = [
        format!("{}{}{}", t("Aa"), br("page"), t("Bb")),
        format!(
            r#"<w:ins w:id="9" w:author="Ada">{}{}{}</w:ins>"#,
            t("Aa"),
            br("page"),
            t("Bb")
        ),
        format!(
            r#"<w:hyperlink w:anchor="target">{}{}{}</w:hyperlink>"#,
            t("Aa"),
            br("page"),
            t("Bb")
        ),
    ];
    for runs in runs {
        let body = format!(
            "{}{}{}",
            p("", &t("Lead")),
            p(&format!(r#"<w:spacing w:before="480"/>{INDENT}"#), &runs),
            p("", &t("Cc"))
        );
        let pages = lines(&lay_out(&document(&body, ONE_COLUMN)));
        assert_eq!(
            texts(&pages),
            [vec!["Lead", "Aa"], vec!["Bb", "Cc"]],
            "{runs}"
        );
        let top = line(&pages, "Lead").2;
        assert_eq!(line(&pages, "Aa").1, LEFT + 2.0 * HALF_INCH);
        assert_eq!(line(&pages, "Bb").1, LEFT + HALF_INCH, "{runs}");
        assert_eq!(line(&pages, "Bb").2, top, "{runs}");
        assert_eq!(line(&pages, "Cc").2, top + 16.0, "{runs}");
    }
}

/// `w:pageBreakBefore` turns the page before the paragraph, not before the
/// text after its column break.
#[test]
fn a_paragraph_with_page_break_before_continues_in_the_next_column() {
    let body = format!(
        "{}{}{}",
        p("", &t("Lead")),
        p(
            r#"<w:pageBreakBefore/><w:spacing w:before="480"/>"#,
            &format!("{}{}{}", t("Aa"), br("column"), t("Bb"))
        ),
        p("", &t("Cc"))
    );
    let pages = lines(&lay_out(&document(&body, TWO_COLUMNS)));
    assert_eq!(texts(&pages), [vec!["Lead"], vec!["Aa", "Bb", "Cc"]]);
    assert_eq!(line(&pages, "Bb").1, LEFT + 3.0 * 96.0 + HALF_INCH);
    assert_eq!(line(&pages, "Bb").2, line(&pages, "Lead").2);
}

#[test]
fn text_after_a_mid_paragraph_column_break_starts_the_next_column() {
    let body = format!(
        "{}{}{}",
        p("", &t("Lead")),
        p(
            &format!(r#"<w:spacing w:before="480"/>{INDENT}"#),
            &format!("{}{}{}", t("Aa"), br("column"), t("Bb"))
        ),
        p("", &t("Cc"))
    );
    let pages = lines(&lay_out(&document(&body, TWO_COLUMNS)));
    assert_eq!(texts(&pages), [vec!["Lead", "Aa", "Bb", "Cc"]]);
    // Two 3-inch columns half an inch apart.
    let second_column = LEFT + 3.0 * 96.0 + HALF_INCH;
    assert_eq!(line(&pages, "Bb").1, second_column + HALF_INCH);
    assert_eq!(line(&pages, "Bb").2, line(&pages, "Lead").2);
    assert_eq!(line(&pages, "Cc").1, second_column);
}

#[test]
fn a_list_item_continues_after_its_break_without_a_number() {
    let item = r#"<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>"#;
    let body = format!(
        "{}{}{}",
        p("", &t("Lead")),
        p(item, &format!("{}{}{}", t("Aa"), br("page"), t("Bb"))),
        p(item, &t("Cc"))
    );
    let pages = lines(&lay_out(&document(&body, ONE_COLUMN)));
    // A number and its item's text share a line.
    assert_eq!(texts(&pages), [vec!["Lead", "1.Aa"], vec!["Bb", "2.Cc"]]);
    assert_eq!(line(&pages, "Bb").1, LEFT + HALF_INCH);
    assert_eq!(line(&pages, "Bb").2, line(&pages, "Lead").2);
}

#[test]
fn a_break_ending_the_paragraph_text_starts_the_next_paragraph_on_the_next_page() {
    let body = format!(
        "{}{}{}",
        p("", &t("Lead")),
        p("", &format!("{}{}", t("Aa"), br("page"))),
        p("", &t("Cc"))
    );
    let pages = lines(&lay_out(&document(&body, ONE_COLUMN)));
    assert_eq!(texts(&pages), [vec!["Lead", "Aa"], vec!["Cc"]]);
    assert_eq!(line(&pages, "Cc").2, line(&pages, "Lead").2);
}

/// Typing before the break relays out only what changed; the reused page
/// after the break must still paint the text after it.
#[test]
fn typing_before_a_mid_paragraph_break_keeps_the_text_after_it() {
    let body = format!(
        "{}{}",
        p("", &format!("{}{}{}", t("Aa"), br("page"), t("Bb"))),
        p("", &t("Cc"))
    );
    let laid = lay_out(&document(&body, ONE_COLUMN));
    laid.engine
        .doc()
        .insert_text(
            &EditCtx::local("", ""),
            Position::new("body", 2),
            "x",
            FormatPolicy::Inherit,
        )
        .unwrap();
    laid.engine.apply_and_layout("body", 1).unwrap();
    assert_eq!(texts(&lines(&laid)), [vec!["Aax"], vec!["Bb", "Cc"]]);
    laid.engine
        .build_display_list_frame(&laid.extras, 2)
        .unwrap();
    assert_eq!(texts(&lines(&laid)), [vec!["Aax"], vec!["Bb", "Cc"]]);
}
