use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

/// A body with word parts related from the main document.
mod fixture {
    pub const NS: &str = concat!(
        r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
        r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#
    );
    const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const OFFICE: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml";

    pub struct Package {
        body: String,
        parts: Vec<(String, String, String, String)>,
    }

    impl Package {
        pub fn new(body: &str) -> Self {
            Self {
                body: body.to_owned(),
                parts: Vec::new(),
            }
        }

        /// Adds `word/{name}`, related from the main document as `id`.
        pub fn part(
            mut self,
            name: &str,
            id: &str,
            kind: &str,
            content_type: &str,
            xml: &str,
        ) -> Self {
            self.parts.push((
                name.to_owned(),
                format!(r#"<Relationship Id="{id}" Type="{REL}/{kind}" Target="{name}"/>"#),
                format!(r#"<Override PartName="/word/{name}" ContentType="{OFFICE}.{content_type}+xml"/>"#),
                xml.to_owned(),
            ));
            self
        }

        pub fn bytes(&self) -> Vec<u8> {
            let overrides: String = self.parts.iter().map(|part| part.2.as_str()).collect();
            let rels: String = self.parts.iter().map(|part| part.1.as_str()).collect();
            let relationships = |entries: &str| {
                format!(
                    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{entries}</Relationships>"#
                )
            };
            let mut parts = vec![
                (
                    "[Content_Types].xml".to_owned(),
                    format!(
                        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{OFFICE}.document.main+xml"/>{overrides}</Types>"#
                    ),
                ),
                (
                    "_rels/.rels".to_owned(),
                    relationships(&format!(
                        r#"<Relationship Id="rId1" Type="{REL}/officeDocument" Target="word/document.xml"/>"#
                    )),
                ),
                (
                    "word/_rels/document.xml.rels".to_owned(),
                    relationships(&rels),
                ),
                (
                    "word/document.xml".to_owned(),
                    format!(
                        r#"<w:document {NS}><w:body>{}</w:body></w:document>"#,
                        self.body
                    ),
                ),
            ];
            for (name, _, _, xml) in &self.parts {
                parts.push((format!("word/{name}"), xml.clone()));
            }
            let parts: Vec<_> = parts
                .into_iter()
                .map(|(name, xml)| (name, xml.into_bytes()))
                .collect();
            ooxml_opc::rezip_parts(&parts).unwrap()
        }
    }
}

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn document(header: Option<&str>, footer: Option<&str>, title: bool, even: bool) -> Vec<u8> {
    let references: String = [("header", header), ("footer", footer)]
        .into_iter()
        .filter_map(|(kind, variant)| {
            variant
                .map(|variant| format!(r#"<w:{kind}Reference w:type="{variant}" r:id="{kind}"/>"#))
        })
        .collect();
    let title = if title { "<w:titlePg/>" } else { "" };
    let section = format!(
        r#"<w:sectPr>{references}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="450" w:right="1440" w:bottom="450" w:left="1440" w:header="708" w:footer="708"/>{title}</w:sectPr>"#
    );
    let paragraph = |text: &str, page_break: bool| {
        let page_break = if page_break {
            "<w:pageBreakBefore/>"
        } else {
            ""
        };
        format!(
            r#"<w:p><w:pPr>{page_break}<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"#
        )
    };
    let body = format!(
        "{}{}{}{section}",
        paragraph("First", false),
        paragraph("Second", true),
        paragraph("Third", true),
    );
    let band = paragraph("Band", false);
    let even = if even { "<w:evenAndOddHeaders/>" } else { "" };
    fixture::Package::new(&body)
        .part(
            "header.xml",
            "header",
            "header",
            "header",
            &format!(r#"<w:hdr {}>{band}</w:hdr>"#, fixture::NS),
        )
        .part(
            "footer.xml",
            "footer",
            "footer",
            "footer",
            &format!(r#"<w:ftr {}>{band}</w:ftr>"#, fixture::NS),
        )
        .part(
            "settings.xml",
            "settings",
            "settings",
            "settings",
            &format!(r#"<w:settings {}>{even}</w:settings>"#, fixture::NS),
        )
        .bytes()
}

fn layout(bytes: &[u8]) -> Value {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let engine = EngineSession::new(76610);
    seed_from_docx(engine.doc(), bytes).unwrap();
    serde_json::from_str(
        &engine
            .layout_document_with_regions_json(
                &json!({
                    "bodyStory": "body", "renderEnv": {},
                    "regions": {
                        "sections": [{"properties": package.document.final_section_properties}],
                        "settings": package.settings
                    },
                    "measurement": {
                        "fontChains": {"calibri|0|0": [font]},
                        "defaults": {"fontFamily": "Calibri", "fontSize": 12}
                    }
                })
                .to_string(),
            )
            .unwrap(),
    )
    .unwrap()
}

fn assert_body_margins(output: &Value, expected: [(f64, f64); 3]) {
    let pages = output["layout"]["pages"].as_array().unwrap();
    assert_eq!(pages.len(), expected.len());
    for (page, (top, bottom)) in pages.iter().zip(expected) {
        let margins = if page["bodyMargins"].is_object() {
            &page["bodyMargins"]
        } else {
            &page["margins"]
        };
        assert!((margins["top"].as_f64().unwrap() - top).abs() < 1e-6);
        assert!((margins["bottom"].as_f64().unwrap() - bottom).abs() < 1e-6);
        assert!((page["fragments"][0]["y"].as_f64().unwrap() - top).abs() < 1e-6);
        assert_eq!(margins["header"].as_f64(), Some(47.2));
        assert_eq!(margins["footer"].as_f64(), Some(47.2));
    }
}

#[test]
fn no_header_reference_keeps_the_top_margin_with_a_default_footer() {
    let output = layout(&document(None, Some("default"), false, false));
    assert_body_margins(&output, [(30.0, 63.2); 3]);
}

#[test]
fn default_header_and_footer_keep_the_existing_extension() {
    let output = layout(&document(Some("default"), Some("default"), false, false));
    assert_body_margins(&output, [(63.2, 63.2); 3]);
}

#[test]
fn no_footer_reference_keeps_the_bottom_margin_with_a_default_header() {
    let output = layout(&document(Some("default"), None, false, false));
    assert_body_margins(&output, [(63.2, 30.0); 3]);
}

#[test]
fn unreferenced_parts_leave_both_body_margins_unchanged() {
    let output = layout(&document(None, None, false, false));
    assert_body_margins(&output, [(30.0, 30.0); 3]);
}
