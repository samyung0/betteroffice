// Linear memory grows in large steps (crates/wasm-alloc).
#[cfg(target_arch = "wasm32")]
use wasm_alloc as _;

use docx_edit::{EngineSession, package_media, parse_docx_for_edit, seed_parsed_docx_in_place};
use docx_layout::display_list::{ResidentDisplayInput, px};
use docx_parse::S9PackageWire;
use docx_parse::section::SectionProperties;
use serde_json::{Value, json};
use wasm_bindgen::prelude::*;

const DEFAULT_PAGE_GAP: f64 = 24.0;

/// Opens in steps so the host registers the fonts the layout needs first:
/// `open`, `layoutRequestJson`, `fontRequirementsJson`, `layout`, then
/// `displayPagesJson` a batch of pages at a time (`pageSizesJson` sizes them
/// all before any is built). `layoutFirst` before `layout` lays out the first
/// pages alone, to build them while `layout` lays out the rest.
#[wasm_bindgen]
pub struct DocxViewDocument {
    engine: EngineSession,
    request: String,
    /// What the display list is built with once laid out, parsed once.
    display: Option<ResidentDisplayInput>,
    /// Each laid-out page's `[width, height]`, as its display page gives them.
    page_sizes: Vec<[serde_json::Number; 2]>,
    /// The widest page of the whole body, while only its first pages are laid out.
    document_width: f64,
}

#[wasm_bindgen]
impl DocxViewDocument {
    pub fn open(bytes: &[u8]) -> Result<DocxViewDocument, JsValue> {
        let envelope = parse_docx_for_edit(bytes).map_err(js_error)?;
        let request = layout_request(&envelope.document.package).to_string();
        let engine = EngineSession::new(1);
        engine.set_media(package_media(&envelope));
        seed_parsed_docx_in_place(engine.doc(), envelope).map_err(js_error)?;
        Ok(Self {
            engine,
            request,
            display: None,
            page_sizes: Vec::new(),
            document_width: 0.0,
        })
    }

    /// The region-layout request without `measurement`.
    #[wasm_bindgen(js_name = layoutRequestJson)]
    pub fn layout_request_json(&self) -> String {
        self.request.clone()
    }

    /// Font families and scripts `request` measures with.
    #[wasm_bindgen(js_name = fontRequirementsJson)]
    pub fn font_requirements_json(&self, request: &str) -> Result<String, JsValue> {
        self.engine
            .layout_font_requirements_json(request)
            .map_err(js_error)
    }

    /// Lays out `request` (with its `measurement`) and returns the page count.
    /// The measurement's font chains go to the display list too, so measured
    /// text becomes joined glyph runs as in the editor.
    pub fn layout(&mut self, request: &str) -> Result<u32, JsValue> {
        let sizes = self
            .engine
            .layout_document_with_regions(request)
            .map_err(js_error)?;
        self.laid_out(request, sizes)
    }

    /// Lays out only the body's first `pages` pages, each as [`Self::layout`]
    /// lays it out, with page-count fields showing their saved results, and
    /// returns `pages`; 0 when the body can't be cut after them, leaving
    /// nothing laid out. `layout` lays out the whole body after either.
    #[wasm_bindgen(js_name = layoutFirst)]
    pub fn layout_first(&mut self, request: &str, pages: u32) -> Result<u32, JsValue> {
        let Some(first) = self
            .engine
            .layout_document_with_regions_first_pages(request, pages as usize)
            .map_err(js_error)?
        else {
            self.display = None;
            return Ok(0);
        };
        self.document_width = first.document_width;
        self.laid_out(request, first.sizes)
    }

    /// The widest page of the whole body, once [`Self::layout_first`] laid out
    /// its first pages.
    #[wasm_bindgen(js_name = documentWidth)]
    pub fn document_width(&self) -> f64 {
        self.document_width
    }

    /// Readies the pages just laid out for display.
    fn laid_out(&mut self, request: &str, sizes: Vec<(f64, f64)>) -> Result<u32, JsValue> {
        let extras = match serde_json::from_str::<Value>(request)
            .map_err(js_error)?
            .pointer("/measurement/fontChains")
        {
            Some(chains) => json!({ "fontChains": chains }),
            None => json!({}),
        };
        self.display = Some(
            self.engine
                .region_display_input(&extras.to_string())
                .map_err(js_error)?,
        );
        self.page_sizes = sizes.into_iter().map(|(w, h)| [px(w), px(h)]).collect();
        u32::try_from(self.page_sizes.len()).map_err(js_error)
    }

    /// `[[width, height], …]` of every laid-out page, before its display page
    /// is built.
    #[wasm_bindgen(js_name = pageSizesJson)]
    pub fn page_sizes_json(&self) -> String {
        json!(self.page_sizes).to_string()
    }

    /// `{ pages }` display list JSON of pages `start..end`, built from the
    /// engine's typed state; the host frees the document after the last one.
    #[wasm_bindgen(js_name = displayPagesJson)]
    pub fn display_pages_json(&mut self, start: u32, end: u32) -> Result<String, JsValue> {
        let display = self
            .display
            .as_mut()
            .ok_or_else(|| js_error("DOCX view is not laid out"))?;
        self.engine
            .display_pages_json(display, start as usize..end as usize)
            .map_err(js_error)
    }

    pub fn version() -> String {
        env!("CARGO_PKG_VERSION").to_owned()
    }
}

/// Mirrors `buildResidentRegionLayoutRequest` (packages/docx/src/editor/computeLayout.ts).
fn layout_request(package: &S9PackageWire) -> Value {
    let body = &package.document;
    let body_sections = body.sections.as_deref().unwrap_or_default();
    let mut sections = body_sections
        .iter()
        .map(|section| {
            section_entry(
                section
                    .id
                    .as_ref()
                    .or(section.properties.section_id.as_ref()),
                &section.properties,
            )
        })
        .collect::<Vec<_>>();
    let final_id = sections
        .last()
        .and_then(|section| section.get("sectionId"))
        .and_then(Value::as_str)
        .map(str::to_owned)
        .or_else(|| {
            body.final_section_properties
                .as_ref()
                .and_then(|properties| properties.section_id.clone())
        });
    let final_properties = body_sections
        .last()
        .map(|section| &section.properties)
        .or(body.final_section_properties.as_ref())
        .cloned()
        .unwrap_or_default();
    sections.push(section_entry(final_id.as_ref(), &final_properties));

    let notes = package
        .footnotes
        .as_deref()
        .unwrap_or_default()
        .iter()
        .filter(|note| note.note_type.is_empty() || note.note_type == "normal")
        // Note ids parse as f64; the layout request reads them as integers.
        .map(|note| json!({ "id": note.id as i64, "noteKind": "footnote", "height": 0 }))
        .chain(
            package
                .endnotes
                .as_deref()
                .unwrap_or_default()
                .iter()
                .filter(|note| note.note_type.is_empty() || note.note_type == "normal")
                .map(|note| json!({ "id": note.id as i64, "noteKind": "endnote", "height": 0 })),
        )
        .collect::<Vec<_>>();

    let styles = package
        .styles
        .as_ref()
        .map(|styles| styles.styles.as_slice())
        .unwrap_or_default();
    let paragraph_styles = || {
        styles
            .iter()
            .filter(|style| style.style_type == "paragraph" && !style.style_id.is_empty())
    };
    let mut toc_style_ids = Vec::<&str>::new();
    for style in paragraph_styles() {
        let toc = std::iter::once(style.style_id.as_str())
            .chain(style.name.as_deref())
            .any(is_toc_style_name);
        if toc && !toc_style_ids.contains(&style.style_id.as_str()) {
            toc_style_ids.push(&style.style_id);
        }
    }
    let default_paragraph_style_id = paragraph_styles()
        .find(|style| style.default == Some(true))
        .map(|style| style.style_id.as_str())
        .or_else(|| {
            paragraph_styles()
                .any(|style| style.style_id == "Normal")
                .then_some("Normal")
        });

    let theme_colors = serde_json::to_value(&package.theme.color_scheme).unwrap_or_default();
    let settings = serde_json::to_value(&package.settings).unwrap_or_default();
    let default_tab_stop = settings
        .get("defaultTabStop")
        .cloned()
        .unwrap_or(Value::Null);
    let mut render_env = json!({
        "themeColors": theme_colors,
        "defaultTabStopTwips": default_tab_stop,
        "numericIds": {},
        "tocStyleIds": toc_style_ids,
    });
    if let Some(flags) = settings.get("compatibilityFlags") {
        render_env["compatibilityFlags"] = flags.clone();
    }
    if let Some(id) = default_paragraph_style_id {
        render_env["defaultParagraphStyleId"] = json!(id);
    }
    // What the editor's host sends too: the theme whose fonts win over run font
    // names (#1921), the document's default language (#2246), each paragraph
    // style chain's size for unsized runs (#1880) and the first picture
    // bullet's box (#1936).
    render_env["fontTheme"] = json!(package.theme);
    let doc_defaults = package
        .styles
        .as_ref()
        .and_then(|styles| styles.doc_defaults.as_ref())
        .and_then(|defaults| defaults.r_pr.as_ref());
    if let Some(language) = doc_defaults.and_then(|r_pr| r_pr.language.as_ref()) {
        let slots: serde_json::Map<String, Value> = [
            ("latin", &language.latin),
            ("eastAsia", &language.east_asia),
            ("bidi", &language.bidi),
        ]
        .into_iter()
        .filter_map(|(slot, tag)| {
            tag.as_deref()
                .filter(|tag| !tag.is_empty())
                .map(|tag| (slot.to_owned(), json!(tag)))
        })
        .collect();
        if !slots.is_empty() {
            render_env["docDefaultLanguage"] = Value::Object(slots);
        }
    }
    let doc_default_size = doc_defaults.and_then(|r_pr| r_pr.font_size).unwrap_or(20.0);
    let chain_sizes: serde_json::Map<String, Value> = paragraph_styles()
        .map(|style| {
            let mut seen = Vec::new();
            let mut current = Some(style);
            let mut size = None;
            while let Some(entry) = current.filter(|entry| !seen.contains(&&entry.style_id)) {
                seen.push(&entry.style_id);
                if let Some(found) = entry.r_pr.as_ref().and_then(|r_pr| r_pr.font_size) {
                    size = Some(found);
                    break;
                }
                current = entry
                    .based_on
                    .as_deref()
                    .and_then(|id| styles.iter().find(|style| style.style_id == id));
            }
            (
                style.style_id.clone(),
                json!(size.unwrap_or(doc_default_size) / 2.0),
            )
        })
        .collect();
    if !chain_sizes.is_empty() {
        render_env["paragraphStyleFontSizes"] = Value::Object(chain_sizes);
    }
    render_env["docDefaultFontSize"] = json!(doc_default_size / 2.0);
    if let Some(bullet) = &package.numbering.first_picture_bullet {
        render_env["firstPictureBullet"] = json!(bullet);
    }
    // The default paragraph style's size, else the document default, else
    // Word's own 10pt: what a reading layout eases sizes around.
    let body_half_points = default_paragraph_style_id
        .and_then(|id| styles.iter().find(|style| style.style_id == id))
        .and_then(|style| style.r_pr.as_ref()?.font_size)
        .or_else(|| {
            package
                .styles
                .as_ref()?
                .doc_defaults
                .as_ref()?
                .r_pr
                .as_ref()?
                .font_size
        })
        .filter(|size| *size > 0.0)
        .unwrap_or(20.0);
    json!({
        "bodyStory": "body",
        "bodyFontSize": body_half_points / 2.0,
        "options": { "pageGap": DEFAULT_PAGE_GAP },
        "regions": {
            "sections": sections,
            "settings": settings,
        },
        "notes": { "contents": notes },
        "renderEnv": render_env,
    })
}

fn section_entry(section_id: Option<&String>, properties: &SectionProperties) -> Value {
    let mut entry = json!({ "properties": properties });
    if let Some(id) = section_id {
        entry["sectionId"] = json!(id);
    }
    entry
}

/// `/^TOC\s*\d+$/i`
fn is_toc_style_name(name: &str) -> bool {
    let Some(rest) = name
        .get(..3)
        .filter(|prefix| prefix.eq_ignore_ascii_case("toc"))
        .map(|_| name[3..].trim_start())
    else {
        return false;
    };
    !rest.is_empty() && rest.bytes().all(|byte| byte.is_ascii_digit())
}

fn js_error(error: impl std::fmt::Display) -> JsValue {
    JsValue::from_str(&error.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_document_with_footnotes_lays_out() {
        let bytes = include_bytes!("../../docx-edit/tests/fixtures/footnote-anchor.docx");
        let mut document = DocxViewDocument::open(bytes).unwrap();
        let request: Value = serde_json::from_str(&document.layout_request_json()).unwrap();
        let notes = request["notes"]["contents"].as_array().unwrap();
        assert!(!notes.is_empty());
        assert!(notes.iter().all(|note| note["id"].is_i64()));
        document
            .font_requirements_json(&document.layout_request_json())
            .unwrap();
        let pages = document.layout(&document.layout_request_json()).unwrap();
        assert!(pages > 0);
        document.display_pages_json(0, pages).unwrap();
    }

    /// Every page is sized right after layout as its display page is built.
    #[test]
    fn page_sizes_are_the_display_pages_sizes() {
        for bytes in [
            &include_bytes!("../../../poc/fixtures/book-30p.docx")[..],
            include_bytes!("../../../poc/fixtures/wordprocessingml-comprehensive.docx"),
        ] {
            let mut document = DocxViewDocument::open(bytes).unwrap();
            let pages = document.layout(&document.layout_request_json()).unwrap();
            let sizes: Value = serde_json::from_str(&document.page_sizes_json()).unwrap();
            let list: Value =
                serde_json::from_str(&document.display_pages_json(0, pages).unwrap()).unwrap();
            let built = list["pages"]
                .as_array()
                .unwrap()
                .iter()
                .map(|page| json!([page["width"], page["height"]]))
                .collect::<Vec<_>>();
            assert!(pages > 1);
            assert_eq!(sizes, Value::Array(built));
        }
    }

    /// The viewer's list as the host joins it, from batches of `batch` pages.
    fn paged_display_list(document: &mut DocxViewDocument, request: &str, batch: u32) -> String {
        let pages = document.layout(request).unwrap();
        let joined = (0..pages)
            .step_by(batch as usize)
            .map(|start| {
                let json = document
                    .display_pages_json(start, (start + batch).min(pages))
                    .unwrap();
                json.strip_prefix(r#"{"pages":["#)
                    .and_then(|pages| pages.strip_suffix("]}"))
                    .unwrap()
                    .to_owned()
            })
            .collect::<Vec<_>>()
            .join(",");
        format!(r#"{{"pages":[{joined}]}}"#)
    }

    /// Equal apart from numbers at most 4 ULP apart.
    fn ulp_equal(a: &Value, b: &Value) -> bool {
        match (a, b) {
            (Value::Number(x), Value::Number(y)) => {
                let (x, y) = (x.as_f64().unwrap(), y.as_f64().unwrap());
                x == y || (x.to_bits() as i64 - y.to_bits() as i64).abs() <= 4
            }
            (Value::Array(x), Value::Array(y)) => {
                x.len() == y.len() && x.iter().zip(y).all(|(x, y)| ulp_equal(x, y))
            }
            (Value::Object(x), Value::Object(y)) => {
                x.len() == y.len()
                    && x.iter()
                        .all(|(key, value)| y.get(key).is_some_and(|other| ulp_equal(value, other)))
            }
            _ => a == b,
        }
    }

    /// The paged build writes the list the whole-layout JSON round trip
    /// (`layout_document_with_regions_json` then `build_display_list_json`)
    /// wrote. The round trip read floats back without `float_roundtrip`, so
    /// in two fixtures the text of a few numbers differs from the engine's
    /// own values, which the paged build keeps (and the editor paints); those
    /// compare within 4 ULP once parsed, every other fixture byte-identical.
    #[test]
    fn paged_display_list_matches_the_json_round_trip() {
        let fixtures: [(&str, &[u8], bool); 8] = [
            (
                "footnote-anchor",
                include_bytes!("../../docx-edit/tests/fixtures/footnote-anchor.docx"),
                false,
            ),
            (
                "stories",
                include_bytes!("../../office-service/tests/fixtures/stories.docx"),
                false,
            ),
            (
                "probe-linked-header",
                include_bytes!(
                    "../../../packages/docx-react/src/components/DocxEditor/hooks/__fixtures__/probe-linked-header.docx"
                ),
                false,
            ),
            (
                "feature-rich",
                include_bytes!("../../../poc/fixtures/feature-rich.docx"),
                false,
            ),
            (
                "exchange-plan",
                include_bytes!("../../../poc/fixtures/exchange-plan.docx"),
                false,
            ),
            (
                "book-30p",
                include_bytes!("../../../poc/fixtures/book-30p.docx"),
                false,
            ),
            (
                "opaque-objects",
                include_bytes!("../../../poc/fixtures/opaque-objects.docx"),
                true,
            ),
            (
                "wordprocessingml-comprehensive",
                include_bytes!("../../../poc/fixtures/wordprocessingml-comprehensive.docx"),
                true,
            ),
        ];
        for (name, bytes, ulp_only) in fixtures {
            let mut document = DocxViewDocument::open(bytes).unwrap();
            let request = document.layout_request_json();
            let paged = paged_display_list(&mut document, &request, 3);

            let envelope = parse_docx_for_edit(bytes).unwrap();
            let engine = EngineSession::new(1);
            engine.set_media(package_media(&envelope));
            seed_parsed_docx_in_place(engine.doc(), envelope).unwrap();
            let layout = engine.layout_document_with_regions_json(&request).unwrap();
            let round_trip = engine.build_display_list_json(&layout).unwrap();

            if ulp_only {
                assert!(
                    ulp_equal(
                        &serde_json::from_str(&paged).unwrap(),
                        &serde_json::from_str(&round_trip).unwrap()
                    ),
                    "{name}"
                );
            } else {
                assert!(paged == round_trip, "{name}");
            }
        }
    }

    /// `document`'s request measuring every font it requires with `font`, as
    /// the host registers the faces it requires.
    fn measured_request(document: &DocxViewDocument, font: u32) -> String {
        let mut request: Value = serde_json::from_str(&document.layout_request_json()).unwrap();
        let requirements: Vec<Value> = serde_json::from_str(
            &document
                .font_requirements_json(&request.to_string())
                .unwrap(),
        )
        .unwrap();
        let chains = requirements
            .iter()
            .map(|requirement| {
                (
                    requirement["key"].as_str().unwrap().to_owned(),
                    json!([font]),
                )
            })
            .collect::<serde_json::Map<_, _>>();
        request["measurement"] = json!({
            "fontChains": chains,
            "defaults": { "fontSize": 11, "fontFamily": "Calibri" },
            "authoritativeShaping": true,
        });
        request.to_string()
    }

    const LIBERATION: &[u8] =
        include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

    const NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape""#;
    const RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const TYPE: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml";
    const TEXT: &str = "The committee reviews every application against the published criteria before the spring deadline, and late submissions wait for the autumn round while the panel settles the budget.";
    /// Letter pages with one-inch margins.
    const LETTER: &str = r#"<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>"#;

    /// A 1×1 PNG.
    const PIXEL: &[u8] = &[
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f,
        0x15, 0xc4, 0x89, 0x00, 0x00, 0x00, 0x0a, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00,
        0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49,
        0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
    ];

    /// About fifteen pages whose paragraphs carry footnotes, with floating
    /// tables beside their paragraph and on the page, pictures wrapped at
    /// their paragraph, keep-with-next headings and a footer whose page count
    /// saves a stale 99, so notes and floats fall on both sides of a cut.
    fn notes_and_floats_docx() -> Vec<u8> {
        let floating = |anchor: &str, y: u32| {
            format!(
                r#"<w:tbl><w:tblPr><w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="{anchor}" w:horzAnchor="margin" w:tblpXSpec="right" w:tblpY="{y}"/><w:tblW w:w="2400" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Floating {anchor} table</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#
            )
        };
        let picture = |id: usize| {
            format!(
                r#"<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="{id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:align>left</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="1371600" cy="914400"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="{id}" name="Picture {id}"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="{id}" name="pixel.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdPixel"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1371600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"#
            )
        };
        let mut body = String::new();
        let mut notes = String::new();
        for index in 1..=120 {
            if index % 9 == 1 {
                body += &format!(
                    r#"<w:p><w:pPr><w:keepNext/><w:spacing w:before="240"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="32"/></w:rPr><w:t>Part {index}</w:t></w:r></w:p>"#
                );
            }
            let mut runs = String::new();
            if index % 8 == 4 {
                runs += &picture(index);
            }
            runs +=
                &format!(r#"<w:r><w:t xml:space="preserve">{index}. {TEXT} {TEXT}</w:t></w:r>"#);
            if index % 3 == 0 {
                let id = index / 3;
                runs += &format!(
                    r#"<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="{id}"/></w:r>"#
                );
                notes += &format!(
                    r#"<w:footnote w:id="{id}"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Note {id}: {TEXT}</w:t></w:r></w:p></w:footnote>"#
                );
            }
            body += &format!("<w:p>{runs}</w:p>");
            if index % 5 == 0 {
                body += &floating(if index % 10 == 0 { "page" } else { "text" }, 400);
            }
        }
        body += r#"<w:sectPr><w:footerReference w:type="default" r:id="rIdFooter"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>"#;
        let parts = [
            (
                "[Content_Types].xml",
                format!(
                    r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="{TYPE}.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="{TYPE}.footnotes+xml"/><Override PartName="/word/footer1.xml" ContentType="{TYPE}.footer+xml"/></Types>"#
                ),
            ),
            (
                "_rels/.rels",
                format!(
                    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{RELS}/officeDocument" Target="word/document.xml"/></Relationships>"#
                ),
            ),
            (
                "word/_rels/document.xml.rels",
                format!(
                    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdNotes" Type="{RELS}/footnotes" Target="footnotes.xml"/><Relationship Id="rIdFooter" Type="{RELS}/footer" Target="footer1.xml"/><Relationship Id="rIdPixel" Type="{RELS}/image" Target="media/pixel.png"/></Relationships>"#
                ),
            ),
            (
                "word/footnotes.xml",
                format!(
                    r#"<w:footnotes {NS}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>{notes}</w:footnotes>"#
                ),
            ),
            (
                "word/footer1.xml",
                format!(
                    r#"<w:ftr {NS}><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve"> of </w:t></w:r><w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>99</w:t></w:r></w:fldSimple></w:p></w:ftr>"#
                ),
            ),
            (
                "word/document.xml",
                format!(r#"<w:document {NS}><w:body>{body}</w:body></w:document>"#),
            ),
        ];
        let mut entries = parts
            .into_iter()
            .map(|(path, xml)| (path.to_owned(), xml.into_bytes()))
            .collect::<Vec<_>>();
        entries.push(("word/media/pixel.png".to_owned(), PIXEL.to_vec()));
        ooxml_opc::rezip_parts(&entries).unwrap()
    }

    /// The text of the page-count (NUMPAGES) fields on a page, its header and
    /// footer included.
    fn page_counts(page: &Value) -> Vec<String> {
        [
            &page["primitives"],
            &page["header"]["primitives"],
            &page["footer"]["primitives"],
        ]
        .into_iter()
        .filter_map(Value::as_array)
        .flatten()
        .filter(|primitive| primitive["field"]["category"] == "NUMPAGES")
        .map(|primitive| primitive["text"].as_str().unwrap().to_owned())
        .collect()
    }

    /// Body paragraphs `range`, each two sentences long.
    fn text_paragraphs(range: std::ops::Range<usize>) -> String {
        range
            .map(|index| {
                format!(
                    r#"<w:p><w:r><w:t xml:space="preserve">{index}. {TEXT} {TEXT}</w:t></w:r></w:p>"#
                )
            })
            .collect()
    }

    /// A paragraph ending a section with the letter pages under the page-number
    /// footer, with `more` (columns, say).
    fn letter_section_end(more: &str) -> String {
        format!(
            r#"<w:p><w:pPr><w:sectPr><w:footerReference w:type="default" r:id="rIdFooter"/>{LETTER}{more}</w:sectPr></w:pPr></w:p>"#
        )
    }

    /// The final section's properties: letter pages under the page-number
    /// footer, with `more`.
    fn final_section(more: &str) -> String {
        format!(
            r#"<w:sectPr><w:footerReference w:type="default" r:id="rIdFooter"/>{LETTER}{more}</w:sectPr>"#
        )
    }

    /// A package whose body is `body` (its final `w:sectPr` included), with a
    /// footer showing the page number and `endnotes` (`w:endnote` elements).
    fn body_docx(body: &str, endnotes: &str) -> Vec<u8> {
        let parts = [
            (
                "[Content_Types].xml",
                format!(
                    r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{TYPE}.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="{TYPE}.endnotes+xml"/><Override PartName="/word/footer1.xml" ContentType="{TYPE}.footer+xml"/></Types>"#
                ),
            ),
            (
                "_rels/.rels",
                format!(
                    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{RELS}/officeDocument" Target="word/document.xml"/></Relationships>"#
                ),
            ),
            (
                "word/_rels/document.xml.rels",
                format!(
                    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdNotes" Type="{RELS}/endnotes" Target="endnotes.xml"/><Relationship Id="rIdFooter" Type="{RELS}/footer" Target="footer1.xml"/></Relationships>"#
                ),
            ),
            (
                "word/endnotes.xml",
                format!(
                    r#"<w:endnotes {NS}><w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>{endnotes}</w:endnotes>"#
                ),
            ),
            (
                "word/footer1.xml",
                format!(
                    r#"<w:ftr {NS}><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>"#
                ),
            ),
            (
                "word/document.xml",
                format!(r#"<w:document {NS}><w:body>{body}</w:body></w:document>"#),
            ),
        ];
        let entries = parts
            .into_iter()
            .map(|(path, xml)| (path.to_owned(), xml.into_bytes()))
            .collect::<Vec<_>>();
        ooxml_opc::rezip_parts(&entries).unwrap()
    }

    /// Liberation Sans, the font measuring every font a request requires.
    fn measure_font() -> u32 {
        docx_layout::clear_measure_fonts();
        docx_layout::register_measure_font(LIBERATION).unwrap()
    }

    /// Cuts `bytes` after three pages, which must lay out as the whole body
    /// does, page-count fields showing `saved` (the file's count) until the
    /// whole layout; the whole layout after the cut must be a fresh one's.
    fn assert_cut_as_whole(name: &str, bytes: &[u8], saved: Option<&str>, font: u32) {
        let mut cut = DocxViewDocument::open(bytes).unwrap();
        let request = measured_request(&cut, font);
        assert_eq!(cut.layout_first(&request, 3).unwrap(), 3, "{name}");
        let first: Value = serde_json::from_str(&cut.display_pages_json(0, 3).unwrap()).unwrap();
        let after = paged_display_list(&mut cut, &request, 8);
        let mut fresh = DocxViewDocument::open(bytes).unwrap();
        let whole = paged_display_list(&mut fresh, &request, 8);
        assert!(after == whole, "{name}: the whole layout after the cut");
        let whole: Value = serde_json::from_str(&whole).unwrap();
        let total = whole["pages"].as_array().unwrap().len().to_string();
        for index in 0..3 {
            let (shown, last) = (&first["pages"][index], &whole["pages"][index]);
            match saved {
                None => assert_eq!(shown, last, "{name} page {index}"),
                Some(saved) => {
                    assert_eq!(page_counts(shown), [saved], "{name} page {index}");
                    assert_eq!(page_counts(last), [total.as_str()], "{name} page {index}");
                    // Only the header and footer lines holding the count move.
                    let bare = |page: &Value| {
                        let mut page = page.clone();
                        let page = page.as_object_mut().unwrap();
                        page.remove("header");
                        page.remove("footer");
                        Value::Object(page.clone())
                    };
                    assert_eq!(bare(shown), bare(last), "{name} page {index}");
                }
            }
        }
    }

    /// `bytes`, several pages long, is not cut: the caller lays it out whole.
    fn assert_not_cut(name: &str, bytes: &[u8], font: u32) {
        let mut document = DocxViewDocument::open(bytes).unwrap();
        let request = measured_request(&document, font);
        assert_eq!(document.layout_first(&request, 3).unwrap(), 0, "{name}");
        assert!(document.layout(&request).unwrap() > 5, "{name}");
    }

    /// A cut inside a landscape section between portrait ones, its pages
    /// numbered from 5, lays the section out with its own page size, margins
    /// and numbering, not the final section's.
    #[test]
    fn a_cut_between_sections_is_the_whole_layouts() {
        let font = measure_font();
        let body = [
            text_paragraphs(0..8),
            letter_section_end(""),
            text_paragraphs(8..128),
            r#"<w:p><w:pPr><w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/><w:pgNumType w:start="5"/></w:sectPr></w:pPr></w:p>"#.to_owned(),
            text_paragraphs(128..168),
            r#"<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="2160" w:right="2160" w:bottom="2160" w:left="2160" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>"#.to_owned(),
        ]
        .concat();
        assert_cut_as_whole("sections", &body_docx(&body, ""), None, font);
    }

    /// The endnotes of a body cut after its first pages land past them, as
    /// the whole body's land at its end.
    #[test]
    fn endnotes_stay_off_the_first_pages() {
        let font = measure_font();
        let mut body = String::new();
        let mut endnotes = String::new();
        for index in 1..=150 {
            let mut runs =
                format!(r#"<w:r><w:t xml:space="preserve">{index}. {TEXT} {TEXT}</w:t></w:r>"#);
            if index % 4 == 0 {
                let id = index / 4;
                runs += &format!(
                    r#"<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="{id}"/></w:r>"#
                );
                endnotes += &format!(
                    r#"<w:endnote w:id="{id}"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteRef/></w:r><w:r><w:t xml:space="preserve"> Note {id}: {TEXT}</w:t></w:r></w:p></w:endnote>"#
                );
            }
            body += &format!("<w:p>{runs}</w:p>");
        }
        body += &final_section("");
        assert_cut_as_whole("endnotes", &body_docx(&body, &endnotes), None, font);
    }

    /// Columns are balanced over their whole section: a cut skips a section
    /// in columns (the first prefix ends inside it, the second in the
    /// one-column section after it) and a body all in columns is laid out
    /// whole.
    #[test]
    fn a_cut_skips_sections_in_columns() {
        let font = measure_font();
        let columns = r#"<w:cols w:num="2" w:space="720"/>"#;
        let body = [
            text_paragraphs(0..60),
            letter_section_end(columns),
            text_paragraphs(60..180),
            final_section(""),
        ]
        .concat();
        assert_cut_as_whole("columns then one", &body_docx(&body, ""), None, font);
        let body = text_paragraphs(0..300) + &final_section(columns);
        assert_not_cut("all columns", &body_docx(&body, ""), font);
    }

    /// A shape wrapped on the side of an inside margin is settled over the
    /// whole body, and a section break inside a content control has no
    /// request section: both bodies are laid out whole.
    #[test]
    fn margin_side_shapes_and_sections_in_controls_lay_out_whole() {
        let font = measure_font();
        let shape = r#"<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="insideMargin"><wp:align>left</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="914400"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="Shape 1"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"#;
        let body = [
            text_paragraphs(0..100),
            format!(r#"<w:p>{shape}<w:r><w:t>Beside the shape</w:t></w:r></w:p>"#),
            text_paragraphs(100..200),
            final_section(""),
        ]
        .concat();
        assert_not_cut("inside margin shape", &body_docx(&body, ""), font);
        let body = [
            text_paragraphs(0..100),
            format!(
                r#"<w:sdt><w:sdtPr><w:alias w:val="Part"/></w:sdtPr><w:sdtContent>{}</w:sdtContent></w:sdt>"#,
                letter_section_end("")
            ),
            text_paragraphs(100..200),
            final_section(""),
        ]
        .concat();
        assert_not_cut("section in a control", &body_docx(&body, ""), font);
    }

    /// Single-section bodies cut after their first pages lay them out as the
    /// whole body does: footnotes, floating tables, wrapped pictures and
    /// keep-with-next chains on both sides of the cut, page-count fields
    /// showing the saved count until the whole layout.
    #[test]
    fn first_pages_are_the_whole_layouts() {
        let font = measure_font();
        let fixtures: [(&str, Vec<u8>, Option<&str>); 3] = [
            ("notes-and-floats", notes_and_floats_docx(), Some("99")),
            (
                "long-handbook",
                include_bytes!("../../../poc/fixtures/long-handbook.docx").to_vec(),
                Some("60"),
            ),
            (
                "book-30p",
                include_bytes!("../../../poc/fixtures/book-30p.docx").to_vec(),
                None,
            ),
        ];
        for (name, bytes, saved) in fixtures {
            assert_cut_as_whole(name, &bytes, saved, font);
        }
    }

    /// A short file, here 15 pages of pictures, tables and little text, is
    /// laid out whole: its first pages would come no sooner (Q822 A).
    #[test]
    fn a_short_file_is_laid_out_whole() {
        assert_not_cut(
            "exchange-plan",
            include_bytes!("../../../poc/fixtures/exchange-plan.docx"),
            measure_font(),
        );
    }

    /// A body too short to lay out two pages past the ones asked for is not
    /// cut: the caller lays it out whole.
    #[test]
    fn a_short_body_is_not_cut() {
        let mut document = DocxViewDocument::open(include_bytes!(
            "../../docx-edit/tests/fixtures/footnote-anchor.docx"
        ))
        .unwrap();
        let request = document.layout_request_json();
        assert_eq!(document.layout_first(&request, 3).unwrap(), 0);
        assert!(document.layout(&request).unwrap() > 0);
    }

    /// With real fonts the viewer writes the list the editor's resident
    /// engine keeps: measured text as joined glyph runs.
    #[test]
    fn measured_text_becomes_the_editors_glyph_runs() {
        const FONT: &[u8] =
            include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
        docx_layout::clear_measure_fonts();
        let font = docx_layout::register_measure_font(FONT).unwrap();
        let bytes = include_bytes!("../../../poc/fixtures/feature-rich.docx");
        let mut document = DocxViewDocument::open(bytes).unwrap();
        let mut request: Value = serde_json::from_str(&document.layout_request_json()).unwrap();
        let requirements: Vec<Value> = serde_json::from_str(
            &document
                .font_requirements_json(&request.to_string())
                .unwrap(),
        )
        .unwrap();
        let chains = requirements
            .iter()
            .map(|requirement| {
                (
                    requirement["key"].as_str().unwrap().to_owned(),
                    json!([font]),
                )
            })
            .collect::<serde_json::Map<_, _>>();
        request["measurement"] = json!({
            "fontChains": chains,
            "defaults": { "fontSize": 11, "fontFamily": "Calibri" },
            "authoritativeShaping": true,
        });
        let request = request.to_string();
        let viewer = paged_display_list(&mut document, &request, 3);

        let envelope = parse_docx_for_edit(bytes).unwrap();
        let engine = EngineSession::new(1);
        engine.set_media(package_media(&envelope));
        seed_parsed_docx_in_place(engine.doc(), envelope).unwrap();
        engine.layout_document_with_regions_json(&request).unwrap();
        engine
            .build_display_list_frame(&json!({ "fontChains": chains }).to_string(), 0)
            .unwrap();
        let editor = engine
            .with_display_list(|list| serde_json::to_string(list).unwrap())
            .unwrap();

        assert!(viewer == editor);
        let list: Value = serde_json::from_str(&viewer).unwrap();
        let runs = list["pages"]
            .as_array()
            .unwrap()
            .iter()
            .flat_map(|page| page["primitives"].as_array().unwrap())
            .filter(|primitive| primitive["kind"] == "glyphRun")
            .map(|run| run["text"].as_str().unwrap().chars().count())
            .collect::<Vec<_>>();
        assert!(runs.iter().any(|&chars| chars > 1));
    }
}
