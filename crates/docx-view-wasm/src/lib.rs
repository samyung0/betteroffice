// Linear memory grows in large steps (crates/wasm-alloc).
#[cfg(target_arch = "wasm32")]
use wasm_alloc as _;

use docx_edit::{EngineSession, package_media, parse_docx_for_edit, seed_parsed_docx_in_place};
use docx_parse::S9PackageWire;
use docx_parse::section::SectionProperties;
use serde_json::{Value, json};
use wasm_bindgen::prelude::*;

const DEFAULT_PAGE_GAP: f64 = 24.0;

/// Opens in steps so the host registers the fonts the layout needs first:
/// `open`, `layoutRequestJson`, `fontRequirementsJson`, `layout`, then
/// `displayPagesJson` a batch of pages at a time.
#[wasm_bindgen]
pub struct DocxViewDocument {
    engine: EngineSession,
    request: String,
    /// What the display list is built with once laid out.
    extras: Option<String>,
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
            extras: None,
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
        let extras = match serde_json::from_str::<Value>(request)
            .map_err(js_error)?
            .pointer("/measurement/fontChains")
        {
            Some(chains) => json!({ "fontChains": chains }),
            None => json!({}),
        };
        let pages = self
            .engine
            .layout_document_with_regions(request)
            .map_err(js_error)?;
        self.extras = Some(
            self.engine
                .region_display_extras(&extras.to_string())
                .map_err(js_error)?,
        );
        u32::try_from(pages).map_err(js_error)
    }

    /// `{ pages }` display list JSON of pages `start..end`, built from the
    /// engine's typed state; the host frees the document after the last one.
    #[wasm_bindgen(js_name = displayPagesJson)]
    pub fn display_pages_json(&self, start: u32, end: u32) -> Result<String, JsValue> {
        let extras = self
            .extras
            .as_deref()
            .ok_or_else(|| js_error("DOCX view is not laid out"))?;
        self.engine
            .display_pages_json(extras, start as usize..end as usize)
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
    if let Some(id) = default_paragraph_style_id {
        render_env["defaultParagraphStyleId"] = json!(id);
    }
    json!({
        "bodyStory": "body",
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
    /// wrote. The round trip parses floats without `float_roundtrip`, so in
    /// two fixtures a few numbers it read back sit up to 4 ULP from the
    /// engine's own values, which the paged build keeps (and the editor
    /// paints); every other fixture is byte-identical.
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
