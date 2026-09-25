use docx_edit::{EngineSession, parse_docx_for_edit, seed_parsed_docx};
use docx_parse::S9PackageWire;
use docx_parse::section::SectionProperties;
use serde_json::{Value, json};
use wasm_bindgen::prelude::*;

const DEFAULT_PAGE_GAP: f64 = 24.0;

/// Opens in steps so the host registers the fonts the layout needs first:
/// `open`, `layoutRequestJson`, `fontRequirementsJson`, then `layout`.
#[wasm_bindgen]
pub struct DocxViewDocument {
    engine: Option<EngineSession>,
    request: String,
    display_list: Option<String>,
}

#[wasm_bindgen]
impl DocxViewDocument {
    pub fn open(bytes: &[u8]) -> Result<DocxViewDocument, JsValue> {
        let envelope = parse_docx_for_edit(bytes).map_err(js_error)?;
        let request = layout_request(&envelope.document.package).to_string();
        let engine = EngineSession::new(1);
        seed_parsed_docx(engine.doc(), envelope).map_err(js_error)?;
        Ok(Self {
            engine: Some(engine),
            request,
            display_list: None,
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
            .as_ref()
            .ok_or_else(|| js_error("DOCX view is already laid out"))?
            .layout_font_requirements_json(request)
            .map_err(js_error)
    }

    /// Lays out `request` (with its `measurement`) and keeps only the display
    /// list, so view mode holds no second document graph.
    pub fn layout(&mut self, request: &str) -> Result<(), JsValue> {
        let engine = self
            .engine
            .take()
            .ok_or_else(|| js_error("DOCX view is already laid out"))?;
        let layout = engine
            .layout_document_with_regions_json(request)
            .map_err(js_error)?;
        self.display_list = Some(engine.build_display_list_json(&layout).map_err(js_error)?);
        Ok(())
    }

    #[wasm_bindgen(js_name = displayListJson)]
    pub fn display_list_json(&self, page_gap: Option<f64>) -> Result<String, JsValue> {
        if page_gap.is_some_and(|gap| (gap - DEFAULT_PAGE_GAP).abs() > f64::EPSILON) {
            return Err(JsValue::from_str(
                "DOCX view page gap is fixed when the document opens",
            ));
        }
        self.display_list
            .clone()
            .ok_or_else(|| js_error("DOCX view is not laid out"))
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
        .map(|note| json!({ "id": note.id, "noteKind": "footnote", "height": 0 }))
        .chain(
            package
                .endnotes
                .as_deref()
                .unwrap_or_default()
                .iter()
                .filter(|note| note.note_type.is_empty() || note.note_type == "normal")
                .map(|note| json!({ "id": note.id, "noteKind": "endnote", "height": 0 })),
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
