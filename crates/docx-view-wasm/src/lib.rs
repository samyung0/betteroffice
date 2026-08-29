use docx_edit::{EngineSession, parse_docx_for_edit, seed_parsed_docx};
use docx_parse::S9PackageWire;
use serde_json::{Value, json};
use wasm_bindgen::prelude::*;

const DEFAULT_PAGE_GAP: f64 = 24.0;

#[wasm_bindgen]
pub struct DocxViewDocument {
    display_list: String,
}

#[wasm_bindgen]
impl DocxViewDocument {
    pub fn open(bytes: &[u8]) -> Result<DocxViewDocument, JsValue> {
        let envelope = parse_docx_for_edit(bytes).map_err(js_error)?;
        let request = layout_request(&envelope.document.package);
        let engine = EngineSession::new(1);
        seed_parsed_docx(engine.doc(), envelope).map_err(js_error)?;
        let layout = engine
            .layout_document_with_regions_json(&request.to_string())
            .map_err(js_error)?;
        let display_list = engine.build_display_list_json(&layout).map_err(js_error)?;
        // The editing/Yrs projection is currently shared code used while
        // lowering OOXML. It is deliberately not retained by the viewer: only
        // the immutable display list survives `open`, so view mode does not
        // keep a second editable document graph resident.
        Ok(Self { display_list })
    }

    #[wasm_bindgen(js_name = displayListJson)]
    pub fn display_list_json(&self, page_gap: Option<f64>) -> Result<String, JsValue> {
        if page_gap.is_some_and(|gap| (gap - DEFAULT_PAGE_GAP).abs() > f64::EPSILON) {
            return Err(JsValue::from_str(
                "DOCX view page gap is fixed when the document opens",
            ));
        }
        Ok(self.display_list.clone())
    }

    pub fn version() -> String {
        env!("CARGO_PKG_VERSION").to_owned()
    }
}

fn layout_request(package: &S9PackageWire) -> Value {
    let mut sections = package
        .document
        .sections
        .as_deref()
        .unwrap_or_default()
        .iter()
        .map(|section| {
            json!({
                "sectionId": section.id,
                "properties": section.properties,
            })
        })
        .collect::<Vec<_>>();
    sections.push(json!({
        "properties": package.document.final_section_properties.clone().unwrap_or_default(),
    }));

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

    let theme_colors = serde_json::to_value(&package.theme.color_scheme).unwrap_or_default();
    let settings = serde_json::to_value(&package.settings).unwrap_or_default();
    let default_tab_stop = settings
        .get("defaultTabStop")
        .cloned()
        .unwrap_or(Value::Null);
    json!({
        "bodyStory": "body",
        "options": { "pageGap": DEFAULT_PAGE_GAP },
        "regions": {
            "sections": sections,
            "settings": settings,
        },
        "notes": { "contents": notes },
        "renderEnv": {
            "themeColors": theme_colors,
            "defaultTabStopTwips": default_tab_stop,
            "numericIds": {},
        },
        "measurement": {
            "fontChains": {},
            "defaults": { "fontSize": 11, "fontFamily": "Calibri" },
            "compat": { "noLeading": false, "doNotExpandShiftReturn": false },
            "authoritativeShaping": true,
        },
    })
}

fn js_error(error: impl std::fmt::Display) -> JsValue {
    JsValue::from_str(&error.to_string())
}
