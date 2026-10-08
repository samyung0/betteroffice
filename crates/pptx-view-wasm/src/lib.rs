// Linear memory grows in large steps (crates/wasm-alloc).
#[cfg(target_arch = "wasm32")]
use wasm_alloc as _;

use pptx_edit::{DeckSnapshot, baseline_snapshot};
use wasm_bindgen::prelude::*;

#[wasm_bindgen]
pub struct PptxViewDocument {
    package: pptx_parse::PptxPackage,
    snapshot: DeckSnapshot,
}

#[wasm_bindgen]
impl PptxViewDocument {
    pub fn open(bytes: &[u8]) -> Result<PptxViewDocument, JsValue> {
        let package = pptx_parse::parse_pptx(bytes).map_err(js_error)?;
        let snapshot = baseline_snapshot(&package).map_err(js_error)?;
        Ok(Self { package, snapshot })
    }

    #[wasm_bindgen(js_name = snapshotJson)]
    pub fn snapshot_json(&self) -> Result<String, JsValue> {
        serde_json::to_string(&self.snapshot).map_err(js_error)
    }

    #[wasm_bindgen(js_name = mediaBytes)]
    pub fn media_bytes(&self, part_path: &str) -> Result<Vec<u8>, JsValue> {
        self.package
            .media
            .iter()
            .find(|media| media.part_path == part_path)
            .map(|media| media.bytes.clone())
            .ok_or_else(|| js_error(format!("media part {part_path} was not found")))
    }

    pub fn version() -> String {
        env!("CARGO_PKG_VERSION").to_owned()
    }
}

#[wasm_bindgen]
pub struct PptxViewRenderer {
    renderer: pptx_render::SlideRenderer,
    rendered: Option<pptx_render::RenderedSlide>,
}

#[wasm_bindgen]
impl PptxViewRenderer {
    #[wasm_bindgen(constructor)]
    pub fn new() -> PptxViewRenderer {
        Self {
            renderer: pptx_render::SlideRenderer::new(),
            rendered: None,
        }
    }

    #[wasm_bindgen(js_name = registerFont)]
    pub fn register_font(
        &mut self,
        family: &str,
        bold: bool,
        italic: bool,
        bytes: &[u8],
    ) -> Result<u32, JsValue> {
        self.renderer
            .register_font(family, bold, italic, bytes)
            .map_err(js_error)
    }

    /// Registers the deck's embedded faces, one per part, as
    /// `[{"family" (its alias), "typeface", "bold", "italic", "fontId"}]`.
    #[wasm_bindgen(js_name = registerEmbeddedFontsJson)]
    pub fn register_embedded_fonts_json(
        &mut self,
        document: &PptxViewDocument,
    ) -> Result<String, JsValue> {
        let faces = self.renderer.register_embedded_fonts(&document.package);
        serde_json::to_string(&faces).map_err(js_error)
    }

    /// Drops an embedded face the page could not load from layout.
    #[wasm_bindgen(js_name = refuseEmbeddedFont)]
    pub fn refuse_embedded_font(&mut self, font_id: u32) {
        self.renderer.refuse_embedded_font(font_id);
    }

    #[wasm_bindgen(js_name = fontBytes)]
    pub fn font_bytes(&self, font_id: u32) -> Result<Vec<u8>, JsValue> {
        self.renderer
            .font_bytes(font_id)
            .map(<[u8]>::to_vec)
            .ok_or_else(|| js_error(format!("font {font_id} is not registered")))
    }

    #[wasm_bindgen(js_name = layoutSlideJson)]
    pub fn layout_slide_json(
        &mut self,
        document: &PptxViewDocument,
        slide_index: u32,
    ) -> Result<String, JsValue> {
        let rendered = self
            .renderer
            .layout_slide(&document.package, &document.snapshot, slide_index as usize)
            .map_err(js_error)?;
        let json = serde_json::to_string(&rendered.display_list).map_err(js_error)?;
        self.rendered = Some(rendered);
        Ok(json)
    }

    #[wasm_bindgen(js_name = hitTestJson)]
    pub fn hit_test_json(&self, x: f32, y: f32) -> Result<String, JsValue> {
        let result = self
            .rendered
            .as_ref()
            .and_then(|rendered| rendered.hit_test(x, y));
        serde_json::to_string(&result).map_err(js_error)
    }
}

impl Default for PptxViewRenderer {
    fn default() -> Self {
        Self::new()
    }
}

fn js_error(error: impl std::fmt::Display) -> JsValue {
    JsValue::from_str(&error.to_string())
}
