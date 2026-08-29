use std::collections::BTreeMap;

use ooxml_drawingml::{
    ShapeFill, ShapeOutline, Theme, preset_geometry_default_adjustments,
    resolve_color_value_to_hex_with_theme,
};
use pptx_parse::{
    GraphicFrameData, Placeholder, PptxPackage, RunProperties, ShapeNode, Slide, TextBody,
};
use serde::{Deserialize, Serialize};
use thiserror::Error;

#[derive(Clone, Copy, Debug, Default, Eq, Hash, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum EditOrigin {
    #[default]
    Local,
    Agent,
    Remote,
    System,
}

#[derive(Clone, Debug, Default, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct EditCtx {
    pub origin: EditOrigin,
    pub author: String,
}

impl EditCtx {
    pub fn local(author: impl Into<String>) -> Self {
        Self {
            origin: EditOrigin::Local,
            author: author.into(),
        }
    }
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextStyle {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub font_size_pt: Option<f64>,
    pub color: Option<String>,
    pub font_family: Option<String>,
    pub underline: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextStylePatch {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub font_size_pt: Option<f64>,
    pub color: Option<String>,
    pub font_family: Option<String>,
    pub underline: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextRunSnapshot {
    pub text: String,
    pub style: TextStyle,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ParagraphSnapshot {
    pub id: String,
    pub alignment: Option<String>,
    pub level: u32,
    pub bullet_json: Option<String>,
    pub runs: Vec<TextRunSnapshot>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct StorySnapshot {
    pub id: String,
    pub length: u32,
    pub paragraphs: Vec<ParagraphSnapshot>,
}

impl StorySnapshot {
    pub fn plain_text(&self) -> String {
        self.paragraphs
            .iter()
            .map(|paragraph| {
                paragraph
                    .runs
                    .iter()
                    .map(|run| run.text.as_str())
                    .collect::<String>()
            })
            .collect::<Vec<_>>()
            .join("\n")
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeSnapshot {
    pub id: String,
    pub source_id: u32,
    pub kind: ShapeKind,
    pub name: String,
    pub x: i64,
    pub y: i64,
    pub width: i64,
    pub height: i64,
    pub rotation_deg: f64,
    pub flip_h: bool,
    pub flip_v: bool,
    pub geometry: String,
    pub adjust_values: BTreeMap<String, f64>,
    pub placeholder: Option<Placeholder>,
    pub fill: Option<ShapeFill>,
    pub resolved_fill_color: Option<String>,
    pub outline: Option<ShapeOutline>,
    pub resolved_outline_color: Option<String>,
    pub media_part_path: Option<String>,
    pub graphic: Option<GraphicFrameData>,
    pub text_stories: Vec<StorySnapshot>,
    pub children: Vec<ShapeSnapshot>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ShapeKind {
    Shape,
    Picture,
    GraphicFrame,
    Group,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SlideSnapshot {
    pub id: String,
    pub source_part_path: Option<String>,
    pub layout_part_path: Option<String>,
    pub name: Option<String>,
    pub shapes: Vec<ShapeSnapshot>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DeckSnapshot {
    pub width_emu: i64,
    pub height_emu: i64,
    pub slides: Vec<SlideSnapshot>,
}

pub fn snapshot_package(package: &PptxPackage) -> EditResult<DeckSnapshot> {
    Ok(DeckSnapshot {
        width_emu: package.presentation.width_emu,
        height_emu: package.presentation.height_emu,
        slides: package
            .slides
            .iter()
            .enumerate()
            .map(|(slide_index, slide)| snapshot_slide(package, slide_index, slide))
            .collect::<EditResult<Vec<_>>>()?,
    })
}

fn snapshot_slide(
    package: &PptxPackage,
    slide_index: usize,
    slide: &Slide,
) -> EditResult<SlideSnapshot> {
    let reference = package
        .presentation
        .slides
        .get(slide_index)
        .ok_or_else(|| {
            EditError::InvalidState(format!("slide {slide_index} has no presentation reference"))
        })?;
    let id = format!("slide:{slide_index}:{}", reference.id);
    let theme = slide_theme(package, slide);
    Ok(SlideSnapshot {
        id: id.clone(),
        source_part_path: Some(slide.part_path.clone()),
        layout_part_path: slide.layout_part_path.clone(),
        name: slide.name.clone(),
        shapes: slide
            .shapes
            .iter()
            .enumerate()
            .map(|(shape_index, shape)| {
                snapshot_parsed_shape(&id, &shape_index.to_string(), shape, theme)
            })
            .collect(),
    })
}

fn snapshot_parsed_shape(
    slide_id: &str,
    path: &str,
    shape: &ShapeNode,
    theme: Option<&Theme>,
) -> ShapeSnapshot {
    let id = format!("{slide_id}:shape:{path}");
    let base = match shape {
        ShapeNode::Shape(shape) => &shape.base,
        ShapeNode::Picture(shape) => &shape.base,
        ShapeNode::GraphicFrame(shape) => &shape.base,
        ShapeNode::Group(shape) => &shape.base,
    };
    let (
        kind,
        geometry,
        adjust_values,
        fill,
        outline,
        media_part_path,
        graphic,
        text_stories,
        children,
    ) = match shape {
        ShapeNode::Shape(shape) => {
            let mut adjust_values = preset_geometry_default_adjustments(&shape.geometry)
                .into_iter()
                .collect::<BTreeMap<_, _>>();
            adjust_values.extend(shape.adjust_values.clone());
            let text_stories = shape
                .text
                .as_ref()
                .map(|body| vec![snapshot_text_body(&format!("story:{id}:0"), body, theme)])
                .unwrap_or_default();
            (
                ShapeKind::Shape,
                shape.geometry.clone(),
                adjust_values,
                shape.fill.clone(),
                shape.outline.clone(),
                None,
                None,
                text_stories,
                Vec::new(),
            )
        }
        ShapeNode::Picture(picture) => (
            ShapeKind::Picture,
            "rect".to_owned(),
            BTreeMap::new(),
            picture.fill.clone(),
            picture.outline.clone(),
            picture.media_part_path.clone(),
            None,
            Vec::new(),
            Vec::new(),
        ),
        ShapeNode::GraphicFrame(frame) => {
            let mut text_stories = Vec::new();
            if let GraphicFrameData::Table { rows } = &frame.data {
                for (row_index, row) in rows.iter().enumerate() {
                    for (cell_index, body) in row.iter().enumerate() {
                        text_stories.push(snapshot_text_body(
                            &format!("story:{id}:table:{row_index}:{cell_index}"),
                            body,
                            theme,
                        ));
                    }
                }
            }
            (
                ShapeKind::GraphicFrame,
                "rect".to_owned(),
                BTreeMap::new(),
                None,
                None,
                None,
                Some(frame.data.clone()),
                text_stories,
                Vec::new(),
            )
        }
        ShapeNode::Group(group) => (
            ShapeKind::Group,
            "group".to_owned(),
            BTreeMap::new(),
            None,
            None,
            None,
            None,
            Vec::new(),
            group
                .children
                .iter()
                .enumerate()
                .map(|(child_index, child)| {
                    snapshot_parsed_shape(slide_id, &format!("{path}.{child_index}"), child, theme)
                })
                .collect(),
        ),
    };
    let resolved_fill_color = fill
        .as_ref()
        .filter(|fill| fill.fill_type != "none")
        .and_then(|fill| resolve_color_value_to_hex_with_theme(fill.color.as_ref(), theme));
    let resolved_outline_color = outline
        .as_ref()
        .and_then(|outline| resolve_color_value_to_hex_with_theme(outline.color.as_ref(), theme));
    ShapeSnapshot {
        id,
        source_id: base.id,
        kind,
        name: base.name.clone(),
        x: base.transform.x,
        y: base.transform.y,
        width: base.transform.width,
        height: base.transform.height,
        rotation_deg: base.transform.rotation_deg,
        flip_h: base.transform.flip_h,
        flip_v: base.transform.flip_v,
        geometry,
        adjust_values,
        placeholder: base.placeholder.clone(),
        fill,
        resolved_fill_color,
        outline,
        resolved_outline_color,
        media_part_path,
        graphic,
        text_stories,
        children,
    }
}

fn snapshot_text_body(story_id: &str, body: &TextBody, theme: Option<&Theme>) -> StorySnapshot {
    let paragraphs = if body.paragraphs.is_empty() {
        vec![ParagraphSnapshot {
            id: format!("para:{story_id}:0"),
            alignment: None,
            level: 0,
            bullet_json: None,
            runs: Vec::new(),
        }]
    } else {
        body.paragraphs
            .iter()
            .enumerate()
            .map(|(paragraph_index, paragraph)| ParagraphSnapshot {
                id: format!("para:{story_id}:{paragraph_index}"),
                alignment: paragraph.properties.alignment.clone(),
                level: paragraph.properties.level,
                bullet_json: paragraph
                    .properties
                    .bullet
                    .as_ref()
                    .and_then(|bullet| serde_json::to_string(bullet).ok()),
                runs: paragraph
                    .runs
                    .iter()
                    .filter(|run| !run.text.is_empty())
                    .map(|run| TextRunSnapshot {
                        text: run.text.clone(),
                        style: style_from_run_properties(&run.properties, theme),
                    })
                    .collect(),
            })
            .collect()
    };
    let text_length = paragraphs
        .iter()
        .flat_map(|paragraph| &paragraph.runs)
        .map(|run| run.text.encode_utf16().count() as u32)
        .sum::<u32>();
    StorySnapshot {
        id: story_id.to_owned(),
        length: text_length.saturating_add(paragraphs.len() as u32),
        paragraphs,
    }
}

fn style_from_run_properties(properties: &RunProperties, theme: Option<&Theme>) -> TextStyle {
    TextStyle {
        bold: properties.bold,
        italic: properties.italic,
        font_size_pt: properties.font_size_pt,
        color: resolve_color_value_to_hex_with_theme(properties.color.as_ref(), theme),
        font_family: properties.font_family.clone(),
        underline: properties.underline.clone(),
    }
}

fn slide_theme<'a>(package: &'a PptxPackage, slide: &Slide) -> Option<&'a Theme> {
    let layout = slide
        .layout_part_path
        .as_deref()
        .and_then(|path| {
            package
                .layouts
                .iter()
                .find(|layout| layout.part_path == path)
        })
        .or_else(|| package.layouts.first());
    let master = layout
        .and_then(|layout| layout.master_part_path.as_deref())
        .and_then(|path| {
            package
                .masters
                .iter()
                .find(|master| master.part_path == path)
        })
        .or_else(|| {
            layout.and_then(|layout| {
                package.masters.iter().find(|master| {
                    master
                        .layout_part_paths
                        .iter()
                        .any(|path| path == &layout.part_path)
                })
            })
        })
        .or_else(|| package.masters.first());
    master
        .and_then(|master| master.theme_part_path.as_deref())
        .and_then(|path| package.themes.iter().find(|theme| theme.part_path == path))
        .map(|part| &part.theme)
        .or_else(|| package.themes.first().map(|part| &part.theme))
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SlideReceipt {
    pub slide_id: String,
    pub from_index: Option<u32>,
    pub to_index: Option<u32>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeReceipt {
    pub slide_id: String,
    pub shape_id: String,
    pub index: u32,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TransformReceipt {
    pub slide_id: String,
    pub shape_id: String,
    pub before: ShapeRect,
    pub after: ShapeRect,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct ShapeRect {
    pub x: i64,
    pub y: i64,
    pub width: i64,
    pub height: i64,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextReceipt {
    pub story_id: String,
    pub start: u32,
    pub end: u32,
    pub text: String,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeDraft {
    pub name: String,
    pub rect: ShapeRect,
    pub text: String,
    pub style: TextStyle,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PresetShapeDraft {
    pub name: String,
    pub geometry: String,
    pub rect: ShapeRect,
    pub fill: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeStroke {
    pub color: Option<String>,
    pub width_pt: Option<f64>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeFillReceipt {
    pub slide_id: String,
    pub shape_id: String,
    pub before: Option<String>,
    pub after: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeStrokeReceipt {
    pub slide_id: String,
    pub shape_id: String,
    pub before: Option<ShapeStroke>,
    pub after: Option<ShapeStroke>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeAdjustReceipt {
    pub slide_id: String,
    pub shape_id: String,
    pub before: BTreeMap<String, f64>,
    pub after: BTreeMap<String, f64>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum UpdateOrigin {
    Local,
    Remote,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct UpdateEvent {
    pub update: Vec<u8>,
    pub origin: UpdateOrigin,
}

#[derive(Debug, Error)]
pub enum EditError {
    #[error("invalid client ID {0}")]
    InvalidClientId(u64),
    #[error("could not parse PPTX: {0}")]
    Parse(String),
    #[error("invalid deck state: {0}")]
    InvalidState(String),
    #[error("invalid yrs update: {0}")]
    InvalidUpdate(String),
    #[error("invalid yrs state vector: {0}")]
    InvalidStateVector(String),
    #[error("slide {0:?} was not found")]
    SlideNotFound(String),
    #[error("shape {0:?} was not found")]
    ShapeNotFound(String),
    #[error("story {0:?} was not found")]
    StoryNotFound(String),
    #[error("index {index} is outside length {length}")]
    OutOfBounds { index: u32, length: u32 },
    #[error("text range {start}..{end} crosses a paragraph boundary")]
    ParagraphBoundary { start: u32, end: u32 },
    #[error("invalid shape geometry: {0}")]
    InvalidGeometry(String),
    #[error("invalid shape adjustment: {0}")]
    InvalidAdjustment(String),
    #[error("invalid text: {0}")]
    InvalidText(String),
    #[error("update observer failed: {0}")]
    Observer(String),
    #[error("JSON boundary error: {0}")]
    Json(String),
    #[error("could not write PPTX: {0}")]
    Write(String),
}

pub type EditResult<T> = Result<T, EditError>;

/// Rejects characters XML 1.0 cannot carry, so a bad edit fails loudly
/// instead of producing an unopenable file at save time.
pub(crate) fn validate_xml_text(value: &str) -> EditResult<()> {
    match value
        .chars()
        .find(|character| !legal_xml_character(*character))
    {
        Some(character) => Err(EditError::InvalidText(format!(
            "character U+{:04X} cannot be stored in a PPTX file",
            character as u32
        ))),
        None => Ok(()),
    }
}

fn legal_xml_character(character: char) -> bool {
    matches!(character, '\u{9}' | '\u{a}' | '\u{d}')
        || ('\u{20}'..='\u{d7ff}').contains(&character)
        || ('\u{e000}'..='\u{fffd}').contains(&character)
        || ('\u{10000}'..='\u{10ffff}').contains(&character)
}
