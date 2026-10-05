//! Paragraphs and tracked inline structures.

use serde::{Deserialize, Serialize};

use crate::block::transparent_children;
use crate::chart::{ChartPartsMap, DrawingChart, parse_chart_from_drawing};
use crate::formatting::{
    ParagraphFormatting, ParagraphFrame, SpacingExplicit, parse_paragraph_properties,
    unmodeled_paragraph_children,
};
use crate::image::{is_text_box_drawing, parse_drawing};
use crate::inline::{
    ContentPosition, Hyperlink, InlineNode, InlineSdt, InlineSdtType, MathEquation, MathType,
    OpenComplexField, RawInlineXml, RawInlineXmlType, Run, RunContent, SimpleField,
    SimpleFieldType, StructuredFieldContent, StructuredFieldTree, parse_bookmark_end,
    parse_bookmark_start, parse_field_type, parse_hyperlink, parse_run, parse_sdt_properties,
};
use crate::media::MediaMap;
use crate::numbering::{ListRendering, NumberingMap, compute_list_rendering};
use crate::relationships::RelationshipMap;
use crate::section::{SectionProperties, parse_section_properties};
use crate::shape::{
    Shape, is_shape_drawing, parse_shape_from_drawing, resolve_shape_fill_pictures,
};
use crate::smart_art::{SmartArtContext, is_smart_art_drawing, parse_smart_art_from_drawing};
use crate::styles::{DocDefaults, StyleMap};
use crate::theme::Theme;
use crate::vml::{parse_horizontal_rule, parse_vml_image_content};
use crate::xml::{
    ParseBudget, ParseError, ParseLimits, XmlElement, XmlNode, parse_javascript_integer_prefix,
    parse_xml_strict,
};

const MAX_FIELD_NESTING: usize = 32;
const MAX_HEX_ID_EXCLUSIVE: u32 = 0x7fff_ffff;

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct TrackedChangeInfo {
    pub id: f64,
    pub author: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub date: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ParagraphPropertyChange {
    #[serde(rename = "type")]
    pub node_type: String,
    pub info: TrackedChangeInfo,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub previous_formatting: Option<ParagraphFormatting>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub current_formatting: Option<ParagraphFormatting>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct TrackedInline {
    #[serde(rename = "type")]
    pub node_type: String,
    pub info: TrackedChangeInfo,
    pub content: Vec<InlineNode>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct RangeStart {
    #[serde(rename = "type")]
    pub node_type: String,
    pub id: f64,
    pub name: String,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct RangeEnd {
    #[serde(rename = "type")]
    pub node_type: String,
    pub id: f64,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct CommentRange {
    #[serde(rename = "type")]
    pub node_type: String,
    pub id: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub offset: Option<f64>,
}

#[derive(Clone, Debug, PartialEq, Serialize)]
#[serde(untagged)]
pub enum ParagraphContent {
    Inline(InlineNode),
    Tracked(TrackedInline),
    RangeStart(RangeStart),
    RangeEnd(RangeEnd),
    CommentRange(CommentRange),
}

impl<'de> Deserialize<'de> for ParagraphContent {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        let value = serde_json::Value::deserialize(deserializer)?;
        let node_type = value
            .get("type")
            .and_then(serde_json::Value::as_str)
            .ok_or_else(|| serde::de::Error::custom("paragraph content requires a string type"))?;
        let decoded = match node_type {
            "insertion" | "deletion" | "moveFrom" | "moveTo" => {
                serde_json::from_value(value).map(Self::Tracked)
            }
            "moveFromRangeStart" | "moveToRangeStart" => {
                serde_json::from_value(value).map(Self::RangeStart)
            }
            "moveFromRangeEnd" | "moveToRangeEnd" => {
                serde_json::from_value(value).map(Self::RangeEnd)
            }
            "commentRangeStart" | "commentRangeEnd" => {
                serde_json::from_value(value).map(Self::CommentRange)
            }
            _ => serde_json::from_value(value).map(Self::Inline),
        };
        decoded.map_err(serde::de::Error::custom)
    }
}

impl ParagraphContent {
    pub fn node_type(&self) -> &str {
        match self {
            Self::Inline(node) => node.node_type(),
            Self::Tracked(node) => &node.node_type,
            Self::RangeStart(node) => &node.node_type,
            Self::RangeEnd(node) => &node.node_type,
            Self::CommentRange(node) => &node.node_type,
        }
    }
}

/// One attribute kept exactly as authored, prefixed name and all.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct RawAttribute {
    pub name: String,
    pub value: String,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Paragraph {
    #[serde(rename = "type")]
    pub node_type: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub para_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_id: Option<String>,
    /// Attributes the model does not type, kept in authored form so a save
    /// re-emits them; dropping an attribute the parser never read is a loss.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub extra_attributes: Vec<RawAttribute>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formatting: Option<ParagraphFormatting>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub property_changes: Option<Vec<ParagraphPropertyChange>>,
    #[serde(rename = "pPrIns", skip_serializing_if = "Option::is_none")]
    pub p_pr_ins: Option<TrackedChangeInfo>,
    #[serde(rename = "pPrDel", skip_serializing_if = "Option::is_none")]
    pub p_pr_del: Option<TrackedChangeInfo>,
    pub content: Vec<ParagraphContent>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub list_rendering: Option<ListRendering>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rendered_page_break_before: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub section_properties: Option<SectionProperties>,
}

#[derive(Clone, Debug)]
pub struct HexIdAllocator {
    state: u32,
}

pub struct DrawingContext<'a> {
    pub media: &'a MediaMap,
    pub charts: &'a ChartPartsMap,
    pub smart_art: &'a mut SmartArtContext,
}

impl HexIdAllocator {
    pub fn from_sha256(seed: &str) -> Result<Self, ParseError> {
        if seed.len() != 64 || !seed.bytes().all(|byte| byte.is_ascii_hexdigit()) {
            return Err(ParseError::Canonical(
                "canonical determinism seed must be a SHA-256 hex digest".to_owned(),
            ));
        }
        let mut state = 0u32;
        for offset in (0..64).step_by(8) {
            let word = u32::from_str_radix(&seed[offset..offset + 8], 16)
                .map_err(|error| ParseError::Canonical(error.to_string()))?;
            state ^= word;
        }
        if state == 0 {
            state = 0x6d2b_79f5;
        }
        Ok(Self { state })
    }

    pub fn allocate(&mut self) -> String {
        self.state ^= self.state << 13;
        self.state ^= self.state >> 17;
        self.state ^= self.state << 5;
        format!("{:08X}", self.state % MAX_HEX_ID_EXCLUSIVE)
    }
}

/// Everything on `w:p` the model does not consume itself.
fn extra_paragraph_attributes(element: &XmlElement) -> Vec<RawAttribute> {
    const CONSUMED: [&str; 4] = ["w14:paraId", "w:paraId", "w14:textId", "w:textId"];
    element
        .attributes
        .iter()
        .filter(|(name, _)| !CONSUMED.contains(&name.as_str()))
        .map(|(name, value)| RawAttribute {
            name: name.clone(),
            value: value.clone(),
        })
        .collect()
}

#[allow(clippy::too_many_arguments)]
pub fn parse_paragraph(
    element: &XmlElement,
    relationships: Option<&RelationshipMap>,
    theme: Option<&Theme>,
    styles: Option<&StyleMap>,
    doc_defaults: Option<&DocDefaults>,
    numbering: Option<&NumberingMap>,
    part: &str,
    budget: &mut ParseBudget<'_>,
    ids: &mut HexIdAllocator,
    drawing: Option<&mut DrawingContext<'_>>,
    in_header_footer: bool,
    depth: usize,
) -> Result<Paragraph, ParseError> {
    budget.check_nesting_depth(depth, part)?;
    budget.charge_paragraph(part)?;
    let mut paragraph = Paragraph {
        node_type: "paragraph".to_owned(),
        para_id: normalize_hex_id(
            element
                .attribute(Some("w14"), "paraId")
                .or_else(|| element.attribute(Some("w"), "paraId")),
            ids,
        ),
        text_id: normalize_hex_id(
            element
                .attribute(Some("w14"), "textId")
                .or_else(|| element.attribute(Some("w"), "textId")),
            ids,
        ),
        extra_attributes: extra_paragraph_attributes(element),
        formatting: None,
        property_changes: None,
        p_pr_ins: None,
        p_pr_del: None,
        content: Vec::new(),
        list_rendering: None,
        rendered_page_break_before: (!in_header_footer
            && paragraph_starts_with_rendered_page_break(element))
        .then_some(true),
        section_properties: None,
    };
    let properties = element.child("w", "pPr");
    if let Some(properties) = properties {
        paragraph.formatting = parse_document_paragraph_properties(properties, theme, styles);
        paragraph.property_changes = parse_paragraph_property_changes(
            properties,
            theme,
            styles,
            paragraph.formatting.as_ref(),
        );
        if let Some(mark) = properties.child("w", "rPr") {
            paragraph.p_pr_ins = mark.child("w", "ins").and_then(parse_paragraph_mark_change);
            paragraph.p_pr_del = mark.child("w", "del").and_then(parse_paragraph_mark_change);
        }
        paragraph.section_properties = properties
            .child("w", "sectPr")
            .map(|section| parse_section_properties(Some(section)));
    }
    paragraph.content = parse_paragraph_contents(
        element,
        relationships,
        theme,
        styles,
        doc_defaults,
        part,
        budget,
        drawing,
        depth,
        false,
    )?;
    apply_list_rendering(&mut paragraph, properties, styles, numbering);
    Ok(paragraph)
}

pub fn parse_document_paragraph_properties(
    properties: &XmlElement,
    theme: Option<&Theme>,
    styles: Option<&StyleMap>,
) -> Option<ParagraphFormatting> {
    let mut value = parse_paragraph_properties(Some(properties), theme).unwrap_or_default();
    if let Some(spacing) = properties.child("w", "spacing") {
        let before = spacing
            .parse_numeric_attribute(Some("w"), "before", 1.0)
            .or(value.space_before_lines)
            .map(|_| true);
        let after = spacing
            .parse_numeric_attribute(Some("w"), "after", 1.0)
            .or(value.space_after_lines)
            .map(|_| true);
        if before.is_some() || after.is_some() {
            value.spacing_explicit = Some(SpacingExplicit { before, after });
        }
    }
    if let Some(indent) = properties.child("w", "ind") {
        if value.indent_left.is_none() {
            value.indent_left = indent.parse_numeric_attribute(Some("w"), "start", 1.0);
        }
        if value.indent_right.is_none() {
            value.indent_right = indent.parse_numeric_attribute(Some("w"), "end", 1.0);
        }
    }
    if let Some(frame) = properties.child("w", "framePr") {
        value.frame = parse_frame(frame);
    }
    value.extra_children = unmodeled_paragraph_children(properties);
    if let Some(run_properties) = properties.child("w", "rPr") {
        let shell = XmlElement {
            name: "w:r".to_owned(),
            attributes: Default::default(),
            children: vec![XmlNode::Element(run_properties.clone())],
        };
        value.run_properties = parse_run(&shell, theme, styles, None).run.formatting;
    }
    let creates_empty_bag = properties.child("w", "shd").is_some()
        || properties.child("w", "tabs").is_some()
        || properties.child("w", "framePr").is_some()
        || properties.child("w", "rPr").is_some();
    (value != ParagraphFormatting::default() || creates_empty_bag).then_some(value)
}

fn parse_frame(element: &XmlElement) -> Option<ParagraphFrame> {
    let value = ParagraphFrame {
        width: element.parse_numeric_attribute(Some("w"), "w", 1.0),
        height: element.parse_numeric_attribute(Some("w"), "h", 1.0),
        h_anchor: enum_attribute(element, "hAnchor", &["text", "margin", "page"]),
        v_anchor: enum_attribute(element, "vAnchor", &["text", "margin", "page"]),
        x: element.parse_numeric_attribute(Some("w"), "x", 1.0),
        y: element.parse_numeric_attribute(Some("w"), "y", 1.0),
        x_align: nonempty_attribute(element, "xAlign"),
        y_align: nonempty_attribute(element, "yAlign"),
        wrap: nonempty_attribute(element, "wrap"),
        drop_cap: nonempty_attribute(element, "dropCap"),
        lines: element.parse_numeric_attribute(Some("w"), "lines", 1.0),
        h_space: element.parse_numeric_attribute(Some("w"), "hSpace", 1.0),
        v_space: element.parse_numeric_attribute(Some("w"), "vSpace", 1.0),
        h_rule: nonempty_attribute(element, "hRule"),
        anchor_lock: nonempty_attribute(element, "anchorLock"),
    };
    (value != ParagraphFrame::default()).then_some(value)
}

fn parse_paragraph_property_changes(
    properties: &XmlElement,
    theme: Option<&Theme>,
    styles: Option<&StyleMap>,
    current: Option<&ParagraphFormatting>,
) -> Option<Vec<ParagraphPropertyChange>> {
    let changes: Vec<_> = properties
        .children_named("w", "pPrChange")
        .filter_map(|change| {
            let previous = change.child("w", "pPr").and_then(|properties| {
                parse_document_paragraph_properties(properties, theme, styles)
            });
            let current = current.cloned();
            (previous.is_some() || current.is_some()).then(|| ParagraphPropertyChange {
                node_type: "paragraphPropertyChange".to_owned(),
                info: parse_tracked_change_info(change),
                previous_formatting: previous,
                current_formatting: current,
            })
        })
        .collect();
    (!changes.is_empty()).then_some(changes)
}

fn parse_paragraph_mark_change(element: &XmlElement) -> Option<TrackedChangeInfo> {
    let id = parse_javascript_integer_prefix(element.attribute(Some("w"), "id")?)?;
    Some(TrackedChangeInfo {
        id,
        author: element
            .attribute(Some("w"), "author")
            .unwrap_or_default()
            .to_owned(),
        date: element
            .attribute(Some("w"), "date")
            .filter(|date| !date.is_empty())
            .map(str::to_owned),
    })
}

fn parse_tracked_change_info(element: &XmlElement) -> TrackedChangeInfo {
    let parsed = element
        .attribute(Some("w"), "id")
        .and_then(parse_javascript_integer_prefix)
        .filter(|id| id.fract() == 0.0 && *id >= 0.0)
        .unwrap_or(0.0);
    let author = element
        .attribute(Some("w"), "author")
        .unwrap_or_default()
        .trim();
    let date = element
        .attribute(Some("w"), "date")
        .unwrap_or_default()
        .trim();
    TrackedChangeInfo {
        id: parsed,
        author: if author.is_empty() { "Unknown" } else { author }.to_owned(),
        date: (!date.is_empty()).then(|| date.to_owned()),
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum TrackedContext {
    Default,
    Deletion,
}

#[allow(clippy::too_many_arguments)]
fn parse_paragraph_contents(
    element: &XmlElement,
    relationships: Option<&RelationshipMap>,
    theme: Option<&Theme>,
    styles: Option<&StyleMap>,
    doc_defaults: Option<&DocDefaults>,
    part: &str,
    budget: &mut ParseBudget<'_>,
    mut drawing: Option<&mut DrawingContext<'_>>,
    depth: usize,
    deletion: bool,
) -> Result<Vec<ParagraphContent>, ParseError> {
    budget.check_nesting_depth(depth, part)?;
    let tracked_context = if deletion {
        TrackedContext::Deletion
    } else {
        TrackedContext::Default
    };
    let mut output = Vec::new();
    let mut fields: Vec<OpenComplexField> = Vec::new();
    // Range ends met inside a complex field go after it, as starts go before it.
    let mut field_ends = Vec::new();
    for child in transparent_children(element, false) {
        // Deleted text reads as text, in runs and in the fields and links a
        // deletion holds.
        let normalized;
        let child = if tracked_context == TrackedContext::Deletion {
            normalized = normalize_deletion_element(child);
            &normalized
        } else {
            child
        };
        match child.local_name() {
            "r" => {
                let run = parse_run_composed(
                    child,
                    relationships,
                    theme,
                    styles,
                    doc_defaults,
                    budget,
                    drawing.as_deref_mut(),
                )?;
                process_field_run(run, &mut fields, &mut output, part)?;
                if fields.is_empty() {
                    output.append(&mut field_ends);
                }
            }
            "hyperlink" => {
                let hyperlink = parse_hyperlink_composed(
                    child,
                    relationships,
                    theme,
                    styles,
                    doc_defaults,
                    part,
                    budget,
                    drawing.as_deref_mut(),
                    depth + 1,
                )?;
                let (starts, ends) = nested_comment_ranges(child);
                output.extend(starts);
                if let Some(active) = fields.last_mut() {
                    let runs: Vec<Run> = hyperlink
                        .children
                        .iter()
                        .filter_map(|node| match node {
                            InlineNode::Run(run) => Some(run.clone()),
                            _ => None,
                        })
                        .collect();
                    active.absorb(InlineNode::Hyperlink(Box::new(hyperlink)), runs);
                } else {
                    output.push(ParagraphContent::Inline(InlineNode::Hyperlink(Box::new(
                        hyperlink,
                    ))));
                }
                close_ranges(ends, &fields, &mut output, &mut field_ends);
            }
            "bookmarkStart" => {
                let node = InlineNode::BookmarkStart(parse_bookmark_start(child));
                match fields.last_mut() {
                    Some(active) => active.absorb(node, Vec::new()),
                    None => output.push(ParagraphContent::Inline(node)),
                }
            }
            "bookmarkEnd" => {
                let node = InlineNode::BookmarkEnd(parse_bookmark_end(child));
                match fields.last_mut() {
                    Some(active) => active.absorb(node, Vec::new()),
                    None => output.push(ParagraphContent::Inline(node)),
                }
            }
            "fldSimple" => {
                let field = parse_simple_field_composed(
                    child,
                    relationships,
                    theme,
                    styles,
                    doc_defaults,
                    part,
                    budget,
                    drawing.as_deref_mut(),
                    depth + 1,
                )?;
                let (starts, ends) = nested_comment_ranges(child);
                output.extend(starts);
                if let Some(active) = fields.last_mut() {
                    let runs = field.content.clone();
                    active.absorb(InlineNode::SimpleField(Box::new(field)), runs);
                } else {
                    output.push(ParagraphContent::Inline(InlineNode::SimpleField(Box::new(
                        field,
                    ))));
                }
                close_ranges(ends, &fields, &mut output, &mut field_ends);
            }
            "sdt" => {
                if let Some(container) = child.child("w", "sdtContent") {
                    let parsed = parse_paragraph_contents(
                        container,
                        relationships,
                        theme,
                        styles,
                        doc_defaults,
                        part,
                        budget,
                        drawing.as_deref_mut(),
                        depth + 1,
                        tracked_context == TrackedContext::Deletion,
                    )?;
                    let (starts, ends) = nested_comment_ranges(container);
                    output.extend(starts);
                    let sdt = InlineNode::InlineSdt(Box::new(InlineSdt {
                        node_type: InlineSdtType::InlineSdt,
                        properties: parse_sdt_properties(child.child("w", "sdtPr"), None, theme),
                        content: filter_field_inline(parsed),
                    }));
                    match fields.last_mut().filter(|field| field.in_result()) {
                        Some(active) => {
                            let runs = shown_runs(std::slice::from_ref(&sdt));
                            active.absorb(sdt, runs);
                        }
                        None => output.push(ParagraphContent::Inline(sdt)),
                    }
                    close_ranges(ends, &fields, &mut output, &mut field_ends);
                }
            }
            "ins" | "del" | "moveFrom" | "moveTo" => {
                let is_deletion = matches!(child.local_name(), "del" | "moveFrom");
                let content = parse_paragraph_contents(
                    child,
                    relationships,
                    theme,
                    styles,
                    doc_defaults,
                    part,
                    budget,
                    drawing.as_deref_mut(),
                    depth + 1,
                    is_deletion,
                )?;
                let content = filter_field_inline(content);
                let node_type = match child.local_name() {
                    "ins" => "insertion",
                    "del" => "deletion",
                    "moveFrom" => "moveFrom",
                    _ => "moveTo",
                };
                let (starts, ends) = nested_comment_ranges(child);
                output.extend(starts);
                let change = TrackedInline {
                    node_type: node_type.to_owned(),
                    info: parse_tracked_change_info(child),
                    content,
                };
                // A field has no tracked change in its grammar: a change whose field
                // characters balance stays inside it as its markup, and what it
                // inserts shows as the field's code or result.
                let shown = shown_change_runs(&change);
                let kept = balanced_field_chars(child).then(|| raw_change(child, shown.clone()));
                match (fields.last_mut(), kept) {
                    (Some(active), Some(raw)) => {
                        if !active.in_result() {
                            active.append_instruction(&instruction_text(&shown));
                        }
                        active.absorb(raw, shown);
                    }
                    (None, Some(raw)) if element.local_name() == "fldSimple" => {
                        output.push(ParagraphContent::Inline(raw));
                    }
                    _ => output.push(ParagraphContent::Tracked(change)),
                }
                close_ranges(ends, &fields, &mut output, &mut field_ends);
            }
            "moveFromRangeStart" | "moveToRangeStart" => {
                output.push(ParagraphContent::RangeStart(RangeStart {
                    node_type: child.local_name().to_owned(),
                    id: parse_range_id(child),
                    name: child
                        .attribute(Some("w"), "name")
                        .unwrap_or_default()
                        .to_owned(),
                }));
            }
            "moveFromRangeEnd" | "moveToRangeEnd" => {
                output.push(ParagraphContent::RangeEnd(RangeEnd {
                    node_type: child.local_name().to_owned(),
                    id: parse_range_id(child),
                }));
            }
            "commentRangeStart" | "commentRangeEnd" => {
                let marker = ParagraphContent::CommentRange(CommentRange {
                    node_type: child.local_name().to_owned(),
                    id: parse_range_id(child),
                    offset: None,
                });
                if child.local_name() == "commentRangeEnd" {
                    close_ranges(vec![marker], &fields, &mut output, &mut field_ends);
                } else {
                    output.push(marker);
                }
            }
            "oMath" | "oMathPara" => output.push(ParagraphContent::Inline(InlineNode::Math(
                parse_math(child),
            ))),
            // Paragraph properties are handled by the orchestrator.
            "pPr" | "proofErr" | "permStart" | "permEnd" => {}
            _ => {
                if let Some(node) = crate::inline::raw_foreign_inline(child) {
                    match fields.last_mut() {
                        Some(active) => active.absorb(node, Vec::new()),
                        None => output.push(ParagraphContent::Inline(node)),
                    }
                }
            }
        }
    }
    while let Some(field) = fields.pop() {
        let completed = field.finish();
        if let Some(parent) = fields.last_mut() {
            parent.absorb_nested(completed);
        } else {
            output.push(ParagraphContent::Inline(InlineNode::ComplexField(
                Box::new(completed),
            )));
        }
    }
    output.append(&mut field_ends);
    assign_marker_offsets(&mut output);
    Ok(output)
}

fn process_field_run(
    run: Run,
    fields: &mut Vec<OpenComplexField>,
    output: &mut Vec<ParagraphContent>,
    part: &str,
) -> Result<(), ParseError> {
    // Word can put text and several field boundaries in the same run.
    // Split at each boundary so every piece stays on its original side.
    let mut piece = Run {
        content: Vec::new(),
        ..run.clone()
    };
    for content in &run.content {
        let RunContent::FieldChar { char_type, .. } = content else {
            piece.content.push(content.clone());
            continue;
        };
        if !piece.content.is_empty() {
            emit_field_run(piece.clone(), fields, output);
            piece.content.clear();
        }
        match char_type.as_str() {
            "begin" => {
                if fields.len() >= MAX_FIELD_NESTING {
                    return Err(ParseError::ResourceLimit {
                        kind: "fieldDepth",
                        part: part.to_owned(),
                    });
                }
                fields.push(OpenComplexField::opening(&Run {
                    content: vec![content.clone()],
                    ..run.clone()
                }));
            }
            "separate" if !fields.is_empty() => {
                fields.last_mut().unwrap().switch_to_result();
            }
            "end" if !fields.is_empty() => {
                let completed = fields.pop().unwrap().finish();
                if let Some(parent) = fields.last_mut() {
                    parent.absorb_nested(completed);
                } else {
                    output.push(ParagraphContent::Inline(InlineNode::ComplexField(
                        Box::new(completed),
                    )));
                }
            }
            _ => piece.content.push(content.clone()),
        }
    }
    if !piece.content.is_empty() || run.content.is_empty() {
        emit_field_run(piece, fields, output);
    }
    Ok(())
}

fn emit_field_run(run: Run, fields: &mut [OpenComplexField], output: &mut Vec<ParagraphContent>) {
    if let Some(active) = fields.last_mut() {
        active.append_instruction(&instruction_text(std::slice::from_ref(&run)));
        active.adopt_formatting(&run);
        active.absorb(InlineNode::Run(run.clone()), vec![run]);
    } else {
        output.push(ParagraphContent::Inline(InlineNode::Run(run)));
    }
}

#[allow(clippy::too_many_arguments)]
fn parse_simple_field_composed(
    element: &XmlElement,
    relationships: Option<&RelationshipMap>,
    theme: Option<&Theme>,
    styles: Option<&StyleMap>,
    doc_defaults: Option<&DocDefaults>,
    part: &str,
    budget: &mut ParseBudget<'_>,
    drawing: Option<&mut DrawingContext<'_>>,
    depth: usize,
) -> Result<SimpleField, ParseError> {
    let instruction = element
        .attribute(Some("w"), "instr")
        .unwrap_or_default()
        .to_owned();
    let mut content = Vec::new();
    let mut result = Vec::new();
    for item in parse_paragraph_contents(
        element,
        relationships,
        theme,
        styles,
        doc_defaults,
        part,
        budget,
        drawing,
        depth,
        false,
    )? {
        match item {
            ParagraphContent::Inline(
                node @ (InlineNode::BookmarkStart(_) | InlineNode::BookmarkEnd(_)),
            ) => result.push(node),
            ParagraphContent::Inline(node) => {
                content.extend(shown_runs(std::slice::from_ref(&node)));
                result.push(node);
            }
            _ => {}
        }
    }
    let structured_result = (!result.is_empty()).then(|| StructuredFieldContent {
        inline: Some(result),
        blocks: None,
    });
    let field_tree = structured_result.clone().map(|result| StructuredFieldTree {
        version: Some(1.0),
        code: None,
        result: Some(result),
        children: None,
        display_mode: None,
    });
    Ok(SimpleField {
        node_type: SimpleFieldType::SimpleField,
        field_type: parse_field_type(&instruction),
        instruction,
        content,
        fld_lock: matches!(element.attribute(Some("w"), "fldLock"), Some("1" | "true"))
            .then_some(true),
        dirty: matches!(element.attribute(Some("w"), "dirty"), Some("1" | "true")).then_some(true),
        structured_result,
        field_tree,
    })
}

#[allow(clippy::too_many_arguments)]
fn parse_hyperlink_composed(
    element: &XmlElement,
    relationships: Option<&RelationshipMap>,
    theme: Option<&Theme>,
    styles: Option<&StyleMap>,
    doc_defaults: Option<&DocDefaults>,
    part: &str,
    budget: &mut ParseBudget<'_>,
    drawing: Option<&mut DrawingContext<'_>>,
    depth: usize,
) -> Result<Hyperlink, ParseError> {
    budget.check_nesting_depth(depth, part)?;
    let mut hyperlink = parse_hyperlink(
        element,
        relationships,
        theme,
        styles,
        doc_defaults,
        part,
        budget,
    )?;
    let structured = filter_field_inline(parse_paragraph_contents(
        element,
        relationships,
        theme,
        styles,
        doc_defaults,
        part,
        budget,
        drawing,
        depth,
        false,
    )?);
    hyperlink.children = structured
        .iter()
        .flat_map(|node| match node {
            InlineNode::BookmarkStart(_) | InlineNode::BookmarkEnd(_) => vec![node.clone()],
            _ => shown_runs(std::slice::from_ref(node))
                .into_iter()
                .map(InlineNode::Run)
                .collect(),
        })
        .collect();
    hyperlink.structured_children = structured
        .iter()
        .any(|node| !matches!(node, InlineNode::Run(_)))
        .then_some(structured);
    Ok(hyperlink)
}

/// The runs `nodes` show, in order: what a field result containing them reads as.
fn shown_runs(nodes: &[InlineNode]) -> Vec<Run> {
    let mut runs = Vec::new();
    for node in nodes {
        match node {
            InlineNode::Run(run) => runs.push(run.clone()),
            InlineNode::Hyperlink(hyperlink) => runs.extend(shown_runs(&hyperlink.children)),
            InlineNode::SimpleField(field) => runs.extend(field.content.iter().cloned()),
            InlineNode::ComplexField(field) => runs.extend(field.field_result.iter().cloned()),
            InlineNode::InlineSdt(sdt) => runs.extend(shown_runs(&sdt.content)),
            InlineNode::RawXml(raw) => runs.extend(raw.shown.iter().cloned()),
            InlineNode::Tracked(change) => runs.extend(shown_change_runs(change)),
            _ => {}
        }
    }
    runs
}

/// What a tracked change shows: its insertions, not its deletions nor the
/// comment references `raw_change` leaves out.
fn shown_change_runs(change: &TrackedInline) -> Vec<Run> {
    match change.node_type.as_str() {
        "deletion" | "moveFrom" => Vec::new(),
        _ => shown_runs(&change.content)
            .into_iter()
            .filter_map(|mut run| {
                let referenced = run
                    .content
                    .iter()
                    .any(|content| matches!(content, RunContent::CommentReference { .. }));
                run.content
                    .retain(|content| !matches!(content, RunContent::CommentReference { .. }));
                (!referenced || !run.content.is_empty()).then_some(run)
            })
            .collect(),
    }
}

/// The field instruction text `runs` carry.
fn instruction_text(runs: &[Run]) -> String {
    runs.iter()
        .flat_map(|run| &run.content)
        .filter_map(|content| match content {
            RunContent::InstrText { text } => Some(text.as_str()),
            _ => None,
        })
        .collect()
}

/// Whether the field characters inside `element` open and close whole fields,
/// so keeping it as markup leaves its own field balanced.
fn balanced_field_chars(element: &XmlElement) -> bool {
    fn visit(element: &XmlElement, depth: &mut usize) -> bool {
        if element.local_name() == "fldChar" {
            match element.attribute(Some("w"), "fldCharType") {
                Some("begin") => *depth += 1,
                Some("separate") if *depth == 0 => return false,
                Some("end") if *depth == 0 => return false,
                Some("end") => *depth -= 1,
                _ => {}
            }
        }
        element.child_elements().all(|child| visit(child, depth))
    }
    let mut depth = 0;
    visit(element, &mut depth) && depth == 0
}

/// A tracked change as the raw markup a field keeps, less the comment markers
/// `nested_comment_ranges` moved to the field's edges and the comment
/// references, which the export writes with their comments.
fn raw_change(element: &XmlElement, shown: Vec<Run>) -> InlineNode {
    /// A run without its comment references; `None` when nothing else was in it.
    fn without_references(run: &XmlElement) -> Option<XmlElement> {
        let mut kept = run.clone();
        kept.children.retain(
            |node| !matches!(node, XmlNode::Element(child) if child.local_name() == "commentReference"),
        );
        let referenced = kept.children.len() < run.children.len();
        let left = kept
            .child_elements()
            .any(|child| child.local_name() != "rPr");
        (!referenced || left).then_some(kept)
    }
    fn without_comment_ranges(element: &XmlElement) -> XmlElement {
        let children = element
            .children
            .iter()
            .filter_map(|node| match node {
                XmlNode::Element(child) => match child.local_name() {
                    "commentRangeStart" | "commentRangeEnd" => None,
                    "r" => without_references(child).map(XmlNode::Element),
                    "customXml" | "smartTag" | "sdt" | "sdtContent" | "hyperlink" | "ins"
                    | "del" | "moveFrom" | "moveTo" | "fldSimple" => {
                        Some(XmlNode::Element(without_comment_ranges(child)))
                    }
                    _ => Some(node.clone()),
                },
                _ => Some(node.clone()),
            })
            .collect();
        XmlElement {
            name: element.name.clone(),
            attributes: element.attributes.clone(),
            children,
        }
    }
    InlineNode::RawXml(Box::new(RawInlineXml {
        node_type: RawInlineXmlType::RawXml,
        xml: without_comment_ranges(element).to_raw_inline_xml(),
        shown,
    }))
}

fn raw_change_kind(element: &XmlElement) -> Option<bool> {
    match element.local_name() {
        "ins" | "moveTo" => Some(false),
        "del" | "moveFrom" => Some(true),
        _ => None,
    }
}

fn holds_change(element: &XmlElement) -> bool {
    raw_change_kind(element).is_some()
        || element.local_name() != "rPr" && element.child_elements().any(holds_change)
}

fn parse_raw_root(xml: &str) -> Option<XmlElement> {
    let limits = ParseLimits::default();
    parse_xml_strict(xml.as_bytes(), "raw-inline", &mut ParseBudget::new(&limits))
        .ok()?
        .root()
        .cloned()
}

/// Whether raw inline markup a field keeps holds a tracked change.
pub fn raw_holds_changes(xml: &str) -> bool {
    parse_raw_root(xml).is_some_and(|root| holds_change(&root))
}

/// Raw inline markup a field keeps, its tracked changes resolved (accepting
/// keeps insertions and removes deletions, rejecting the reverse) and parsed,
/// with its part's `relationships` and the package `theme` but no styles, into
/// what a field holding it in a file reads as; `None` when it holds no change.
pub fn resolve_raw_inline(
    xml: &str,
    accept: bool,
    relationships: Option<&RelationshipMap>,
    theme: Option<&Theme>,
) -> Option<Vec<InlineNode>> {
    fn resolve(element: &XmlElement, accept: bool) -> Vec<XmlElement> {
        match raw_change_kind(element) {
            Some(deletion) if deletion == accept => Vec::new(),
            Some(deletion) => {
                let content = if deletion {
                    normalize_deletion_element(element)
                } else {
                    element.clone()
                };
                content
                    .child_elements()
                    .flat_map(|child| resolve(child, accept))
                    .collect()
            }
            None if element.local_name() == "rPr" => vec![element.clone()],
            None => vec![XmlElement {
                name: element.name.clone(),
                attributes: element.attributes.clone(),
                children: element
                    .children
                    .iter()
                    .flat_map(|node| match node {
                        XmlNode::Element(child) => resolve(child, accept)
                            .into_iter()
                            .map(XmlNode::Element)
                            .collect(),
                        node => vec![node.clone()],
                    })
                    .collect(),
            }],
        }
    }
    let root = parse_raw_root(xml).filter(holds_change)?;
    let paragraph = XmlElement {
        name: "w:p".to_owned(),
        attributes: Default::default(),
        children: resolve(&root, accept)
            .into_iter()
            .map(XmlNode::Element)
            .collect(),
    };
    let limits = ParseLimits::default();
    let content = parse_paragraph_contents(
        &paragraph,
        relationships,
        theme,
        None,
        None,
        "raw-inline",
        &mut ParseBudget::new(&limits),
        None,
        0,
        false,
    )
    .ok()?;
    Some(
        content
            .into_iter()
            .filter_map(|content| match content {
                ParagraphContent::Inline(node) => Some(node),
                _ => None,
            })
            .collect(),
    )
}

/// Raw inline markup a field keeps, parsed as the same markup reads outside a
/// field: tracked changes stay changes. With no relationships, theme or
/// styles; `None` when it does not parse.
pub fn parse_raw_inline(xml: &str) -> Option<Vec<InlineNode>> {
    let paragraph = XmlElement {
        name: "w:p".to_owned(),
        attributes: Default::default(),
        children: vec![XmlNode::Element(parse_raw_root(xml)?)],
    };
    let limits = ParseLimits::default();
    parse_paragraph_contents(
        &paragraph,
        None,
        None,
        None,
        None,
        "raw-inline",
        &mut ParseBudget::new(&limits),
        None,
        0,
        false,
    )
    .ok()
    .map(filter_field_inline)
}

fn filter_field_inline(content: Vec<ParagraphContent>) -> Vec<InlineNode> {
    content
        .into_iter()
        .filter_map(|content| match content {
            ParagraphContent::Inline(node) => Some(node),
            ParagraphContent::Tracked(change) => Some(InlineNode::Tracked(Box::new(change))),
            _ => None,
        })
        .collect()
}

fn normalize_deletion_element(element: &XmlElement) -> XmlElement {
    let local = element.local_name();
    let mapped = match local {
        "delText" => Some("t"),
        "delInstrText" => Some("instrText"),
        _ => None,
    };
    let name = mapped.map_or_else(
        || element.name.clone(),
        |local| match element.name.split_once(':') {
            Some((prefix, _)) => format!("{prefix}:{local}"),
            None => local.to_owned(),
        },
    );
    XmlElement {
        name,
        attributes: element.attributes.clone(),
        children: element
            .children
            .iter()
            .map(|node| match node {
                XmlNode::Element(child) => XmlNode::Element(normalize_deletion_element(child)),
                _ => node.clone(),
            })
            .collect(),
    }
}

/// Places range ends after the complex field still open around them, if any.
fn close_ranges(
    ends: Vec<ParagraphContent>,
    fields: &[OpenComplexField],
    output: &mut Vec<ParagraphContent>,
    field_ends: &mut Vec<ParagraphContent>,
) {
    if fields.is_empty() {
        output.extend(ends);
    } else {
        field_ends.extend(ends);
    }
}

/// Comment range markers inside a hyperlink, tracked change, inline content
/// control or simple field, which cannot hold them: starts go before it and
/// ends after it, so the range holds the container whole (complex fields:
/// `close_ranges`).
fn nested_comment_ranges(container: &XmlElement) -> (Vec<ParagraphContent>, Vec<ParagraphContent>) {
    let (mut starts, mut ends) = (Vec::new(), Vec::new());
    for child in transparent_children(container, true) {
        match child.local_name() {
            name @ ("commentRangeStart" | "commentRangeEnd") => {
                let marker = ParagraphContent::CommentRange(CommentRange {
                    node_type: name.to_owned(),
                    id: parse_range_id(child),
                    offset: None,
                });
                if name == "commentRangeStart" {
                    starts.push(marker);
                } else {
                    ends.push(marker);
                }
            }
            "hyperlink" | "ins" | "del" | "moveFrom" | "moveTo" | "fldSimple" => {
                let (inner_starts, inner_ends) = nested_comment_ranges(child);
                starts.extend(inner_starts);
                ends.extend(inner_ends);
            }
            _ => {}
        }
    }
    (starts, ends)
}

fn parse_range_id(element: &XmlElement) -> f64 {
    element
        .attribute(Some("w"), "id")
        .and_then(parse_javascript_integer_prefix)
        .unwrap_or(0.0)
}

fn parse_math(element: &XmlElement) -> MathEquation {
    let mut text = String::new();
    append_math_text(element, &mut text);
    MathEquation {
        node_type: MathType::MathEquation,
        display: if element.local_name() == "oMathPara" {
            "block"
        } else {
            "inline"
        }
        .to_owned(),
        omml_xml: element.to_raw_inline_xml(),
        plain_text: (!text.is_empty()).then_some(text),
    }
}

fn append_math_text(element: &XmlElement, output: &mut String) {
    for child in element.child_elements() {
        if child.local_name() == "t" {
            for node in &child.children {
                if let XmlNode::Text(text) = node {
                    output.push_str(text);
                }
            }
        } else {
            append_math_text(child, output);
        }
    }
}

#[allow(clippy::too_many_arguments)]
pub(crate) fn parse_run_composed(
    element: &XmlElement,
    relationships: Option<&RelationshipMap>,
    theme: Option<&Theme>,
    styles: Option<&StyleMap>,
    doc_defaults: Option<&DocDefaults>,
    budget: &mut ParseBudget<'_>,
    mut drawing: Option<&mut DrawingContext<'_>>,
) -> Result<Run, ParseError> {
    let mut run = parse_run(element, theme, styles, doc_defaults).run;
    // `None` keeps the run's opaque drawing: a `w:pict`, `w:object` or
    // `mc:AlternateContent` the model cannot carry is replayed verbatim.
    let mut replacements = Vec::new();
    for child in element.child_elements() {
        match child.local_name() {
            "drawing" => replacements.push(Some(parse_drawing_owned(
                child,
                relationships,
                budget,
                drawing.as_deref_mut(),
            )?)),
            "pict" | "object" => {
                let parsed =
                    parse_drawing_owned(child, relationships, budget, drawing.as_deref_mut())?;
                replacements.push((!parsed.is_empty()).then_some(parsed));
            }
            "AlternateContent" if contains_drawing_owned_content(child) => {
                let parsed =
                    parse_alternate_content(child, relationships, budget, drawing.as_deref_mut())?;
                // Text boxes enter the paragraph through block enrichment instead.
                replacements.push((!parsed.is_empty() || holds_text_box(child)).then_some(parsed));
            }
            _ => {}
        }
    }
    let mut replacement = replacements.into_iter();
    let mut content = Vec::new();
    for item in run.content {
        if matches!(item, RunContent::OpaqueDrawing { .. }) {
            match replacement.next() {
                Some(Some(mut parsed)) => {
                    if let (
                        RunContent::OpaqueDrawing { kind, xml },
                        [
                            RunContent::Drawing { source_xml, .. }
                            | RunContent::Shape { source_xml, .. },
                        ],
                    ) = (&item, parsed.as_mut_slice())
                        && kind != "drawing"
                    {
                        *source_xml = Some(xml.clone());
                    }
                    content.extend(parsed);
                }
                Some(None) => content.push(item),
                None => {}
            }
        } else {
            content.push(item);
        }
    }
    run.content = content;
    Ok(run)
}

fn parse_drawing_owned(
    element: &XmlElement,
    relationships: Option<&RelationshipMap>,
    budget: &mut ParseBudget<'_>,
    drawing: Option<&mut DrawingContext<'_>>,
) -> Result<Vec<RunContent>, ParseError> {
    if matches!(element.local_name(), "pict" | "object") {
        if let Some(rule) = parse_horizontal_rule(element) {
            return Ok(vec![RunContent::HorizontalRule {
                rule: Box::new(rule),
            }]);
        }
        let media = drawing.as_ref().map(|context| context.media);
        return Ok(parse_vml_image_content(element, relationships, media)
            .map(|image| RunContent::Drawing {
                image: Box::new(image),
                source_xml: None,
            })
            .into_iter()
            .collect());
    }
    if is_text_box_drawing(element) {
        return Ok(Vec::new());
    }
    let charts = drawing.as_ref().map(|context| context.charts);
    match parse_chart_from_drawing(element, relationships, charts)? {
        DrawingChart::Chart(chart) => return Ok(vec![RunContent::Chart { chart }]),
        DrawingChart::Unread => {
            return Ok(vec![RunContent::OpaqueDrawing {
                kind: "drawing".to_owned(),
                xml: element.to_raw_inline_xml(),
            }]);
        }
        DrawingChart::None => {}
    }
    if is_smart_art_drawing(element) {
        let shape = match drawing {
            Some(context) => parse_smart_art_from_drawing(
                element,
                relationships,
                Some(&mut *context.smart_art),
                budget,
            )?,
            None => None,
        };
        return Ok(shape
            .map(|mut shape| {
                apply_shape_metadata(&mut shape, element);
                RunContent::Shape {
                    shape: Box::new(shape),
                    source_xml: None,
                }
            })
            .into_iter()
            .collect());
    }
    if is_shape_drawing(element) {
        if let Some(mut shape) = parse_shape_from_drawing(element) {
            let media = drawing.as_ref().map(|context| context.media);
            resolve_shape_fill_pictures(&mut shape, relationships, media);
            apply_shape_metadata(&mut shape, element);
            return Ok(vec![RunContent::Shape {
                shape: Box::new(shape),
                source_xml: None,
            }]);
        }
        return Ok(Vec::new());
    }
    let media = drawing.as_ref().map(|context| context.media);
    Ok(parse_drawing(element, relationships, media)
        .map(|image| RunContent::Drawing {
            image: Box::new(image),
            source_xml: None,
        })
        .into_iter()
        .collect())
}

fn parse_alternate_content(
    element: &XmlElement,
    relationships: Option<&RelationshipMap>,
    budget: &mut ParseBudget<'_>,
    mut drawing: Option<&mut DrawingContext<'_>>,
) -> Result<Vec<RunContent>, ParseError> {
    for branch_name in ["Choice", "Fallback"] {
        for branch in element
            .child_elements()
            .filter(|branch| branch.local_name() == branch_name)
        {
            let mut parsed = Vec::new();
            for child in branch.child_elements() {
                if matches!(child.local_name(), "drawing" | "pict" | "object") {
                    parsed.extend(parse_drawing_owned(
                        child,
                        relationships,
                        budget,
                        drawing.as_deref_mut(),
                    )?);
                }
            }
            if !parsed.is_empty() {
                return Ok(parsed);
            }
        }
    }
    Ok(Vec::new())
}

fn contains_drawing_owned_content(element: &XmlElement) -> bool {
    element.child_elements().any(|child| {
        matches!(child.local_name(), "drawing" | "pict" | "object")
            || contains_drawing_owned_content(child)
    })
}

/// The branch block enrichment reads (`Choice`, else `Fallback`) holds a text box.
fn holds_text_box(alternate: &XmlElement) -> bool {
    alternate
        .child_elements()
        .find(|branch| branch.local_name() == "Choice")
        .or_else(|| {
            alternate
                .child_elements()
                .find(|branch| branch.local_name() == "Fallback")
        })
        .is_some_and(|branch| {
            branch
                .child_elements()
                .filter(|child| child.local_name() == "drawing")
                .any(|drawing| crate::text_box::parse_text_box(drawing).is_some())
        })
}

fn apply_shape_metadata(shape: &mut Shape, drawing: &XmlElement) {
    let doc_properties = find_descendant_by_name(drawing, "wp:docPr", 0);
    let non_visual = find_descendant_by_name(drawing, "a:cNvPr", 0);
    let title = doc_properties
        .and_then(|element| element.attribute(None, "title"))
        .or_else(|| non_visual.and_then(|element| element.attribute(None, "title")))
        .filter(|value| !value.is_empty())
        .map(|value| value.chars().take(4096).collect::<String>());
    let description = doc_properties
        .and_then(|element| element.attribute(None, "descr"))
        .or_else(|| non_visual.and_then(|element| element.attribute(None, "descr")))
        .filter(|value| !value.is_empty())
        .map(|value| value.chars().take(4096).collect::<String>());
    let decorative = doc_properties
        .is_some_and(|element| element.attribute(None, "decorative") == Some("1"))
        || find_descendant_by_name(drawing, "adec:decorative", 0).is_some();
    let hidden = doc_properties
        .is_some_and(|element| element.attribute(None, "hidden") == Some("1"))
        || non_visual.is_some_and(|element| element.attribute(None, "hidden") == Some("1"));
    let relative_height = find_descendant_by_name(drawing, "wp:anchor", 0)
        .and_then(|anchor| anchor.attribute(None, "relativeHeight"))
        .and_then(|raw| raw.parse::<f64>().ok())
        .filter(|value| value.is_finite());
    if let Some(title) = title {
        shape.title = Some(title);
    }
    if let Some(description) = description {
        shape.description = Some(description);
    }
    if decorative {
        shape.decorative = Some(true);
    }
    if hidden {
        shape.hidden = Some(true);
    }
    if let Some(relative_height) = relative_height {
        shape.relative_height = Some(relative_height);
    }
    let mut scene = shape.scene.take().map(|scene| *scene).unwrap_or_default();
    scene.version = Some(1.0);
    scene.title.clone_from(&shape.title);
    scene.description.clone_from(&shape.description);
    scene.decorative = shape.decorative;
    scene.hidden = shape.hidden;
    shape.scene = Some(Box::new(scene));
}

fn find_descendant_by_name<'a>(
    element: &'a XmlElement,
    full_name: &str,
    depth: usize,
) -> Option<&'a XmlElement> {
    if depth > 64 {
        return None;
    }
    if element.name == full_name {
        return Some(element);
    }
    element
        .child_elements()
        .find_map(|child| find_descendant_by_name(child, full_name, depth + 1))
}

fn assign_marker_offsets(content: &mut [ParagraphContent]) {
    let mut offset = 0usize;
    for node in content {
        match node {
            ParagraphContent::Inline(InlineNode::BookmarkStart(bookmark)) => {
                bookmark.position = Some(ContentPosition {
                    offset: Some(offset as f64),
                });
            }
            ParagraphContent::Inline(InlineNode::BookmarkEnd(bookmark)) => {
                bookmark.position = Some(ContentPosition {
                    offset: Some(offset as f64),
                });
            }
            ParagraphContent::CommentRange(marker) => marker.offset = Some(offset as f64),
            _ => offset = offset.saturating_add(paragraph_content_length(node)),
        }
    }
}

fn paragraph_content_length(content: &ParagraphContent) -> usize {
    match content {
        ParagraphContent::Inline(node) => inline_node_length(node),
        ParagraphContent::Tracked(change) => change.content.iter().map(inline_node_length).sum(),
        ParagraphContent::RangeStart(_)
        | ParagraphContent::RangeEnd(_)
        | ParagraphContent::CommentRange(_) => 0,
    }
}

fn inline_node_length(node: &InlineNode) -> usize {
    match node {
        InlineNode::RawXml(_) => 0,
        InlineNode::Run(run) => run
            .content
            .iter()
            .map(|content| match content {
                RunContent::Text { text, .. } | RunContent::InstrText { text } => {
                    text.encode_utf16().count()
                }
                RunContent::Tab
                | RunContent::PositionalTab { .. }
                | RunContent::SoftHyphen
                | RunContent::NoBreakHyphen
                | RunContent::Symbol { .. } => 1,
                _ => 0,
            })
            .sum(),
        InlineNode::Hyperlink(hyperlink) => hyperlink
            .children
            .iter()
            .filter(|child| matches!(child, InlineNode::Run(_)))
            .map(inline_node_length)
            .sum(),
        InlineNode::SimpleField(field) => field
            .content
            .iter()
            .cloned()
            .map(InlineNode::Run)
            .map(|node| inline_node_length(&node))
            .sum(),
        InlineNode::ComplexField(field) => field
            .field_result
            .iter()
            .cloned()
            .map(InlineNode::Run)
            .map(|node| inline_node_length(&node))
            .sum(),
        InlineNode::InlineSdt(sdt) => sdt.content.iter().map(inline_node_length).sum(),
        InlineNode::Tracked(change) => change.content.iter().map(inline_node_length).sum(),
        InlineNode::Math(math) => math
            .plain_text
            .as_deref()
            .map(|text| text.encode_utf16().count())
            .unwrap_or(0),
        InlineNode::BookmarkStart(_) | InlineNode::BookmarkEnd(_) => 0,
    }
}

fn apply_list_rendering(
    paragraph: &mut Paragraph,
    direct_properties: Option<&XmlElement>,
    styles: Option<&StyleMap>,
    numbering: Option<&NumberingMap>,
) {
    let mut effective = paragraph
        .formatting
        .as_ref()
        .and_then(|formatting| formatting.num_pr.clone());
    let mut from_style = false;
    if effective.is_none()
        && let (Some(style_id), Some(styles)) = (
            paragraph
                .formatting
                .as_ref()
                .and_then(|formatting| formatting.style_id.as_deref()),
            styles,
        )
        && let Some(numbering) = styles
            .get(style_id)
            .and_then(|style| style.p_pr.as_ref())
            .and_then(|formatting| formatting.num_pr.clone())
    {
        from_style = true;
        paragraph
            .formatting
            .get_or_insert_with(ParagraphFormatting::default)
            .num_pr = Some(numbering.clone());
        paragraph.formatting.as_mut().unwrap().num_pr_from_style = Some(numbering.clone());
        effective = Some(numbering);
    }
    let (Some(properties), Some(numbering)) = (effective, numbering) else {
        return;
    };
    let Some(mut rendering) = compute_list_rendering(properties.num_id, properties.ilvl, numbering)
    else {
        return;
    };
    if rendering.is_bullet {
        rendering.marker = convert_bullet_to_unicode(&rendering.marker);
    }
    let level = numbering.get_level(rendering.num_id, rendering.level);
    if let Some(level_properties) = level.and_then(|level| level.p_pr) {
        let style_indents = if from_style {
            style_chain_ind(
                paragraph
                    .formatting
                    .as_ref()
                    .and_then(|formatting| formatting.style_id.as_deref()),
                styles,
            )
        } else {
            (false, false)
        };
        let direct_indent = direct_properties.and_then(|properties| properties.child("w", "ind"));
        let direct_left = direct_indent.is_some_and(|indent| {
            ["left", "start", "leftChars", "startChars"]
                .iter()
                .any(|name| indent.attribute(Some("w"), name).is_some())
        });
        let direct_first = direct_indent.is_some_and(|indent| {
            ["firstLine", "hanging", "firstLineChars", "hangingChars"]
                .iter()
                .any(|name| {
                    indent.attribute(Some("w"), name).is_some_and(|raw| {
                        parse_javascript_integer_prefix(raw).is_none_or(|value| value != 0.0)
                    })
                })
        });
        if !direct_left && !style_indents.0 {
            rendering.indent_left = level_properties.indent_left;
        }
        if !direct_first && !style_indents.1 {
            rendering.indent_first_line = level_properties.indent_first_line;
            rendering.hanging_indent = level_properties.hanging_indent;
        }
    }
    paragraph.list_rendering = Some(rendering);
}

fn style_chain_ind(style_id: Option<&str>, styles: Option<&StyleMap>) -> (bool, bool) {
    let (Some(mut current), Some(styles)) = (style_id, styles) else {
        return (false, false);
    };
    let mut seen = Vec::new();
    let mut result = (false, false);
    while !seen.contains(&current) {
        seen.push(current);
        let Some(style) = styles.get(current) else {
            break;
        };
        if let Some(properties) = &style.p_pr {
            result.0 |= properties.indent_left.is_some() || properties.indent_left_chars.is_some();
            result.1 |= properties.indent_first_line.is_some()
                || properties.hanging_indent.is_some()
                || properties.indent_first_line_chars.is_some();
        }
        if result.0 && result.1 {
            break;
        }
        let Some(parent) = style.based_on.as_deref() else {
            break;
        };
        current = parent;
    }
    result
}

pub fn convert_bullet_to_unicode(value: &str) -> String {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        return "•".to_owned();
    }
    let first = value.encode_utf16().next().unwrap_or_default();
    match first {
        0x00b7 | 0xf0b7 | 0x2022 => "•".to_owned(),
        0x006f | 0x0071 | 0xf06f | 0x25cb => "○".to_owned(),
        0x00a7 | 0x006e | 0xf06e | 0xf0a7 | 0x25a0 => "■".to_owned(),
        0x00fc | 0x00a8 | 0x00fb | 0x00fe | 0xf0fc => "✓".to_owned(),
        0x0075 | 0x25c6 => "◆".to_owned(),
        0x0076 => "❖".to_owned(),
        0x25cf => "●".to_owned(),
        0x25a1 => "□".to_owned(),
        0x25c7 => "◇".to_owned(),
        0x2013 => "–".to_owned(),
        0x2014 => "—".to_owned(),
        0x003e => ">".to_owned(),
        0x002d => "-".to_owned(),
        0xe000..=0xf8ff | 0..=31 | 127..=159 => "•".to_owned(),
        _ => value.to_owned(),
    }
}

pub fn paragraph_starts_with_rendered_page_break(element: &XmlElement) -> bool {
    fn visit(element: &XmlElement, saw: &mut bool) -> Option<bool> {
        for child in element.child_elements() {
            let name = child.local_name();
            if matches!(
                name,
                "pPr"
                    | "proofErr"
                    | "bookmarkStart"
                    | "bookmarkEnd"
                    | "commentRangeStart"
                    | "commentRangeEnd"
                    | "commentReference"
                    | "permStart"
                    | "permEnd"
                    | "rsidR"
            ) {
                continue;
            }
            if name == "lastRenderedPageBreak" {
                *saw = true;
                continue;
            }
            if name == "r" {
                for run_child in child.child_elements() {
                    let run_name = run_child.local_name();
                    if run_name == "rPr" {
                        continue;
                    }
                    if run_name == "lastRenderedPageBreak" {
                        *saw = true;
                        continue;
                    }
                    if run_name == "br" && run_child.attribute(Some("w"), "type") == Some("page") {
                        return Some(true);
                    }
                    if matches!(
                        run_name,
                        "t" | "tab"
                            | "br"
                            | "cr"
                            | "sym"
                            | "drawing"
                            | "pict"
                            | "object"
                            | "softHyphen"
                            | "noBreakHyphen"
                            | "fldChar"
                            | "instrText"
                            | "pgNum"
                            | "separator"
                            | "continuationSeparator"
                            | "footnoteRef"
                            | "endnoteRef"
                            | "footnoteReference"
                            | "endnoteReference"
                            | "ptab"
                            | "monthShort"
                            | "monthLong"
                            | "yearShort"
                            | "yearLong"
                            | "dayShort"
                            | "dayLong"
                    ) {
                        return Some(false);
                    }
                }
                continue;
            }
            if matches!(
                name,
                "hyperlink"
                    | "smartTag"
                    | "sdt"
                    | "sdtContent"
                    | "fldSimple"
                    | "customXml"
                    | "ins"
                    | "del"
                    | "moveFrom"
                    | "moveTo"
            ) && let Some(result) = visit(child, saw)
            {
                return Some(result);
            }
        }
        None
    }
    let mut saw = false;
    match visit(element, &mut saw) {
        Some(true) => true,
        Some(false) => saw,
        None => false,
    }
}

fn normalize_hex_id(value: Option<&str>, ids: &mut HexIdAllocator) -> Option<String> {
    let value = value?;
    let valid = value.len() == 8
        && value.bytes().all(|byte| byte.is_ascii_hexdigit())
        && u32::from_str_radix(value, 16).is_ok_and(|value| value < MAX_HEX_ID_EXCLUSIVE);
    Some(if valid {
        value.to_owned()
    } else {
        ids.allocate()
    })
}

fn enum_attribute(element: &XmlElement, name: &str, allowed: &[&str]) -> Option<String> {
    element
        .attribute(Some("w"), name)
        .filter(|value| allowed.contains(value))
        .map(str::to_owned)
}

fn nonempty_attribute(element: &XmlElement, name: &str) -> Option<String> {
    element
        .attribute(Some("w"), name)
        .filter(|value| !value.is_empty())
        .map(str::to_owned)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::{ParseLimits, parse_xml};

    fn parse(xml: &str) -> Paragraph {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let document = parse_xml(xml.as_bytes(), "word/document.xml", &mut budget).unwrap();
        let mut ids = HexIdAllocator::from_sha256(&"0".repeat(64)).unwrap();
        parse_paragraph(
            document.root().unwrap(),
            None,
            None,
            None,
            None,
            None,
            "word/document.xml",
            &mut budget,
            &mut ids,
            None,
            false,
            0,
        )
        .unwrap()
    }

    #[test]
    fn pins_paragraph_mark_changes_and_inline_section_properties() {
        let paragraph = parse(
            r#"<w:p xmlns:w="w" xmlns:r="r">
              <w:pPr>
                <w:rPr>
                  <w:ins w:id="12tail" w:author="  A  " w:date="2024-01-02"/>
                  <w:del w:author="missing-id"/>
                </w:rPr>
                <w:sectPr><w:type w:val="continuous"/><w:headerReference w:type="default" r:id="rH"/></w:sectPr>
              </w:pPr>
            </w:p>"#,
        );
        let insertion = paragraph.p_pr_ins.unwrap();
        assert_eq!(insertion.id, 12.0);
        // Paragraph-mark metadata is deliberately not trimmed/defaulted.
        assert_eq!(insertion.author, "  A  ");
        assert!(paragraph.p_pr_del.is_none());
        let section = paragraph.section_properties.unwrap();
        assert_eq!(section.section_start.as_deref(), Some("continuous"));
        assert_eq!(
            section.header_references.as_ref().unwrap()[0].relationship_id,
            "rH"
        );
    }

    #[test]
    fn parses_deletion_fields_and_utf16_marker_offsets() {
        let paragraph = parse(
            r#"<w:p xmlns:w="w">
              <w:bookmarkStart w:id="7" w:name="start"/>
              <w:r><w:t>A😀</w:t></w:r>
              <w:del w:id="8"><w:r><w:delText>B</w:delText></w:r></w:del>
              <w:commentRangeStart w:id="9"/>
              <w:r><w:fldChar w:fldCharType="begin"/></w:r>
              <w:r><w:instrText> PAGE </w:instrText></w:r>
              <w:r><w:fldChar w:fldCharType="separate"/></w:r>
              <w:r><w:t>4</w:t></w:r>
              <w:r><w:fldChar w:fldCharType="end"/></w:r>
              <w:bookmarkEnd w:id="7"/>
            </w:p>"#,
        );
        let ParagraphContent::Inline(InlineNode::BookmarkStart(start)) = &paragraph.content[0]
        else {
            panic!("bookmark start")
        };
        assert_eq!(start.position.as_ref().unwrap().offset, Some(0.0));
        let ParagraphContent::Tracked(deletion) = &paragraph.content[2] else {
            panic!("deletion")
        };
        let InlineNode::Run(run) = &deletion.content[0] else {
            panic!("deleted run")
        };
        assert!(matches!(
            &run.content[0],
            RunContent::Text { text, .. } if text == "B"
        ));
        let ParagraphContent::CommentRange(comment) = &paragraph.content[3] else {
            panic!("comment marker")
        };
        // A, an astral emoji, and deleted B span four UTF-16 code units.
        assert_eq!(comment.offset, Some(4.0));
        let ParagraphContent::Inline(InlineNode::ComplexField(field)) = &paragraph.content[4]
        else {
            panic!("complex field")
        };
        assert_eq!(field.instruction, "PAGE");
        assert_eq!(field.field_type, "PAGE");
        assert_eq!(field.field_result.len(), 1);
        let ParagraphContent::Inline(InlineNode::BookmarkEnd(end)) = &paragraph.content[5] else {
            panic!("bookmark end")
        };
        assert_eq!(end.position.as_ref().unwrap().offset, Some(5.0));
    }

    #[test]
    fn comment_markers_inside_a_container_move_to_its_edges() {
        let paragraph = parse(
            r#"<w:p xmlns:w="w">
              <w:hyperlink w:anchor="a"><w:r><w:t>li</w:t></w:r><w:commentRangeStart w:id="1"/><w:r><w:t>nk</w:t></w:r></w:hyperlink>
              <w:ins w:id="5"><w:r><w:t>in</w:t></w:r><w:commentRangeEnd w:id="1"/><w:commentRangeStart w:id="2"/></w:ins>
              <w:del w:id="6"><w:commentRangeEnd w:id="2"/><w:r><w:delText>de</w:delText></w:r></w:del>
              <w:sdt><w:sdtContent><w:r><w:t>ct</w:t></w:r><w:commentRangeStart w:id="3"/><w:r><w:t>rl</w:t></w:r><w:commentRangeEnd w:id="3"/></w:sdtContent></w:sdt>
            </w:p>"#,
        );
        let layout: Vec<String> = paragraph
            .content
            .iter()
            .map(|content| match content {
                ParagraphContent::CommentRange(marker) => {
                    format!("{}{}", &marker.node_type[12..13], marker.id)
                }
                ParagraphContent::Inline(InlineNode::Hyperlink(_)) => "link".into(),
                ParagraphContent::Inline(InlineNode::InlineSdt(_)) => "sdt".into(),
                ParagraphContent::Tracked(change) => change.node_type.clone(),
                _ => "other".into(),
            })
            .collect();
        assert_eq!(
            layout,
            [
                "S1",
                "link",
                "S2",
                "insertion",
                "E1",
                "deletion",
                "E2",
                "S3",
                "sdt",
                "E3"
            ]
        );
    }

    #[test]
    fn comment_markers_inside_a_field_result_move_to_its_edges() {
        let paragraph = parse(
            r#"<w:p xmlns:w="w">
              <w:r><w:fldChar w:fldCharType="begin"/></w:r>
              <w:r><w:instrText> DATE </w:instrText></w:r>
              <w:r><w:fldChar w:fldCharType="separate"/></w:r>
              <w:r><w:t>20</w:t></w:r><w:commentRangeStart w:id="1"/><w:r><w:t>26</w:t></w:r><w:commentRangeEnd w:id="1"/>
              <w:r><w:fldChar w:fldCharType="end"/></w:r>
              <w:r><w:t> tail</w:t></w:r>
            </w:p>"#,
        );
        let layout: Vec<&str> = paragraph
            .content
            .iter()
            .map(|content| match content {
                ParagraphContent::CommentRange(marker) => marker.node_type.as_str(),
                ParagraphContent::Inline(InlineNode::ComplexField(_)) => "field",
                _ => "other",
            })
            .collect();
        assert_eq!(
            layout,
            ["commentRangeStart", "field", "commentRangeEnd", "other"]
        );
    }

    #[test]
    fn comment_markers_in_a_simple_field_or_a_change_in_a_field_result_move_to_the_field_edges() {
        let paragraph = parse(
            r#"<w:p xmlns:w="w">
              <w:fldSimple w:instr=" DATE "><w:r><w:t>20</w:t></w:r><w:commentRangeStart w:id="1"/><w:r><w:t>26</w:t></w:r><w:commentRangeEnd w:id="1"/></w:fldSimple>
              <w:r><w:fldChar w:fldCharType="begin"/></w:r>
              <w:r><w:instrText> DATE </w:instrText></w:r>
              <w:r><w:fldChar w:fldCharType="separate"/></w:r>
              <w:ins w:id="5"><w:commentRangeStart w:id="2"/><w:r><w:t>20</w:t></w:r><w:commentRangeEnd w:id="2"/></w:ins>
              <w:r><w:fldChar w:fldCharType="end"/></w:r>
            </w:p>"#,
        );
        let layout: Vec<String> = paragraph
            .content
            .iter()
            .map(|content| match content {
                ParagraphContent::CommentRange(marker) => {
                    format!("{}{}", &marker.node_type[12..13], marker.id)
                }
                ParagraphContent::Inline(InlineNode::SimpleField(_)) => "simple".into(),
                ParagraphContent::Inline(InlineNode::ComplexField(_)) => "field".into(),
                ParagraphContent::Tracked(change) => change.node_type.clone(),
                _ => "other".into(),
            })
            .collect();
        assert_eq!(layout, ["S1", "simple", "E1", "S2", "field", "E2"]);
        // The insertion stays in the result as its markup, less the markers.
        let ParagraphContent::Inline(InlineNode::ComplexField(field)) = &paragraph.content[4]
        else {
            panic!("complex field")
        };
        let result = field
            .structured_result
            .as_ref()
            .unwrap()
            .inline
            .as_ref()
            .unwrap();
        let [InlineNode::RawXml(raw)] = result.as_slice() else {
            panic!("raw insertion")
        };
        assert_eq!(
            raw.xml,
            r#"<w:ins w:id="5"><w:r><w:t>20</w:t></w:r></w:ins>"#
        );
        assert_eq!(field.field_result.len(), 1);
    }

    #[test]
    fn kept_changes_resolve_to_what_their_accepted_or_rejected_markup_parses_as() {
        let texts = |nodes: Vec<InlineNode>| -> Vec<String> {
            nodes
                .iter()
                .map(|node| match node {
                    InlineNode::Run(run) => run_text(run),
                    InlineNode::Hyperlink(link) => {
                        format!("H({})", run_text(&shown_runs(&link.children)[0]))
                    }
                    InlineNode::ComplexField(field) => format!(
                        "[{}|{}]",
                        field.instruction,
                        run_text(&field.field_result[0])
                    ),
                    other => other.node_type().to_owned(),
                })
                .collect()
        };
        let xml = r#"<w:del w:id="1"><w:r><w:delText>old</w:delText></w:r></w:del>"#;
        assert_eq!(
            resolve_raw_inline(xml, true, None, None).map(texts),
            Some(Vec::new())
        );
        assert_eq!(
            resolve_raw_inline(xml, false, None, None).map(texts),
            Some(vec!["old".to_owned()])
        );
        let link = r#"<w:hyperlink w:anchor="a"><w:ins w:id="2"><w:r><w:t>new</w:t></w:r></w:ins></w:hyperlink>"#;
        assert_eq!(
            resolve_raw_inline(link, true, None, None).map(texts),
            Some(vec!["H(new)".to_owned()])
        );
        let field = r#"<w:ins w:id="3"><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:ins>"#;
        assert_eq!(
            resolve_raw_inline(field, true, None, None).map(texts),
            Some(vec!["[PAGE|2]".to_owned()])
        );
        assert!(!raw_holds_changes("<w:r><w:t>plain</w:t></w:r>"));
        assert_eq!(
            resolve_raw_inline("<w:r><w:t>plain</w:t></w:r>", true, None, None),
            None
        );
    }

    fn run_text(run: &Run) -> String {
        run.content
            .iter()
            .filter_map(|content| match content {
                RunContent::Text { text, .. } => Some(text.as_str()),
                _ => None,
            })
            .collect()
    }

    #[test]
    fn a_change_holding_its_fields_own_field_character_moves_out_as_before() {
        let paragraph = parse(
            r#"<w:p xmlns:w="w">
              <w:r><w:fldChar w:fldCharType="begin"/></w:r>
              <w:r><w:instrText> DATE </w:instrText></w:r>
              <w:r><w:fldChar w:fldCharType="separate"/></w:r>
              <w:r><w:t>20</w:t></w:r>
              <w:del w:id="5"><w:r><w:delText>26</w:delText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:del>
            </w:p>"#,
        );
        let ParagraphContent::Tracked(change) = &paragraph.content[0] else {
            panic!("the deletion moves in front of its field")
        };
        assert_eq!(change.node_type, "deletion");
    }

    fn run_texts(paragraph: &Paragraph) -> Vec<&str> {
        paragraph
            .content
            .iter()
            .filter_map(|content| match content {
                ParagraphContent::Inline(InlineNode::Run(run)) => {
                    run.content.iter().find_map(|content| match content {
                        RunContent::Text { text, .. } => Some(text.as_str()),
                        _ => None,
                    })
                }
                _ => None,
            })
            .collect()
    }

    #[test]
    fn smart_tag_wrapped_runs_keep_their_text() {
        let paragraph = parse(
            r#"<w:p xmlns:w="w">
              <w:r><w:t>in </w:t></w:r>
              <w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="place">
                <w:smartTagPr><w:attr w:name="k" w:val="v"/></w:smartTagPr>
                <w:r><w:t>Paris</w:t></w:r>
              </w:smartTag>
            </w:p>"#,
        );
        assert_eq!(run_texts(&paragraph), ["in ", "Paris"]);
        assert_eq!(paragraph.content.len(), 2);
    }

    #[test]
    fn custom_xml_wrapped_runs_and_nested_smart_tags_keep_document_order() {
        let paragraph = parse(
            r#"<w:p xmlns:w="w">
              <w:r><w:t>a</w:t></w:r>
              <w:customXml w:uri="urn:x" w:element="e">
                <w:customXmlPr><w:placeholder w:val="p"/></w:customXmlPr>
                <w:r><w:t>b</w:t></w:r>
                <w:smartTag w:element="date"><w:r><w:t>c</w:t></w:r></w:smartTag>
              </w:customXml>
              <w:r><w:t>d</w:t></w:r>
            </w:p>"#,
        );
        assert_eq!(run_texts(&paragraph), ["a", "b", "c", "d"]);
        assert_eq!(paragraph.content.len(), 4);
    }

    #[test]
    fn paragraph_content_wire_uses_the_type_discriminator_for_same_shape_markers() {
        let comment: ParagraphContent =
            serde_json::from_str(r#"{"type":"commentRangeStart","id":7}"#).unwrap();
        assert!(matches!(comment, ParagraphContent::CommentRange(_)));
        let movement: ParagraphContent =
            serde_json::from_str(r#"{"type":"moveFromRangeEnd","id":8}"#).unwrap();
        assert!(matches!(movement, ParagraphContent::RangeEnd(_)));
    }

    #[test]
    fn enforces_paragraph_budget_before_allocating_content() {
        let mut limits = ParseLimits::default();
        limits.max_paragraphs = 0;
        let mut budget = ParseBudget::new(&limits);
        let document = parse_xml(
            br#"<w:p xmlns:w="w"><w:r><w:t>x</w:t></w:r></w:p>"#,
            "word/document.xml",
            &mut budget,
        )
        .unwrap();
        let mut ids = HexIdAllocator::from_sha256(&"0".repeat(64)).unwrap();
        let error = parse_paragraph(
            document.root().unwrap(),
            None,
            None,
            None,
            None,
            None,
            "word/document.xml",
            &mut budget,
            &mut ids,
            None,
            false,
            0,
        )
        .unwrap_err();
        assert!(matches!(
            error,
            ParseError::ResourceLimit {
                kind: "paragraphs",
                ..
            }
        ));
    }

    #[test]
    fn keeps_drawings_the_model_cannot_carry_as_raw_xml() {
        let paragraph = parse(
            r#"<w:p xmlns:w="w" xmlns:v="v" xmlns:o="o" xmlns:mc="mc" xmlns:wp="wp" xmlns:a="a" xmlns:wps="wps">
              <w:r><w:pict><v:rect id="r1"><v:textbox><w:txbxContent><w:p><w:r><w:t>Callout</w:t></w:r></w:p></w:txbxContent></v:textbox></v:rect></w:pict></w:r>
              <w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor><wp:extent cx="10" cy="10"/><a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent><w:p/></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape/></w:pict></mc:Fallback></mc:AlternateContent></w:r>
            </w:p>"#,
        );
        let opaque = paragraph
            .content
            .iter()
            .filter_map(|item| match item {
                ParagraphContent::Inline(InlineNode::Run(run)) => Some(run),
                _ => None,
            })
            .flat_map(|run| &run.content)
            .filter_map(|item| match item {
                RunContent::OpaqueDrawing { kind, xml } => Some((kind.as_str(), xml.as_str())),
                _ => None,
            })
            .collect::<Vec<_>>();
        // The text box enters through block enrichment, so only the VML rect stays raw.
        assert_eq!(opaque.len(), 1);
        assert_eq!(opaque[0].0, "pict");
        assert!(opaque[0].1.contains("<v:rect") && opaque[0].1.contains("Callout"));
    }
}
