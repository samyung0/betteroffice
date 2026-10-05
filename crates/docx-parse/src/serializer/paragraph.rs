//! Paragraph properties and inline-content serialization.

use crate::block::BlockContent;
use crate::borders::Borders;
use crate::formatting::{NumberingProperties, ParagraphFormatting, ParagraphFrame};
use crate::inline::{
    BookmarkEnd, BookmarkStart, ComplexField, Hyperlink, InlineNode, InlineSdt, MathEquation, Run,
    RunContent, SdtProperties, SimpleField,
};
use crate::paragraph::{
    CommentRange, Paragraph, ParagraphContent, ParagraphPropertyChange, RangeEnd, RangeStart,
    TrackedChangeInfo, TrackedInline,
};
use crate::section::SectionProperties;
use crate::xml::ParseError;

use super::context::SerializerContext;
use super::foundation::{BorderSide, write_border};
use super::raw::{validate_math_subtree, validate_raw_subtree, validate_replayed_fragment};
use super::run::{
    append_generated, deleted_inline_xml, nonempty, nonempty_trimmed, normalized_tracked_id,
    serialize_deleted_run, serialize_run, serialize_text_formatting, write_shading,
};
use super::section::serialize_section_properties;
use super::xml_writer::{XmlWriter, int_attr, js_number};

/// A replayed attribute name must be exactly one XML name, or it could
/// smuggle markup into the tag.
pub(crate) fn is_safe_attribute_name(name: &str) -> bool {
    !name.is_empty()
        && name.chars().all(|character| {
            character.is_ascii_alphanumeric() || matches!(character, ':' | '-' | '_' | '.')
        })
}

/// Serialize one complete `w:p` element.
pub fn serialize_paragraph(
    paragraph: &Paragraph,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    context.enter_paragraph(paragraph.rendered_page_break_before == Some(true));
    let result = serialize_paragraph_inner(paragraph, context);
    context.leave_paragraph();
    result
}

fn serialize_paragraph_inner(
    paragraph: &Paragraph,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let mut writer = XmlWriter::with_capacity(512);
    writer.start_element("w:p");
    if let Some(value) = nonempty(paragraph.para_id.as_deref()) {
        writer.attribute("w14:paraId", value);
    }
    if let Some(value) = nonempty(paragraph.text_id.as_deref()) {
        writer.attribute("w14:textId", value);
    }
    for attribute in &paragraph.extra_attributes {
        if is_safe_attribute_name(&attribute.name) {
            writer.dynamic_attribute(&attribute.name, &attribute.value);
        }
    }
    let properties = serialize_paragraph_formatting(
        paragraph.formatting.as_ref(),
        paragraph.property_changes.as_deref(),
        paragraph.p_pr_ins.as_ref(),
        paragraph.p_pr_del.as_ref(),
        false,
        paragraph.section_properties.as_ref(),
    )?;
    append_generated(&mut writer, &properties);
    for content in &paragraph.content {
        append_generated(&mut writer, &serialize_paragraph_content(content, context)?);
        if let ParagraphContent::CommentRange(marker) = content
            && marker.node_type == "commentRangeEnd"
            && context.keeps_comment(marker.id)
            && !context.references_comment(marker.id)
            && !paragraph_has_comment_reference(paragraph, marker.id)
        {
            writer
                .start_element("w:r")
                .start_element("w:rPr")
                .start_element("w:rStyle")
                .attribute("w:val", "CommentReference")
                .end_element()
                .end_element()
                .start_element("w:commentReference")
                .attribute("w:id", &js_number(marker.id))
                .end_element()
                .end_element();
            context.note_comment_reference(marker.id);
        }
    }
    writer.end_element();
    Ok(writer.finish())
}

#[allow(clippy::too_many_arguments)]
pub fn serialize_paragraph_formatting(
    formatting: Option<&ParagraphFormatting>,
    property_changes: Option<&[ParagraphPropertyChange]>,
    p_pr_ins: Option<&TrackedChangeInfo>,
    p_pr_del: Option<&TrackedChangeInfo>,
    base_only: bool,
    section_properties: Option<&SectionProperties>,
) -> Result<String, ParseError> {
    let mut body = XmlWriter::with_capacity(512);
    if let Some(formatting) = formatting {
        // CT_PPrBase order (ECMA-376 §17.3.1.26).
        let kept = |body: &mut XmlWriter, name: &'static str| {
            write_unmodeled(body, formatting, name);
        };
        if let Some(value) = nonempty(formatting.style_id.as_deref()) {
            empty_attr(&mut body, "w:pStyle", "w:val", value);
        }
        on_off(&mut body, "w:keepNext", formatting.keep_next);
        on_off(&mut body, "w:keepLines", formatting.keep_lines);
        on_off(&mut body, "w:pageBreakBefore", formatting.page_break_before);
        write_frame(&mut body, formatting.frame.as_ref());
        on_off(&mut body, "w:widowControl", formatting.widow_control);
        if formatting.num_pr != formatting.num_pr_from_style
            || formatting.num_pr_from_style.is_none()
        {
            write_numbering(&mut body, formatting.num_pr.as_ref());
        }
        on_off(
            &mut body,
            "w:suppressLineNumbers",
            formatting.suppress_line_numbers,
        );
        write_paragraph_borders(&mut body, formatting.borders.as_ref());
        write_shading(&mut body, formatting.shading.as_ref());
        write_tabs(&mut body, formatting.tabs.as_deref());
        on_off(
            &mut body,
            "w:suppressAutoHyphens",
            formatting.suppress_auto_hyphens,
        );
        kept(&mut body, "w:kinsoku");
        kept(&mut body, "w:wordWrap");
        kept(&mut body, "w:overflowPunct");
        kept(&mut body, "w:topLinePunct");
        on_off(&mut body, "w:autoSpaceDE", formatting.auto_space_de);
        on_off(&mut body, "w:autoSpaceDN", formatting.auto_space_dn);
        on_off(&mut body, "w:bidi", formatting.bidi);
        kept(&mut body, "w:adjustRightInd");
        on_off(&mut body, "w:snapToGrid", formatting.snap_to_grid);
        write_spacing(&mut body, formatting);
        write_indentation(&mut body, formatting);
        on_off(
            &mut body,
            "w:contextualSpacing",
            formatting.contextual_spacing,
        );
        kept(&mut body, "w:mirrorIndents");
        kept(&mut body, "w:suppressOverlap");
        if let Some(value) = nonempty(formatting.alignment.as_deref()) {
            empty_attr(&mut body, "w:jc", "w:val", value);
        }
        kept(&mut body, "w:textDirection");
        kept(&mut body, "w:textAlignment");
        kept(&mut body, "w:textboxTightWrap");
        if let Some(value) = formatting.outline_level {
            empty_attr(&mut body, "w:outlineLvl", "w:val", &js_number(value));
        }
        kept(&mut body, "w:divId");
        kept(&mut body, "w:cnfStyle");
    }

    if !base_only {
        write_paragraph_mark_properties(&mut body, formatting, p_pr_ins, p_pr_del);
        if let Some(properties) = section_properties {
            append_generated(&mut body, &serialize_section_properties(Some(properties)));
        }
        if let Some(change) = property_changes.and_then(|changes| changes.first()) {
            append_generated(&mut body, &serialize_paragraph_property_change(change)?);
        }
    }
    let body = body.finish();
    if body.is_empty() {
        Ok(String::new())
    } else {
        Ok(format!("<w:pPr>{body}</w:pPr>"))
    }
}

fn serialize_paragraph_property_change(
    change: &ParagraphPropertyChange,
) -> Result<String, ParseError> {
    let previous = serialize_paragraph_formatting(
        change.previous_formatting.as_ref(),
        None,
        None,
        None,
        true,
        None,
    )?;
    let previous = if previous.is_empty() {
        "<w:pPr/>".to_owned()
    } else {
        previous
    };
    let mut writer = XmlWriter::with_capacity(previous.len() + 128);
    writer
        .start_element("w:pPrChange")
        .attribute("w:id", &normalized_tracked_id(change.info.id))
        .attribute(
            "w:author",
            nonempty_trimmed(&change.info.author).unwrap_or("Unknown"),
        );
    if let Some(date) = change.info.date.as_deref().and_then(nonempty_trimmed) {
        writer.attribute("w:date", date);
    }
    append_generated(&mut writer, &previous);
    writer.end_element();
    Ok(writer.finish())
}

pub fn serialize_paragraph_content(
    content: &ParagraphContent,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    match content {
        ParagraphContent::Inline(node) => serialize_inline_node(node, context),
        ParagraphContent::Tracked(change) => serialize_tracked_change(change, context),
        ParagraphContent::RangeStart(marker) => serialize_range_start(marker),
        ParagraphContent::RangeEnd(marker) => serialize_range_end(marker),
        ParagraphContent::CommentRange(marker) if !context.keeps_comment(marker.id) => {
            Ok(String::new())
        }
        ParagraphContent::CommentRange(marker) => serialize_comment_range(marker),
    }
}

pub(crate) fn serialize_inline_node(
    node: &InlineNode,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    match node {
        InlineNode::Run(run) => serialize_run(run, context),
        InlineNode::Hyperlink(hyperlink) => serialize_hyperlink(hyperlink, context),
        InlineNode::BookmarkStart(bookmark) => Ok(serialize_bookmark_start(bookmark)),
        InlineNode::BookmarkEnd(bookmark) => Ok(serialize_bookmark_end(bookmark)),
        InlineNode::SimpleField(field) => serialize_simple_field(field, context),
        InlineNode::ComplexField(field) => serialize_complex_field(field, context),
        InlineNode::InlineSdt(sdt) => serialize_inline_sdt(sdt, context),
        InlineNode::Math(math) => serialize_math(math),
        InlineNode::Tracked(change) => serialize_tracked_change(change, context),
        InlineNode::RawXml(raw) => {
            validate_replayed_fragment(&raw.xml)?;
            Ok(raw.xml.clone())
        }
    }
}

fn serialize_bookmark_start(bookmark: &BookmarkStart) -> String {
    let mut writer = XmlWriter::with_capacity(96);
    writer
        .start_element("w:bookmarkStart")
        .attribute("w:id", &js_number(bookmark.id))
        .attribute("w:name", &bookmark.name);
    if let Some(value) = bookmark.col_first {
        writer.attribute("w:colFirst", &js_number(value));
    }
    if let Some(value) = bookmark.col_last {
        writer.attribute("w:colLast", &js_number(value));
    }
    writer.end_element();
    writer.finish()
}

fn serialize_bookmark_end(bookmark: &BookmarkEnd) -> String {
    let mut writer = XmlWriter::with_capacity(40);
    writer
        .start_element("w:bookmarkEnd")
        .attribute("w:id", &js_number(bookmark.id))
        .end_element();
    writer.finish()
}

fn serialize_hyperlink(
    hyperlink: &Hyperlink,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let mut children = String::new();
    for child in hyperlink.written_children() {
        children.push_str(&serialize_inline_node(child, context)?);
    }
    let has_attributes = nonempty(hyperlink.relationship_id.as_deref()).is_some()
        || nonempty(hyperlink.anchor.as_deref()).is_some()
        || nonempty(hyperlink.tooltip.as_deref()).is_some()
        || nonempty(hyperlink.target.as_deref()).is_some()
        || hyperlink.history.is_some()
        || nonempty(hyperlink.doc_location.as_deref()).is_some();
    if !has_attributes && nonempty(hyperlink.href.as_deref()).is_none() {
        return Ok(children);
    }
    let mut writer = XmlWriter::with_capacity(children.len() + 128);
    writer.start_element("w:hyperlink");
    optional_attr(&mut writer, "r:id", hyperlink.relationship_id.as_deref());
    optional_attr(&mut writer, "w:anchor", hyperlink.anchor.as_deref());
    optional_attr(&mut writer, "w:tooltip", hyperlink.tooltip.as_deref());
    optional_attr(&mut writer, "w:tgtFrame", hyperlink.target.as_deref());
    if let Some(history) = hyperlink.history {
        writer.attribute("w:history", if history { "1" } else { "0" });
    }
    optional_attr(
        &mut writer,
        "w:docLocation",
        hyperlink.doc_location.as_deref(),
    );
    append_generated(&mut writer, &children);
    writer.end_element();
    Ok(writer.finish())
}

fn serialize_simple_field(
    field: &SimpleField,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let mut output = String::new();
    output.push_str("<w:fldSimple w:instr=\"");
    output.push_str(&super::xml_writer::escape_xml(&field.instruction));
    output.push('"');
    if field.fld_lock == Some(true) {
        output.push_str(" w:fldLock=\"true\"");
    }
    if field.dirty == Some(true) {
        output.push_str(" w:dirty=\"true\"");
    }
    output.push('>');
    if let Some(nodes) = field.written_result() {
        for node in nodes {
            output.push_str(&serialize_inline_node(node, context)?);
        }
    } else {
        for run in &field.content {
            output.push_str(&serialize_run(run, context)?);
        }
    }
    output.push_str("</w:fldSimple>");
    Ok(output)
}

fn serialize_complex_field(
    field: &ComplexField,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let formatting = field
        .field_result
        .first()
        .map(|run| run.formatting.as_ref())
        .unwrap_or(field.formatting.as_ref());
    let properties = serialize_text_formatting(formatting);
    let mut output = String::new();
    output.push_str("<w:r>");
    output.push_str(&properties);
    output.push_str("<w:fldChar w:fldCharType=\"begin\"");
    if field.fld_lock == Some(true) {
        output.push_str(" w:fldLock=\"true\"");
    }
    if field.dirty == Some(true) {
        output.push_str(" w:dirty=\"true\"");
    }
    output.push_str("/></w:r>");
    if let Some(nodes) = field.written_code() {
        for node in nodes {
            output.push_str(&serialize_inline_node(node, context)?);
        }
    } else if field.field_code.is_empty() {
        output.push_str("<w:r>");
        output.push_str(&properties);
        output.push_str("<w:instrText");
        if needs_preserve(&field.instruction) {
            output.push_str(" xml:space=\"preserve\"");
        }
        output.push('>');
        output.push_str(&super::xml_writer::escape_xml(&field.instruction));
        output.push_str("</w:instrText></w:r>");
    } else {
        for run in &field.field_code {
            output.push_str(&serialize_run(run, context)?);
        }
    }
    if !field
        .continuation
        .as_ref()
        .is_some_and(|value| value.separate)
    {
        output.push_str("<w:r>");
        output.push_str(&properties);
        output.push_str("<w:fldChar w:fldCharType=\"separate\"/></w:r>");
    }
    // Run-level fallback: multi-block results and fields rebuilt by the edit path.
    if let Some(nodes) = field.written_result() {
        for node in nodes {
            output.push_str(&serialize_inline_node(node, context)?);
        }
    } else {
        for run in &field.field_result {
            output.push_str(&serialize_run(run, context)?);
        }
    }
    if !field.continuation.as_ref().is_some_and(|value| value.end) {
        output.push_str("<w:r>");
        output.push_str(&properties);
        output.push_str("<w:fldChar w:fldCharType=\"end\"/></w:r>");
    }
    Ok(output)
}

pub fn serialize_inline_sdt(
    sdt: &InlineSdt,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let (properties, end_properties) = serialize_sdt_properties(&sdt.properties)?;
    let mut writer = XmlWriter::with_capacity(256);
    writer.start_element("w:sdt");
    append_generated(&mut writer, &properties);
    append_generated(&mut writer, &end_properties);
    writer.start_element("w:sdtContent");
    for item in &sdt.content {
        append_generated(&mut writer, &serialize_inline_node(item, context)?);
    }
    writer.end_element().end_element();
    Ok(writer.finish())
}

pub(crate) fn serialize_sdt_properties(
    properties: &SdtProperties,
) -> Result<(String, String), ParseError> {
    let properties_xml = if let Some(raw) = properties.raw_properties_xml.as_deref() {
        validate_raw_subtree(raw, "w", "sdtPr")?;
        raw.to_owned()
    } else {
        synthesize_sdt_properties(properties)
    };
    let end_properties = if let Some(raw) = properties.raw_end_properties_xml.as_deref() {
        validate_raw_subtree(raw, "w", "sdtEndPr")?;
        raw.to_owned()
    } else {
        String::new()
    };
    Ok((properties_xml, end_properties))
}

pub fn synthesize_sdt_properties(properties: &SdtProperties) -> String {
    let mut body = XmlWriter::with_capacity(256);
    if let Some(value) = nonempty(properties.alias.as_deref()) {
        empty_attr(&mut body, "w:alias", "w:val", value);
    }
    if let Some(value) = nonempty(properties.tag.as_deref()) {
        empty_attr(&mut body, "w:tag", "w:val", value);
    }
    if let Some(value) = properties.id {
        empty_attr(&mut body, "w:id", "w:val", &js_number(value));
    }
    if let Some(value) = nonempty(properties.lock.as_deref()).filter(|value| *value != "unlocked") {
        empty_attr(&mut body, "w:lock", "w:val", value);
    }
    if let Some(value) = nonempty(properties.placeholder.as_deref()) {
        body.start_element("w:placeholder");
        empty_attr(&mut body, "w:docPart", "w:val", value);
        body.end_element();
    }
    if properties.showing_placeholder == Some(true) {
        body.start_element("w:showingPlcHdr").end_element();
    }
    match properties.sdt_type.as_str() {
        "plainText" => {
            body.start_element("w:text").end_element();
        }
        "date" => {
            body.start_element("w:date");
            if let Some(value) = nonempty(properties.date_format.as_deref()) {
                empty_attr(&mut body, "w:dateFormat", "w:val", value);
            }
            body.end_element();
        }
        "dropDownList" | "comboBox" => {
            let name = if properties.sdt_type == "dropDownList" {
                "w:dropDownList"
            } else {
                "w:comboBox"
            };
            body.start_element(name).attribute("w:lastValue", "");
            for item in properties.list_items.as_deref().unwrap_or_default() {
                body.start_element("w:listItem")
                    .attribute("w:displayText", &item.display_text)
                    .attribute("w:value", &item.value)
                    .end_element();
            }
            body.end_element();
        }
        "checkbox" => {
            body.start_element("w14:checkbox");
            body.start_element("w14:checked")
                .attribute(
                    "w14:val",
                    if properties.checked == Some(true) {
                        "1"
                    } else {
                        "0"
                    },
                )
                .end_element()
                .start_element("w14:checkedState")
                .attribute("w14:val", "2612")
                .attribute("w14:font", "MS Gothic")
                .end_element()
                .start_element("w14:uncheckedState")
                .attribute("w14:val", "2610")
                .attribute("w14:font", "MS Gothic")
                .end_element()
                .end_element();
        }
        "picture" => {
            body.start_element("w:picture").end_element();
        }
        _ => {}
    }
    format!("<w:sdtPr>{}</w:sdtPr>", body.finish())
}

fn serialize_tracked_change(
    change: &TrackedInline,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let (element, deletion) = match change.node_type.as_str() {
        "insertion" => ("w:ins", false),
        "deletion" => ("w:del", true),
        "moveFrom" => ("w:moveFrom", true),
        "moveTo" => ("w:moveTo", false),
        _ => return Ok(String::new()),
    };
    let mut writer = XmlWriter::with_capacity(256);
    writer
        .start_element(element)
        .attribute("w:id", &normalized_tracked_id(change.info.id))
        .attribute(
            "w:author",
            nonempty_trimmed(&change.info.author).unwrap_or("Unknown"),
        );
    if let Some(date) = change.info.date.as_deref().and_then(nonempty_trimmed) {
        writer.attribute("w:date", date);
    }
    // Tracked wrappers use explicit start and end tags.
    writer.text("");
    for item in &change.content {
        match item {
            InlineNode::Run(run) => append_generated(
                &mut writer,
                &if deletion {
                    serialize_deleted_run(run, context)?
                } else {
                    serialize_run(run, context)?
                },
            ),
            _ => {
                let xml = serialize_inline_node(item, context)?;
                append_generated(
                    &mut writer,
                    &if deletion {
                        deleted_inline_xml(xml)
                    } else {
                        xml
                    },
                );
            }
        };
    }
    writer.end_element();
    Ok(writer.finish())
}

fn serialize_range_start(marker: &RangeStart) -> Result<String, ParseError> {
    let element = match marker.node_type.as_str() {
        "moveFromRangeStart" => "w:moveFromRangeStart",
        "moveToRangeStart" => "w:moveToRangeStart",
        _ => return Ok(String::new()),
    };
    let mut writer = XmlWriter::with_capacity(80);
    writer
        .start_element(element)
        .attribute("w:id", &js_number(marker.id))
        .attribute("w:name", &marker.name)
        .end_element();
    Ok(writer.finish())
}

fn serialize_range_end(marker: &RangeEnd) -> Result<String, ParseError> {
    let element = match marker.node_type.as_str() {
        "moveFromRangeEnd" => "w:moveFromRangeEnd",
        "moveToRangeEnd" => "w:moveToRangeEnd",
        _ => return Ok(String::new()),
    };
    let mut writer = XmlWriter::with_capacity(48);
    writer
        .start_element(element)
        .attribute("w:id", &js_number(marker.id))
        .end_element();
    Ok(writer.finish())
}

fn run_comment_references(run: &Run, found: &mut impl FnMut(f64)) {
    for content in &run.content {
        if let RunContent::CommentReference { id: Some(id) } = content {
            found(*id);
        }
    }
}

/// Calls `found` with the id of each comment reference `node` writes.
fn comment_references(node: &InlineNode, found: &mut impl FnMut(f64)) {
    match node {
        InlineNode::Run(run) => run_comment_references(run, found),
        InlineNode::Hyperlink(hyperlink) => {
            for node in hyperlink.written_children() {
                if matches!(node, InlineNode::Run(_) | InlineNode::SimpleField(_)) {
                    comment_references(node, found);
                }
            }
        }
        InlineNode::InlineSdt(sdt) => {
            for node in &sdt.content {
                comment_references(node, found);
            }
        }
        InlineNode::SimpleField(field) => match field.written_result() {
            Some(nodes) => nodes
                .iter()
                .for_each(|node| comment_references(node, found)),
            None => field
                .content
                .iter()
                .for_each(|run| run_comment_references(run, found)),
        },
        InlineNode::ComplexField(field) => {
            match field.written_code() {
                Some(nodes) => nodes
                    .iter()
                    .for_each(|node| comment_references(node, found)),
                None => field
                    .field_code
                    .iter()
                    .for_each(|run| run_comment_references(run, found)),
            }
            match field.written_result() {
                Some(nodes) => nodes
                    .iter()
                    .for_each(|node| comment_references(node, found)),
                None => field
                    .field_result
                    .iter()
                    .for_each(|run| run_comment_references(run, found)),
            }
        }
        _ => {}
    }
}

fn paragraph_comment_references(paragraph: &Paragraph, found: &mut impl FnMut(f64)) {
    for content in &paragraph.content {
        match content {
            ParagraphContent::Inline(node) => comment_references(node, found),
            ParagraphContent::Tracked(change) => {
                let deletion = matches!(change.node_type.as_str(), "deletion" | "moveFrom");
                for node in &change.content {
                    if matches!(node, InlineNode::Run(_) | InlineNode::Hyperlink(_))
                        || matches!(node, InlineNode::SimpleField(_) if !deletion)
                    {
                        comment_references(node, found);
                    }
                }
            }
            _ => {}
        }
    }
}

fn paragraph_has_comment_reference(paragraph: &Paragraph, id: f64) -> bool {
    let mut held = false;
    paragraph_comment_references(paragraph, &mut |reference| held |= reference == id);
    held
}

/// Calls `found` with the id of each comment reference `blocks` write.
pub(crate) fn block_comment_references(blocks: &[BlockContent], found: &mut impl FnMut(f64)) {
    for block in blocks {
        match block {
            BlockContent::Paragraph(paragraph) => paragraph_comment_references(paragraph, found),
            BlockContent::Table(table) => {
                for row in &table.rows {
                    for cell in &row.cells {
                        block_comment_references(&cell.content, found);
                    }
                }
            }
            BlockContent::BlockSdt(sdt) => block_comment_references(&sdt.content, found),
            BlockContent::RawXml(_) => {}
        }
    }
}

fn serialize_comment_range(marker: &CommentRange) -> Result<String, ParseError> {
    let mut writer = XmlWriter::with_capacity(180);
    match marker.node_type.as_str() {
        "commentRangeStart" => {
            writer
                .start_element("w:commentRangeStart")
                .attribute("w:id", &js_number(marker.id))
                .end_element();
        }
        "commentRangeEnd" => {
            writer
                .start_element("w:commentRangeEnd")
                .attribute("w:id", &js_number(marker.id))
                .end_element();
        }
        _ => {}
    }
    Ok(writer.finish())
}

fn serialize_math(math: &MathEquation) -> Result<String, ParseError> {
    if math.omml_xml.is_empty() {
        return Ok(String::new());
    }
    validate_math_subtree(&math.omml_xml)?;
    Ok(math.omml_xml.clone())
}

fn write_paragraph_mark_properties(
    writer: &mut XmlWriter,
    formatting: Option<&ParagraphFormatting>,
    insertion: Option<&TrackedChangeInfo>,
    deletion: Option<&TrackedChangeInfo>,
) {
    let formatting = formatting
        .and_then(|formatting| formatting.run_properties.as_ref())
        .map(|formatting| serialize_text_formatting(Some(formatting)))
        .unwrap_or_default();
    let inner = formatting
        .strip_prefix("<w:rPr>")
        .and_then(|value| value.strip_suffix("</w:rPr>"))
        .unwrap_or_default();
    if insertion.is_none() && deletion.is_none() && inner.is_empty() {
        return;
    }
    writer.start_element("w:rPr");
    if let Some(info) = insertion {
        write_mark_change(writer, "w:ins", info);
    }
    if let Some(info) = deletion {
        write_mark_change(writer, "w:del", info);
    }
    append_generated(writer, inner);
    writer.end_element();
}

fn write_mark_change(writer: &mut XmlWriter, element: &'static str, info: &TrackedChangeInfo) {
    writer
        .start_element(element)
        .attribute("w:id", &normalized_tracked_id(info.id))
        .attribute(
            "w:author",
            if info.author.is_empty() {
                "Unknown"
            } else {
                &info.author
            },
        );
    if let Some(date) = nonempty(info.date.as_deref()) {
        writer.attribute("w:date", date);
    }
    writer.end_element();
}

fn write_paragraph_borders(writer: &mut XmlWriter, borders: Option<&Borders>) {
    let Some(borders) = borders else {
        return;
    };
    let entries = [
        (borders.top.as_ref(), BorderSide::Top),
        (borders.left.as_ref(), BorderSide::Left),
        (borders.bottom.as_ref(), BorderSide::Bottom),
        (borders.right.as_ref(), BorderSide::Right),
        (borders.between.as_ref(), BorderSide::Between),
        (borders.bar.as_ref(), BorderSide::Bar),
    ];
    if !entries.iter().any(|(border, _)| border.is_some()) {
        return;
    }
    writer.start_element("w:pBdr");
    for (border, side) in entries {
        if let Some(border) = border {
            write_border(writer, border, side);
        }
    }
    writer.end_element();
}

fn write_tabs(writer: &mut XmlWriter, tabs: Option<&[crate::tabs::TabStop]>) {
    let Some(tabs) = tabs.filter(|tabs| !tabs.is_empty()) else {
        return;
    };
    writer.start_element("w:tabs");
    for tab in tabs {
        writer
            .start_element("w:tab")
            .attribute("w:val", &tab.alignment)
            .attribute("w:pos", &int_attr(Some(tab.position)));
        if let Some(leader) = nonempty(tab.leader.as_deref()).filter(|value| *value != "none") {
            writer.attribute("w:leader", leader);
        }
        writer.end_element();
    }
    writer.end_element();
}

fn write_spacing(writer: &mut XmlWriter, formatting: &ParagraphFormatting) {
    if formatting.space_before.is_none()
        && formatting.space_after.is_none()
        && formatting.space_before_lines.is_none()
        && formatting.space_after_lines.is_none()
        && formatting.line_spacing.is_none()
        && nonempty(formatting.line_spacing_rule.as_deref()).is_none()
        && formatting.before_autospacing.is_none()
        && formatting.after_autospacing.is_none()
    {
        return;
    }
    writer.start_element("w:spacing");
    optional_int(writer, "w:before", formatting.space_before);
    optional_int(writer, "w:after", formatting.space_after);
    optional_int(writer, "w:beforeLines", formatting.space_before_lines);
    optional_int(writer, "w:afterLines", formatting.space_after_lines);
    optional_int(writer, "w:line", formatting.line_spacing);
    optional_attr(
        writer,
        "w:lineRule",
        formatting.line_spacing_rule.as_deref(),
    );
    if let Some(auto) = formatting.before_autospacing {
        writer.attribute("w:beforeAutospacing", if auto { "1" } else { "0" });
    }
    if let Some(auto) = formatting.after_autospacing {
        writer.attribute("w:afterAutospacing", if auto { "1" } else { "0" });
    }
    writer.end_element();
}

fn write_indentation(writer: &mut XmlWriter, formatting: &ParagraphFormatting) {
    let hanging = formatting.hanging_indent == Some(true);
    if formatting.indent_left.is_none()
        && formatting.indent_right.is_none()
        && formatting.indent_first_line.is_none()
        && formatting.indent_left_chars.is_none()
        && formatting.indent_right_chars.is_none()
        && formatting.indent_first_line_chars.is_none()
    {
        return;
    }
    writer.start_element("w:ind");
    optional_int(writer, "w:left", formatting.indent_left);
    optional_int(writer, "w:leftChars", formatting.indent_left_chars);
    optional_int(writer, "w:right", formatting.indent_right);
    optional_int(writer, "w:rightChars", formatting.indent_right_chars);
    if let Some(first) = formatting.indent_first_line {
        let name = if hanging { "w:hanging" } else { "w:firstLine" };
        writer.attribute(
            name,
            &int_attr(Some(if hanging { first.abs() } else { first })),
        );
    }
    if let Some(first) = formatting.indent_first_line_chars {
        let hanging = formatting
            .hanging_indent_chars
            .unwrap_or_else(|| first.is_sign_negative());
        let name = if hanging {
            "w:hangingChars"
        } else {
            "w:firstLineChars"
        };
        writer.attribute(name, &int_attr(Some(first.abs())));
    }
    writer.end_element();
}

fn write_numbering(writer: &mut XmlWriter, numbering: Option<&NumberingProperties>) {
    let Some(numbering) = numbering else {
        return;
    };
    if numbering.ilvl.is_none() && numbering.num_id.is_none() {
        return;
    }
    writer.start_element("w:numPr");
    if let Some(value) = numbering.ilvl {
        empty_attr(writer, "w:ilvl", "w:val", &int_attr(Some(value)));
    }
    if let Some(value) = numbering.num_id {
        empty_attr(writer, "w:numId", "w:val", &int_attr(Some(value)));
    }
    writer.end_element();
}

fn write_frame(writer: &mut XmlWriter, frame: Option<&ParagraphFrame>) {
    let Some(frame) = frame else {
        return;
    };
    let has_attributes = frame.width.is_some()
        || frame.height.is_some()
        || nonempty(frame.h_anchor.as_deref()).is_some()
        || nonempty(frame.v_anchor.as_deref()).is_some()
        || frame.x.is_some()
        || frame.y.is_some()
        || nonempty(frame.x_align.as_deref()).is_some()
        || nonempty(frame.y_align.as_deref()).is_some()
        || nonempty(frame.wrap.as_deref()).is_some()
        || nonempty(frame.drop_cap.as_deref()).is_some()
        || frame.lines.is_some()
        || frame.h_space.is_some()
        || frame.v_space.is_some()
        || nonempty(frame.h_rule.as_deref()).is_some()
        || nonempty(frame.anchor_lock.as_deref()).is_some();
    if !has_attributes {
        return;
    }
    writer.start_element("w:framePr");
    optional_int(writer, "w:w", frame.width);
    optional_int(writer, "w:h", frame.height);
    optional_attr(writer, "w:hAnchor", frame.h_anchor.as_deref());
    optional_attr(writer, "w:vAnchor", frame.v_anchor.as_deref());
    optional_int(writer, "w:x", frame.x);
    optional_int(writer, "w:y", frame.y);
    optional_attr(writer, "w:xAlign", frame.x_align.as_deref());
    optional_attr(writer, "w:yAlign", frame.y_align.as_deref());
    optional_attr(writer, "w:wrap", frame.wrap.as_deref());
    optional_attr(writer, "w:dropCap", frame.drop_cap.as_deref());
    optional_int(writer, "w:lines", frame.lines);
    optional_int(writer, "w:hSpace", frame.h_space);
    optional_int(writer, "w:vSpace", frame.v_space);
    optional_attr(writer, "w:hRule", frame.h_rule.as_deref());
    optional_attr(writer, "w:anchorLock", frame.anchor_lock.as_deref());
    writer.end_element();
}

/// One kept [`crate::formatting::UNMODELED_PPR`] child, as the source wrote it.
fn write_unmodeled(writer: &mut XmlWriter, formatting: &ParagraphFormatting, name: &'static str) {
    let Some(attributes) = name
        .strip_prefix("w:")
        .and_then(|local| formatting.extra_children.get(local))
    else {
        return;
    };
    writer.start_element(name);
    for (attribute, value) in attributes {
        let attribute = format!("w:{attribute}");
        if is_safe_attribute_name(&attribute) {
            writer.dynamic_attribute(&attribute, value);
        }
    }
    writer.end_element();
}

fn needs_preserve(value: &str) -> bool {
    value.starts_with(' ') || value.ends_with(' ') || value.contains("  ")
}

fn optional_attr(writer: &mut XmlWriter, name: &'static str, value: Option<&str>) {
    if let Some(value) = nonempty(value) {
        writer.attribute(name, value);
    }
}

fn optional_int(writer: &mut XmlWriter, name: &'static str, value: Option<f64>) {
    if value.is_some() {
        writer.attribute(name, &int_attr(value));
    }
}

fn empty_attr(writer: &mut XmlWriter, element: &'static str, name: &'static str, value: &str) {
    writer
        .start_element(element)
        .attribute(name, value)
        .end_element();
}

fn on_off(writer: &mut XmlWriter, name: &'static str, value: Option<bool>) {
    match value {
        Some(true) => {
            writer.start_element(name).end_element();
        }
        Some(false) => {
            writer
                .start_element(name)
                .attribute("w:val", "0")
                .end_element();
        }
        None => {}
    }
}

#[cfg(test)]
mod tests {
    use crate::inline::{Run, RunContent, RunType};
    use crate::serializer::s10::SerializerDeterminism;

    use super::*;

    fn context() -> SerializerContext {
        SerializerContext::new(&SerializerDeterminism {
            seed: "0".repeat(64),
            now: "2000-01-01T00:00:00.000Z".to_owned(),
        })
        .unwrap()
    }

    #[test]
    fn a_simple_field_saves_back_as_a_simple_field() {
        let field = SimpleField {
            node_type: crate::inline::SimpleFieldType::SimpleField,
            field_type: "PAGE".to_owned(),
            instruction: " PAGE ".to_owned(),
            content: vec![Run {
                node_type: RunType::Run,
                formatting: None,
                property_changes: None,
                content: vec![RunContent::Text {
                    text: "3".to_owned(),
                    preserve_space: None,
                }],
            }],
            fld_lock: Some(true),
            dirty: Some(true),
            structured_result: None,
            field_tree: None,
        };
        let output = serialize_simple_field(&field, &mut context()).unwrap();
        assert_eq!(
            output,
            "<w:fldSimple w:instr=\" PAGE \" w:fldLock=\"true\" w:dirty=\"true\">\
             <w:r><w:t>3</w:t></w:r></w:fldSimple>"
        );
    }

    #[test]
    fn paragraph_bytes_pin_ids_properties_and_rendered_break_injection() {
        let paragraph = Paragraph {
            node_type: "paragraph".to_owned(),
            para_id: Some("AA&BB\"CC".to_owned()),
            text_id: None,
            extra_attributes: Vec::new(),
            formatting: Some(ParagraphFormatting {
                keep_next: Some(false),
                alignment: Some("center".to_owned()),
                ..ParagraphFormatting::default()
            }),
            property_changes: None,
            p_pr_ins: None,
            p_pr_del: None,
            content: vec![ParagraphContent::Inline(InlineNode::Run(Run {
                node_type: RunType::Run,
                formatting: None,
                property_changes: None,
                content: vec![RunContent::Text {
                    text: "hello & goodbye".to_owned(),
                    preserve_space: None,
                }],
            }))],
            list_rendering: None,
            rendered_page_break_before: Some(true),
            section_properties: None,
        };
        assert_eq!(
            serialize_paragraph(&paragraph, &mut context()).unwrap(),
            "<w:p w14:paraId=\"AA&amp;BB&quot;CC\"><w:pPr><w:keepNext w:val=\"0\"/><w:jc w:val=\"center\"/></w:pPr><w:r><w:lastRenderedPageBreak/><w:t>hello &amp; goodbye</w:t></w:r></w:p>"
        );
    }

    /// Word reads pPr children in schema order; the ones the model has no
    /// field for come back where they stood.
    #[test]
    fn paragraph_properties_round_trip_in_schema_order_with_unmodeled_children() {
        let source = concat!(
            "<w:pPr><w:pStyle w:val=\"Body\"/><w:keepNext/><w:keepLines w:val=\"0\"/>",
            "<w:pageBreakBefore/><w:framePr w:w=\"2000\" w:hAnchor=\"text\" w:wrap=\"around\" ",
            "w:dropCap=\"drop\" w:lines=\"3\" w:hSpace=\"72\" w:vSpace=\"0\" w:hRule=\"exact\" ",
            "w:anchorLock=\"1\"/><w:widowControl w:val=\"0\"/>",
            "<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"2\"/></w:numPr><w:suppressLineNumbers/>",
            "<w:suppressAutoHyphens/><w:kinsoku w:val=\"0\"/><w:wordWrap w:val=\"0\"/>",
            "<w:overflowPunct w:val=\"0\"/><w:topLinePunct/><w:autoSpaceDE w:val=\"0\"/>",
            "<w:autoSpaceDN w:val=\"0\"/><w:bidi/><w:adjustRightInd w:val=\"0\"/>",
            "<w:snapToGrid w:val=\"0\"/><w:spacing w:after=\"0\"/><w:ind w:left=\"100\"/>",
            "<w:contextualSpacing/><w:mirrorIndents/><w:suppressOverlap/><w:jc w:val=\"both\"/>",
            "<w:textDirection w:val=\"tbRl\"/><w:textAlignment w:val=\"center\"/>",
            "<w:textboxTightWrap w:val=\"allLines\"/><w:outlineLvl w:val=\"8\"/>",
            "<w:divId w:val=\"123\"/><w:cnfStyle w:firstRow=\"1\" w:val=\"100000000000\"/></w:pPr>"
        );
        let limits = crate::xml::ParseLimits::default();
        let document = crate::xml::parse_xml(
            format!("<w:p xmlns:w=\"urn:w\">{source}</w:p>").as_bytes(),
            "p.xml",
            &mut crate::xml::ParseBudget::new(&limits),
        )
        .unwrap();
        let p_pr = document.root().unwrap().child("w", "pPr").unwrap().clone();
        let formatting =
            crate::paragraph::parse_document_paragraph_properties(&p_pr, None, None).unwrap();
        let saved =
            serialize_paragraph_formatting(Some(&formatting), None, None, None, false, None)
                .unwrap();
        assert_eq!(saved, source);

        // An unprefixed attribute, which the parser also reads, keeps its value.
        let document = crate::xml::parse_xml(
            b"<w:p xmlns:w=\"urn:w\"><w:pPr><w:kinsoku val=\"0\"/></w:pPr></w:p>",
            "p.xml",
            &mut crate::xml::ParseBudget::new(&limits),
        )
        .unwrap();
        let p_pr = document.root().unwrap().child("w", "pPr").unwrap().clone();
        let formatting =
            crate::paragraph::parse_document_paragraph_properties(&p_pr, None, None).unwrap();
        assert_eq!(
            serialize_paragraph_formatting(Some(&formatting), None, None, None, false, None)
                .unwrap(),
            "<w:pPr><w:kinsoku w:val=\"0\"/></w:pPr>"
        );
    }

    /// A hanging character indent of zero has no sign to carry its kind, so
    /// the flag has to.
    #[test]
    fn a_zero_hanging_character_indent_stays_hanging() {
        let mut writer = XmlWriter::with_capacity(64);
        write_indentation(
            &mut writer,
            &ParagraphFormatting {
                indent_first_line_chars: Some(0.0),
                hanging_indent_chars: Some(true),
                ..ParagraphFormatting::default()
            },
        );
        assert_eq!(writer.finish(), "<w:ind w:hangingChars=\"0\"/>");
    }

    #[test]
    fn indentation_keeps_character_units_and_an_explicit_zero_first_line() {
        let mut writer = XmlWriter::with_capacity(128);
        write_indentation(
            &mut writer,
            &ParagraphFormatting {
                indent_left: Some(200.0),
                indent_left_chars: Some(100.0),
                indent_first_line_chars: Some(200.0),
                indent_first_line: Some(420.0),
                ..ParagraphFormatting::default()
            },
        );
        assert_eq!(
            writer.finish(),
            "<w:ind w:left=\"200\" w:leftChars=\"100\" w:firstLine=\"420\" w:firstLineChars=\"200\"/>"
        );
        let mut writer = XmlWriter::with_capacity(128);
        write_indentation(
            &mut writer,
            &ParagraphFormatting {
                indent_first_line: Some(0.0),
                indent_first_line_chars: Some(0.0),
                ..ParagraphFormatting::default()
            },
        );
        assert_eq!(
            writer.finish(),
            "<w:ind w:firstLine=\"0\" w:firstLineChars=\"0\"/>"
        );
        let mut writer = XmlWriter::with_capacity(128);
        write_indentation(
            &mut writer,
            &ParagraphFormatting {
                hanging_indent: Some(true),
                indent_first_line: Some(-315.0),
                indent_first_line_chars: Some(-150.0),
                ..ParagraphFormatting::default()
            },
        );
        assert_eq!(
            writer.finish(),
            "<w:ind w:hanging=\"315\" w:hangingChars=\"150\"/>"
        );
        let mut writer = XmlWriter::with_capacity(128);
        write_indentation(
            &mut writer,
            &ParagraphFormatting {
                indent_first_line: Some(420.0),
                indent_first_line_chars: Some(-200.0),
                ..ParagraphFormatting::default()
            },
        );
        assert_eq!(
            writer.finish(),
            "<w:ind w:firstLine=\"420\" w:hangingChars=\"200\"/>"
        );
        let mut writer = XmlWriter::with_capacity(128);
        write_indentation(
            &mut writer,
            &ParagraphFormatting {
                indent_first_line_chars: Some(-0.0),
                ..ParagraphFormatting::default()
            },
        );
        assert_eq!(writer.finish(), "<w:ind w:hangingChars=\"0\"/>");
    }

    #[test]
    fn empty_tracked_wrappers_use_explicit_end_tags() {
        let change = TrackedInline {
            node_type: "deletion".to_owned(),
            info: TrackedChangeInfo {
                id: 7.0,
                author: "Ada".to_owned(),
                date: None,
            },
            content: Vec::new(),
        };
        assert_eq!(
            serialize_tracked_change(&change, &mut context()).unwrap(),
            "<w:del w:id=\"7\" w:author=\"Ada\"></w:del>"
        );
    }

    #[test]
    fn rejects_unvalidated_math_and_sdt_raw_xml() {
        let mut properties = SdtProperties {
            sdt_type: "richText".to_owned(),
            id: None,
            alias: None,
            tag: None,
            lock: None,
            placeholder: None,
            showing_placeholder: None,
            date_format: None,
            list_items: None,
            checked: None,
            run_properties: None,
            temporary: None,
            label: None,
            tab_index: None,
            multi_line: None,
            date_state: None,
            list_last_value: None,
            checked_state: None,
            unchecked_state: None,
            gallery: None,
            appearance: None,
            color: None,
            control_state: None,
            repeating_section: None,
            repeating_section_item: None,
            data_binding: None,
            raw_properties_xml: Some("<w:sdtPr/><evil/>".to_owned()),
            raw_end_properties_xml: None,
        };
        assert!(serialize_sdt_properties(&properties).is_err());
        properties.raw_properties_xml = None;
        assert!(serialize_sdt_properties(&properties).is_ok());
    }
}
