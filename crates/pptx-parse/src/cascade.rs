//! Where a placeholder's paragraph properties come from. The renderer and the
//! editor fold these sources in order, each with its own merge.

use crate::{ParagraphProperties, Placeholder, ShapeNode, SlideMaster, TextBody, TextParagraph};

/// The node among `nodes` a `target` placeholder inherits from: the same
/// `idx`, else the same type.
pub fn find_placeholder<'a>(nodes: &'a [ShapeNode], target: &Placeholder) -> Option<&'a ShapeNode> {
    for node in nodes {
        if node_placeholder(node).is_some_and(|value| placeholders_match(value, target)) {
            return Some(node);
        }
        if let ShapeNode::Group(group) = node
            && let Some(found) = find_placeholder(&group.children, target)
        {
            return Some(found);
        }
    }
    None
}

/// `p:ph@type`, with `ctrTitle` as `title` and `obj` or none as `body`.
fn normalize_placeholder_type(value: Option<&str>) -> &str {
    match value.unwrap_or("body") {
        "ctrTitle" => "title",
        "obj" => "body",
        value => value,
    }
}

/// The master's `p:txStyles` entry a shape's paragraph at `level` starts from.
pub fn master_text_style<'a>(
    master: &'a SlideMaster,
    placeholder: Option<&Placeholder>,
    level: u32,
) -> Option<&'a ParagraphProperties> {
    let styles = match placeholder
        .map(|placeholder| normalize_placeholder_type(placeholder.placeholder_type.as_deref()))
    {
        Some("title") => &master.text_styles.title,
        Some("body" | "subTitle") => &master.text_styles.body,
        _ => &master.text_styles.other,
    };
    styles.get(level as usize).or_else(|| styles.first())
}

/// A layout or master body's paragraph for a laid-out one: by position,
/// else by level.
pub fn inherited_paragraph(
    body: Option<&TextBody>,
    index: usize,
    level: u32,
) -> Option<&TextParagraph> {
    let body = body?;
    body.paragraphs
        .get(index)
        .or_else(|| body.paragraphs.get(level as usize))
}

/// Paragraph properties after the master text style, least specific first:
/// for each `(body, paragraph)`, its `a:defPPr`, its `a:lvlNpPr`, then the
/// paragraph's own.
pub fn paragraph_cascade<'a>(
    bodies: [(Option<&'a TextBody>, Option<&'a TextParagraph>); 3],
    level: u32,
) -> impl Iterator<Item = &'a ParagraphProperties> {
    bodies
        .into_iter()
        .filter_map(|(body, paragraph)| Some((body?, paragraph)))
        .flat_map(move |(body, paragraph)| {
            body.default_list_style
                .as_deref()
                .into_iter()
                .chain(body.list_style.get(level as usize))
                .chain(paragraph.map(|paragraph| &paragraph.properties))
        })
}

fn placeholders_match(left: &Placeholder, right: &Placeholder) -> bool {
    match (left.index, right.index) {
        (Some(left), Some(right)) => left == right,
        _ => {
            normalize_placeholder_type(left.placeholder_type.as_deref())
                == normalize_placeholder_type(right.placeholder_type.as_deref())
        }
    }
}

fn node_placeholder(node: &ShapeNode) -> Option<&Placeholder> {
    match node {
        ShapeNode::Shape(shape) => shape.base.placeholder.as_ref(),
        ShapeNode::Picture(shape) => shape.base.placeholder.as_ref(),
        ShapeNode::GraphicFrame(shape) => shape.base.placeholder.as_ref(),
        ShapeNode::Group(shape) => shape.base.placeholder.as_ref(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn placeholder(placeholder_type: &str, index: Option<u32>) -> Placeholder {
        Placeholder {
            placeholder_type: Some(placeholder_type.to_owned()),
            index,
            orientation: None,
            size: None,
        }
    }

    #[test]
    fn placeholder_matching_prefers_indices_and_normalizes_common_types() {
        assert!(placeholders_match(
            &placeholder("body", Some(4)),
            &placeholder("title", Some(4))
        ));
        assert!(placeholders_match(
            &placeholder("ctrTitle", None),
            &placeholder("title", None)
        ));
    }
}
