//! The chat agent's PPTX commands and read (Epo 2026-10-20): a shape's
//! whole text, slides added from a layout (with its text placeholders),
//! deleted and moved, text boxes, and a paragraph's text formatting and
//! alignment; the read lists slides with their layouts and shapes, or one
//! slide's shapes and paragraphs. Live Undo comes from the room, so no
//! command returns an inverse.

use pptx_edit::{
    EditCtx, ParagraphSnapshot, ShapeDraft, ShapeKind, ShapeRect, ShapeSnapshot, SlideSnapshot,
    StorySnapshot, TextStyle, TextStylePatch,
};
use pptx_parse::{
    GraphicFrameData, PptxPackage, ShapeNode, ShapeTransform, SlideLayout, find_placeholder,
    master_for_layout,
};
use serde_json::{Value, json};

use super::{PptxSession, STORIES, refused};
use crate::common::{check_paragraph_text, len16};
use crate::error::{EditCode, Error, Result};
use crate::session::Applied;
use crate::touched::{self, Kept, Touched};
use crate::types::{Alignment, Command, InspectQuery, TextFormat};

const EMU_PER_POINT: f64 = 12_700.0;
/// A shape's text one command writes (the engine survey's bounds).
const MAX_SHAPE_PARAGRAPHS: usize = 50;
const MAX_SHAPE_CHARS: usize = 5_000;
/// Slides, or one slide's paragraphs, a read lists.
const MAX_READ: u32 = 200;
/// Characters of a shape's first line the slide list shows.
const FIRST_LINE_CHARS: usize = 80;

fn invalid(message: impl Into<String>) -> Error {
    Error::edit(EditCode::InvalidInput, message)
}

fn ctx() -> EditCtx {
    EditCtx::local("wasm")
}

fn missing_slide(id: &str) -> Error {
    Error::edit(
        EditCode::UnavailableTarget,
        format!("slide {id} is no longer in the deck"),
    )
}

/// A refusal carrying what its target reads now.
fn stale(message: &str, id: &str, kind: &str, text: String) -> Error {
    Error::Edit {
        code: EditCode::StaleTarget,
        message: message.to_owned(),
        details: Some(json!({ "block": { "id": id, "type": kind, "text": text } })),
    }
}

fn paragraph_text(paragraph: &ParagraphSnapshot) -> String {
    paragraph.runs.iter().map(|run| run.text.as_str()).collect()
}

/// A shape's text as `set_shape_text` expects it: its paragraphs joined by
/// line breaks.
fn story_text(story: &StorySnapshot) -> String {
    story
        .paragraphs
        .iter()
        .map(paragraph_text)
        .collect::<Vec<_>>()
        .join("\n")
}

/// Every paragraph of `shapes` in the read's order: each shape's, then its
/// children's.
fn walk<'a>(shapes: &'a [ShapeSnapshot], visit: &mut impl FnMut(&'a ShapeSnapshot)) {
    for shape in shapes {
        visit(shape);
        walk(&shape.children, visit);
    }
}

/// A slide's text as `delete_slide` compares it: its paragraphs' trimmed
/// text, blank ones left out, one a line.
fn slide_text(slide: &SlideSnapshot) -> String {
    let mut lines = Vec::new();
    walk(&slide.shapes, &mut |shape| {
        for story in &shape.text_stories {
            lines.extend(story.paragraphs.iter().map(paragraph_text));
        }
    });
    canonical_lines(&lines.join("\n"))
}

fn canonical_lines(text: &str) -> String {
    text.lines()
        .map(str::trim)
        .filter(|line| !line.is_empty())
        .collect::<Vec<_>>()
        .join("\n")
}

/// A shape's new text: 1 to [`MAX_SHAPE_PARAGRAPHS`] paragraphs of at most
/// [`MAX_SHAPE_CHARS`] characters together, none holding a line break.
fn check_shape_text(paragraphs: &[String]) -> Result<()> {
    if paragraphs.is_empty() {
        return Err(invalid(
            "paragraphs holds the shape's text, one string a paragraph; [\"\"] empties it",
        ));
    }
    if paragraphs.len() > MAX_SHAPE_PARAGRAPHS {
        return Err(invalid(format!(
            "paragraphs holds {} paragraphs; a shape takes at most {MAX_SHAPE_PARAGRAPHS} a command",
            paragraphs.len()
        )));
    }
    let chars: usize = paragraphs.iter().map(|text| text.chars().count()).sum();
    if chars > MAX_SHAPE_CHARS {
        return Err(invalid(format!(
            "paragraphs hold {chars} characters; a shape takes at most {MAX_SHAPE_CHARS} a command"
        )));
    }
    paragraphs
        .iter()
        .try_for_each(|text| check_paragraph_text(text))
}

/// `after` or `before` (at most one) as (id, after).
fn slide_anchor<'a>(
    after: Option<&'a str>,
    before: Option<&'a str>,
    command: &str,
) -> Result<Option<(&'a str, bool)>> {
    match (after, before) {
        (None, None) => Ok(None),
        (Some(id), None) => Ok(Some((id, true))),
        (None, Some(id)) => Ok(Some((id, false))),
        _ => Err(invalid(format!(
            "{command} names at most one of after_slide and before_slide"
        ))),
    }
}

fn points(emu: i64) -> f64 {
    (emu as f64 / EMU_PER_POINT * 10.0).round() / 10.0
}

fn emu(points: f64) -> i64 {
    (points * EMU_PER_POINT).round() as i64
}

fn node_transform(node: &ShapeNode) -> &ShapeTransform {
    match node {
        ShapeNode::Shape(shape) => &shape.base.transform,
        ShapeNode::Picture(shape) => &shape.base.transform,
        ShapeNode::GraphicFrame(shape) => &shape.base.transform,
        ShapeNode::Group(shape) => &shape.base.transform,
    }
}

fn slide_layout<'a>(package: &'a PptxPackage, slide: &SlideSnapshot) -> Option<&'a SlideLayout> {
    let path = slide.layout_part_path.as_deref()?;
    package
        .layouts
        .iter()
        .find(|layout| layout.part_path == path)
}

fn layout_name(layout: &SlideLayout) -> &str {
    layout.name.as_deref().unwrap_or(&layout.part_path)
}

fn kind(shape: &ShapeSnapshot) -> &'static str {
    match shape.kind {
        ShapeKind::Shape if shape.text_stories.is_empty() => "shape",
        ShapeKind::Shape => "text",
        ShapeKind::Picture => "picture",
        ShapeKind::Group => "group",
        ShapeKind::GraphicFrame => match &shape.graphic {
            Some(GraphicFrameData::Table(_)) => "table",
            Some(GraphicFrameData::Chart { .. }) => "chart",
            Some(GraphicFrameData::Diagram { .. }) => "diagram",
            _ => "object",
        },
    }
}

/// A table cell's (row, column) from its story id.
fn cell_of(story_id: &str) -> Option<(u32, u32)> {
    let (_, cell) = story_id.rsplit_once(":table:")?;
    let (row, column) = cell.split_once(':')?;
    Some((row.parse().ok()?, column.parse().ok()?))
}

impl PptxSession {
    /// [`crate::CommandResult::touched`] of the agent's shape and slide
    /// commands: the text and marks of a `set_shape_text` shape's story
    /// (what it can take out), a `delete_slide` slide whole (its entry, its
    /// shapes', their stories' and its comments'); nothing for the rest,
    /// which add what they write.
    pub(super) fn touched_agent(&self, command: &Command) -> Touched {
        let doc = self.deck.yrs_doc();
        match command {
            Command::SetShapeText { shape_id, .. } => {
                let story = format!("story:{shape_id}:0");
                let Ok(snapshot) = self.deck.story(&story) else {
                    return Touched::default();
                };
                let end: u32 = snapshot
                    .paragraphs
                    .iter()
                    .map(|paragraph| len16(&paragraph_text(paragraph)) + 1)
                    .sum();
                touched::story_range(doc, STORIES, &story, 0, end, Kept::Content)
            }
            Command::DeleteSlide { slide_id, .. } => {
                let mut touched = touched::story(doc, "pptx:slides", slide_id);
                let Ok(deck) = self.deck.snapshot() else {
                    return touched;
                };
                if let Some(slide) = deck.slides.iter().find(|slide| slide.id == *slide_id) {
                    walk(&slide.shapes, &mut |shape| {
                        touched += touched::story(doc, "pptx:shapes", &shape.id);
                        for story in &shape.text_stories {
                            touched += touched::story(doc, STORIES, &story.id);
                        }
                    });
                }
                for comment in deck
                    .comments
                    .iter()
                    .filter(|comment| comment.slide_id == *slide_id)
                {
                    touched += touched::story(doc, "pptx:comments", &comment.id);
                }
                touched
            }
            _ => Touched::default(),
        }
    }

    pub(super) fn apply_agent(&mut self, command: &Command) -> Result<Applied> {
        match command {
            Command::SetShapeText {
                shape_id,
                expected_text,
                paragraphs,
            } => self.set_shape_text(shape_id, expected_text, paragraphs),
            Command::AddSlide {
                after_slide,
                before_slide,
                layout,
            } => self.add_slide(
                slide_anchor(after_slide.as_deref(), before_slide.as_deref(), "add_slide")?,
                layout.as_deref(),
            ),
            Command::DeleteSlide {
                slide_id,
                expected_text,
            } => self.delete_slide(slide_id, expected_text),
            Command::MoveSlide {
                slide_id,
                after_slide,
                before_slide,
            } => match slide_anchor(
                after_slide.as_deref(),
                before_slide.as_deref(),
                "move_slide",
            )? {
                Some(anchor) => self.move_slide(slide_id, anchor),
                None => Err(invalid("move_slide names after_slide or before_slide")),
            },
            Command::AddTextBox {
                slide_id,
                paragraphs,
                x,
                y,
                w,
                h,
            } => self.add_text_box(slide_id, paragraphs, [*x, *y, *w, *h]),
            Command::FormatText {
                target_id,
                expected_text,
                format,
            } => self.format_text(target_id, expected_text, format),
            Command::SetAlignment {
                target_id,
                alignment,
            } => self.set_alignment(target_id, *alignment),
            _ => Err(Error::edit(
                EditCode::UnsupportedOperation,
                "PPTX files take replace_text, insert_paragraph, delete_paragraph, set_shape_text, \
                 format_text, set_alignment, add_slide, delete_slide, move_slide and add_text_box",
            )),
        }?;
        if !matches!(
            command,
            Command::FormatText { .. } | Command::SetAlignment { .. }
        ) {
            self.stories = None;
        }
        Ok(Applied {
            ids: Vec::new(),
            inverse: None,
        })
    }

    /// No slide: `count` slides from `start`, each with its layout and its
    /// shapes (id, kind, placeholder type, first line of text), the first
    /// page also listing the deck's layouts. A slide (its number): its
    /// shapes (id, kind, name, placeholder type, box in points for a shape
    /// not in a group) and `count` of its paragraphs from `start`, in shape
    /// order.
    pub(crate) fn inspect_scoped(&mut self, query: &InspectQuery) -> Result<Value> {
        let count = query
            .count
            .filter(|count| (1..=MAX_READ).contains(count))
            .ok_or_else(|| invalid(format!("a PPTX read names a count of 1 to {MAX_READ}")))?
            as usize;
        let start = query.start.unwrap_or(0) as usize;
        let deck = self.deck.snapshot().map_err(Error::engine)?;
        let package = self.deck.package();
        let total = deck.slides.len();
        let mut answer = json!({
            "width": points(deck.width_emu),
            "height": points(deck.height_emu),
        });
        let Some(number) = query.slide else {
            let end = (start + count).min(total);
            let slides: Vec<Value> = deck.slides[start.min(end)..end]
                .iter()
                .enumerate()
                .map(|(offset, slide)| self.listed_slide(package, slide, start + offset + 1))
                .collect();
            answer["slides"] = json!(slides);
            answer["total"] = json!(total);
            answer["nextStart"] = json!((end < total).then_some(end));
            if start == 0 {
                answer["layouts"] = package
                    .layouts
                    .iter()
                    .map(|layout| json!({ "id": layout.part_path, "name": layout_name(layout) }))
                    .collect();
            }
            return Ok(answer);
        };
        let slide = (number as usize)
            .checked_sub(1)
            .and_then(|index| deck.slides.get(index))
            .ok_or_else(|| {
                invalid(format!(
                    "slide {number} is not in the deck, which has {total} slides"
                ))
            })?;
        let layout = slide_layout(package, slide);
        let master = master_for_layout(package, layout);
        let mut shapes = Vec::new();
        let mut paragraphs = Vec::new();
        for shape in &slide.shapes {
            // A group's children are placed in its own coordinates: only
            // top-level shapes show a box.
            let mut transform = (shape.x, shape.y, shape.width, shape.height);
            if (transform.2 <= 0 || transform.3 <= 0)
                && let Some(placeholder) = &shape.placeholder
            {
                let inherited = [
                    layout.map(|layout| &layout.shapes),
                    master.map(|m| &m.shapes),
                ]
                .into_iter()
                .flatten()
                .filter_map(|nodes| find_placeholder(nodes, placeholder))
                .map(node_transform)
                .find(|found| found.width > 0 && found.height > 0);
                if let Some(found) = inherited {
                    transform = (found.x, found.y, found.width, found.height);
                }
            }
            let mut entry = self.read_shape(shape, &mut paragraphs);
            entry["box"] = json!([
                points(transform.0),
                points(transform.1),
                points(transform.2),
                points(transform.3)
            ]);
            shapes.push(entry);
            walk(&shape.children, &mut |child| {
                let mut entry = self.read_shape(child, &mut paragraphs);
                entry["inGroup"] = json!(true);
                shapes.push(entry);
            });
        }
        let end = (start + count).min(paragraphs.len());
        answer["slide"] = json!({
            "id": slide.id,
            "number": number,
            "layout": layout.map(|layout| json!({ "id": layout.part_path, "name": layout_name(layout) })),
        });
        answer["slides"] = json!(total);
        answer["shapes"] = json!(shapes);
        answer["total"] = json!(paragraphs.len());
        answer["nextStart"] = json!((end < paragraphs.len()).then_some(end));
        answer["paragraphs"] = Value::from(paragraphs[start.min(end)..end].to_vec());
        Ok(answer)
    }

    /// A slide of the slide list.
    fn listed_slide(&self, package: &PptxPackage, slide: &SlideSnapshot, number: usize) -> Value {
        let mut shapes = Vec::new();
        walk(&slide.shapes, &mut |shape| {
            let mut entry = json!({ "id": shape.id, "kind": kind(shape) });
            self.mark(shape, &mut entry);
            let first = shape
                .text_stories
                .iter()
                .flat_map(|story| &story.paragraphs)
                .map(paragraph_text)
                .find(|text| !text.trim().is_empty());
            if let Some(text) = first {
                let text = text.trim();
                entry["text"] = json!(match text.char_indices().nth(FIRST_LINE_CHARS) {
                    Some((at, _)) => format!("{}…", &text[..at]),
                    None => text.to_owned(),
                });
            }
            shapes.push(entry);
        });
        json!({
            "id": slide.id,
            "number": number,
            "layout": slide_layout(package, slide).map(layout_name),
            "shapes": shapes,
        })
    }

    /// A shape of the slide read, its paragraphs added to `paragraphs`.
    fn read_shape(&self, shape: &ShapeSnapshot, paragraphs: &mut Vec<Value>) -> Value {
        let mut entry = json!({ "id": shape.id, "kind": kind(shape), "name": shape.name });
        self.mark(shape, &mut entry);
        for story in &shape.text_stories {
            let cell = cell_of(&story.id);
            for paragraph in &story.paragraphs {
                let mut item = json!({
                    "id": paragraph.id,
                    "shape": shape.id,
                    "text": paragraph_text(paragraph),
                });
                if paragraph.level > 0 {
                    item["level"] = json!(paragraph.level);
                }
                if let Some((row, column)) = cell {
                    item["row"] = json!(row);
                    item["column"] = json!(column);
                }
                paragraphs.push(item);
            }
        }
        entry
    }

    /// A shape's placeholder type (`obj` when the file names none) and
    /// whether its text is an equation's.
    fn mark(&self, shape: &ShapeSnapshot, entry: &mut Value) {
        if let Some(placeholder) = &shape.placeholder {
            entry["placeholder"] = json!(placeholder.placeholder_type.as_deref().unwrap_or("obj"));
        }
        if shape
            .text_stories
            .iter()
            .any(|story| self.deck.is_equation_story(&story.id))
        {
            entry["equation"] = json!(true);
        }
    }

    /// The text story of a shape holding one.
    fn shape_story(&self, shape_id: &str) -> Result<StorySnapshot> {
        if let Ok(story) = self.deck.story(&format!("story:{shape_id}:0")) {
            return Ok(story);
        }
        let deck = self.deck.snapshot().map_err(Error::engine)?;
        let mut found = None;
        for slide in &deck.slides {
            walk(&slide.shapes, &mut |shape| {
                if shape.id == shape_id {
                    found = Some(kind(shape));
                }
            });
        }
        Err(match found {
            None => Error::edit(
                EditCode::UnavailableTarget,
                "the shape is no longer in the deck",
            ),
            Some("table") => invalid(
                "a table's text is edited by paragraph: replace_text, insert_paragraph and delete_paragraph",
            ),
            Some(_) => invalid("the shape holds no text; add a text box instead"),
        })
    }

    /// The shape's paragraphs become `paragraphs`, each kept paragraph
    /// keeping its level, bullet and alignment and taking its first run's
    /// formatting; new ones continue the last, and the last kept one takes
    /// the removed ones' place.
    fn set_shape_text(
        &mut self,
        shape_id: &str,
        expected_text: &str,
        paragraphs: &[String],
    ) -> Result<()> {
        check_shape_text(paragraphs)?;
        let story = self.shape_story(shape_id)?;
        let current: Vec<String> = story.paragraphs.iter().map(paragraph_text).collect();
        if current.join("\n") != expected_text {
            return Err(stale(
                "the shape's text differs from expected_text",
                shape_id,
                "shape",
                story_text(&story),
            ));
        }
        let mut spans = Vec::with_capacity(current.len());
        let mut cursor = 0;
        for text in &current {
            let length = len16(text);
            spans.push((cursor, cursor + length));
            cursor += length + 1;
        }
        let (old, new) = (current.len(), paragraphs.len());
        let kept = old.min(new);
        let Some(&(_, kept_end)) = kept.checked_sub(1).and_then(|last| spans.get(last)) else {
            return Err(Error::engine("a text story without paragraphs"));
        };
        let style = |index: usize, last: bool| {
            let runs = &story.paragraphs[index].runs;
            let run = if last { runs.last() } else { runs.first() };
            run.map(|run| run.style.clone()).unwrap_or_default()
        };
        if new < old {
            self.deck
                .replace_text(
                    &ctx(),
                    &story.id,
                    kept_end,
                    spans[old - 1].1,
                    "",
                    &TextStyle::default(),
                )
                .map_err(refused)?;
        } else if new > old {
            let added = format!("\n{}", paragraphs[old..].join("\n"));
            self.deck
                .replace_text(
                    &ctx(),
                    &story.id,
                    kept_end,
                    kept_end,
                    &added,
                    &style(old - 1, true),
                )
                .map_err(refused)?;
        }
        // From the last, so earlier offsets stay put.
        for index in (0..kept).rev() {
            if current[index] != paragraphs[index] {
                let (start, end) = spans[index];
                self.deck
                    .replace_text(
                        &ctx(),
                        &story.id,
                        start,
                        end,
                        &paragraphs[index],
                        &style(index, false),
                    )
                    .map_err(refused)?;
            }
        }
        Ok(())
    }

    /// The slide's position in deck order.
    fn slide_index(&self, ids: &[String], id: &str) -> Result<usize> {
        ids.iter()
            .position(|other| other == id)
            .ok_or_else(|| missing_slide(id))
    }

    fn add_slide(&mut self, anchor: Option<(&str, bool)>, layout: Option<&str>) -> Result<()> {
        let ids = self.deck.slide_ids().map_err(Error::engine)?;
        let index = match anchor {
            None => ids.len(),
            Some((id, after)) => self.slide_index(&ids, id)? + usize::from(after),
        };
        let layout = match layout {
            Some(id) => {
                if !self
                    .deck
                    .package()
                    .layouts
                    .iter()
                    .any(|layout| layout.part_path == id)
                {
                    return Err(Error::edit(
                        EditCode::UnavailableTarget,
                        format!("the deck has no layout {id}; the deck read lists its layouts"),
                    ));
                }
                Some(id.to_owned())
            }
            // The slide it follows, else the one it precedes.
            None if !ids.is_empty() => {
                self.deck
                    .slide_scope(index.saturating_sub(1))
                    .map_err(Error::engine)?
                    .slide
                    .layout_part_path
            }
            None => return Err(invalid("the deck has no slides; name a layout")),
        };
        let receipt = self
            .deck
            .insert_slide(&ctx(), index as u32, layout.as_deref())
            .map_err(refused)?;
        self.deck
            .add_layout_placeholders(&ctx(), &receipt.slide_id)
            .map_err(refused)?;
        Ok(())
    }

    /// Refused unless the slide's text reads `expected_text` (blank lines
    /// and the spaces around a line aside), and for a deck's only slide.
    fn delete_slide(&mut self, slide_id: &str, expected_text: &str) -> Result<()> {
        let ids = self.deck.slide_ids().map_err(Error::engine)?;
        let index = self.slide_index(&ids, slide_id)?;
        if ids.len() == 1 {
            return Err(invalid("a deck keeps at least one slide"));
        }
        let slide = self.deck.slide_scope(index).map_err(Error::engine)?.slide;
        let current = slide_text(&slide);
        if current != canonical_lines(expected_text) {
            return Err(stale(
                "the slide's text differs from expected_text",
                slide_id,
                "slide",
                current,
            ));
        }
        self.deck.delete_slide(&ctx(), slide_id).map_err(refused)?;
        Ok(())
    }

    fn move_slide(&mut self, slide_id: &str, (anchor, after): (&str, bool)) -> Result<()> {
        let mut ids = self.deck.slide_ids().map_err(Error::engine)?;
        let from = self.slide_index(&ids, slide_id)?;
        if anchor == slide_id {
            return Err(invalid("a slide cannot move next to itself"));
        }
        ids.remove(from);
        let to = self.slide_index(&ids, anchor)? + usize::from(after);
        self.deck
            .move_slide(&ctx(), slide_id, to as u32)
            .map_err(refused)?;
        Ok(())
    }

    /// A text box of `paragraphs` whose box, in points, lies on the slide.
    fn add_text_box(
        &mut self,
        slide_id: &str,
        paragraphs: &[String],
        [x, y, w, h]: [f64; 4],
    ) -> Result<()> {
        check_shape_text(paragraphs)?;
        let presentation = &self.deck.package().presentation;
        let (width, height) = (
            points(presentation.width_emu),
            points(presentation.height_emu),
        );
        let fits = [x, y, w, h].iter().all(|value| value.is_finite())
            && x >= 0.0
            && y >= 0.0
            && w > 0.0
            && h > 0.0
            && x + w <= width
            && y + h <= height;
        if !fits {
            return Err(invalid(format!(
                "the box (x {x}, y {y}, w {w}, h {h} points) must lie on the {width} × {height} point slide, with a positive size"
            )));
        }
        let ids = self.deck.slide_ids().map_err(Error::engine)?;
        self.slide_index(&ids, slide_id)?;
        let draft = ShapeDraft {
            name: "TextBox".into(),
            rect: ShapeRect {
                x: emu(x),
                y: emu(y),
                width: emu(w),
                height: emu(h),
            },
            text: String::new(),
            style: TextStyle::default(),
        };
        let receipt = self
            .deck
            .add_text_box(&ctx(), slide_id, &draft)
            .map_err(refused)?;
        self.deck
            .replace_text(
                &ctx(),
                &format!("story:{}:0", receipt.shape_id),
                0,
                0,
                &paragraphs.join("\n"),
                &TextStyle::default(),
            )
            .map_err(refused)?;
        Ok(())
    }

    /// The one occurrence of `expected_text` in a paragraph, as a range of
    /// its story.
    fn find_text(&mut self, id: &str, expected_text: &str) -> Result<(String, u32, u32)> {
        let (story_id, paragraph) = self.paragraph(id)?;
        if expected_text.is_empty() {
            return Err(invalid("expected_text names the text to format"));
        }
        let found: Vec<usize> = paragraph
            .text
            .match_indices(expected_text)
            .map(|(at, _)| at)
            .collect();
        let refusal = |code, message: String| Error::Edit {
            code,
            message,
            details: Some(json!({
                "block": { "id": id, "type": "paragraph", "text": paragraph.text }
            })),
        };
        let at = match found[..] {
            [at] => at,
            [] => {
                return Err(refusal(
                    EditCode::StaleTarget,
                    "expected text was not found in the paragraph".into(),
                ));
            }
            _ => {
                return Err(refusal(
                    EditCode::InvalidInput,
                    format!(
                        "expected text occurs {} times in the paragraph; include more context",
                        found.len()
                    ),
                ));
            }
        };
        let start = paragraph.start + len16(&paragraph.text[..at]);
        Ok((story_id, start, start + len16(expected_text)))
    }

    fn format_text(&mut self, id: &str, expected_text: &str, format: &TextFormat) -> Result<()> {
        let patch = text_patch(format)?;
        let (story_id, start, end) = self.find_text(id, expected_text)?;
        self.deck
            .format_text(&ctx(), &story_id, start, end, &patch)
            .map_err(refused)?;
        Ok(())
    }

    fn set_alignment(&mut self, id: &str, alignment: Alignment) -> Result<()> {
        let (story_id, paragraph) = self.paragraph(id)?;
        let value = match alignment {
            Alignment::Left => "l",
            Alignment::Center => "ctr",
            Alignment::Right => "r",
            Alignment::Justify => "just",
        };
        self.deck
            .set_paragraph_alignment(
                &ctx(),
                &story_id,
                paragraph.start,
                paragraph.start,
                Some(value),
            )
            .map_err(refused)?;
        Ok(())
    }
}

/// `format_text`'s fields as the deck's run formatting; DOCX's highlight
/// is refused.
fn text_patch(format: &TextFormat) -> Result<TextStylePatch> {
    if format.highlight.is_some() {
        return Err(invalid(
            "highlight formats Word text; PowerPoint text takes bold, italic, underline, \
             strikethrough, color, font_size and font_family",
        ));
    }
    let mut patch = TextStylePatch {
        bold: format.bold,
        italic: format.italic,
        underline: format
            .underline
            .map(|on| if on { "sng" } else { "none" }.to_owned()),
        strike: format
            .strikethrough
            .map(|on| if on { "sngStrike" } else { "noStrike" }.to_owned()),
        ..TextStylePatch::default()
    };
    if let Some(color) = &format.color {
        let hex = color.strip_prefix('#').unwrap_or(color);
        if hex.len() != 6 || !hex.bytes().all(|byte| byte.is_ascii_hexdigit()) {
            return Err(invalid(format!(
                "color {color:?} is not a hex colour like \"#1F4E79\""
            )));
        }
        patch.color = Some(format!("#{}", hex.to_ascii_uppercase()));
    }
    if let Some(size) = format.font_size {
        if !(1.0..=1638.0).contains(&size) {
            return Err(invalid("font_size is 1 to 1638 points"));
        }
        patch.font_size_pt = Some(size);
    }
    if let Some(family) = &format.font_family {
        let family = family.trim();
        if family.is_empty() || family.chars().count() > 31 || family.chars().any(char::is_control)
        {
            return Err(invalid("font_family is a font name of 1 to 31 characters"));
        }
        patch.font_family = Some(family.to_owned());
    }
    if patch == TextStylePatch::default() {
        return Err(invalid("format_text names no formatting"));
    }
    Ok(patch)
}
