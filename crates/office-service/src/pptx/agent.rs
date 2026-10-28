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
use std::borrow::Cow;
use std::rc::Rc;

use pptx_parse::{
    GraphicFrameData, Picture, PptxPackage, ShapeNode, ShapeTransform, SlideLayout,
    find_placeholder, master_for_layout,
};
use serde_json::{Value, json};

use super::{PptxSession, STORIES, missing_paragraph, refused};
use crate::common::{check_paragraph_text, len16, malformed_id};
use crate::error::{EditCode, Error, Result};
use crate::objects::{self, Chart, Found, Head, Parts, Stored, crop_side};
use crate::session::Applied;
use crate::touched::{self, Kept, Touched};
use crate::types::{Alignment, Command, InspectQuery, TextFormat};

const EMU_PER_POINT: f64 = 12_700.0;
/// Slides, or one slide's shapes and paragraphs, a read lists.
const MAX_READ: u32 = 200;
/// Shapes of a slide the slide list shows; its slide read pages the rest.
const MAX_LISTED_SHAPES: usize = 20;
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

fn digits(text: &str) -> bool {
    !text.is_empty() && text.bytes().all(|byte| byte.is_ascii_digit())
}

/// `<client>:<counter>`: an id pptx-edit `next_id` made for a slide, shape
/// or paragraph added since the seed, its prefix taken off.
fn added(rest: &str) -> bool {
    rest.split_once(':')
        .is_some_and(|(client, counter)| digits(client) && digits(counter))
}

/// A slide id's form: `slide:<index>:<sldId>` seeded from the file, else
/// `slide:<client>:<counter>`.
fn slide_form(id: &str) -> bool {
    id.strip_prefix("slide:").is_some_and(added)
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
/// What `delete_slide` compares: each paragraph a line in shape order, then
/// the notes as the slide read shows them (Epo 2026-10-22).
fn slide_text(slide: &SlideSnapshot) -> Result<String> {
    let mut lines = Vec::new();
    walk(&slide.shapes, &mut |shape| {
        for story in &shape.text_stories {
            lines.extend(story.paragraphs.iter().map(paragraph_text));
        }
    });
    lines.push(shown_notes(&slide.notes)?);
    Ok(canonical_lines(&lines.join("\n")))
}

/// A slide's speaker notes as its read shows them: whole up to limits.json's
/// `slideNotes` characters, else cut there, a marker counting the rest.
fn shown_notes(notes: &str) -> Result<String> {
    let most = crate::limits::slide_notes()?;
    Ok(match notes.char_indices().nth(most) {
        Some((at, _)) => format!(
            "{}… [notes cut: {} more characters]",
            &notes[..at],
            notes[at..].chars().count()
        ),
        None => notes.to_owned(),
    })
}

fn canonical_lines(text: &str) -> String {
    text.lines()
        .map(str::trim)
        .filter(|line| !line.is_empty())
        .collect::<Vec<_>>()
        .join("\n")
}

/// A shape's new text: at least one paragraph, none holding a line break,
/// within limits.json's `shapeText` paragraphs and characters (review L2).
fn check_shape_text(paragraphs: &[String]) -> Result<()> {
    if paragraphs.is_empty() {
        return Err(invalid(
            "paragraphs holds the shape's text, one string a paragraph; [\"\"] empties it",
        ));
    }
    let most = crate::limits::shape_text()?;
    if paragraphs.len() > most.paragraphs {
        return Err(invalid(format!(
            "paragraphs holds {} paragraphs; a shape takes at most {} a command",
            paragraphs.len(),
            most.paragraphs
        )));
    }
    let chars: usize = paragraphs.iter().map(|text| text.chars().count()).sum();
    if chars > most.characters {
        return Err(invalid(format!(
            "paragraphs hold {chars} characters; a shape takes at most {} a command",
            most.characters
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

/// The source picture a picture shape of a seeded slide came from: its alt
/// text and crop, which the deck's snapshot leaves out.
fn source_picture<'a>(
    package: &'a PptxPackage,
    slide: &SlideSnapshot,
    shape: &ShapeSnapshot,
) -> Option<&'a Picture> {
    fn find(nodes: &[ShapeNode], id: u32) -> Option<&Picture> {
        nodes.iter().find_map(|node| match node {
            ShapeNode::Picture(picture) if picture.base.id == id => Some(picture),
            ShapeNode::Group(group) => find(&group.children, id),
            _ => None,
        })
    }
    let path = slide.source_part_path.as_deref()?;
    let source = package
        .slides
        .iter()
        .find(|source| source.part_path == path)?;
    find(&source.shapes, shape.source_id)
}

/// `line`'s fields `entry` does not have yet.
fn merge(entry: &mut Value, line: Value) {
    if let (Some(entry), Value::Object(line)) = (entry.as_object_mut(), line) {
        for (key, value) in line {
            entry.entry(key).or_insert(value);
        }
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
    /// which add what they write (`add_slide`'s is [`Self::created`]).
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
            Command::DeleteSlide { slide_id, .. } => self.slide_touched(slide_id),
            _ => Touched::default(),
        }
    }

    /// What the last command wrote that its Redo re-creates, read after it
    /// applied, where the command alone cannot say how much (the Editor
    /// 2026-10-22): an `add_slide`'s slide whole (its layout's
    /// placeholders), a `set_shape_text`'s story whole (its paragraphs take
    /// their runs' formatting, a start and an end marker a key; review M3).
    pub(crate) fn created(&mut self) -> Touched {
        self.created.take().unwrap_or_default()
    }

    /// A slide's entry, its shapes', their stories' and its comments', read
    /// from the slide alone (review L5).
    fn slide_touched(&self, slide_id: &str) -> Touched {
        let doc = self.deck.yrs_doc();
        let mut touched = touched::story(doc, "pptx:slides", slide_id);
        let index = self
            .deck
            .slide_ids()
            .ok()
            .and_then(|ids| ids.iter().position(|id| id == slide_id));
        if let Some(Ok(scope)) = index.map(|index| self.deck.slide_scope(index)) {
            walk(&scope.slide.shapes, &mut |shape| {
                touched += touched::story(doc, "pptx:shapes", &shape.id);
                for story in &shape.text_stories {
                    touched += touched::story(doc, STORIES, &story.id);
                }
            });
        }
        for comment in self.deck.comments().unwrap_or_default() {
            if comment.slide_id == slide_id {
                touched += touched::story(doc, "pptx:comments", &comment.id);
            }
        }
        touched
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
                text,
                format,
            } => self.format_text(target_id, text, format),
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
        Ok(Applied {
            ids: Vec::new(),
            inverse: None,
        })
    }

    /// No slide: `count` slides from `start`, each with its layout and its
    /// first [`MAX_LISTED_SHAPES`] shapes (id, kind, placeholder type, first
    /// line of text) and how many more it holds, the first page also
    /// listing the deck's layouts. A slide (its number): `count` of its
    /// shapes and paragraphs from `start`, in reading order, each shape
    /// followed by its paragraphs (review M1: a slide of hundreds of shapes
    /// pages like its text), a shape with its name, placeholder type and,
    /// outside a group, its box in points; the first page also gives the
    /// slide's speaker notes ([`shown_notes`]). A picture lists its alt
    /// text and whether the model can see it, a chart its type and title,
    /// the slide read adding a picture's pixel size and crop and a chart's
    /// series ([`objects`]). `ids`: those pictures and charts
    /// ([`Self::read_ids`]).
    pub(crate) fn inspect_scoped(&mut self, query: &InspectQuery) -> Result<Value> {
        if let Some(ids) = &query.ids {
            return self.read_ids(ids);
        }
        let count = query
            .count
            .filter(|count| (1..=MAX_READ).contains(count))
            .ok_or_else(|| invalid(format!("a PPTX read names a count of 1 to {MAX_READ}")))?
            as usize;
        let start = query.start.unwrap_or(0) as usize;
        let total = self.deck.slide_ids().map_err(Error::engine)?.len();
        let scope = |index: usize| self.deck.slide_scope(index).map_err(Error::engine);
        let package = self.deck.package();
        let Some(number) = query.slide else {
            let end = (start + count).min(total);
            let scopes = (start.min(end)..end)
                .map(scope)
                .collect::<Result<Vec<_>>>()?;
            let mut parts = Parts::default();
            let slides = scopes
                .iter()
                .map(|scope| self.listed_slide(package, &scope.slide, scope.index + 1, &mut parts))
                .collect::<Result<Vec<_>>>()?;
            let (width, height) = match scopes.first() {
                Some(first) => (first.width_emu, first.height_emu),
                None => {
                    let deck = self.deck.snapshot().map_err(Error::engine)?;
                    (deck.width_emu, deck.height_emu)
                }
            };
            let mut answer = json!({
                "width": points(width),
                "height": points(height),
                "slides": slides,
                "total": total,
                "nextStart": (end < total).then_some(end),
            });
            if start == 0 {
                answer["layouts"] = package
                    .layouts
                    .iter()
                    .map(|layout| json!({ "id": layout.part_path, "name": layout_name(layout) }))
                    .collect();
            }
            return Ok(answer);
        };
        let scope = (number as usize)
            .checked_sub(1)
            .filter(|index| *index < total)
            .map(scope)
            .transpose()?
            .ok_or_else(|| {
                invalid(format!(
                    "slide {number} is not in the deck, which has {total} slides"
                ))
            })?;
        let slide = &scope.slide;
        let layout = slide_layout(package, slide);
        let master = master_for_layout(package, layout);
        // Each shape, then its paragraphs: the units a page counts.
        let mut items = Vec::new();
        let mut parts = Parts::default();
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
            let at = self.read_shape(slide, shape, &mut items, &mut parts)?;
            items[at].1["box"] = json!([
                points(transform.0),
                points(transform.1),
                points(transform.2),
                points(transform.3)
            ]);
            let mut children = Vec::new();
            walk(&shape.children, &mut |child| children.push(child));
            for child in children {
                let at = self.read_shape(slide, child, &mut items, &mut parts)?;
                items[at].1["inGroup"] = json!(true);
            }
        }
        let end = (start + count).min(items.len());
        let (shapes, paragraphs): (Vec<_>, Vec<_>) = items[start.min(end)..end]
            .iter()
            .cloned()
            .partition(|(shape, _)| *shape);
        let mut answer = json!({
            "width": points(scope.width_emu),
            "height": points(scope.height_emu),
            "slide": {
                "id": slide.id,
                "number": number,
                "layout": layout.map(|layout| json!({ "id": layout.part_path, "name": layout_name(layout) })),
            },
            "slides": total,
            "shapes": shapes.into_iter().map(|(_, item)| item).collect::<Vec<_>>(),
            "paragraphs": paragraphs.into_iter().map(|(_, item)| item).collect::<Vec<_>>(),
            "total": items.len(),
            "nextStart": (end < items.len()).then_some(end),
        });
        // The speaker notes, on the first page (Epo 2026-10-22).
        if start == 0 && !slide.notes.is_empty() {
            answer["slide"]["notes"] = json!(shown_notes(&slide.notes)?);
        }
        Ok(answer)
    }

    /// A slide of the slide list.
    fn listed_slide(
        &self,
        package: &PptxPackage,
        slide: &SlideSnapshot,
        number: usize,
        parts: &mut Parts,
    ) -> Result<Value> {
        let mut listed = Vec::new();
        let mut more = 0;
        walk(&slide.shapes, &mut |shape| {
            if listed.len() == MAX_LISTED_SHAPES {
                more += 1;
            } else {
                listed.push(shape);
            }
        });
        let mut shapes = Vec::with_capacity(listed.len());
        for shape in listed {
            let mut entry = json!({ "id": shape.id, "kind": kind(shape) });
            self.mark(shape, &mut entry);
            match kind(shape) {
                "picture" => {
                    let line = self.picture(slide, shape, parts).line()?;
                    for key in ["viewable", "reason"] {
                        if let Some(value) = line.get(key) {
                            entry[key] = value.clone();
                        }
                    }
                    if let Some(alt) = line["alt"].as_str() {
                        entry["alt"] = json!(objects::shown(alt, FIRST_LINE_CHARS));
                    }
                }
                "chart" => {
                    let line = objects::chart_line(None, self.chart(shape, parts).as_deref());
                    for key in ["chartType", "title", "reason"] {
                        if let Some(value) = line.get(key) {
                            entry[key] = value.clone();
                        }
                    }
                }
                _ => {}
            }
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
        }
        let mut entry = json!({
            "id": slide.id,
            "number": number,
            "layout": slide_layout(package, slide).map(layout_name),
            "shapes": shapes,
        });
        if more > 0 {
            entry["moreShapes"] = json!(more);
        }
        Ok(entry)
    }

    /// A picture shape of `slide` as the reads show it: its header read
    /// once a call, its alt text and crop the source picture's while the
    /// shape still shows that picture's media (a picture added since has
    /// pending media and source id 0, which a file picture may also have; a
    /// replaced one other media; review r1 L5).
    fn picture<'a>(
        &'a self,
        slide: &SlideSnapshot,
        shape: &'a ShapeSnapshot,
        parts: &mut Parts,
    ) -> objects::Picture<'a> {
        let package = self.deck.package();
        let (head, stored) = match (&shape.media_part_path, &shape.pending_media) {
            (Some(path), _) => {
                let bytes = package
                    .media
                    .iter()
                    .find(|media| media.part_path == *path)
                    .map_or(&[][..], |media| media.bytes.as_slice());
                let head = parts.head(path, || Head::of(bytes, bytes.len(), path));
                (head, Stored::Bytes(Cow::Borrowed(bytes)))
            }
            (None, Some(pending)) => {
                let part = pending.content_type.replace('/', ".");
                let bytes: &[u8] = &pending.bytes;
                (
                    Head::of(bytes, bytes.len(), &part),
                    Stored::Bytes(Cow::Borrowed(bytes)),
                )
            }
            (None, None) => (Head::linked(), Stored::Nothing),
        };
        let source = shape
            .pending_media
            .is_none()
            .then(|| source_picture(package, slide, shape))
            .flatten()
            .filter(|picture| picture.media_part_path == shape.media_part_path);
        objects::Picture {
            id: Some(shape.id.clone()),
            alt: source.and_then(|picture| picture.base.description.clone()),
            head,
            stored,
            crop: source.map_or([0.0; 4], |picture| {
                let crop = &picture.crop;
                [crop.left, crop.top, crop.right, crop.bottom]
                    .map(|side| crop_side(f64::from(side)))
            }),
        }
    }

    /// A chart shape's chart, read again from its part once a call.
    fn chart(&self, shape: &ShapeSnapshot, parts: &mut Parts) -> Option<Rc<Chart>> {
        let Some(GraphicFrameData::Chart {
            part_path: Some(path),
            ..
        }) = &shape.graphic
        else {
            return None;
        };
        parts.chart(path, || self.deck.package().part_bytes(path))
    }

    /// The by-id read ([`objects::answer`]): picture and chart shape ids from
    /// a read; a gone or malformed id, or one of a shape that is no picture
    /// or chart, is named in the notice (Q381).
    fn read_ids(&self, ids: &[String]) -> Result<Value> {
        let (ids, past) = objects::requested(ids)?;
        let deck = self.deck.snapshot().map_err(Error::engine)?;
        let mut shapes: Vec<(&SlideSnapshot, &ShapeSnapshot)> = Vec::new();
        for slide in &deck.slides {
            walk(&slide.shapes, &mut |shape| shapes.push((slide, shape)));
        }
        let mut parts = Parts::default();
        let mut answered = Vec::new();
        let mut missing = Vec::new();
        for id in ids {
            let Some(&(slide, shape)) = shapes.iter().find(|(_, shape)| shape.id == id) else {
                let why = if self.shape_form(&id) {
                    "is no longer in the deck".to_owned()
                } else {
                    let first = shapes
                        .iter()
                        .position(|(_, shape)| matches!(kind(shape), "picture" | "chart"));
                    objects::malformed(&first.map_or_else(
                        || "the deck holds no picture or chart".to_owned(),
                        |at| {
                            let (slide, shape) = shapes[at];
                            let number = deck.slides.iter().position(|other| other.id == slide.id);
                            format!(
                                "picture and chart ids read like {:?}, slide {}'s",
                                shape.id,
                                number.map_or(0, |number| number + 1)
                            )
                        },
                    ))
                };
                missing.push((id, why));
                continue;
            };
            let found = match kind(shape) {
                "picture" => Found::Picture(self.picture(slide, shape, &mut parts)),
                "chart" => Found::Chart(self.chart(shape, &mut parts)),
                other => {
                    missing.push((
                        id,
                        format!(
                            "is a {other}, not a picture or chart; the slide read shows its text"
                        ),
                    ));
                    continue;
                }
            };
            answered.push((id, found));
        }
        objects::answer(answered, &missing, &past)
    }

    /// A shape of the slide read and then its paragraphs, added to `items`
    /// (`true` marks the shape); the shape's index there. A picture adds its
    /// read line's fields, a chart its type, title and series.
    fn read_shape(
        &self,
        slide: &SlideSnapshot,
        shape: &ShapeSnapshot,
        items: &mut Vec<(bool, Value)>,
        parts: &mut Parts,
    ) -> Result<usize> {
        let mut entry = json!({ "id": shape.id, "kind": kind(shape), "name": shape.name });
        self.mark(shape, &mut entry);
        match kind(shape) {
            "picture" => merge(&mut entry, self.picture(slide, shape, parts).line()?),
            "chart" => merge(
                &mut entry,
                objects::chart_line(None, self.chart(shape, parts).as_deref()),
            ),
            _ => {}
        }
        let at = items.len();
        items.push((true, entry));
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
                items.push((false, item));
            }
        }
        Ok(at)
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

    /// A slide the deck was seeded with: `slide:<index>:<sldId>` of the
    /// file's slide at that index.
    fn seeded_slide(&self, id: &str) -> bool {
        let slides = &self.deck.package().presentation.slides;
        id.strip_prefix("slide:")
            .and_then(|rest| rest.split_once(':'))
            .and_then(|(index, sld)| Some((slides.get(index.parse::<usize>().ok()?)?, sld)))
            .is_some_and(|(slide, sld)| slide.id.to_string() == sld)
    }

    /// A shape id's form: `shape:<client>:<counter>`, else seeded,
    /// `<slide id>:shape:<path>` of a seeded slide, its path the shape's
    /// place (`2`, `2.0` in a group).
    fn shape_form(&self, id: &str) -> bool {
        match id.strip_prefix("shape:") {
            Some(rest) => added(rest),
            None => id.split_once(":shape:").is_some_and(|(slide, path)| {
                self.seeded_slide(slide) && path.split('.').all(digits)
            }),
        }
    }

    /// A paragraph id's form: `para:<client>:<counter>`, else seeded,
    /// `para:<story>:<n>` of a shape's text (`story:<shape id>:0`) or a
    /// table cell's (`story:<shape id>:table:<row>:<column>`).
    fn paragraph_form(&self, id: &str) -> bool {
        let Some(rest) = id.strip_prefix("para:") else {
            return false;
        };
        if added(rest) {
            return true;
        }
        let Some((story, n)) = rest
            .strip_prefix("story:")
            .and_then(|story| story.rsplit_once(':'))
        else {
            return false;
        };
        digits(n)
            && (story
                .strip_suffix(":0")
                .is_some_and(|shape| self.shape_form(shape))
                || story
                    .rsplit_once(":table:")
                    .is_some_and(|(shape, cell)| added(cell) && self.shape_form(shape)))
    }

    /// The refusal of a paragraph id the deck does not hold: gone when it
    /// has a paragraph id's form, else malformed (Epo 2026-10-24, Q223).
    pub(super) fn paragraph_refusal(&self, id: &str) -> Error {
        if self.paragraph_form(id) {
            missing_paragraph()
        } else {
            malformed_id(id, "PPTX paragraph", &self.read_example(true))
        }
    }

    /// The read's ids for a malformed id's refusal: the deck's first slide,
    /// or the first shape (`paragraph`: shape paragraph) of the first slide
    /// holding one.
    fn read_example(&self, paragraph: bool) -> String {
        let total = self.deck.slide_ids().map_or(0, |ids| ids.len());
        for index in 0..total {
            let Ok(scope) = self.deck.slide_scope(index) else {
                continue;
            };
            let mut found = None;
            walk(&scope.slide.shapes, &mut |shape| {
                if found.is_none() {
                    found = if paragraph {
                        shape
                            .text_stories
                            .iter()
                            .flat_map(|story| &story.paragraphs)
                            .map(|paragraph| paragraph.id.clone())
                            .next()
                    } else {
                        Some(shape.id.clone())
                    };
                }
            });
            if let Some(id) = found {
                let kind = if paragraph { "paragraph" } else { "shape" };
                return format!("{kind} ids read like {id:?}, slide {}'s first", index + 1);
            }
        }
        if paragraph {
            "the deck holds no text".to_owned()
        } else {
            "the deck holds no shape".to_owned()
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
            None if !self.shape_form(shape_id) => {
                malformed_id(shape_id, "shape", &self.read_example(false))
            }
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
        self.created = Some(touched::story(self.deck.yrs_doc(), STORIES, &story.id));
        Ok(())
    }

    /// The slide's position in deck order; a slide id the deck does not hold
    /// is gone when it has a slide id's form, else malformed (Q223).
    fn slide_index(&self, ids: &[String], id: &str) -> Result<usize> {
        ids.iter().position(|other| other == id).ok_or_else(|| {
            if slide_form(id) {
                return missing_slide(id);
            }
            let first = self
                .deck
                .slide_ids()
                .ok()
                .and_then(|ids| ids.first().cloned());
            let shown = first.map_or_else(
                || "the deck has no slide".to_owned(),
                |first| format!("slide ids read like {first:?}, slide 1's"),
            );
            malformed_id(id, "slide", &shown)
        })
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
        self.created = Some(self.slide_touched(&receipt.slide_id));
        Ok(())
    }

    /// Refused unless the slide's text and notes read `expected_text` as
    /// [`slide_text`] gives them (blank lines and the spaces around a line
    /// aside), and for a deck's only slide.
    fn delete_slide(&mut self, slide_id: &str, expected_text: &str) -> Result<()> {
        let ids = self.deck.slide_ids().map_err(Error::engine)?;
        let index = self.slide_index(&ids, slide_id)?;
        if ids.len() == 1 {
            return Err(invalid("a deck keeps at least one slide"));
        }
        let slide = self.deck.slide_scope(index).map_err(Error::engine)?.slide;
        let current = slide_text(&slide)?;
        if current != canonical_lines(expected_text) {
            return Err(stale(
                "the slide's text or notes differ from expected_text",
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

    /// The one occurrence of `text` in a paragraph, as a range of its
    /// story.
    fn find_text(&mut self, id: &str, text: &str) -> Result<(String, u32, u32)> {
        let (story_id, paragraph) = self.paragraph(id)?;
        if text.is_empty() {
            return Err(invalid("text names the text to format"));
        }
        let found: Vec<usize> = paragraph
            .text
            .match_indices(text)
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
        Ok((story_id, start, start + len16(text)))
    }

    fn format_text(&mut self, id: &str, text: &str, format: &TextFormat) -> Result<()> {
        let patch = text_patch(format)?;
        let (story_id, start, end) = self.find_text(id, text)?;
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
        // Controls and the XML noncharacters a file cannot store (review L4).
        let unstorable = |c: char| c.is_control() || matches!(c, '\u{fffe}' | '\u{ffff}');
        if family.is_empty() || family.chars().count() > 31 || family.chars().any(unstorable) {
            return Err(invalid("font_family is a font name of 1 to 31 characters"));
        }
        patch.font_family = Some(family.to_owned());
    }
    if patch == TextStylePatch::default() {
        return Err(invalid("format_text names no formatting"));
    }
    Ok(patch)
}
