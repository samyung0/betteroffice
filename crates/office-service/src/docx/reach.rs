//! The stories the save projection reaches and the content of their embeds,
//! read straight from the session (#6): `entries` and `editable` need the
//! reachable stories and the readable text of embeds, not the projected
//! document. Only a block content control's own story is projected, for its
//! text. The walk mirrors `visitReachableStories` over the projected tree:
//! body, header and footer parts, then notes, each story before the stories
//! of its tables' cells and block content controls.

use std::collections::{HashMap, HashSet};

use docx_edit::EditingDoc;

use crate::error::Result;
use crate::jsv::{Obj, V, text};
use crate::obj;

use super::item::{
    content_control_value, is, is_kind, ordinary_content_for_item, restore_projected_field_results,
};
use super::paragraph::{items_with_markers, page_break_paragraph, split_slot_bookmarks};
use super::project::{Hooks, SaveContext, push_text};
use super::read;
use super::sdt::{apply_content_control_value, sdt_attrs_to_props};
use super::styles::map_entries;

/// Reachable stories in projection order, the segments read for them, and
/// `onEmbed` content by `${story}:${offset}`.
pub(super) struct Reached {
    pub stories: Vec<String>,
    pub segments: HashMap<String, Vec<V>>,
    pub embeds: HashMap<String, V>,
}

/// `None` when a story holds an opaque block content control: the
/// projection carries the base block at the same block index, which only the
/// projection knows.
pub(super) fn reach(doc: &EditingDoc, base: &V) -> Result<Option<Reached>> {
    let mut walk = Walk {
        doc,
        base,
        all: read::all_story_segments(doc)?,
        seen: HashSet::new(),
        reached: Reached {
            stories: Vec::new(),
            segments: HashMap::new(),
            embeds: HashMap::new(),
        },
        context: None,
    };
    let package = base.get("package");
    if !walk.visit("body", false)? {
        return Ok(None);
    }
    for parts in [package.get("headers"), package.get("footers")] {
        for (id, _) in map_entries(&parts) {
            if !walk.visit(&format!("hf:{id}"), false)? {
                return Ok(None);
            }
        }
    }
    for (prefix, notes) in [
        ("fn:", package.get("footnotes")),
        ("en:", package.get("endnotes")),
    ] {
        for note in notes.items() {
            // A projected note is reachable even without blocks.
            if !walk.visit(&format!("{prefix}{}", note.get("id").to_js_string()), true)? {
                return Ok(None);
            }
        }
    }
    Ok(Some(walk.reached))
}

/// The paragraph being gathered in a story.
#[derive(Default)]
struct Paragraph {
    items: Vec<V>,
    /// Page or column breaks opened the paragraph's slot.
    slot: bool,
    /// Inline embeds by offset, as their content is read once the story ends.
    inline: Vec<(f64, V)>,
}

struct Walk<'a> {
    doc: &'a EditingDoc,
    base: &'a V,
    /// Every story's segments, reachable or not.
    all: HashMap<String, Vec<V>>,
    seen: HashSet<String>,
    reached: Reached,
    /// For block content controls' stories, made on first use.
    context: Option<SaveContext<'a>>,
}

impl<'a> Walk<'a> {
    /// Adds `story` when it has blocks (or `always`), then the stories its
    /// tables and block content controls hold. False to fall back.
    fn visit(&mut self, story: &str, always: bool) -> Result<bool> {
        if self.seen.contains(story) {
            return Ok(true);
        }
        let Some(segments) = self.all.remove(story) else {
            return Ok(true);
        };
        if !always && segments.is_empty() {
            return Ok(true);
        }
        self.seen.insert(story.to_owned());
        self.reached.stories.push(story.to_owned());
        // Each paragraph's items as `storyToBlocks` gathers them: building
        // its content folds a field's result runs into the field's payload,
        // which the field's text is read from. Inline embeds are copies, so
        // the segments `entries` reads keep their payloads.
        let mut paragraph = Paragraph::default();
        let mut offset = 0.0;
        for segment in &segments {
            if is_kind(segment, "text") {
                let source = segment.get("text").to_js_string();
                push_text(&mut paragraph.items, &source, &segment.get("attributes"));
                offset += text::len(&source) as f64;
                continue;
            }
            if is_kind(segment, "pilcrow") {
                let (properties, _) =
                    split_slot_bookmarks(&segment.get("properties"), paragraph.slot);
                let items = std::mem::take(&mut paragraph.items);
                restore_projected_field_results(items_with_markers(&properties, items))?;
                paragraph.slot = false;
            } else if !self.embed(story, offset, segment, &mut paragraph)? {
                return Ok(false);
            }
            offset += 1.0;
        }
        if !paragraph.items.is_empty() {
            restore_projected_field_results(paragraph.items)?;
        }
        for (offset, item) in paragraph.inline {
            if let Some(content) = ordinary_content_for_item(&item)? {
                let at = format!("{story}:{}", crate::js::number_to_string(offset));
                self.reached.embeds.insert(at, content);
            }
        }
        self.reached.segments.insert(story.to_owned(), segments);
        Ok(true)
    }

    fn embed(
        &mut self,
        story: &str,
        offset: f64,
        segment: &V,
        paragraph: &mut Paragraph,
    ) -> Result<bool> {
        let at = format!("{story}:{}", crate::js::number_to_string(offset));
        let payload = segment.get("payload");
        let kind = segment.get("embedKind").as_str().unwrap_or_default();
        if matches!(&*kind, "table" | "blockSdt") {
            // A block settles the breaks that opened the slot.
            paragraph.slot = false;
        }
        match &*kind {
            "table" => {
                for row in payload.get("rows").items() {
                    for cell in row.get("cells").items() {
                        let story = cell.get("story");
                        if story.truthy() && !self.visit(&story.to_js_string(), false)? {
                            return Ok(false);
                        }
                    }
                }
            }
            "blockSdt" => {
                let applied = content_control_value(&payload.get("value")).and_then(|value| {
                    apply_content_control_value(&sdt_attrs_to_props(&payload), &value).ok()
                });
                let content = match applied {
                    // The value's blocks replace the child story's.
                    Some((_, blocks)) => blocks,
                    None => match payload.get("story").as_str() {
                        Some(child)
                            if self.all.contains_key(&*child) || self.seen.contains(&*child) =>
                        {
                            if !self.visit(&child, false)? {
                                return Ok(false);
                            }
                            self.project(&child)?
                        }
                        _ => Vec::new(),
                    },
                };
                self.reached.embeds.insert(
                    at,
                    obj! { "type": "blockSdt", "content": V::array(content) },
                );
            }
            "opaque" => {
                let blob = payload.get("blob");
                if is(&blob, "pageBreak") {
                    paragraph.slot = false;
                    self.reached.embeds.insert(at, page_break_paragraph());
                } else if is(&blob, "blockSdt") {
                    return Ok(false);
                }
            }
            "pageBreak" | "columnBreak" => {
                if paragraph.items.is_empty() {
                    paragraph.slot = true;
                } else {
                    let kind = if &*kind == "pageBreak" {
                        "page"
                    } else {
                        "column"
                    };
                    paragraph.items.push(obj! {
                        "kind": "embed",
                        "embedKind": "flowBreak",
                        "payload": obj! { "breakType": kind },
                        "attributes": segment.get("attributes"),
                    });
                }
            }
            _ => {
                let copy = V::Obj(segment.obj().map(Obj::spread).unwrap_or_default());
                paragraph.items.push(copy.clone());
                paragraph.inline.push((offset, copy));
            }
        }
        Ok(true)
    }

    fn project(&mut self, story: &str) -> Result<Vec<V>> {
        if self.context.is_none() {
            self.context = Some(SaveContext::new(self.doc, self.base, Hooks::default())?);
        }
        self.context
            .as_mut()
            .expect("made above")
            .story_to_blocks(story)
    }
}
