//! yrsToDocument.ts, part three: `SaveContext` (a story's segments to
//! blocks, the slot breaks settled between paragraphs, tables and content
//! controls) and `projectDocument`. The projection runs with hooks
//! (checkpoint export, entries, rebase), so the session memo is never used.

use std::collections::{HashMap, HashSet};

use docx_edit::EditingDoc;

use crate::error::Result;
use crate::jsv::{Obj, V, stringify, text};
use crate::{arr, obj};

use super::comments::{comment_shared_id, project_comments};
use super::format::{
    table_attrs_to_formatting, table_cell_attrs_to_formatting, table_row_attrs_to_formatting,
};
use super::item::{
    build_paragraph_content, content_control_value, is, is_kind, o, ordinary_content_for_item,
    tracked_info, ty,
};
use super::paragraph::{
    FlowToken, SlotBreak, TokenKind, comment_marker, flow_tokens, page_break_paragraph,
    paragraph_from_story, place_trailing, restore_note_marks, slot_break_content, sort_boundaries,
    split_flow, split_slot_bookmarks, tracked_kind, without_breaks,
};
use super::read;
use super::sdt::{apply_content_control_value, sdt_attrs_to_props};
use super::styles::{
    Numbering, SourceStories, StyleResolver, enclosing_cell_story, map_entries, same_json,
    source_cell_formatting, source_stories, style_list_rendering, table_payload_cell_formatting,
};

/// What a projection reports besides the document.
#[derive(Default)]
pub(crate) struct Hooks {
    pub on_story: bool,
    pub stories: Vec<String>,
    /// `onEmbed(story, offset, content)` calls in order.
    pub embeds: Option<Vec<(String, f64, V)>>,
    /// `onParagraph(story, offset, paragraph, logicalId)` calls in order.
    pub paragraphs: Option<Vec<(String, f64, V, String)>>,
}

const PARA_ID_LIMIT: u32 = 0x7fff_ffff;

fn is_para_id(id: &str) -> bool {
    id.len() == 8
        && id.bytes().all(|byte| byte.is_ascii_hexdigit())
        && u32::from_str_radix(id, 16).is_ok_and(|value| value < PARA_ID_LIMIT)
}

struct SeedSources {
    styles: Option<StyleResolver>,
    numbering: Numbering,
    /// List renderings by `listKey`.
    lists: HashMap<String, V>,
}

/// paragraphSeed.ts `seededParagraphProperties`.
fn seeded_paragraph_properties(formatting: &V, list: &V, style_ppr: Option<&V>) -> V {
    let f = |key: &str| formatting.get(key);
    let (first_line, hanging) = if !f("indentFirstLine").is_undef() {
        (f("indentFirstLine"), f("hangingIndent"))
    } else {
        let first = list.get("indentFirstLine");
        let hanging = if first.is_undef() {
            V::Undef
        } else {
            list.get("hangingIndent")
        };
        (first, hanging)
    };
    let attrs = Obj::new();
    let or_null = |value: V| value.or_else(|| V::Null);
    match style_ppr {
        Some(style) => {
            let pick = |key: &str| or_null(f(key).or_else(|| style.get(key)));
            for key in [
                "alignment",
                "spaceBefore",
                "spaceAfter",
                "spaceBeforeLines",
                "spaceAfterLines",
                "beforeAutospacing",
                "afterAutospacing",
                "lineSpacing",
                "lineSpacingRule",
            ] {
                attrs.set(key, pick(key));
            }
            attrs.set("spacingExplicit", f("spacingExplicit").or(|| V::Null));
            attrs.set(
                "indentLeft",
                or_null(
                    f("indentLeft")
                        .or_else(|| list.get("indentLeft"))
                        .or_else(|| style.get("indentLeft")),
                ),
            );
            attrs.set("indentRight", pick("indentRight"));
            let removed = f("numPr").get("numId").same(&V::Num(0.0))
                && style.get("numPr").truthy()
                && !style.get("numPr").get("numId").same(&V::Num(0.0));
            let style_first = if removed { V::Undef } else { style.clone() };
            attrs.set(
                "indentFirstLine",
                or_null(
                    first_line
                        .clone()
                        .or_else(|| style_first.get("indentFirstLine")),
                ),
            );
            attrs.set(
                "hangingIndent",
                (if !first_line.is_undef() {
                    hanging
                } else {
                    style_first.get("hangingIndent")
                })
                .or_else(|| V::Bool(false)),
            );
            for key in [
                "borders",
                "shading",
                "tabs",
                "pageBreakBefore",
                "keepNext",
                "keepLines",
                "widowControl",
                "contextualSpacing",
                "snapToGrid",
                "autoSpaceDE",
                "autoSpaceDN",
                "outlineLevel",
                "bidi",
            ] {
                attrs.set(key, pick(key));
            }
            let style_num = style.get("numPr");
            if !f("numPr").truthy()
                && style_num.truthy()
                && !style_num.get("numId").same(&V::Num(0.0))
            {
                attrs.set("numPr", style_num.clone());
                attrs.set("numPrFromStyle", style_num);
            }
        }
        None => {
            for key in [
                "alignment",
                "spaceBefore",
                "spaceAfter",
                "spaceBeforeLines",
                "spaceAfterLines",
                "beforeAutospacing",
                "afterAutospacing",
                "lineSpacing",
                "lineSpacingRule",
            ] {
                attrs.set(key, or_null(f(key)));
            }
            attrs.set("spacingExplicit", f("spacingExplicit").or(|| V::Null));
            attrs.set(
                "indentLeft",
                or_null(f("indentLeft").or_else(|| list.get("indentLeft"))),
            );
            attrs.set("indentRight", or_null(f("indentRight")));
            attrs.set("indentFirstLine", or_null(first_line));
            attrs.set("hangingIndent", hanging.or_else(|| V::Bool(false)));
            for key in [
                "borders",
                "shading",
                "tabs",
                "pageBreakBefore",
                "keepNext",
                "keepLines",
                "widowControl",
                "snapToGrid",
                "autoSpaceDE",
                "autoSpaceDN",
                "outlineLevel",
                "bidi",
            ] {
                attrs.set(key, or_null(f(key)));
            }
        }
    }
    V::Obj(attrs)
}

fn list_key(formatting: &V, num_pr: &V) -> String {
    let null = |value: V| value.or_else(|| V::Null);
    stringify(&arr![
        num_pr.get("numId"),
        null(num_pr.get("ilvl")).or_else(|| V::Num(0.0)),
        null(formatting.get("styleId")),
        !formatting.get("numPrFromStyle").nullish(),
        !formatting.get("indentLeft").nullish() || !formatting.get("indentLeftChars").nullish(),
        formatting.get("indentFirstLine").truthy()
            || formatting.get("indentFirstLineChars").truthy(),
    ])
    .unwrap_or_default()
}

pub(crate) struct SaveContext<'a> {
    doc: &'a EditingDoc,
    pub story_ids: HashSet<String>,
    pub projected_comments: Vec<V>,
    base_paragraphs: HashMap<String, V>,
    para_ids: HashSet<u32>,
    base_stories: HashMap<String, Vec<V>>,
    source: SourceStories,
    seed_sources: SeedSources,
    cell_formatting: HashMap<String, V>,
    cell_contexts: HashMap<String, V>,
    style_pprs: HashMap<String, V>,
    comments: HashMap<String, Vec<(f64, f64, f64)>>,
    /// Projected blocks by identity: their story, the block kept alive.
    story_owners: HashMap<usize, (Obj, String)>,
    projected_stories: HashSet<String>,
    referenced_comments: Option<HashSet<u64>>,
    hooks: Hooks,
}

/// One story's projection state (`storyToBlocks` locals).
struct StoryState {
    blocks: Vec<V>,
    slot_breaks: Vec<SlotBreak>,
    previous: Option<usize>,
    previous_base: Option<V>,
}

impl<'a> SaveContext<'a> {
    pub fn new(doc: &'a EditingDoc, base: &V, hooks: Hooks) -> Result<Self> {
        let story_ids: HashSet<String> = read::story_ids(doc).into_iter().collect();
        let base_paragraphs = collect_base_paragraphs(base);
        let para_ids = base_paragraphs
            .keys()
            .filter_map(|id| parse_int_hex(id))
            .collect();
        let source = source_stories(base);
        let base_stories = source.stories.clone();
        let package = base.get("package");
        let mut lists = HashMap::new();
        for story in &source.order {
            for block in &base_stories[story] {
                let formatting = block.get("formatting");
                let num_pr = formatting.get("numPr");
                let rendering = block.get("listRendering");
                if is(block, "paragraph")
                    && formatting.truthy()
                    && num_pr.truthy()
                    && rendering.truthy()
                {
                    lists.insert(list_key(&formatting, &num_pr), rendering);
                }
            }
        }
        let styles = package.get("styles");
        let seed_sources = SeedSources {
            styles: styles.truthy().then(|| StyleResolver::new(&styles)),
            numbering: Numbering::new(&package.get("numbering")),
            lists,
        };
        let projected_comments = project_comments(
            read::list_comments(doc)?,
            &package.get("document").get("comments").items(),
        )?;
        let comments = comment_ranges(doc, &projected_comments);
        Ok(Self {
            doc,
            story_ids,
            projected_comments,
            base_paragraphs,
            para_ids,
            base_stories,
            source,
            seed_sources,
            cell_formatting: HashMap::new(),
            cell_contexts: HashMap::new(),
            style_pprs: HashMap::new(),
            comments,
            story_owners: HashMap::new(),
            projected_stories: HashSet::new(),
            referenced_comments: None,
            hooks,
        })
    }

    pub fn into_hooks(self) -> Hooks {
        self.hooks
    }

    /// The `w14:paraId` a story paragraph saves with (`savedParaId`).
    fn saved_para_id(&mut self, id: &str) -> String {
        if is_para_id(id) {
            return id.to_owned();
        }
        let mut hash: u32 = 2_166_136_261;
        for unit in id.encode_utf16() {
            hash = (hash ^ unit as u32).wrapping_mul(16_777_619);
        }
        let mut value = hash % (PARA_ID_LIMIT - 1) + 1;
        while self.para_ids.contains(&value) {
            value = value % (PARA_ID_LIMIT - 1) + 1;
        }
        self.para_ids.insert(value);
        format!("{value:08X}")
    }

    /// The paragraph properties the seed gave a paragraph (`seededProperties`).
    fn seeded_properties(&mut self, story: &str, attrs: &V) -> Result<V> {
        let original = attrs.get("_originalFormatting");
        let num_from_style = attrs
            .get("numPrFromStyle")
            .or_else(|| original.get("numPrFromStyle"));
        let formatting = o(&original)
            .map(|original| original.spread())
            .unwrap_or_default();
        formatting.set("styleId", attrs.get("styleId").or_else(|| V::Undef));
        formatting.set(
            "numPrFromStyle",
            num_from_style.clone().or_else(|| V::Undef),
        );
        let formatting = V::Obj(formatting);
        let mut style_ppr: Option<V> = None;
        let mut has_styles = false;
        if self.seed_sources.styles.is_some() {
            has_styles = true;
            let cell_story = enclosing_cell_story(story);
            let cell = match &cell_story {
                Some(cell_story) => self.cell_context(cell_story)?,
                None => V::Undef,
            };
            let key = format!(
                "{}|{}",
                if cell.truthy() {
                    cell_story.clone().unwrap_or_default()
                } else {
                    String::new()
                },
                formatting
                    .get("styleId")
                    .or_else(|| V::str(""))
                    .to_js_string()
            );
            if !self.style_pprs.contains_key(&key) {
                let resolved = self
                    .seed_sources
                    .styles
                    .as_ref()
                    .map(|styles| styles.resolve_paragraph_style(&formatting.get("styleId"), &cell))
                    .unwrap_or_default();
                self.style_pprs.insert(key.clone(), resolved);
            }
            style_ppr = self.style_pprs.get(&key).cloned();
        }
        let num_pr = attrs.get("numPr");
        let mut rendering = if num_pr.truthy() {
            self.seed_sources
                .lists
                .get(&list_key(&formatting, &num_pr))
                .cloned()
                .unwrap_or_default()
        } else {
            V::Undef
        };
        // A list no source paragraph had with this style renders as the seed renders it.
        if !rendering.truthy() && num_pr.truthy() {
            let from_style = same_json(&num_from_style, &num_pr);
            let source = if from_style {
                style_ppr.clone().unwrap_or_default()
            } else {
                obj! { "numPr": num_pr.clone() }
            };
            rendering =
                style_list_rendering(&source, &self.seed_sources.numbering).unwrap_or_default();
        }
        let style = if has_styles {
            Some(style_ppr.unwrap_or_default())
        } else {
            None
        };
        Ok(seeded_paragraph_properties(
            &formatting,
            &rendering,
            style.as_ref(),
        ))
    }

    /// The table-style formatting a cell's paragraphs save against (`cellContext`).
    fn cell_context(&mut self, cell_story: &str) -> Result<V> {
        if let Some(context) = self.cell_contexts.get(cell_story) {
            return Ok(context.clone());
        }
        let current = self
            .cell_formatting
            .get(cell_story)
            .cloned()
            .unwrap_or_default();
        let mut context = current.clone();
        if let (Some(base), Some(styles)) =
            (self.source.cells.get(cell_story), &self.seed_sources.styles)
        {
            let seeded = source_cell_formatting(base, styles);
            // Only a cell whose look moved needs the costlier source check.
            if !same_json(&seeded, &current) && self.source_cell(Some(cell_story))? {
                context = seeded;
            }
        }
        self.cell_contexts
            .insert(cell_story.to_owned(), context.clone());
        Ok(context)
    }

    /// Whether `story` still holds one of its source cell's paragraphs (`heldSourceCell`).
    fn source_cell(&self, story: Option<&str>) -> Result<bool> {
        let Some(story) = story else { return Ok(false) };
        if !self.source.cells.contains_key(story) {
            return Ok(false);
        }
        let ids: HashSet<String> = self
            .source
            .stories
            .get(story)
            .map(Vec::as_slice)
            .unwrap_or_default()
            .iter()
            .filter(|block| is(block, "paragraph"))
            .enumerate()
            .map(|(index, block)| {
                let id = block.get("paraId");
                if id.nullish() {
                    format!("{story}:p{index}")
                } else {
                    id.to_js_string()
                }
            })
            .collect();
        Ok(read::paragraphs(self.doc, story)?
            .iter()
            .any(|paragraph| ids.contains(&paragraph.get("paraId").to_js_string())))
    }

    /// The paragraphs a vMerge continuation cell under `anchor` saves with.
    fn continuation_content(
        &self,
        anchor_story: Option<&str>,
        siblings: &[Option<String>],
    ) -> Result<Vec<V>> {
        if !self.source_cell(anchor_story)? {
            return Ok(Vec::new());
        }
        let Some(anchor) = anchor_story.and_then(|story| self.source.cells.get(story)) else {
            return Ok(Vec::new());
        };
        let mut held = Vec::new();
        for story in siblings {
            let story = story.as_deref();
            held.push(if self.source_cell(story)? {
                story.and_then(|story| self.source.cells.get(story))
            } else {
                None
            });
        }
        let sibling = held
            .into_iter()
            .flatten()
            .find(|cell| cell.table.same(&anchor.table));
        let row = sibling.map(|cell| cell.row as isize).unwrap_or(-1);
        let rows = anchor.table.get("rows").items();
        let cell_at = |index: usize| -> Option<V> {
            let mut column = 0.0;
            rows.get(index)?
                .get("cells")
                .items()
                .into_iter()
                .find(|candidate| {
                    let start = column;
                    column += candidate
                        .get("formatting")
                        .get("gridSpan")
                        .or_else(|| V::Num(1.0))
                        .to_number();
                    start == anchor.column
                })
        };
        let mut index = anchor.row as isize + 1;
        while index <= row {
            let merge = cell_at(index as usize)
                .map(|cell| cell.get("formatting").get("vMerge"))
                .unwrap_or_default();
            if merge.as_str().as_deref() != Some("continue") {
                return Ok(Vec::new());
            }
            index += 1;
        }
        if row <= anchor.row as isize {
            return Ok(Vec::new());
        }
        Ok(cell_at(row as usize)
            .map(|cell| cell.get("content").items())
            .unwrap_or_default()
            .into_iter()
            .filter(|block| is(block, "paragraph") && block.get("formatting").truthy())
            .map(|block| obj! { "type": "paragraph", "formatting": block.get("formatting"), "content": arr![] })
            .collect())
    }

    /// Whether a story holds comment `id`'s reference mark.
    fn referenced(&mut self, id: f64) -> bool {
        if self.referenced_comments.is_none() {
            let mut found = HashSet::new();
            let mut stories: Vec<&String> = self.story_ids.iter().collect();
            stories.sort();
            for story in stories {
                for segment in read::story_segments(self.doc, story).unwrap_or_default() {
                    if segment.get("kind").as_str().as_deref() == Some("embed")
                        && segment.get("payload").get("modelKind").as_str().as_deref()
                            == Some("commentReference")
                        && let Some(comment) = segment.get("payload").get("commentId").finite()
                    {
                        found.insert(comment.to_bits());
                    }
                }
            }
            self.referenced_comments = Some(found);
        }
        self.referenced_comments
            .as_ref()
            .is_some_and(|found| found.contains(&id.to_bits()))
    }

    fn cell_contents_story(&mut self, story: &V) -> Result<Vec<V>> {
        match story.as_str() {
            Some(story) if self.story_ids.contains(&*story) => self.story_to_blocks(&story),
            _ => Ok(Vec::new()),
        }
    }

    fn embed(&mut self, story: &str, offset: f64, content: V) {
        if let Some(embeds) = &mut self.hooks.embeds {
            embeds.push((story.to_owned(), offset, content));
        }
    }

    pub fn story_to_blocks(&mut self, story_id: &str) -> Result<Vec<V>> {
        self.projected_stories.insert(story_id.to_owned());
        let base_blocks = self.base_stories.get(story_id).cloned();
        let story_comments = self.comments.get(story_id).cloned().unwrap_or_default();
        let base_paragraph_blocks: Option<Vec<V>> = base_blocks.as_ref().map(|blocks| {
            blocks
                .iter()
                .filter(|block| is(block, "paragraph"))
                .cloned()
                .collect()
        });
        let segments = read::story_segments(self.doc, story_id)?;
        let mut state = StoryState {
            blocks: Vec::new(),
            slot_breaks: Vec::new(),
            previous: None,
            previous_base: None,
        };
        let mut items: Vec<V> = Vec::new();
        let mut paragraph_start = 0.0;
        // Past a table, content control or break that opens the paragraph's slot.
        let mut content_start = 0.0;
        let mut paragraph_index = 0;
        let mut story_offset = 0.0;
        let mut inline_embeds: Vec<(f64, V)> = Vec::new();
        let tracks_embeds = self.hooks.embeds.is_some();
        // A range that leaves the story it opened in covers its final pilcrow.
        let mut last_pilcrow = -1.0;
        let mut offset = 0.0;
        for segment in &segments {
            if segment.get("kind").as_str().as_deref() == Some("pilcrow") {
                last_pilcrow = offset;
            }
            offset += if is_kind(segment, "text") {
                text::len(&segment.get("text").to_js_string()) as f64
            } else {
                1.0
            };
        }
        for segment in &segments {
            let kind = segment.get("kind").as_str().unwrap_or_else(|| "".into());
            if &*kind == "text" {
                let source = segment.get("text").to_js_string();
                push_text(&mut items, &source, &segment.get("attributes"));
                story_offset += text::len(&source) as f64;
                continue;
            }
            if &*kind == "pilcrow" {
                let para_id = segment.get("paraId").to_js_string();
                let generated = format!("{story_id}:p{paragraph_index}");
                let saved = if para_id == generated && !self.base_paragraphs.contains_key(&para_id)
                {
                    String::new()
                } else {
                    self.saved_para_id(&para_id)
                };
                // A seeded id without w14:paraId names its source paragraph by index.
                let seeded: Option<usize> = para_id
                    .strip_prefix(&format!("{story_id}:p"))
                    .and_then(|rest| {
                        let number = V::str(rest).to_number();
                        (number.is_finite() && number.fract() == 0.0 && number >= 0.0)
                            .then_some(number as usize)
                    });
                let base_paragraph = self.base_paragraphs.get(&para_id).cloned().or_else(|| {
                    seeded.and_then(|index| {
                        base_paragraph_blocks
                            .as_ref()
                            .and_then(|blocks| blocks.get(index).cloned())
                    })
                });
                let (boundaries, slot) = paragraph_comment_boundaries(
                    &story_comments,
                    paragraph_start,
                    content_start,
                    story_offset,
                    last_pilcrow,
                );
                // Bookmarks at a slot of breaks are placed among them by settle.
                let (properties, slot_bookmarks) =
                    split_slot_bookmarks(&segment.get("properties"), !state.slot_breaks.is_empty());
                let story = story_id.to_owned();
                let paragraph = paragraph_from_story(
                    &saved,
                    &properties,
                    std::mem::take(&mut items),
                    &boundaries,
                    base_paragraph.as_ref(),
                    &mut |attrs| self.seeded_properties(&story, attrs),
                )?;
                let paragraph = self
                    .settle(&mut state, Some(paragraph.clone()), &slot, &slot_bookmarks)?
                    .unwrap_or(paragraph);
                state.blocks.push(paragraph.clone());
                state.previous = Some(state.blocks.len() - 1);
                state.previous_base = base_paragraph;
                if let Some(paragraphs) = &mut self.hooks.paragraphs {
                    paragraphs.push((story_id.to_owned(), story_offset, paragraph, para_id));
                }
                paragraph_index += 1;
                story_offset += 1.0;
                paragraph_start = story_offset;
                content_start = story_offset;
                continue;
            }
            let embed_kind = segment
                .get("embedKind")
                .as_str()
                .unwrap_or_else(|| "".into());
            let payload = segment.get("payload");
            let mut projected: Option<V> = None;
            match &*embed_kind {
                "table" => {
                    for (cell, formatting) in
                        table_payload_cell_formatting(&payload, self.seed_sources.styles.as_ref())
                    {
                        self.cell_formatting.insert(cell, formatting);
                    }
                    let table = self.table_from_payload(&payload)?;
                    projected = Some(table.clone());
                    self.open_block(&mut state)?;
                    state.blocks.push(table);
                }
                "blockSdt" => {
                    let child_story = payload.get("story").as_str();
                    self.open_block(&mut state)?;
                    let child_content = match &child_story {
                        Some(child) if self.story_ids.contains(&**child) => {
                            self.story_to_blocks(child)?
                        }
                        _ => Vec::new(),
                    };
                    let mut properties = sdt_attrs_to_props(&payload);
                    let mut content = child_content;
                    if let Some(value) = content_control_value(&payload.get("value"))
                        && let Ok((applied, blocks)) =
                            apply_content_control_value(&properties, &value)
                    {
                        properties = applied;
                        content = blocks;
                    }
                    let block = obj! {
                        "type": "blockSdt",
                        "properties": properties,
                        "content": V::array(content),
                    };
                    projected = Some(block.clone());
                    state.blocks.push(block);
                }
                "opaque" => {
                    let blob = payload.get("blob");
                    if is(&blob, "pageBreak") {
                        let paragraph = page_break_paragraph();
                        projected = Some(paragraph.clone());
                        self.open_block(&mut state)?;
                        state.blocks.push(paragraph);
                    } else {
                        // Carry a same-position base block while block SDTs stay opaque.
                        let base_block = base_blocks
                            .as_ref()
                            .and_then(|blocks| blocks.get(state.blocks.len()).cloned());
                        if is(&blob, "blockSdt")
                            && let Some(base_block) =
                                base_block.filter(|block| is(block, "blockSdt"))
                        {
                            projected = Some(base_block.clone());
                            self.open_block(&mut state)?;
                            state.blocks.push(base_block);
                        }
                    }
                }
                "pageBreak" | "columnBreak" => {
                    let entry = SlotBreak {
                        kind: if &*embed_kind == "pageBreak" {
                            "page"
                        } else {
                            "column"
                        },
                        at: story_offset,
                        leading: payload.get("leading").same(&V::Bool(true)),
                        trailing: payload.get("trailing").same(&V::Bool(true)),
                        attributes: segment.get("attributes"),
                    };
                    if items.is_empty() {
                        state.slot_breaks.push(entry);
                    } else {
                        items.push(obj! {
                            "kind": "embed",
                            "embedKind": "flowBreak",
                            "payload": obj! { "breakType": entry.kind },
                            "attributes": entry.attributes,
                        });
                    }
                }
                _ => {
                    items.push(segment.clone());
                    if tracks_embeds {
                        inline_embeds.push((story_offset, segment.clone()));
                    }
                }
            }
            if let Some(projected) = projected {
                self.embed(story_id, story_offset, projected);
            }
            story_offset += 1.0;
            if items.is_empty() {
                content_start = story_offset;
            }
        }
        self.settle(&mut state, None, &[], &[])?;
        // Defensive recovery for malformed stories without a final pilcrow.
        if !items.is_empty() {
            let trailing = build_paragraph_content(items)?;
            if !trailing.is_empty() {
                state
                    .blocks
                    .push(obj! { "type": "paragraph", "content": V::array(trailing) });
            }
        }
        for (offset, item) in inline_embeds {
            if let Some(content) = ordinary_content_for_item(&item)? {
                self.embed(story_id, offset, content);
            }
        }
        let projected =
            restore_raw_blocks(state.blocks, base_blocks.as_deref().unwrap_or_default());
        for block in &projected {
            if let Some(object) = block.obj() {
                self.story_owners
                    .insert(object.id(), (object.clone(), story_id.to_owned()));
            }
        }
        Ok(projected)
    }

    fn open_block(&mut self, state: &mut StoryState) -> Result<()> {
        self.settle(state, None, &[], &[])?;
        state.previous = None;
        state.previous_base = None;
        Ok(())
    }

    /// Places the slot's breaks between the paragraph before and `next` (`settle`).
    fn settle(
        &mut self,
        state: &mut StoryState,
        next: Option<V>,
        marks: &[V],
        bookmarks: &[V],
    ) -> Result<Option<V>> {
        let expected = std::mem::take(&mut state.slot_breaks);
        let before = state.previous.map(|index| state.blocks[index].clone());
        let before_flow = match &before {
            Some(before) => {
                let tokens = flow_tokens(&before.get("content").items());
                Some(split_flow(&tokens, &mut |id| self.referenced(id)))
            }
            None => None,
        };
        let next_tokens = next
            .as_ref()
            .map(|next| flow_tokens(&next.get("content").items()))
            .unwrap_or_default();
        let last_held_break = next_tokens.iter().rposition(FlowToken::is_break);
        let after_last_held = |index: usize| last_held_break.is_none_or(|last| index > last);
        let next_flow = next
            .as_ref()
            .map(|_| split_flow(&next_tokens, &mut |id| self.referenced(id)));
        // Content for the breaks to lead: text, or a bookmark after them.
        let content = next.is_some()
            && (next_tokens.iter().enumerate().any(|(index, token)| {
                token.kind == TokenKind::Visible
                    || token.kind == TokenKind::Mark
                    || (token.kind == TokenKind::Reference && after_last_held(index))
            }) || bookmarks
                .iter()
                .any(|bookmark| !bookmark.get("breaksAfter").truthy()));
        let last_break = expected.last().map(|entry| entry.at);
        let text = content
            || (next.is_some()
                && last_break.is_some()
                && (next_tokens
                    .iter()
                    .filter(|token| token.kind == TokenKind::End)
                    .map(|token| token.id.unwrap_or(f64::NAN))
                    .collect::<Vec<_>>()
                    .into_iter()
                    .any(|id| !self.referenced(id))
                    || marks
                        .iter()
                        .filter(|mark| {
                            mark.get("kind").as_str().as_deref() == Some("end")
                                && mark.get("offset").to_number() <= last_break.unwrap_or(f64::NAN)
                        })
                        .map(|mark| mark.get("id").to_number())
                        .collect::<Vec<_>>()
                        .into_iter()
                        .any(|id| !self.referenced(id))));
        let owned = expected.last().is_some_and(|entry| entry.kind == "column")
            && !expected.iter().any(|entry| entry.trailing)
            && (bookmarks.is_empty() || before.is_none());
        let flagged = expected.iter().position(|entry| entry.leading);
        let leading: isize = match flagged {
            Some(flagged) => flagged as isize,
            None if content && owned => 0,
            None => -1,
        };
        let own = !content
            && next.is_some()
            && (bookmarks.is_empty() || before.is_none())
            && (before.is_none() || !expected.last().is_some_and(|entry| entry.trailing))
            && expected.last().is_some_and(|entry| entry.kind == "column");
        let trailing_marks = own && !bookmarks.is_empty();
        // A comment boundary in a paragraph without text sits after its breaks.
        let commented = !text
            && next.as_ref().is_some_and(|next| {
                next.get("content")
                    .items()
                    .iter()
                    .any(|child| matches!(&*ty(child), "commentRangeStart" | "commentRangeEnd"))
            });
        let marked: isize = if text {
            expected
                .iter()
                .position(|entry| {
                    entry.kind == "column"
                        && marks
                            .iter()
                            .any(|mark| mark.get("offset").to_number() <= entry.at)
                })
                .map(|index| index as isize)
                .unwrap_or(-1)
        } else {
            -1
        };
        let split: usize = if own {
            0
        } else if !text || (before.is_some() && leading < 0) {
            if marked < 0 {
                expected.len()
            } else {
                marked as usize
            }
        } else if before.is_some() {
            if marked < 0 {
                leading as usize
            } else {
                leading.min(marked) as usize
            }
        } else {
            0
        };
        let trailing = before_flow
            .as_ref()
            .map(|(_, trailing)| trailing.len())
            .unwrap_or(0);
        let mut have: Vec<FlowToken> = before_flow
            .as_ref()
            .map(|(_, trailing)| trailing.clone())
            .unwrap_or_default();
        have.extend(
            next_flow
                .as_ref()
                .map(|(leading, _)| leading.clone())
                .unwrap_or_default(),
        );
        let held_before_break = next_tokens.iter().enumerate().any(|(index, token)| {
            matches!(token.kind, TokenKind::Start | TokenKind::End)
                && next_tokens[index + 1..].iter().any(FlowToken::is_break)
        });
        if have.len() == expected.len()
            && have.iter().enumerate().all(|(index, token)| {
                let entry = &expected[index];
                let kind = if token.kind == TokenKind::Page {
                    "page"
                } else {
                    "column"
                };
                kind == entry.kind
                    && token.tracked == tracked_kind(&entry.attributes)
                    && ((token.kind == TokenKind::Column && !entry.leading && !entry.trailing)
                        || (!text && !commented)
                        || (index < trailing) == (index < split))
            })
            && marks.is_empty()
            && bookmarks.is_empty()
            && !commented
            && !held_before_break
        {
            return Ok(next);
        }
        // Breaks the content holds ahead of its text are the slot's.
        let first = next_tokens.iter().enumerate().position(|(index, token)| {
            token.kind == TokenKind::Visible
                || (token.kind == TokenKind::Reference && after_last_held(index))
        });
        let held: Vec<FlowToken> = if !bookmarks.is_empty() {
            next_tokens[..first.unwrap_or(next_tokens.len())]
                .iter()
                .filter(|token| token.is_break())
                .cloned()
                .collect()
        } else {
            next_flow
                .as_ref()
                .map(|(leading, _)| leading.clone())
                .unwrap_or_default()
        };
        let rest = match &next {
            Some(next) if !held.is_empty() => Some(V::Obj(o(next).unwrap_or_default().with(&[(
                "content",
                V::array(without_breaks(&next.get("content").items(), &held)),
            )]))),
            other => other.clone(),
        };
        if let (Some(before), Some(before_flow)) = (&before, &before_flow) {
            let kept = without_breaks(&before.get("content").items(), &before_flow.1);
            let placed = place_trailing(
                kept,
                &expected[..split.min(expected.len())],
                state.previous_base.as_ref(),
            )?;
            let index = state.previous.unwrap_or_default();
            state.blocks[index] = V::Obj(
                o(before)
                    .unwrap_or_default()
                    .with(&[("content", V::array(placed))]),
            );
            return self.lead(
                rest,
                &expected[split.min(expected.len())..],
                marks,
                bookmarks,
                trailing_marks,
            );
        }
        if text || own {
            return self.lead(rest, &expected, marks, bookmarks, trailing_marks);
        }
        // No paragraph before them and no text to lead: a paragraph of their own.
        if !expected.is_empty() {
            let content = expected
                .iter()
                .map(slot_break_content)
                .collect::<Result<Vec<_>>>()?;
            state
                .blocks
                .push(obj! { "type": "paragraph", "content": V::array(content) });
        }
        self.lead(rest, &[], marks, bookmarks, trailing_marks)
    }

    /// The paragraph with the slot's breaks, boundaries and bookmarks ahead of
    /// its content (`lead` in `settle`).
    fn lead(
        &mut self,
        paragraph: Option<V>,
        entries: &[SlotBreak],
        marks: &[V],
        bookmarks: &[V],
        trailing_marks: bool,
    ) -> Result<Option<V>> {
        let Some(paragraph) = paragraph else {
            return Ok(None);
        };
        if entries.is_empty() && marks.is_empty() && bookmarks.is_empty() {
            return Ok(Some(paragraph));
        }
        let mut head: Vec<V> = Vec::new();
        let mut mark = 0;
        let mut placed: HashSet<usize> = HashSet::new();
        let mut references: Vec<f64> = Vec::new();
        let infinite = f64::INFINITY;
        let mut mark_to =
            |this: &mut Self, head: &mut Vec<V>, references: &mut Vec<f64>, at: f64, index: f64| {
                for (position, bookmark) in bookmarks.iter().enumerate() {
                    let due = if trailing_marks {
                        index == infinite
                    } else {
                        entries.len() as f64
                            - bookmark
                                .get("breaksAfter")
                                .or_else(|| V::Num(0.0))
                                .to_number()
                            <= index
                    };
                    if !placed.contains(&position) && due {
                        placed.insert(position);
                        head.push(super::paragraph::bookmark_node(bookmark));
                    }
                }
                while mark < marks.len() && marks[mark].get("offset").to_number() <= at {
                    let boundary = &marks[mark];
                    mark += 1;
                    head.push(comment_marker(boundary));
                    let id = boundary.get("id").to_number();
                    if boundary.get("kind").as_str().as_deref() == Some("end")
                        && index < entries.len() as f64
                        && !this.referenced(id)
                    {
                        references.push(id);
                    }
                }
            };
        for (index, entry) in entries.iter().enumerate() {
            mark_to(self, &mut head, &mut references, entry.at, index as f64);
            head.push(slot_break_content(entry)?);
        }
        for id in &references {
            head.push(obj! {
                "type": "run",
                "formatting": obj! { "styleId": "CommentReference" },
                "content": arr![obj! { "type": "commentReference", "id": *id }],
            });
        }
        mark_to(self, &mut head, &mut references, infinite, infinite);
        head.extend(paragraph.get("content").items());
        Ok(Some(V::Obj(
            o(&paragraph)
                .unwrap_or_default()
                .with(&[("content", V::array(head))]),
        )))
    }

    fn table_cell_from_payload(&mut self, payload: &V) -> Result<V> {
        let attrs = Obj::new();
        for (key, value) in [
            ("colspan", V::Num(1.0)),
            ("rowspan", V::Num(1.0)),
            ("colwidth", V::Null),
            ("width", V::Null),
            ("widthType", V::Null),
            ("verticalAlign", V::Null),
            ("backgroundColor", V::Null),
            ("borders", V::Null),
            ("margins", V::Null),
            ("textDirection", V::Null),
            ("noWrap", V::Bool(false)),
            ("_originalFormatting", V::Null),
            ("_originalResolvedFill", V::Null),
            ("cellMarker", V::Null),
            ("tcPrChange", V::Null),
        ] {
            attrs.set(key, value);
        }
        if let Some(tc_pr) = o(&payload.get("tcPr")) {
            attrs.assign(&tc_pr);
        }
        let attrs = V::Obj(attrs);
        let story = payload.get("story");
        let content = if story.truthy() {
            self.cell_contents_story(&story)?
        } else {
            Vec::new()
        };
        let cell = Obj::new();
        cell.set("type", V::str("tableCell"));
        cell.set("formatting", table_cell_attrs_to_formatting(&attrs));
        cell.set("content", V::array(content));
        let marker = attrs.get("cellMarker");
        if marker.truthy() {
            let info = tracked_info(&marker.get("info"))
                .unwrap_or_else(|| obj! { "id": 0.0, "author": "Unknown" });
            let change = match marker.get("kind").as_str().as_deref() {
                Some("ins") => obj! { "type": "tableCellInsertion", "info": info },
                Some("del") => obj! { "type": "tableCellDeletion", "info": info },
                _ => {
                    let change = Obj::new();
                    change.set("type", V::str("tableCellMerge"));
                    change.set("info", info);
                    if marker.get("vMerge").truthy() {
                        change.set("vMerge", marker.get("vMerge"));
                    }
                    if marker.get("vMergeOrig").truthy() {
                        change.set("vMergeOrig", marker.get("vMergeOrig"));
                    }
                    V::Obj(change)
                }
            };
            cell.set("structuralChange", change);
        }
        if attrs
            .get("tcPrChange")
            .arr()
            .is_some_and(|changes| !changes.is_empty())
        {
            cell.set("propertyChanges", attrs.get("tcPrChange"));
        }
        Ok(V::Obj(cell))
    }

    fn table_from_payload(&mut self, payload: &V) -> Result<V> {
        let row_payloads = payload.get("rows").items();
        let mut occupied: Vec<Vec<bool>> = Vec::new();
        struct Anchor {
            row: usize,
            col: usize,
            rowspan: usize,
            colspan: usize,
            cell: V,
            story: V,
        }
        let mut anchors: Vec<Anchor> = Vec::new();
        let mut total_columns = 0;
        for (row_index, row) in row_payloads.iter().enumerate() {
            let mut column = 0;
            for cell_payload in row.get("cells").items() {
                while occupied
                    .get(row_index)
                    .and_then(|row| row.get(column))
                    .copied()
                    .unwrap_or(false)
                {
                    column += 1;
                }
                let tc_pr = cell_payload.get("tcPr").or_else(|| obj! {});
                let rowspan = finite_or_one(&tc_pr.get("rowspan"));
                let colspan = finite_or_one(&tc_pr.get("colspan"));
                anchors.push(Anchor {
                    row: row_index,
                    col: column,
                    rowspan,
                    colspan,
                    cell: self.table_cell_from_payload(&cell_payload)?,
                    story: cell_payload.get("story"),
                });
                for r in row_index..row_index + rowspan {
                    if occupied.len() <= r {
                        occupied.resize(r + 1, Vec::new());
                    }
                    for c in column..column + colspan {
                        if occupied[r].len() <= c {
                            occupied[r].resize(c + 1, false);
                        }
                        occupied[r][c] = true;
                    }
                }
                column += colspan;
                total_columns = total_columns.max(column);
            }
        }
        let by_start: HashMap<(usize, usize), usize> = anchors
            .iter()
            .enumerate()
            .map(|(index, anchor)| ((anchor.row, anchor.col), index))
            .collect();
        let mut by_covered: HashMap<(usize, usize), usize> = HashMap::new();
        for (index, anchor) in anchors.iter().enumerate() {
            for row in anchor.row..anchor.row + anchor.rowspan {
                for col in anchor.col..anchor.col + anchor.colspan {
                    by_covered.insert((row, col), index);
                }
            }
        }
        let grid: Vec<V> = payload
            .get("grid")
            .items()
            .into_iter()
            .filter(|width| width.as_num().is_some())
            .collect();
        let mut rows = Vec::new();
        for (row_index, row_payload) in row_payloads.iter().enumerate() {
            let mut cells = Vec::new();
            let mut col = 0;
            while col < total_columns {
                if let Some(&index) = by_start.get(&(row_index, col)) {
                    let anchor = &anchors[index];
                    let formatting = o(&anchor.cell.get("formatting"))
                        .map(|f| f.spread())
                        .unwrap_or_default();
                    if anchor.colspan > 1 {
                        formatting.set("gridSpan", V::Num(anchor.colspan as f64));
                    } else {
                        formatting.delete("gridSpan");
                    }
                    let merge = formatting.get("vMerge").as_str();
                    if anchor.rowspan > 1 {
                        formatting.set("vMerge", V::str("restart"));
                    } else if !matches!(merge.as_deref(), Some("restart" | "continue")) {
                        formatting.delete("vMerge");
                    }
                    let value = if formatting.is_empty() {
                        V::Undef
                    } else {
                        V::Obj(formatting)
                    };
                    cells.push(V::Obj(
                        o(&anchor.cell)
                            .unwrap_or_default()
                            .with(&[("formatting", value)]),
                    ));
                    col += anchor.colspan;
                    continue;
                }
                let Some(&index) = by_covered.get(&(row_index, col)) else {
                    col += 1;
                    continue;
                };
                let covering = &anchors[index];
                let formatting = o(&covering.cell.get("formatting"))
                    .map(|f| f.spread())
                    .unwrap_or_default();
                if covering.colspan > 1 {
                    formatting.set("gridSpan", V::Num(covering.colspan as f64));
                } else {
                    formatting.delete("gridSpan");
                }
                formatting.set("vMerge", V::str("continue"));
                let siblings: Vec<Option<String>> = row_payload
                    .get("cells")
                    .items()
                    .iter()
                    .map(|cell| cell.get("story").as_str().map(|story| story.to_string()))
                    .collect();
                let story = covering.story.as_str().map(|story| story.to_string());
                let content = self.continuation_content(story.as_deref(), &siblings)?;
                cells.push(V::Obj(o(&covering.cell).unwrap_or_default().with(&[
                    ("content", V::array(content)),
                    ("formatting", V::Obj(formatting)),
                ])));
                col += covering.colspan;
            }
            let attrs = Obj::new();
            for (key, value) in [
                ("height", V::Null),
                ("heightRule", V::Null),
                ("isHeader", V::Bool(false)),
                ("_originalFormatting", V::Null),
                ("trIns", V::Null),
                ("trDel", V::Null),
                ("trPrChange", V::Null),
            ] {
                attrs.set(key, value);
            }
            if let Some(tr_pr) = o(&row_payload.get("trPr")) {
                attrs.assign(&tr_pr);
            }
            let attrs = V::Obj(attrs);
            let formatting = table_row_attrs_to_formatting(&attrs);
            // A row keeps the grid columns it skips only while they still fit the grid.
            if formatting.truthy()
                && (formatting.get("gridBefore").truthy() || formatting.get("gridAfter").truthy())
            {
                let spans: f64 = cells
                    .iter()
                    .map(|cell| {
                        cell.get("formatting")
                            .get("gridSpan")
                            .or_else(|| V::Num(1.0))
                            .to_number()
                    })
                    .sum();
                let columns = if grid.is_empty() {
                    total_columns as f64
                } else {
                    grid.len() as f64
                };
                let before = formatting
                    .get("gridBefore")
                    .or_else(|| V::Num(0.0))
                    .to_number();
                let after = formatting
                    .get("gridAfter")
                    .or_else(|| V::Num(0.0))
                    .to_number();
                if spans + before + after != columns
                    && let Some(object) = formatting.obj()
                {
                    for key in ["gridBefore", "gridAfter", "widthBefore", "widthAfter"] {
                        object.delete(key);
                    }
                }
            }
            let row = Obj::new();
            row.set("type", V::str("tableRow"));
            row.set("formatting", formatting);
            row.set("cells", V::array(cells));
            if let Some(info) = tracked_info(&attrs.get("trIns")) {
                row.set(
                    "structuralChange",
                    obj! { "type": "tableRowInsertion", "info": info },
                );
            } else if let Some(info) = tracked_info(&attrs.get("trDel")) {
                row.set(
                    "structuralChange",
                    obj! { "type": "tableRowDeletion", "info": info },
                );
            }
            if attrs
                .get("trPrChange")
                .arr()
                .is_some_and(|changes| !changes.is_empty())
            {
                row.set("propertyChanges", attrs.get("trPrChange"));
            }
            rows.push(V::Obj(row));
        }
        normalize_vmerge_runs(&rows);
        let attrs = Obj::new();
        for key in [
            "styleId",
            "width",
            "widthType",
            "justification",
            "columnWidths",
            "tableLayout",
            "floating",
            "cellMargins",
            "look",
            "bidi",
            "_originalFormatting",
            "tblPrChange",
        ] {
            attrs.set(key, V::Null);
        }
        if let Some(tbl_pr) = o(&payload.get("tblPr")) {
            attrs.assign(&tbl_pr);
        }
        attrs.set(
            "columnWidths",
            if grid.is_empty() {
                V::Undef
            } else {
                V::array(grid)
            },
        );
        let attrs = V::Obj(attrs);
        let mut formatting = table_attrs_to_formatting(&attrs);
        if !formatting.get("borders").truthy()
            && let Some(borders) = infer_table_borders(&rows)
        {
            let object = o(&formatting).map(|f| f.spread()).unwrap_or_default();
            object.set("borders", borders);
            formatting = V::Obj(object);
        }
        let table = Obj::new();
        table.set("type", V::str("table"));
        table.set("columnWidths", attrs.get("columnWidths").or(|| V::Undef));
        table.set("formatting", formatting);
        table.set("rows", V::array(rows));
        if attrs
            .get("tblPrChange")
            .arr()
            .is_some_and(|changes| !changes.is_empty())
        {
            table.set("propertyChanges", attrs.get("tblPrChange"));
        }
        Ok(V::Obj(table))
    }

    /// The stories whose projected blocks remain reachable (`visitReachableStories`).
    pub fn visit_reachable_stories(&mut self, document: &V) {
        let mut reachable: Vec<String> = Vec::new();
        let mut seen: HashSet<String> = HashSet::new();
        let owners = &self.story_owners;
        let mut add = |story: &String, reachable: &mut Vec<String>| {
            if seen.insert(story.clone()) {
                reachable.push(story.clone());
            }
        };
        fn walk(
            blocks: &[V],
            owners: &HashMap<usize, (Obj, String)>,
            add: &mut dyn FnMut(&String, &mut Vec<String>),
            reachable: &mut Vec<String>,
        ) {
            for block in blocks {
                if let Some((_, owner)) = block.obj().and_then(|object| owners.get(&object.id())) {
                    add(owner, reachable);
                }
                if is(block, "table") {
                    for row in block.get("rows").items() {
                        for cell in row.get("cells").items() {
                            walk(&cell.get("content").items(), owners, add, reachable);
                        }
                    }
                } else if is(block, "blockSdt") {
                    walk(&block.get("content").items(), owners, add, reachable);
                }
            }
        }
        let package = document.get("package");
        walk(
            &package.get("document").get("content").items(),
            owners,
            &mut add,
            &mut reachable,
        );
        for parts in [package.get("headers"), package.get("footers")] {
            for (_, part) in map_entries(&parts) {
                walk(
                    &part.get("content").items(),
                    owners,
                    &mut add,
                    &mut reachable,
                );
            }
        }
        for (prefix, notes) in [
            ("fn:", package.get("footnotes")),
            ("en:", package.get("endnotes")),
        ] {
            for note in notes.items() {
                let story = format!("{prefix}{}", note.get("id").to_js_string());
                if self.projected_stories.contains(&story) {
                    add(&story, &mut reachable);
                }
                walk(
                    &note.get("content").items(),
                    owners,
                    &mut add,
                    &mut reachable,
                );
            }
        }
        self.hooks.stories = reachable;
    }
}

fn parse_int_hex(id: &str) -> Option<u32> {
    let digits: String = id.chars().take_while(char::is_ascii_hexdigit).collect();
    u32::from_str_radix(&digits, 16).ok()
}

/// `asFiniteNumber(value) || 1` for a span.
fn finite_or_one(value: &V) -> usize {
    match value.finite() {
        Some(number) if number != 0.0 => number.max(0.0) as usize,
        _ => 1,
    }
}

fn push_text(items: &mut Vec<V>, source: &str, attributes: &V) {
    let mut cursor = String::new();
    for ch in source.chars() {
        if ch == '\t' {
            if !cursor.is_empty() {
                items.push(obj! { "kind": "text", "text": std::mem::take(&mut cursor), "attributes": attributes.clone() });
            }
            items.push(obj! {
                "kind": "embed",
                "embedKind": "tab",
                "payload": obj! {},
                "attributes": attributes.clone(),
            });
        } else {
            cursor.push(ch);
        }
    }
    if !cursor.is_empty() {
        items.push(obj! { "kind": "text", "text": cursor, "attributes": attributes.clone() });
    }
}

/// Boundaries in a paragraph's text, and those at the units opening its slot.
fn paragraph_comment_boundaries(
    ranges: &[(f64, f64, f64)],
    paragraph_start: f64,
    content_start: f64,
    end: f64,
    last_pilcrow: f64,
) -> (Vec<V>, Vec<V>) {
    let mut boundaries = Vec::new();
    let mut slot = Vec::new();
    let mut add = |id: f64, kind: &str, at: f64| {
        if at < paragraph_start || at > end {
            return;
        }
        if at < content_start {
            slot.push(obj! { "id": id, "kind": kind, "offset": at });
        } else {
            boundaries.push(obj! { "id": id, "kind": kind, "offset": at - content_start });
        }
    };
    for &(id, start, stop) in ranges {
        add(id, "start", start);
        add(id, "end", stop.min(last_pilcrow));
    }
    (boundaries, sort_boundaries(&slot))
}

fn collect_base_paragraphs(document: &V) -> HashMap<String, V> {
    let mut paragraphs: HashMap<String, V> = HashMap::new();
    fn visit(paragraphs: &mut HashMap<String, V>, blocks: &[V]) {
        for block in blocks {
            match &*ty(block) {
                "paragraph" => {
                    let id = block.get("paraId");
                    if id.truthy() {
                        paragraphs
                            .entry(id.to_js_string())
                            .or_insert_with(|| block.clone());
                    }
                }
                "table" => {
                    for row in block.get("rows").items() {
                        for cell in row.get("cells").items() {
                            visit(paragraphs, &cell.get("content").items());
                        }
                    }
                }
                "blockSdt" => visit(paragraphs, &block.get("content").items()),
                _ => {}
            }
        }
    }
    let package = document.get("package");
    visit(
        &mut paragraphs,
        &package.get("document").get("content").items(),
    );
    for parts in [package.get("headers"), package.get("footers")] {
        for (_, part) in map_entries(&parts) {
            visit(&mut paragraphs, &part.get("content").items());
        }
    }
    for notes in [package.get("footnotes"), package.get("endnotes")] {
        for note in notes.items() {
            visit(&mut paragraphs, &note.get("content").items());
        }
    }
    paragraphs
}

fn comment_ranges(doc: &EditingDoc, comments: &[V]) -> HashMap<String, Vec<(f64, f64, f64)>> {
    let mut by_story: HashMap<String, Vec<(f64, f64, f64)>> = HashMap::new();
    for comment in comments {
        let Ok(anchors) = read::resolve_comment(doc, &comment_shared_id(comment)) else {
            continue;
        };
        let mut groups: Vec<(String, Vec<V>)> = Vec::new();
        for anchor in anchors {
            let story = anchor.get("story").to_js_string();
            match groups.iter_mut().find(|(key, _)| *key == story) {
                Some((_, group)) => group.push(anchor),
                None => groups.push((story, vec![anchor])),
            }
        }
        for (story, group) in groups {
            if group.is_empty() {
                continue;
            }
            let start = group
                .iter()
                .map(|anchor| anchor.get("start").to_number())
                .fold(f64::INFINITY, f64::min);
            let end = group
                .iter()
                .map(|anchor| anchor.get("end").to_number())
                .fold(f64::NEG_INFINITY, f64::max);
            by_story
                .entry(story)
                .or_default()
                .push((comment.get("id").to_number(), start, end));
        }
    }
    by_story
}

fn restore_raw_blocks(mut projected: Vec<V>, base: &[V]) -> Vec<V> {
    let mut offset: isize = 0;
    for (index, block) in base.iter().enumerate() {
        if !is(block, "rawXml") {
            continue;
        }
        let following = base[index + 1..].iter().find(|candidate| {
            is(candidate, "paragraph")
                && candidate.get("paraId").truthy()
                && projected.iter().any(|entry| {
                    is(entry, "paragraph") && entry.get("paraId").same(&candidate.get("paraId"))
                })
        });
        let anchor = following.and_then(|following| {
            projected.iter().position(|entry| {
                is(entry, "paragraph") && entry.get("paraId").same(&following.get("paraId"))
            })
        });
        let position = match anchor {
            Some(anchor) => anchor,
            None => ((index as isize + offset).max(0) as usize).min(projected.len()),
        };
        projected.insert(position, block.clone());
        offset = (position as isize - index as isize).max(0);
    }
    projected
}

/// Lifts a `w:tblBorders` back out of the per-cell edges seeding pushed down.
fn infer_table_borders(rows: &[V]) -> Option<V> {
    let first_row = rows.first()?.get("cells").items();
    let last_row = rows.last()?.get("cells").items();
    let corner = first_row.first()?.get("formatting").get("borders");
    if !corner.truthy() {
        return None;
    }
    let borders = Obj::new();
    let entries = [
        ("top", corner.get("top")),
        ("left", corner.get("left")),
        (
            "bottom",
            last_row
                .first()
                .map(|cell| cell.get("formatting").get("borders").get("bottom"))
                .unwrap_or_default(),
        ),
        (
            "right",
            first_row[first_row.len() - 1]
                .get("formatting")
                .get("borders")
                .get("right"),
        ),
        (
            "insideH",
            if rows.len() > 1 {
                corner.get("bottom")
            } else {
                V::Undef
            },
        ),
        (
            "insideV",
            if first_row.len() > 1 {
                corner.get("right")
            } else {
                V::Undef
            },
        ),
    ];
    for (key, value) in entries {
        if !value.is_undef() {
            borders.set(key, value);
        }
    }
    (!borders.is_empty()).then_some(V::Obj(borders))
}

fn normalize_vmerge_runs(rows: &[V]) {
    let mut columns: Vec<(i64, Vec<(usize, V)>)> = Vec::new();
    for (row_index, row) in rows.iter().enumerate() {
        let mut column = 0_i64;
        for cell in row.get("cells").items() {
            let start = column;
            column += cell
                .get("formatting")
                .get("gridSpan")
                .or_else(|| V::Num(1.0))
                .to_number() as i64;
            if cell.get("formatting").get("vMerge").truthy() {
                match columns.iter_mut().find(|(key, _)| *key == start) {
                    Some((_, entries)) => entries.push((row_index, cell)),
                    None => columns.push((start, vec![(row_index, cell)])),
                }
            }
        }
    }
    let clear = |cell: &V| {
        let formatting = cell.get("formatting");
        let Some(object) = formatting.obj() else {
            return;
        };
        object.delete("vMerge");
        if object.is_empty()
            && let Some(cell) = cell.obj()
        {
            cell.set("formatting", V::Undef);
        }
    };
    for (_, entries) in columns {
        let mut start: Option<V> = None;
        let mut length = 0;
        let mut last_row: isize = -1;
        let close = |start: &mut Option<V>, length: &mut usize| {
            if let Some(cell) = start.as_ref()
                && *length < 2
            {
                clear(cell);
            }
            *start = None;
            *length = 0;
        };
        for (row_index, cell) in entries {
            match cell.get("formatting").get("vMerge").as_str().as_deref() {
                Some("restart") => {
                    close(&mut start, &mut length);
                    start = Some(cell);
                    length = 1;
                    last_row = row_index as isize;
                }
                Some("continue") => {
                    if start.is_some() && row_index as isize == last_row + 1 {
                        length += 1;
                        last_row = row_index as isize;
                    } else {
                        close(&mut start, &mut length);
                        clear(&cell);
                    }
                }
                _ => {}
            }
        }
        close(&mut start, &mut length);
    }
}

/// `projectDocument`: every story projected over the base document.
pub(crate) fn project_document(context: &mut SaveContext, base: &V) -> Result<V> {
    let package = base.get("package");
    let body = if context.story_ids.contains("body") {
        V::array(context.story_to_blocks("body")?)
    } else {
        package.get("document").get("content")
    };
    let project_parts = |context: &mut SaveContext, parts: &V| -> Result<V> {
        let mut out = Vec::new();
        for entry in parts.items() {
            let pair = entry.arr().cloned().unwrap_or_default();
            let (id, part) = (pair.get(0), pair.get(1));
            let story = format!("hf:{}", id.to_js_string());
            let part = if context.story_ids.contains(&story) {
                V::Obj(
                    o(&part)
                        .unwrap_or_default()
                        .with(&[("content", V::array(context.story_to_blocks(&story)?))]),
                )
            } else {
                part
            };
            out.push(arr![id, part]);
        }
        Ok(V::array(out))
    };
    let headers = package.get("headers");
    let headers = if headers.truthy() {
        project_parts(context, &headers)?
    } else {
        headers
    };
    let footers = package.get("footers");
    let footers = if footers.truthy() {
        project_parts(context, &footers)?
    } else {
        footers
    };
    let project_notes = |context: &mut SaveContext, notes: &V, prefix: &str| -> Result<V> {
        if notes.arr().is_none() {
            return Ok(notes.clone());
        }
        let mut out = Vec::new();
        for note in notes.items() {
            let story = format!("{prefix}{}", note.get("id").to_js_string());
            if context.story_ids.contains(&story) {
                let projected = context.story_to_blocks(&story)?;
                let content = restore_note_marks(projected, &note.get("content").items());
                out.push(V::Obj(o(&note).unwrap_or_default().with(&[
                    ("content", V::array(content)),
                    ("verbatimXml", V::Undef),
                ])));
            } else {
                out.push(note);
            }
        }
        Ok(V::array(out))
    };
    let footnotes = project_notes(context, &package.get("footnotes"), "fn:")?;
    let endnotes = project_notes(context, &package.get("endnotes"), "en:")?;
    let document_body = o(&package.get("document")).unwrap_or_default().with(&[
        ("content", body),
        ("comments", V::array(context.projected_comments.clone())),
    ]);
    let projected_package = o(&package)
        .unwrap_or_default()
        .with(&[("document", V::Obj(document_body))]);
    for (key, value) in [
        ("headers", headers),
        ("footers", footers),
        ("footnotes", footnotes),
        ("endnotes", endnotes),
    ] {
        if value.truthy() {
            projected_package.set(key, value);
        }
    }
    let document = V::Obj(
        o(base)
            .unwrap_or_default()
            .with(&[("package", V::Obj(projected_package))]),
    );
    if context.hooks.on_story {
        context.visit_reachable_stories(&document);
    }
    Ok(document)
}
