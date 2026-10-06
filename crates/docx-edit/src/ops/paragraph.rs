//! Paragraph structure and properties: split, merge, attribute deltas, tab
//! stops, indent steps, style application, and paraId maintenance.
//!
//! A paragraph is not a container here. Its text is a plain run of story
//! units, and its identity and properties live on the pilcrow embed that
//! TERMINATES it — so splitting a paragraph is inserting one pilcrow, merging
//! two is removing one, and every property op writes to a pilcrow's map.
//!
//! Two consequences follow. First, splitting has to decide which half keeps
//! the original paraId: the new pilcrow takes it and terminates the FIRST
//! half, while the original pilcrow — now ending the second half — is
//! re-minted. Second, merging has to decide whose properties survive; a plain
//! merge lets the deleted mark's properties and paraId win, so the earlier
//! paragraph's identity carries over. Tracked-change resolution uses the
//! opposite rule, for the reasons its own module docs give.
//!
//! Style resolution stays outside the CRDT. Ops that apply a style take
//! host-resolved values (a [`ResolvedStyleProjection`] for a split's next
//! style) rather than reading `styles.xml`.
//!
//! Spacing, indent and tab values are authored OOXML units — twips and
//! line-spacing units — never pixels.

use std::collections::{BTreeMap, HashMap, HashSet};
use std::sync::Arc;

use yrs::types::Attrs;
use yrs::{Any, Map, MapPrelim, MapRef, Out, ReadTxn, Text, TextRef, TransactionMut};

use crate::format::Patch;
use crate::op::{OpError, OpResult, ParaBounds, Receipt, SplitReceipt, para_bounds};
use crate::ops::text::suggest_delete;
use crate::ops::{
    Chunk, ChunkKind, adjacent_paragraph_change_revision_id, adjacent_revision_id, adopt_pilcrow,
    capture_pilcrow, revision_id_in_range, snapshot_range,
};
use crate::segments::is_block_embed;
use crate::{
    DEL, EditCtx, EditingDoc, KIND_KEY, PARA_ID, PPR_CHANGE, PPR_DEL, PPR_INS, ParagraphId,
    Position, StoryRange, check_position, insertion_attrs, map_string, next_pilcrow,
    revision_value, story_ref,
};

/// The paragraph attributes a style definition owns: every pPr property the
/// seed resolves from a style but numbering ([`STYLE_NUMBERING_ATTRS`], which a
/// paragraph's own numbering overrides). Applying a style resets every one of
/// them to the style's value, or clears it when the style has none — an
/// attribute here is never left over from the previous style.
pub const STYLE_CONTROLLED_PARA_ATTRS: [&str; 27] = [
    "alignment",
    "spaceBefore",
    "spaceBeforeLines",
    "spaceAfterLines",
    "beforeAutospacing",
    "afterAutospacing",
    "spaceAfter",
    "lineSpacing",
    "lineSpacingRule",
    "indentLeft",
    "indentRight",
    "indentFirstLine",
    "hangingIndent",
    "contextualSpacing",
    "keepNext",
    "keepLines",
    "widowControl",
    "pageBreakBefore",
    "outlineLevel",
    "borders",
    "shading",
    "tabs",
    "bidi",
    "snapToGrid",
    "autoSpaceDE",
    "autoSpaceDN",
    "defaultTextFormatting",
];

/// The indents a paragraph's numbering level gives it.
const LIST_INDENT_ATTRS: [&str; 3] = ["indentLeft", "indentFirstLine", "hangingIndent"];

/// A paragraph's numbering and the list rendering the seed derives from it.
/// Applying a style replaces them where the numbering came from the old style
/// (or there is none); numbering set on the paragraph itself stays.
pub const STYLE_NUMBERING_ATTRS: [&str; 15] = [
    "numPr",
    "numPrFromStyle",
    "listNumFmt",
    "listIsBullet",
    "listMarker",
    "listMarkerHidden",
    "listMarkerFontFamily",
    "listMarkerFontSize",
    "listMarkerBold",
    "listMarkerItalic",
    "listMarkerColor",
    "listMarkerSuffix",
    "listLevelNumFmts",
    "listAbstractNumId",
    "listStartOverride",
];

const BORDERS: &str = "borders";
/// The source formatting a save compares against; it names borders too.
const ORIGINAL_FORMATTING: &str = "_originalFormatting";
/// The source paragraph's runs, which a save restores when its text is unchanged.
const ORIGINAL_RUN_BOUNDARIES: &str = "_originalRunBoundaries";
/// The properties of a mark that ends a section. A split leaves them on the
/// original mark, which still ends the section, and never copies them.
const SECTION_KEYS: [&str; 2] = ["sectPr", "sectionBreakType"];
const TABS: &str = "tabs";
const INDENT_LEFT: &str = "indentLeft";
const DEFAULT_TEXT_FORMATTING: &str = "defaultTextFormatting";

/// Default half-inch indent step in twips.
pub const INDENT_STEP_TWIPS: f64 = 720.0;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum MergeDirection {
    /// Delete THIS paragraph's mark, merging it with the following paragraph.
    Forward,
    /// Delete the PREVIOUS paragraph's mark, merging this paragraph into it.
    Backward,
}

/// Which paragraphs an op targets. Every variant is resolved and validated
/// before any mutation, so an unknown id never leaves a partial edit.
#[derive(Clone, Debug, PartialEq)]
pub enum ParaSelector {
    One(ParagraphId),
    /// Resolved in the order given; an unknown id in the list is an error.
    Many(Vec<ParagraphId>),
    /// Every paragraph whose content or mark intersects the range.
    Range(StoryRange),
}

/// One tab stop; `pos` is in twips, `alignment` is the `w:tab` val (`left`, `center`, `right`,
/// `decimal`, `bar`, `clear`), `leader` the optional leader character name. Stored in the
/// seed's shape (`position`, `alignment`, `leader`).
#[derive(Clone, Debug, PartialEq)]
pub struct TabStop {
    pub pos: f64,
    pub alignment: String,
    pub leader: Option<String>,
}

impl TabStop {
    fn to_any(&self) -> Any {
        let mut map = HashMap::from([
            ("position".into(), Any::Number(self.pos)),
            ("alignment".into(), Any::from(self.alignment.as_str())),
        ]);
        if let Some(leader) = &self.leader {
            map.insert("leader".into(), Any::from(leader.as_str()));
        }
        Any::Map(Arc::new(map))
    }
}

/// A tri-state paragraph-property delta: [`Patch::Keep`] leaves the current
/// value, [`Patch::Clear`] removes the property, [`Patch::Set`] writes it.
/// Spacing and indent values are authored OOXML units, never pixels.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct ParaAttrDelta {
    pub alignment: Patch<String>,
    pub line_spacing: Patch<f64>,
    pub line_spacing_rule: Patch<String>,
    pub space_before: Patch<f64>,
    pub space_after: Patch<f64>,
    pub indent_left: Patch<f64>,
    pub indent_right: Patch<f64>,
    pub indent_first_line: Patch<f64>,
    pub hanging_indent: Patch<bool>,
    pub bidi: Patch<bool>,
    pub tabs: Patch<Vec<TabStop>>,
    /// The paragraph-mark run defaults (`defaultTextFormatting`), as an opaque attr map.
    pub default_text_formatting: Patch<BTreeMap<String, Any>>,
    /// Every paragraph property without a typed field above, written as given;
    /// `None` clears the key and `Some(Any::Null)` writes an explicit null, which
    /// readers treat as unset but which wins over a concurrent write as a value
    /// does. Schema-managed identity keys are rejected.
    pub other: BTreeMap<String, Option<Any>>,
}

/// A split's next style (`w:next`) already resolved by the host, injected
/// because the `styles.xml` cascade lives outside the CRDT.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct ResolvedStyleProjection {
    pub style_id: String,
    /// Host-verified existence. `false` yields [`OpError::UnknownStyle`]
    /// before any mutation.
    pub known: bool,
    /// Values for the [`STYLE_CONTROLLED_PARA_ATTRS`] keys — a missing or null
    /// entry clears that attribute — plus any list attributes when the style
    /// defines numbering.
    pub paragraph_attrs: BTreeMap<String, Any>,
}

struct TargetPara {
    story_id: String,
    story: TextRef,
    bounds: ParaBounds,
    map: MapRef,
}

fn all_targets<T: ReadTxn>(txn: &T) -> Vec<TargetPara> {
    let Some(stories) = txn.get_map(crate::STORIES) else {
        return Vec::new();
    };
    let mut story_ids: Vec<String> = stories.keys(txn).map(|key| key.to_string()).collect();
    story_ids.sort();
    let mut result = Vec::new();
    for story_id in story_ids {
        let Some(Out::YText(story)) = stories.get(txn, &story_id) else {
            continue;
        };
        let pilcrow_maps: HashMap<u32, MapRef> = crate::pilcrows(&story, txn).into_iter().collect();
        for bounds in para_bounds(&story, txn) {
            if let Some(map) = pilcrow_maps.get(&bounds.pilcrow) {
                result.push(TargetPara {
                    story_id: story_id.clone(),
                    story: story.clone(),
                    bounds,
                    map: map.clone(),
                });
            }
        }
    }
    result
}

/// Moves the bookmarks of paragraph mark `from` onto `to`, at the start of its
/// content and ahead of `breaks` leading breaks, skipping those `to` holds.
fn move_bookmarks(txn: &mut TransactionMut, from: &MapRef, to: &MapRef, breaks: usize) {
    let Some(Out::Any(Any::Array(moving))) = from.get(txn, "bookmarks") else {
        return;
    };
    let kept = match to.get(txn, "bookmarks") {
        Some(Out::Any(Any::Array(list))) => list.to_vec(),
        _ => Vec::new(),
    };
    let key = |bookmark: &Any| match bookmark {
        Any::Map(map) => Some((map.get("id").cloned(), map.get("kind").cloned())),
        _ => None,
    };
    let mut bookmarks: Vec<Any> = moving
        .iter()
        .filter(|bookmark| !kept.iter().any(|held| key(held) == key(bookmark)))
        .filter_map(|bookmark| {
            let Any::Map(map) = bookmark else {
                return None;
            };
            let mut map = (**map).clone();
            map.insert("offset".to_owned(), Any::from(0_i64));
            map.remove("breaksAfter");
            if breaks > 0 {
                map.insert("breaksAfter".to_owned(), Any::from(breaks as i64));
            }
            Some(Any::Map(Arc::new(map)))
        })
        .collect();
    if bookmarks.is_empty() {
        return;
    }
    bookmarks.extend(kept);
    to.insert(txn, "bookmarks", Any::Array(bookmarks.into()));
}

/// Resolves a selector to pilcrow targets, validating BEFORE any mutation.
fn resolve_selector<T: ReadTxn>(txn: &T, selector: &ParaSelector) -> OpResult<Vec<TargetPara>> {
    let all = all_targets(txn);
    match selector {
        ParaSelector::One(id) => {
            let target = all
                .into_iter()
                .find(|target| target.bounds.para_id == *id)
                .ok_or_else(|| OpError::UnknownPara(id.clone()))?;
            Ok(vec![target])
        }
        ParaSelector::Many(ids) => {
            let mut by_id: HashMap<String, TargetPara> = all
                .into_iter()
                .map(|target| (target.bounds.para_id.clone(), target))
                .collect();
            ids.iter()
                .map(|id| {
                    by_id
                        .remove(id)
                        .ok_or_else(|| OpError::UnknownPara(id.clone()))
                })
                .collect()
        }
        ParaSelector::Range(range) => {
            if range.end < range.start {
                return Err(OpError::InvalidRange {
                    start: range.start,
                    end: range.end,
                });
            }
            let targets: Vec<TargetPara> = all
                .into_iter()
                .filter(|target| {
                    target.story_id == range.story
                        && target.bounds.start <= range.end
                        && target.bounds.pilcrow >= range.start
                })
                .collect();
            if targets.is_empty() {
                return Err(OpError::UnknownStory(range.story.clone()));
            }
            Ok(targets)
        }
    }
}

fn set_or_remove(txn: &mut TransactionMut<'_>, map: &MapRef, key: &str, value: Option<Any>) {
    match value {
        Some(value) if value != Any::Null => {
            map.insert(txn, key.to_owned(), value);
        }
        _ => {
            map.remove(txn, key);
        }
    }
}

/// Writes a style's paragraph-attribute projection: each
/// [`STYLE_CONTROLLED_PARA_ATTRS`] key is reset to the projection's value or
/// cleared when it has none, and any extra key (list attributes) is applied as
/// given. Errors on a schema-managed identity key.
fn apply_paragraph_attr_projection(
    txn: &mut TransactionMut<'_>,
    map: &MapRef,
    attrs: &BTreeMap<String, Any>,
) -> OpResult<()> {
    for key in STYLE_CONTROLLED_PARA_ATTRS {
        set_or_remove(txn, map, key, attrs.get(key).cloned());
    }
    for (key, value) in attrs {
        if STYLE_CONTROLLED_PARA_ATTRS.contains(&key.as_str()) {
            continue;
        }
        if matches!(key.as_str(), PARA_ID | KIND_KEY) {
            return Err(OpError::ReservedKey(key.clone()));
        }
        set_or_remove(txn, map, key, Some(value.clone()));
    }
    if let Some(Out::Any(Any::Map(original))) = map.get(txn, "_originalFormatting") {
        let mut original = (*original).clone();
        for key in [
            "spaceBefore",
            "spaceAfter",
            "spaceBeforeLines",
            "spaceAfterLines",
            "beforeAutospacing",
            "afterAutospacing",
        ] {
            match attrs.get(key) {
                Some(value) if *value != Any::Null => {
                    original.insert(key.to_owned(), value.clone());
                }
                _ => {
                    original.remove(key);
                }
            }
        }
        map.insert(txn, "_originalFormatting", Any::Map(Arc::new(original)));
    }
    Ok(())
}

/// `pPrChange` records under new revision ids.
fn with_fresh_revision_ids(changes: &Any, mut next_id: impl FnMut() -> String) -> Any {
    let Any::Array(changes) = changes else {
        return changes.clone();
    };
    Any::Array(
        changes
            .iter()
            .map(|change| {
                let Any::Map(change) = change else {
                    return change.clone();
                };
                let mut change = (**change).clone();
                if let Some(Any::Map(info)) = change.get("info") {
                    let mut info = (**info).clone();
                    info.remove("revisionId");
                    info.insert("id".to_owned(), Any::from(next_id()));
                    change.insert("info".to_owned(), Any::Map(Arc::new(info)));
                }
                Any::Map(Arc::new(change))
            })
            .collect(),
    )
}

/// `_originalFormatting` without its borders.
fn original_without_borders(value: &Any) -> Option<Any> {
    match value {
        Any::Map(original) if original.contains_key(BORDERS) => {
            let mut original = (**original).clone();
            original.remove(BORDERS);
            Some(Any::Map(Arc::new(original)))
        }
        _ => None,
    }
}

/// Drops a mark's borders, in its properties and its source formatting alike,
/// so a save does not write them back from the source.
fn remove_borders(txn: &mut TransactionMut<'_>, map: &MapRef) {
    map.remove(txn, BORDERS);
    if let Some(Out::Any(original)) = map.get(txn, ORIGINAL_FORMATTING)
        && let Some(original) = original_without_borders(&original)
    {
        map.insert(txn, ORIGINAL_FORMATTING, original);
    }
}

/// Whether `chunk` is the reference field of a comment, which shows nothing.
fn is_comment_reference<T: ReadTxn>(chunk: &Chunk, txn: &T) -> bool {
    matches!(&chunk.kind, ChunkKind::Embed(Some(map))
        if map_string(map, txn, "modelKind").as_deref() == Some("commentReference"))
}

impl EditingDoc {
    /// Splits a paragraph by inserting exactly ONE pilcrow at `at`.
    ///
    /// The new pilcrow terminates the FIRST half, carrying the source
    /// paragraph's properties (but the section it ends) and its ORIGINAL
    /// paraId. Where the source mark keeps its own properties (a split
    /// mid-paragraph or before a block), the new mark leaves out the source
    /// mark's tracked insertion or deletion, and its copy of a tracked
    /// property change takes new revision ids. The original
    /// pilcrow is re-minted with a fresh paraId and becomes the second half's
    /// mark. What the second half then keeps depends on where the split fell:
    ///
    /// - mid-paragraph: it keeps its own properties but borders;
    /// - at the paragraph end, so the second half is empty: it keeps all of
    ///   them, as Word copies the paragraph mark, but the mark's tracked
    ///   insertion or deletion and the source's run cache, which stay with
    ///   the text, and its copy of a tracked property change takes new
    ///   revision ids;
    /// - at the end WITH a `next_style`: it switches to that style's
    ///   projection outright instead, without borders.
    ///
    /// A section the paragraph ends stays
    /// with the second half's mark, which still ends it: the new mark never
    /// takes `sectPr` or `sectionBreakType`. Suggesting mode stamps the inserted
    /// pilcrow `ins` and `pPrIns`, reusing an adjacent revision by the same
    /// author when there is one.
    ///
    /// At the start of a slot that opens with a table, block content control
    /// or break, the split inserts an empty paragraph before the block
    /// instead, as Word does: the new mark takes the fresh paraId and the
    /// properties but borders, and the block's paragraph keeps its own, so
    /// removing the new paragraph restores the document.
    ///
    /// Errors when `next_style` is not known, before any mutation, and when
    /// `at` does not address a position inside a story.
    pub fn split_paragraph(
        &self,
        ctx: &EditCtx,
        mut at: Position,
        next_style: Option<&ResolvedStyleProjection>,
    ) -> OpResult<SplitReceipt> {
        if let Some(projection) = next_style
            && !projection.known
        {
            return Err(OpError::UnknownStyle(projection.style_id.clone()));
        }
        let second_para_id = self.next_id();
        let slot = self
            .segment_index(&at.story)?
            .para_at(at.index)
            .filter(|para| para.start == at.index && para.node_start > at.index)
            .map(|para| (para.node_start, para.pilcrow));
        let mut before_block = slot.is_some();
        let mut txn = self.transact_for(ctx);
        let story = story_ref(&txn, &at.story)?;
        check_position(&story, &txn, at.index)?;
        if let Some((text_start, end)) = slot
            && at.index > 0
            && text_start < end
        {
            let prefix = snapshot_range(&story, &txn, at.index, text_start);
            if prefix.iter().all(|chunk| {
                matches!(&chunk.kind, ChunkKind::Embed(Some(map))
                if map_string(map, &txn, KIND_KEY).as_deref() == Some("columnBreak")
                    && matches!(map.get(&txn, "trailing"), Some(Out::Any(Any::Bool(true)))))
            }) {
                at.index = text_start;
                before_block = false;
            }
        }
        let chunks = snapshot_range(
            &story,
            &txn,
            at.index.saturating_sub(1),
            at.index.saturating_add(1),
        );
        let revision_id = ctx.is_suggesting().then(|| {
            adjacent_revision_id(&chunks, at.index, crate::INS, &ctx.author)
                .or_else(|| {
                    adjacent_paragraph_change_revision_id(&chunks, at.index, &txn, &ctx.author)
                })
                .unwrap_or_else(|| self.next_id())
        });
        let (orig_index, orig_map) =
            next_pilcrow(&story, &txn, at.index).ok_or(OpError::ExpectedPilcrow {
                story: at.story.clone(),
                index: at.index,
            })?;
        let (first_para_id, props) = capture_pilcrow(&orig_map, &txn);
        let second_half_empty = orig_index == at.index;

        let ins = revision_id
            .as_ref()
            .map(|id| revision_value(id, &ctx.revision_author()));
        let new_pilcrow = story.insert_embed_with_attributes(
            &mut txn,
            at.index,
            MapPrelim::default(),
            insertion_attrs(ins, None),
        );
        new_pilcrow.insert(&mut txn, KIND_KEY, crate::PILCROW_KIND);
        let new_para_id = if before_block {
            &second_para_id
        } else {
            &first_para_id
        };
        new_pilcrow.insert(&mut txn, PARA_ID, new_para_id.as_str());
        // Where the source mark keeps its properties, the new mark is a plain
        // one, as in Word: the source mark's own insertion or deletion stays
        // on it, and the new mark's copy of a tracked property change is a
        // change of its own.
        let source_keeps = before_block || !second_half_empty;
        for (key, value) in &props {
            if SECTION_KEYS.contains(&key.as_str())
                || (before_block && key == BORDERS)
                || (source_keeps && (key == PPR_INS || key == PPR_DEL))
            {
                continue;
            }
            let value = if before_block && key == ORIGINAL_FORMATTING {
                original_without_borders(value).unwrap_or_else(|| value.clone())
            } else if source_keeps && key == PPR_CHANGE {
                with_fresh_revision_ids(value, || self.next_id())
            } else {
                value.clone()
            };
            new_pilcrow.insert(&mut txn, key.clone(), value);
        }
        if let Some(id) = revision_id.as_ref() {
            new_pilcrow.insert(
                &mut txn,
                PPR_INS,
                revision_value(id, &ctx.revision_author()),
            );
        }
        if before_block {
            return Ok(SplitReceipt {
                first_para_id: second_para_id,
                second_para_id: first_para_id,
                revision_ids: revision_id.into_iter().collect(),
            });
        }

        // The original pilcrow now terminates the second half: re-mint its identity, then apply
        // post-split inheritance.
        orig_map.insert(&mut txn, PARA_ID, second_para_id.as_str());
        if second_half_empty {
            if let Some(next) = next_style {
                // A `w:next` switch starts from nothing: drop the source
                // properties before writing the projection.
                for (key, _) in &props {
                    if !SECTION_KEYS.contains(&key.as_str()) {
                        orig_map.remove(&mut txn, key);
                    }
                }
                orig_map.insert(&mut txn, "pStyle", next.style_id.as_str());
                apply_paragraph_attr_projection(&mut txn, &orig_map, &next.paragraph_attrs)?;
                orig_map.remove(&mut txn, BORDERS);
            } else {
                // The new paragraph is a copy, as in Word: the mark revision went
                // to the text's mark with the run cache, and its property change
                // becomes one of its own.
                for key in [PPR_INS, PPR_DEL, ORIGINAL_RUN_BOUNDARIES] {
                    orig_map.remove(&mut txn, key);
                }
                if let Some((_, changes)) = props.iter().find(|(key, _)| key == PPR_CHANGE) {
                    let changes = with_fresh_revision_ids(changes, || self.next_id());
                    orig_map.insert(&mut txn, PPR_CHANGE, changes);
                }
            }
        } else {
            // A mid-paragraph split keeps the second half's pPr but its borders.
            remove_borders(&mut txn, &orig_map);
        }
        if !ctx.is_suggesting() {
            crate::ops::field_changes::split_field(
                &mut txn,
                &story,
                &at.story,
                at.index,
                &format!("field:{second_para_id}"),
                self.package().as_deref(),
            )?;
        }
        Ok(SplitReceipt {
            first_para_id,
            second_para_id,
            revision_ids: revision_id.into_iter().collect(),
        })
    }

    /// Merges the paragraph with its neighbour by deleting the boundary
    /// pilcrow. The survivor adopts the deleted mark's properties and paraId,
    /// so the EARLIER paragraph's identity wins.
    ///
    /// Suggesting mode normally retains the mark, stamping it `del` and
    /// `pPrDel`. The exception is backspacing over this same author's still
    /// pending split: that retracts the suggestion, physically removing the
    /// pilcrow rather than authoring a second, contradictory revision.
    ///
    /// A survivor that opens with a table, block content control or page or
    /// column break is not merged into, so text never goes ahead of such a
    /// block in one paragraph slot. As in Word the op removes a leading page
    /// or column break instead, and before a table or content control it
    /// removes the paragraph when that is empty (nothing but its mark and
    /// comment reference fields, which show nothing; the survivor keeps its
    /// own properties) and changes nothing otherwise. The paragraph between
    /// two tables is never empty here: it belongs to the first table's slot,
    /// so the tables are never joined. Suggesting mode marks what it removes
    /// deleted, but for this author's own pending paragraph mark, which goes.
    ///
    /// The receipt's range is the caret position after the merge: where the
    /// caret was, unless the paragraph it was in is gone. Errors when
    /// the paragraph is unknown, when merging forward from a story's last
    /// paragraph, and when merging backward from its first.
    pub fn merge_paragraphs(
        &self,
        ctx: &EditCtx,
        para: &str,
        direction: MergeDirection,
    ) -> OpResult<Receipt> {
        let mut txn = self.transact_for(ctx);
        let targets = all_targets(&txn);
        let index = targets
            .iter()
            .position(|target| target.bounds.para_id == para)
            .ok_or_else(|| OpError::UnknownPara(para.to_owned()))?;
        let boundary_index = match direction {
            MergeDirection::Forward => {
                let is_last_in_story = targets
                    .get(index + 1)
                    .is_none_or(|next| next.story_id != targets[index].story_id);
                if is_last_in_story {
                    return Err(OpError::CannotMergeFinalParagraph(para.to_owned()));
                }
                index
            }
            MergeDirection::Backward => {
                let has_previous =
                    index > 0 && targets[index - 1].story_id == targets[index].story_id;
                if !has_previous {
                    return Err(OpError::NoParagraphBefore(para.to_owned()));
                }
                index - 1
            }
        };
        let boundary = &targets[boundary_index];
        let survivor = &targets[boundary_index + 1];
        let story = boundary.story.clone();
        let pilcrow_index = boundary.bounds.pilcrow;
        let own_insert = ctx
            .is_suggesting()
            .then(|| paragraph_revision_id(&boundary.map, &txn, PPR_INS, &ctx.author))
            .flatten();
        let lead = snapshot_range(&story, &txn, pilcrow_index + 1, pilcrow_index + 2)
            .first()
            .and_then(|chunk| match &chunk.kind {
                ChunkKind::Embed(Some(map)) => {
                    map_string(map, &txn, KIND_KEY).filter(|kind| is_block_embed(kind))
                }
                _ => None,
            });
        if let Some(lead) = lead {
            // Text never goes ahead of a block in one slot, so as in Word the
            // merge takes out what separates the two paragraphs instead.
            let start = boundary.bounds.start;
            let chunks = snapshot_range(&story, &txn, start, pilcrow_index + 2);
            let stays = match direction {
                MergeDirection::Forward => pilcrow_index,
                MergeDirection::Backward => pilcrow_index + 1,
            };
            // A paragraph whose mark ends a section is never empty here: the
            // section break stays.
            let empty = chunks
                .iter()
                .filter(|chunk| chunk.start < pilcrow_index)
                .all(|chunk| is_comment_reference(chunk, &txn))
                && !SECTION_KEYS
                    .iter()
                    .any(|key| boundary.map.contains_key(&txn, key));
            let (from, to) = match lead.as_str() {
                // Before a break it removes the break, unless the paragraph
                // before is empty: that paragraph goes, as before a table, so
                // Enter at the slot then Delete or Backspace restores.
                "pageBreak" | "columnBreak" if !empty => (pilcrow_index + 1, pilcrow_index + 2),
                // An empty paragraph before a table or content control goes;
                // one with content stays, and so do the paragraph and table
                // between two tables, which are never joined.
                _ if empty => (start, pilcrow_index + 1),
                _ => (pilcrow_index, pilcrow_index),
            };
            // A removed empty paragraph's bookmarks move to the paragraph that
            // stays, ahead of its slot's breaks (a suggested deletion keeps them).
            // At a story's start, a survivor holding only breaks that end in a
            // column break takes them after its breaks, where its save writes
            // them, so the file does not depend on a publication in between.
            if from == start && from < to && (!ctx.is_suggesting() || own_insert.is_some()) {
                let kinds: Vec<Option<String>> =
                    snapshot_range(&story, &txn, survivor.bounds.start, survivor.bounds.pilcrow)
                        .iter()
                        .map(|chunk| match &chunk.kind {
                            ChunkKind::Embed(Some(map)) => map_string(map, &txn, KIND_KEY),
                            _ => None,
                        })
                        .collect();
                let breaks = kinds
                    .iter()
                    .map_while(|kind| kind.as_deref().filter(|kind| is_block_embed(kind)))
                    .filter(|kind| kind.ends_with("Break"))
                    .count();
                let only_breaks = breaks == kinds.len()
                    && kinds
                        .last()
                        .is_some_and(|kind| kind.as_deref() == Some("columnBreak"));
                let ahead = if start == 0 && only_breaks { 0 } else { breaks };
                move_bookmarks(&mut txn, &boundary.map, &survivor.map, ahead);
                let moved = crate::bookmarks::move_range(
                    &mut txn,
                    &story,
                    &boundary.story_id,
                    from,
                    to,
                    survivor.bounds.start + (breaks - ahead) as u32,
                )?;
                // The save writes them ahead of the survivor's fields.
                crate::ops::field_changes::renumber_fields(
                    &mut txn,
                    &story,
                    survivor.bounds.start,
                    survivor.bounds.pilcrow,
                    moved as i64,
                );
            }
            let mut revision_id = None;
            // Units removed ahead of the paragraph mark: the caret shifts by them.
            let mut removed = 0;
            if from < to && !ctx.is_suggesting() {
                story.remove_range(&mut txn, from, to - from);
                removed = pilcrow_index.saturating_sub(from);
            } else if from < to {
                let id = adjacent_revision_id(&chunks, from, DEL, &ctx.author)
                    .unwrap_or_else(|| self.next_id());
                let revision = revision_value(&id, &ctx.revision_author());
                let mut end = to;
                if own_insert.is_some() && to == pilcrow_index + 1 {
                    // Backspacing over this author's pending paragraph retracts it.
                    story.remove_range(&mut txn, pilcrow_index, 1);
                    end = pilcrow_index;
                }
                let outcome = suggest_delete(&mut txn, &story, ctx, &revision, from, end, &chunks);
                if from < pilcrow_index {
                    removed = outcome.removed;
                }
                revision_id = (outcome.removed < end - from).then_some(id);
            }
            let caret = if from == start && from < to {
                pilcrow_index - removed
            } else {
                stays
            };
            let caret =
                crate::op::loc_range_in_txn(&boundary.story_id, &story, &txn, caret, caret)?;
            return Ok(Receipt {
                new_para_ids: Vec::new(),
                revision_ids: revision_id.into_iter().collect(),
                range: Some(caret),
            });
        }
        let revision_id = (ctx.is_suggesting() && own_insert.is_none()).then(|| {
            let chunks = snapshot_range(
                &story,
                &txn,
                pilcrow_index.saturating_sub(1),
                pilcrow_index.saturating_add(1),
            );
            adjacent_revision_id(&chunks, pilcrow_index, DEL, &ctx.author)
                .unwrap_or_else(|| self.next_id())
        });

        if own_insert.is_some() {
            // Backspacing over this author's still-pending split retracts the
            // suggestion itself; it must not author a second pPrDel revision.
            let (donor_id, mut donor_props) = capture_pilcrow(&boundary.map, &txn);
            donor_props.retain(|(key, _)| !matches!(key.as_str(), PPR_INS | PPR_DEL));
            story.remove_range(&mut txn, pilcrow_index, 1);
            adopt_pilcrow(&mut txn, &survivor.map, &donor_id, &donor_props);
        } else if let Some(id) = revision_id.as_ref() {
            let revision = revision_value(id, &ctx.revision_author());
            story.format(
                &mut txn,
                pilcrow_index,
                1,
                Attrs::from([(Arc::from(DEL), revision.clone())]),
            );
            boundary.map.insert(&mut txn, PPR_DEL, revision);
        } else {
            let (donor_id, donor_props) = capture_pilcrow(&boundary.map, &txn);
            story.remove_range(&mut txn, pilcrow_index, 1);
            adopt_pilcrow(&mut txn, &survivor.map, &donor_id, &donor_props);
            crate::ops::field_changes::rejoin_fields(
                &mut txn,
                &story,
                &boundary.story_id,
                pilcrow_index,
            )?;
        }
        let caret = crate::op::loc_range_in_txn(
            &boundary.story_id,
            &story,
            &txn,
            pilcrow_index,
            pilcrow_index,
        )?;
        Ok(Receipt {
            new_para_ids: Vec::new(),
            revision_ids: own_insert.into_iter().chain(revision_id).collect(),
            range: Some(caret),
        })
    }

    /// Applies a tri-state paragraph attribute delta to the selected
    /// paragraphs in one transaction. In suggesting mode a `pPrChange` record
    /// capturing the before and after property maps is appended to every
    /// paragraph the delta actually changed, and the receipt carries its
    /// revision id; a delta that changes nothing stamps nothing. Errors when
    /// the delta names a schema-managed identity key, before any mutation.
    pub fn set_paragraph_attrs(
        &self,
        ctx: &EditCtx,
        selector: &ParaSelector,
        delta: &ParaAttrDelta,
    ) -> OpResult<Receipt> {
        self.set_paragraph_attrs_each(ctx, selector, |_, _| Ok(delta.clone()))
    }

    /// [`Self::set_paragraph_attrs`] with a delta per paragraph, read from its
    /// properties before the change.
    fn set_paragraph_attrs_each(
        &self,
        ctx: &EditCtx,
        selector: &ParaSelector,
        delta_for: impl Fn(&MapRef, &TransactionMut<'_>) -> OpResult<ParaAttrDelta>,
    ) -> OpResult<Receipt> {
        let mut txn = self.transact_for(ctx);
        let targets = resolve_selector(&txn, selector)?;
        let deltas = targets
            .iter()
            .map(|target| delta_for(&target.map, &txn))
            .collect::<OpResult<Vec<_>>>()?;
        for delta in &deltas {
            for key in delta.other.keys() {
                if matches!(key.as_str(), PARA_ID | KIND_KEY) {
                    return Err(OpError::ReservedKey(key.clone()));
                }
            }
        }
        let revision_id = ctx.is_suggesting().then(|| {
            targets
                .iter()
                .find_map(|target| {
                    revision_id_in_range(
                        &snapshot_range(
                            &target.story,
                            &txn,
                            target.bounds.start,
                            target.bounds.pilcrow + 1,
                        ),
                        target.bounds.start,
                        target.bounds.pilcrow + 1,
                        crate::INS,
                        &ctx.author,
                    )
                })
                .unwrap_or_else(|| self.next_id())
        });
        let mut changed = false;
        for (target, delta) in targets.iter().zip(&deltas) {
            let previous = paragraph_formatting(&target.map, &txn);
            apply_para_delta(&mut txn, &target.map, delta);
            let current = paragraph_formatting(&target.map, &txn);
            if let Some(id) = revision_id.as_ref()
                && previous != current
            {
                append_paragraph_property_change(
                    &mut txn,
                    &target.map,
                    id,
                    &ctx.revision_author(),
                    previous,
                    current,
                );
                changed = true;
            }
        }
        Ok(Receipt {
            revision_ids: revision_id.filter(|_| changed).into_iter().collect(),
            ..Receipt::default()
        })
    }

    /// Adds a tab stop to the selected paragraphs, replacing any existing stop
    /// at the same position. Stops stay sorted by position.
    pub fn add_tab_stop(
        &self,
        ctx: &EditCtx,
        selector: &ParaSelector,
        stop: &TabStop,
    ) -> OpResult<Receipt> {
        let mut txn = self.transact_for(ctx);
        let targets = resolve_selector(&txn, selector)?;
        for target in &targets {
            let mut stops = read_tab_stops(&target.map, &txn);
            stops.retain(|existing| existing_pos(existing) != Some(stop.pos));
            stops.push(stop.to_any());
            stops.sort_by(|a, b| {
                existing_pos(a)
                    .unwrap_or(f64::MAX)
                    .total_cmp(&existing_pos(b).unwrap_or(f64::MAX))
            });
            target
                .map
                .insert(&mut txn, TABS, Any::Array(Arc::from(stops)));
        }
        Ok(Receipt::default())
    }

    /// Removes the tab stop at `pos` twips from the selected paragraphs,
    /// dropping the `tabs` attribute entirely once none are left.
    pub fn remove_tab_stop(
        &self,
        ctx: &EditCtx,
        selector: &ParaSelector,
        pos: f64,
    ) -> OpResult<Receipt> {
        let mut txn = self.transact_for(ctx);
        let targets = resolve_selector(&txn, selector)?;
        for target in &targets {
            let mut stops = read_tab_stops(&target.map, &txn);
            stops.retain(|existing| existing_pos(existing) != Some(pos));
            if stops.is_empty() {
                target.map.remove(&mut txn, TABS);
            } else {
                target
                    .map
                    .insert(&mut txn, TABS, Any::Array(Arc::from(stops)));
            }
        }
        Ok(Receipt::default())
    }

    /// Increases `indentLeft` by `step` twips, defaulting to
    /// [`INDENT_STEP_TWIPS`].
    pub fn increase_indent(
        &self,
        ctx: &EditCtx,
        selector: &ParaSelector,
        step: Option<f64>,
    ) -> OpResult<Receipt> {
        let step = step.unwrap_or(INDENT_STEP_TWIPS);
        let mut txn = self.transact_for(ctx);
        let targets = resolve_selector(&txn, selector)?;
        for target in &targets {
            let current = number_prop(&target.map, &txn, INDENT_LEFT).unwrap_or(0.0);
            target
                .map
                .insert(&mut txn, INDENT_LEFT, Any::Number(current + step));
        }
        Ok(Receipt::default())
    }

    /// Decreases `indentLeft` by `step` twips, defaulting to
    /// [`INDENT_STEP_TWIPS`] and clamping at zero. Reaching zero removes the
    /// attribute rather than storing it.
    pub fn decrease_indent(
        &self,
        ctx: &EditCtx,
        selector: &ParaSelector,
        step: Option<f64>,
    ) -> OpResult<Receipt> {
        let step = step.unwrap_or(INDENT_STEP_TWIPS);
        let mut txn = self.transact_for(ctx);
        let targets = resolve_selector(&txn, selector)?;
        for target in &targets {
            let current = number_prop(&target.map, &txn, INDENT_LEFT).unwrap_or(0.0);
            let next = (current - step).max(0.0);
            if next > 0.0 {
                target.map.insert(&mut txn, INDENT_LEFT, Any::Number(next));
            } else {
                target.map.remove(&mut txn, INDENT_LEFT);
            }
        }
        Ok(Receipt::default())
    }

    /// Sets (or clears) the paragraph-mark run defaults (`defaultTextFormatting`).
    pub fn set_paragraph_default_format(
        &self,
        ctx: &EditCtx,
        selector: &ParaSelector,
        formatting: Option<&BTreeMap<String, Any>>,
    ) -> OpResult<Receipt> {
        let mut txn = self.transact_for(ctx);
        let targets = resolve_selector(&txn, selector)?;
        for target in &targets {
            match formatting {
                Some(map) if !map.is_empty() => {
                    let value: HashMap<String, Any> =
                        map.iter().map(|(k, v)| (k.clone(), v.clone())).collect();
                    target
                        .map
                        .insert(&mut txn, DEFAULT_TEXT_FORMATTING, Any::Map(Arc::new(value)));
                }
                _ => {
                    target.map.remove(&mut txn, DEFAULT_TEXT_FORMATTING);
                }
            }
        }
        Ok(Receipt::default())
    }

    /// Applies a host-resolved paragraph style to the selected paragraphs in
    /// one transaction, keeping direct paragraph formatting as Word does:
    /// writes `style_id`, and sets each [`STYLE_CONTROLLED_PARA_ATTRS`] key the
    /// paragraph holds nothing for, or holds as its old style gives it
    /// (`previous`, the host's values per style id, `""` for none) and not as
    /// its own source pPr set it, to its
    /// entry in `values` (the new style's), or to an explicit null where
    /// `values` has none, so two peers applying different styles converge on
    /// one style's values. The paragraph mark's run defaults become the new
    /// style's with the mark's own run properties over them. Where the
    /// paragraph's numbering came from its old style, or it has none, the
    /// [`STYLE_NUMBERING_ATTRS`] are taken from `values` the same way. Run
    /// formatting is the host's to apply. In suggesting mode the change is
    /// recorded as for [`Self::set_paragraph_attrs`]. Errors before any
    /// mutation when `previous` lacks a selected paragraph's style.
    pub fn apply_paragraph_style(
        &self,
        ctx: &EditCtx,
        selector: &ParaSelector,
        style_id: &str,
        values: &BTreeMap<String, Any>,
        previous: &BTreeMap<String, BTreeMap<String, Any>>,
    ) -> OpResult<Receipt> {
        let value = |key: &str| Some(values.get(key).cloned().unwrap_or(Any::Null));
        self.set_paragraph_attrs_each(ctx, selector, |map, txn| {
            let property = |key: &str| match map.get(txn, key) {
                Some(Out::Any(Any::Null)) | None => None,
                Some(Out::Any(value)) => Some(value),
                Some(_) => None,
            };
            let old_style = match property("pStyle") {
                Some(Any::String(style)) => style.to_string(),
                _ => String::new(),
            };
            let old = previous
                .get(&old_style)
                .ok_or_else(|| OpError::UnknownStyle(old_style.clone()))?;
            let source = property(ORIGINAL_FORMATTING);
            let source = match &source {
                Some(Any::Map(original)) => Some(original.as_ref()),
                _ => None,
            };
            // The seed keeps a style's numbering only in the source formatting;
            // an earlier style application writes `numPrFromStyle` itself.
            let num_pr = property("numPr");
            let numbering_from_style = num_pr.is_none()
                || num_pr == property("numPrFromStyle")
                || num_pr.as_ref() == source.and_then(|original| original.get("numPrFromStyle"));
            let mut delta = ParaAttrDelta::default();
            delta
                .other
                .insert("pStyle".to_owned(), Some(Any::from(style_id)));
            for key in STYLE_CONTROLLED_PARA_ATTRS {
                // A paragraph's own numbering keeps the indents its level gives.
                if key == DEFAULT_TEXT_FORMATTING
                    || (!numbering_from_style && LIST_INDENT_ATTRS.contains(&key))
                {
                    continue;
                }
                // A value the paragraph's own pPr set and still holds is direct, even
                // where the old style gives the same.
                let from_old_style = match (property(key), old.get(key)) {
                    (None, _) => true,
                    (Some(current), Some(old)) => {
                        same_value(&current, old)
                            && !source
                                .and_then(|original| original.get(key))
                                .is_some_and(|direct| same_value(direct, &current))
                    }
                    (Some(_), None) => false,
                };
                if from_old_style {
                    delta.other.insert(key.to_owned(), value(key));
                }
            }
            delta.other.insert(
                DEFAULT_TEXT_FORMATTING.to_owned(),
                Some(mark_defaults(
                    values.get(DEFAULT_TEXT_FORMATTING),
                    source.and_then(|original| original.get("runProperties")),
                )),
            );
            if numbering_from_style {
                for key in STYLE_NUMBERING_ATTRS {
                    delta.other.insert(key.to_owned(), value(key));
                }
            }
            Ok(delta)
        })
    }

    /// Restores paraId uniqueness after a merge of divergent replicas: every
    /// duplicate is re-minted, with the first occurrence in document order
    /// keeping its id. Runs under a system origin so the pass never enters
    /// undo history. Returns the `(old, new)` pairs.
    pub fn dedupe_para_ids(&self, now_iso: &str) -> OpResult<Vec<(ParagraphId, ParagraphId)>> {
        let ctx = EditCtx::system(now_iso);
        let mut renames = Vec::new();
        let mut txn = self.transact_for(&ctx);
        let targets = all_targets(&txn);
        let mut seen: HashSet<String> = HashSet::new();
        for target in targets {
            let id = target.bounds.para_id.clone();
            if seen.insert(id.clone()) {
                continue;
            }
            let minted = self.next_id();
            target.map.insert(&mut txn, PARA_ID, minted.as_str());
            renames.push((id, minted));
        }
        Ok(renames)
    }
}

fn paragraph_revision_id<T: ReadTxn>(
    map: &MapRef,
    txn: &T,
    key: &str,
    author: &str,
) -> Option<String> {
    let Some(Out::Any(Any::Map(revision))) = map.get(txn, key) else {
        return None;
    };
    if !matches!(revision.get("author"), Some(Any::String(value)) if value.as_ref() == author) {
        return None;
    }
    match revision.get("id").or_else(|| revision.get("revisionId")) {
        Some(Any::String(id)) => Some(id.to_string()),
        Some(Any::Number(id)) if id.is_finite() => Some(id.to_string()),
        Some(Any::BigInt(id)) => Some(id.to_string()),
        _ => None,
    }
}

fn apply_para_delta(txn: &mut TransactionMut<'_>, map: &MapRef, delta: &ParaAttrDelta) {
    fn apply<T, F: Fn(&T) -> Any>(
        txn: &mut TransactionMut<'_>,
        map: &MapRef,
        key: &str,
        patch: &Patch<T>,
        lower: F,
    ) {
        match patch {
            Patch::Keep => {}
            Patch::Clear => {
                map.remove(txn, key);
            }
            Patch::Set(value) => {
                map.insert(txn, key.to_owned(), lower(value));
            }
        }
    }
    apply(txn, map, "alignment", &delta.alignment, |v| {
        Any::from(v.as_str())
    });
    apply(txn, map, "lineSpacing", &delta.line_spacing, |v| {
        Any::Number(*v)
    });
    apply(txn, map, "lineSpacingRule", &delta.line_spacing_rule, |v| {
        Any::from(v.as_str())
    });
    apply(txn, map, "spaceBefore", &delta.space_before, |v| {
        Any::Number(*v)
    });
    apply(txn, map, "spaceAfter", &delta.space_after, |v| {
        Any::Number(*v)
    });
    for (patch, key, lines_key, auto_key) in [
        (
            &delta.space_before,
            "spaceBefore",
            "spaceBeforeLines",
            "beforeAutospacing",
        ),
        (
            &delta.space_after,
            "spaceAfter",
            "spaceAfterLines",
            "afterAutospacing",
        ),
    ] {
        match patch {
            Patch::Keep => continue,
            Patch::Clear => {
                map.remove(txn, lines_key);
                map.remove(txn, auto_key);
            }
            Patch::Set(_) => {
                map.insert(txn, lines_key, Any::Number(0.0));
                map.insert(txn, auto_key, Any::Bool(false));
            }
        }
        if let Some(Out::Any(Any::Map(original))) = map.get(txn, "_originalFormatting") {
            let mut original = (*original).clone();
            for key in [key, lines_key, auto_key] {
                match map.get(txn, key) {
                    Some(Out::Any(value)) => {
                        original.insert(key.to_owned(), value);
                    }
                    _ => {
                        original.remove(key);
                    }
                }
            }
            map.insert(txn, "_originalFormatting", Any::Map(Arc::new(original)));
        }
    }
    apply(txn, map, INDENT_LEFT, &delta.indent_left, |v| {
        Any::Number(*v)
    });
    apply(txn, map, "indentRight", &delta.indent_right, |v| {
        Any::Number(*v)
    });
    apply(txn, map, "indentFirstLine", &delta.indent_first_line, |v| {
        Any::Number(*v)
    });
    apply(txn, map, "hangingIndent", &delta.hanging_indent, |v| {
        Any::Bool(*v)
    });
    apply(txn, map, "bidi", &delta.bidi, |v| Any::Bool(*v));
    apply(txn, map, TABS, &delta.tabs, |stops| {
        Any::Array(Arc::from(
            stops.iter().map(TabStop::to_any).collect::<Vec<_>>(),
        ))
    });
    apply(
        txn,
        map,
        DEFAULT_TEXT_FORMATTING,
        &delta.default_text_formatting,
        |dtf| {
            let value: HashMap<String, Any> =
                dtf.iter().map(|(k, v)| (k.clone(), v.clone())).collect();
            Any::Map(Arc::new(value))
        },
    );
    for (key, value) in &delta.other {
        match value {
            Some(Any::Null) => {
                map.insert(txn, key.clone(), Any::Null);
            }
            _ => set_or_remove(txn, map, key, value.clone()),
        }
    }
}

/// Whether two property values are the same, numbers by value and a map's
/// null entries ignored, as stored and host-resolved values may differ there.
fn same_value(a: &Any, b: &Any) -> bool {
    let number = |value: &Any| match value {
        Any::Number(number) => Some(*number),
        Any::BigInt(number) => Some(*number as f64),
        _ => None,
    };
    let present = |map: &HashMap<String, Any>| {
        map.iter()
            .filter(|(_, value)| **value != Any::Null)
            .map(|(key, value)| (key.clone(), value.clone()))
            .collect::<Vec<_>>()
    };
    match (a, b) {
        (Any::Array(a), Any::Array(b)) => {
            a.len() == b.len() && a.iter().zip(b.iter()).all(|(a, b)| same_value(a, b))
        }
        (Any::Map(a), Any::Map(b)) => {
            let (a, b) = (present(a), present(b));
            a.len() == b.len()
                && a.iter().all(|(key, value)| {
                    b.iter()
                        .any(|(other, b)| other == key && same_value(value, b))
                })
        }
        _ => match (number(a), number(b)) {
            (Some(a), Some(b)) => a == b,
            _ => a == b,
        },
    }
}

/// A paragraph mark's run defaults under a style: the style's (`style`) with
/// the mark's own run properties (`own`, from the source pPr) over them, as
/// the seed merges them.
fn mark_defaults(style: Option<&Any>, own: Option<&Any>) -> Any {
    let json = |value: Option<&Any>| value.and_then(|value| serde_json::to_value(value).ok());
    crate::seed::merge_text_formatting(json(style).as_ref(), json(own).as_ref())
        .filter(|merged| !merged.is_null())
        .and_then(|merged| serde_json::from_value(merged).ok())
        .unwrap_or(Any::Null)
}

fn paragraph_formatting<T: ReadTxn>(map: &MapRef, txn: &T) -> HashMap<String, Any> {
    map.iter(txn)
        .filter_map(|(key, value)| {
            if matches!(
                key.as_ref(),
                KIND_KEY | PARA_ID | PPR_INS | PPR_DEL | PPR_CHANGE
            ) {
                return None;
            }
            match value {
                Out::Any(value) if value != Any::Null => Some((key.to_string(), value)),
                _ => None,
            }
        })
        .collect()
}

/// Appends one `paragraphPropertyChange` record to the pilcrow's `pPrChange`
/// array, holding the revision info plus the complete property maps from
/// before and after — what rejection needs to rewind the paragraph.
fn append_paragraph_property_change(
    txn: &mut TransactionMut<'_>,
    map: &MapRef,
    revision_id: &str,
    author: &crate::Author,
    previous: HashMap<String, Any>,
    current: HashMap<String, Any>,
) {
    let mut changes = match map.get(txn, PPR_CHANGE) {
        Some(Out::Any(Any::Array(changes))) => changes.to_vec(),
        _ => Vec::new(),
    };
    let info = revision_value(revision_id, author);
    changes.push(Any::Map(Arc::new(HashMap::from([
        ("type".to_owned(), Any::from("paragraphPropertyChange")),
        ("info".to_owned(), info),
        (
            "previousFormatting".to_owned(),
            Any::Map(Arc::new(previous)),
        ),
        ("currentFormatting".to_owned(), Any::Map(Arc::new(current))),
    ]))));
    map.insert(txn, PPR_CHANGE, Any::Array(Arc::from(changes)));
}

fn read_tab_stops<T: ReadTxn>(map: &MapRef, txn: &T) -> Vec<Any> {
    match map.get(txn, TABS) {
        Some(Out::Any(Any::Array(stops))) => stops.to_vec(),
        _ => Vec::new(),
    }
}

fn existing_pos(stop: &Any) -> Option<f64> {
    let Any::Map(map) = stop else {
        return None;
    };
    // Stops written before ops stored the seed's shape carry `pos`.
    match map.get("position").or_else(|| map.get("pos")) {
        Some(Any::Number(pos)) => Some(*pos),
        Some(Any::BigInt(pos)) => Some(*pos as f64),
        _ => None,
    }
}

fn number_prop<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> Option<f64> {
    match map.get(txn, key) {
        Some(Out::Any(Any::Number(value))) => Some(value),
        Some(Out::Any(Any::BigInt(value))) => Some(value as f64),
        _ => None,
    }
}
