//! Operation results, errors, and positions.

use std::fmt;

use yrs::{ReadTxn, TextRef, Transact};

use crate::segments::SlotMark;

use crate::{
    CommentId, EditError, ParagraphId, Position, RevisionId, StoryId, StoryRange, pilcrows,
};

pub type OpResult<T> = Result<T, OpError>;

/// Typed operation error.
#[derive(Clone, Debug, PartialEq)]
pub enum OpError {
    UnknownStory(StoryId),
    StoryExists(StoryId),
    UnknownPara(ParagraphId),
    UnknownEmbed(String),
    UnknownComment(CommentId),
    UnknownChange(RevisionId),
    UnknownStyle(String),
    SearchNotFound(String),
    AmbiguousSearch {
        needle: String,
        occurrences: usize,
    },
    EmptyRange,
    InvalidRange {
        start: u32,
        end: u32,
    },
    OutOfBounds {
        index: u32,
        len: u32,
    },
    ExpectedPilcrow {
        story: StoryId,
        index: u32,
    },
    CannotMergeFinalParagraph(ParagraphId),
    NoParagraphBefore(ParagraphId),
    OverlapsTrackedChange,
    /// Operation targets nested story content.
    NotTopLevel,
    LockedSdt,
    /// `insert_text`/`replace_range` text may not contain paragraph or line breaks.
    TextContainsBreak,
    /// The paragraph property key is managed by the schema (`paraId`, `_kind`).
    ReservedKey(String),
    /// An invalid value was supplied for a formatting field (for example a malformed color).
    InvalidFormatValue(String),
    /// The requested table ordinal does not exist in the parent story.
    UnknownTable {
        story: StoryId,
        table_index: u32,
    },
    /// The stored table shape or requested grid operation is invalid.
    InvalidTable(String),
    InvalidComment(String),
    InvalidUpdate(String),
}

impl fmt::Display for OpError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::UnknownStory(id) => write!(f, "story {id:?} was not found"),
            Self::StoryExists(id) => write!(f, "story {id:?} already exists"),
            Self::UnknownPara(id) => write!(f, "paragraph {id:?} was not found"),
            Self::UnknownEmbed(id) => write!(f, "embed {id:?} was not found"),
            Self::UnknownComment(id) => write!(f, "comment {id:?} was not found"),
            Self::UnknownChange(id) => write!(f, "tracked change {id:?} was not found"),
            Self::UnknownStyle(id) => write!(f, "style {id:?} is not defined"),
            Self::SearchNotFound(needle) => write!(f, "search text {needle:?} was not found"),
            Self::AmbiguousSearch {
                needle,
                occurrences,
            } => write!(
                f,
                "search text {needle:?} occurs {occurrences} times; narrow the search"
            ),
            Self::EmptyRange => write!(f, "the range is empty"),
            Self::InvalidRange { start, end } => write!(f, "invalid range {start}..{end}"),
            Self::OutOfBounds { index, len } => {
                write!(f, "index {index} is outside the story length {len}")
            }
            Self::ExpectedPilcrow { story, index } => {
                write!(f, "expected a pilcrow embed at {story}:{index}")
            }
            Self::CannotMergeFinalParagraph(id) => {
                write!(f, "paragraph {id:?} is the final paragraph of its story")
            }
            Self::NoParagraphBefore(id) => {
                write!(f, "paragraph {id:?} is the first paragraph of its story")
            }
            Self::OverlapsTrackedChange => write!(f, "the range overlaps a tracked change"),
            Self::NotTopLevel => write!(f, "the target is not top-level content of its story"),
            Self::LockedSdt => write!(f, "the target is inside a locked content control"),
            Self::TextContainsBreak => {
                write!(f, "inserted text may not contain paragraph or line breaks")
            }
            Self::ReservedKey(key) => {
                write!(f, "paragraph property {key:?} is managed by the schema")
            }
            Self::InvalidFormatValue(message) => write!(f, "invalid format value: {message}"),
            Self::UnknownTable { story, table_index } => {
                write!(f, "table {table_index} was not found in story {story:?}")
            }
            Self::InvalidTable(message) => write!(f, "invalid table: {message}"),
            Self::InvalidComment(message) => write!(f, "invalid comment: {message}"),
            Self::InvalidUpdate(message) => write!(f, "invalid yrs update: {message}"),
        }
    }
}

impl std::error::Error for OpError {}

impl From<EditError> for OpError {
    fn from(error: EditError) -> Self {
        match error {
            EditError::StoryExists(id) => Self::StoryExists(id),
            EditError::StoryNotFound(id) => Self::UnknownStory(id),
            EditError::CommentNotFound(id) => Self::UnknownComment(id),
            EditError::ParagraphNotFound(id) => Self::UnknownPara(id),
            EditError::InvalidRange { start, end } => Self::InvalidRange { start, end },
            EditError::OutOfBounds { index, len } => Self::OutOfBounds { index, len },
            EditError::ExpectedPilcrow { story, index } => Self::ExpectedPilcrow { story, index },
            EditError::CannotMergeFinalParagraph { story, index } => {
                Self::CannotMergeFinalParagraph(format!("{story}:{index}"))
            }
            EditError::InvalidComment(message) => Self::InvalidComment(message),
            EditError::InvalidStateVector(message) => Self::InvalidUpdate(message),
            EditError::InvalidUpdate(message) => Self::InvalidUpdate(message),
            EditError::ReservedParagraphKey(key) => Self::ReservedKey(key),
        }
    }
}

/// Receipt returned by non-split mutations.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct Receipt {
    /// Paragraph IDs minted by this op.
    pub new_para_ids: Vec<ParagraphId>,
    /// Revision IDs stamped by this op (suggesting mode). Compound ops share one ID.
    pub revision_ids: Vec<RevisionId>,
    /// The resulting content range, paragraph-keyed.
    pub range: Option<LocRange>,
}

/// The receipt of [`crate::EditingDoc::split_paragraph`].
#[derive(Clone, Debug, PartialEq)]
pub struct SplitReceipt {
    /// The paragraph ending at the newly inserted pilcrow. Keeps the ORIGINAL paraId.
    pub first_para_id: ParagraphId,
    /// The paragraph ending at the original pilcrow, re-minted with a fresh paraId.
    pub second_para_id: ParagraphId,
    /// Revision IDs stamped by this op (suggesting mode).
    pub revision_ids: Vec<RevisionId>,
    /// The split fell at the paragraph's end (comment references aside), so
    /// the second half is the new empty paragraph.
    pub at_end: bool,
}

/// Paragraph-keyed public position with UTF-16 offsets.
///
/// `offset` lives in `[0, para_len]` where `para_len` excludes the paragraph's own pilcrow;
/// `offset == para_len` addresses the paragraph mark itself.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct Loc {
    pub story: StoryId,
    pub para: ParagraphId,
    pub offset: u32,
}

impl Loc {
    pub fn new(story: impl Into<StoryId>, para: impl Into<ParagraphId>, offset: u32) -> Self {
        Self {
            story: story.into(),
            para: para.into(),
            offset,
        }
    }
}

/// A paragraph-keyed range. `start` and `end` must belong to the same story.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct LocRange {
    pub start: Loc,
    pub end: Loc,
}

impl LocRange {
    pub fn new(start: Loc, end: Loc) -> Self {
        Self { start, end }
    }

    pub fn caret(at: Loc) -> Self {
        Self {
            start: at.clone(),
            end: at,
        }
    }

    pub fn is_caret(&self) -> bool {
        self.start == self.end
    }
}

/// One paragraph's index geometry inside a story, valid only within the current transaction.
#[derive(Clone, Debug)]
pub(crate) struct ParaBounds {
    pub para_id: ParagraphId,
    /// Story-global index of the paragraph's first unit.
    pub start: u32,
    /// Story-global index of the paragraph's pilcrow embed.
    pub pilcrow: u32,
}

impl ParaBounds {
    pub fn len(&self) -> u32 {
        self.pilcrow - self.start
    }
}

pub(crate) fn para_bounds<T: ReadTxn>(story: &TextRef, txn: &T) -> Vec<ParaBounds> {
    let mut start = 0;
    pilcrows(story, txn)
        .into_iter()
        .map(|(pilcrow, map)| {
            let bounds = ParaBounds {
                para_id: crate::map_string(&map, txn, crate::PARA_ID).unwrap_or_default(),
                start,
                pilcrow,
            };
            start = pilcrow + 1;
            bounds
        })
        .collect()
}

/// Resolves a [`Loc`] to a story-global index. Transaction-scoped by construction.
pub(crate) fn global_of_loc<T: ReadTxn>(
    story: &TextRef,
    txn: &T,
    loc: &Loc,
) -> Result<u32, OpError> {
    let bounds = para_bounds(story, txn)
        .into_iter()
        .find(|bounds| bounds.para_id == loc.para)
        .ok_or_else(|| OpError::UnknownPara(loc.para.clone()))?;
    if loc.offset > bounds.len() {
        return Err(OpError::OutOfBounds {
            index: loc.offset,
            len: bounds.len(),
        });
    }
    Ok(bounds.start + loc.offset)
}

/// The paragraphs `indices` resolve into, each the first whose pilcrow is
/// at or after it, else the story's last: read item by item from `from`
/// ([`SlotMark`]), only as far as the paragraph the last index resolves into.
fn bounds_at<T: ReadTxn, const N: usize>(
    story: &TextRef,
    txn: &T,
    from: SlotMark,
    indices: [u32; N],
) -> Option<[ParaBounds; N]> {
    let mut found: [Option<(u32, u32, yrs::MapRef)>; N] = std::array::from_fn(|_| None);
    let mut last = None;
    let (mut next, mut offset) = match from
        .and_then(|(mark, at)| Some((txn.store().get_item(&mark)?, at)))
        .filter(|(mark, _)| !mark.is_deleted())
    {
        Some((mark, at)) => (mark.right(), at),
        None => (AsRef::<yrs::branch::Branch>::as_ref(story).start(), 0),
    };
    let mut start = offset;
    while let Some(item) = next {
        if let Some(map) = crate::pilcrow_map(&item, txn) {
            for (slot, index) in found.iter_mut().zip(indices) {
                if slot.is_none() && index <= offset {
                    *slot = Some((start, offset, map.clone()));
                }
            }
            if found.iter().all(Option::is_some) {
                break;
            }
            last = Some((start, offset, map));
            start = offset + 1;
        }
        offset += crate::segments::units(&item);
        next = item.right();
    }
    let fallback = last.or_else(|| found.iter().flatten().next().cloned())?;
    Some(found.map(|bounds| {
        let (start, pilcrow, map) = bounds.as_ref().unwrap_or(&fallback);
        ParaBounds {
            para_id: crate::map_string(map, txn, crate::PARA_ID).unwrap_or_default(),
            start: *start,
            pilcrow: *pilcrow,
        }
    }))
}

/// `index` as a [`Loc`] in `bounds`, the paragraph it resolves into.
fn loc_in(story_id: &str, bounds: &ParaBounds, index: u32) -> Loc {
    Loc {
        story: story_id.to_owned(),
        para: bounds.para_id.clone(),
        offset: index
            .min(bounds.pilcrow)
            .saturating_sub(bounds.start)
            .min(bounds.len()),
    }
}

/// Maps a story-global index back to a [`Loc`]. Indices past the final pilcrow clamp to the final
/// paragraph mark.
pub(crate) fn loc_of_global<T: ReadTxn>(
    story_id: &str,
    story: &TextRef,
    txn: &T,
    index: u32,
) -> Result<Loc, OpError> {
    let [bounds] = bounds_at(story, txn, None, [index])
        .ok_or_else(|| OpError::UnknownStory(story_id.to_owned()))?;
    Ok(loc_in(story_id, &bounds, index))
}

/// Builds a [`LocRange`] for a story-global span inside an open transaction.
pub(crate) fn loc_range_in_txn<T: ReadTxn>(
    story_id: &str,
    story: &TextRef,
    txn: &T,
    start: u32,
    end: u32,
) -> Result<LocRange, OpError> {
    loc_range_from(story_id, story, txn, None, start, end)
}

/// [`loc_range_in_txn`] read from `from` ([`SlotMark`]).
pub(crate) fn loc_range_from<T: ReadTxn>(
    story_id: &str,
    story: &TextRef,
    txn: &T,
    from: SlotMark,
    start: u32,
    end: u32,
) -> Result<LocRange, OpError> {
    let [first, last] = bounds_at(story, txn, from, [start, end])
        .ok_or_else(|| OpError::UnknownStory(story_id.to_owned()))?;
    Ok(LocRange {
        start: loc_in(story_id, &first, start),
        end: loc_in(story_id, &last, end),
    })
}

impl crate::EditingDoc {
    /// Resolves a public [`Loc`] to a transient story-global [`Position`].
    pub fn locate(&self, loc: &Loc) -> OpResult<Position> {
        let txn = self.yrs_doc().transact();
        let story = crate::story_ref(&txn, &loc.story)?;
        let index = global_of_loc(&story, &txn, loc)?;
        Ok(Position::new(loc.story.clone(), index))
    }

    /// Maps a transient story-global [`Position`] to the public [`Loc`] vocabulary.
    pub fn loc_at(&self, position: &Position) -> OpResult<Loc> {
        let txn = self.yrs_doc().transact();
        let story = crate::story_ref(&txn, &position.story)?;
        loc_of_global(&position.story, &story, &txn, position.index)
    }

    /// Resolves a public [`LocRange`] to a transient [`StoryRange`].
    pub fn locate_range(&self, range: &LocRange) -> OpResult<StoryRange> {
        let txn = self.yrs_doc().transact();
        let story = crate::story_ref(&txn, &range.start.story)?;
        let start = global_of_loc(&story, &txn, &range.start)?;
        let end = global_of_loc(&story, &txn, &range.end)?;
        if end < start {
            return Err(OpError::InvalidRange { start, end });
        }
        Ok(StoryRange::new(range.start.story.clone(), start, end))
    }

    /// Maps a transient [`StoryRange`] to the public [`LocRange`] vocabulary.
    pub fn loc_range_of(&self, range: &StoryRange) -> OpResult<LocRange> {
        let txn = self.yrs_doc().transact();
        let story = crate::story_ref(&txn, &range.story)?;
        Ok(LocRange {
            start: loc_of_global(&range.story, &story, &txn, range.start)?,
            end: loc_of_global(&range.story, &story, &txn, range.end)?,
        })
    }
}
