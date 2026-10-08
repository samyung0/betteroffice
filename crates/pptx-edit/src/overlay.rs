//! Override layout (deck schema 6): the document holds only what users changed
//! over the fingerprinted source package. A source entry (the slide order, one
//! slide, one shape, one story, the comments) enters the document on its first
//! edit, written by a client id derived from the source fingerprint and the
//! entry's key ([`writer`]). Every peer that copies an entry writes the same
//! items, so concurrent first edits of one entry converge on one copy, and an
//! entry was copied exactly when its writer's clock is above zero.

use std::collections::{HashMap, HashSet};

use pptx_parse::{PptxPackage, ShapeNode, Slide};
use sha2::{Digest, Sha256};
use yrs::{
    Array, ArrayPrelim, ClientID, Map, MapPrelim, MapRef, ReadTxn, StateVector, TextRef, Transact,
    WriteTxn,
};

use crate::deck::{baseline_snapshot, map_string_array, required_map};
use crate::{
    CommentSnapshot, EditError, EditResult, META, SHAPES, SLIDE_ORDER, SLIDES, STORIES,
    ShapeSnapshot, SlideSnapshot, StorySnapshot,
};

/// The deck schema of the override layout.
pub(crate) const SCHEMA_OVERRIDES: f64 = 6.0;
/// Writers of copied source entries; above every session client id.
pub(crate) const RESERVED_CLIENTS: u64 = 1 << 52;
pub(crate) const COPY_ORIGIN: &str = "pptx:copy";

/// How a deck's document holds its source package.
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub enum Layout {
    /// Schema 5: the whole deck is seeded into the document.
    #[default]
    Seeded,
    /// Schema 6: only edited entries are in the document.
    Overrides,
}

/// A source entry that is copied into the document as a whole.
#[derive(Clone, Copy, Debug)]
pub(crate) enum Entry<'a> {
    Order,
    Slide(&'a str),
    Shape(&'a str),
    Story(&'a str),
    Comments,
}

/// The client id that copies `kind`/`key` of the source with `fingerprint`.
pub(crate) fn writer(fingerprint: &str, kind: &str, key: &str) -> u64 {
    let digest = Sha256::new()
        .chain_update(b"pptx-override/1\0")
        .chain_update(fingerprint.as_bytes())
        .chain_update(b"\0")
        .chain_update(kind.as_bytes())
        .chain_update(b"\0")
        .chain_update(key.as_bytes())
        .finalize();
    let mut bytes = [0_u8; 8];
    bytes.copy_from_slice(&digest[..8]);
    RESERVED_CLIENTS + u64::from_le_bytes(bytes) % (RESERVED_CLIENTS - 1)
}

pub(crate) struct BaseSlide {
    pub index: usize,
    pub writer: u64,
    /// The slide without its shapes.
    pub snapshot: SlideSnapshot,
    pub shape_ids: Vec<String>,
}

pub(crate) struct BaseShape {
    pub writer: u64,
    pub slide_index: usize,
    /// The slide or group shape whose list holds it.
    pub parent: String,
    /// The shape without its stories and children.
    pub snapshot: ShapeSnapshot,
    pub story_ids: Vec<String>,
    pub child_ids: Vec<String>,
}

pub(crate) struct BaseStory {
    pub writer: u64,
    pub shape_id: String,
    pub snapshot: StorySnapshot,
}

/// The source deck by entry, as the seed would write it.
pub(crate) struct BaseIndex {
    pub fingerprint: String,
    pub order_writer: u64,
    pub comments_writer: u64,
    pub order: Vec<String>,
    pub slides: HashMap<String, BaseSlide>,
    pub shapes: HashMap<String, BaseShape>,
    pub stories: HashMap<String, BaseStory>,
    pub comments: Vec<CommentSnapshot>,
}

impl BaseIndex {
    pub fn new(package: &PptxPackage, fingerprint: &str) -> EditResult<Self> {
        let deck = baseline_snapshot(package)?;
        let mut index = Self {
            fingerprint: fingerprint.to_owned(),
            order_writer: writer(fingerprint, "order", ""),
            comments_writer: writer(fingerprint, "comments", ""),
            order: Vec::with_capacity(deck.slides.len()),
            slides: HashMap::new(),
            shapes: HashMap::new(),
            stories: HashMap::new(),
            comments: deck.comments,
        };
        for (slide_index, mut slide) in deck.slides.into_iter().enumerate() {
            let shapes = std::mem::take(&mut slide.shapes);
            let shape_ids = shapes.iter().map(|shape| shape.id.clone()).collect();
            for shape in shapes {
                index.add_shape(slide_index, &slide.id, shape);
            }
            index.order.push(slide.id.clone());
            index.slides.insert(
                slide.id.clone(),
                BaseSlide {
                    index: slide_index,
                    writer: writer(fingerprint, "slide", &slide.id),
                    snapshot: slide,
                    shape_ids,
                },
            );
        }
        Ok(index)
    }

    fn add_shape(&mut self, slide_index: usize, parent: &str, mut shape: ShapeSnapshot) {
        let stories = std::mem::take(&mut shape.text_stories);
        let children = std::mem::take(&mut shape.children);
        let story_ids = stories.iter().map(|story| story.id.clone()).collect();
        let child_ids = children.iter().map(|child| child.id.clone()).collect();
        for story in stories {
            self.stories.insert(
                story.id.clone(),
                BaseStory {
                    writer: writer(&self.fingerprint, "story", &story.id),
                    shape_id: shape.id.clone(),
                    snapshot: story,
                },
            );
        }
        let shape_id = shape.id.clone();
        for child in children {
            self.add_shape(slide_index, &shape_id, child);
        }
        self.shapes.insert(
            shape_id.clone(),
            BaseShape {
                writer: writer(&self.fingerprint, "shape", &shape_id),
                slide_index,
                parent: parent.to_owned(),
                snapshot: shape,
                story_ids,
                child_ids,
            },
        );
    }

    pub fn entry_writer(&self, entry: Entry<'_>) -> Option<u64> {
        match entry {
            Entry::Order => Some(self.order_writer),
            Entry::Comments => (!self.comments.is_empty()).then_some(self.comments_writer),
            Entry::Slide(id) => self.slides.get(id).map(|slide| slide.writer),
            Entry::Shape(id) => self.shapes.get(id).map(|shape| shape.writer),
            Entry::Story(id) => self.stories.get(id).map(|story| story.writer),
        }
    }
}

/// Reads "the document's record if the entry was copied or made, else the
/// source's". In the seeded layout every record is the document's.
pub(crate) struct Reader<'a, T: ReadTxn> {
    pub txn: &'a T,
    pub index: Option<&'a BaseIndex>,
    sv: StateVector,
    pub slides: MapRef,
    pub shapes: MapRef,
    pub stories: MapRef,
}

pub(crate) enum SlideRec<'a> {
    Live(MapRef),
    Base(&'a BaseSlide),
}

pub(crate) enum ShapeRec<'a> {
    Live(MapRef),
    Base(&'a BaseShape),
}

pub(crate) enum StoryRec<'a> {
    Live(TextRef),
    Base(&'a BaseStory),
}

impl<'a, T: ReadTxn> Reader<'a, T> {
    pub fn new(txn: &'a T, index: Option<&'a BaseIndex>) -> EditResult<Self> {
        Ok(Self {
            sv: if index.is_some() {
                txn.state_vector()
            } else {
                StateVector::default()
            },
            slides: required_map(txn, SLIDES)?,
            shapes: required_map(txn, SHAPES)?,
            stories: required_map(txn, STORIES)?,
            index,
            txn,
        })
    }

    pub fn copied(&self, writer: u64) -> bool {
        self.sv.get(&ClientID::new(writer)) > 0
    }

    /// Slide ids in deck order, first occurrence of each. An id whose slide
    /// record a concurrent delete removed is skipped, as a slide's shape list
    /// skips removed shapes: the delete wins over a peer's move.
    pub fn order(&self) -> EditResult<Vec<String>> {
        let ids = match self.index {
            Some(index) if !self.copied(index.order_writer) => index.order.clone(),
            _ => crate::deck::string_array_ref(&crate::deck::required_order(self.txn)?, self.txn),
        };
        let mut seen = HashSet::new();
        Ok(ids
            .into_iter()
            .filter(|id| self.slide(id).is_some() && seen.insert(id.clone()))
            .collect())
    }

    pub fn slide(&self, id: &str) -> Option<SlideRec<'a>> {
        if let Some(base) = self.index.and_then(|index| index.slides.get(id))
            && !self.copied(base.writer)
        {
            return Some(SlideRec::Base(base));
        }
        self.slides
            .get(self.txn, id)
            .and_then(|value| value.cast::<MapRef>().ok())
            .map(SlideRec::Live)
    }

    pub fn shape(&self, id: &str) -> Option<ShapeRec<'a>> {
        if let Some(base) = self.index.and_then(|index| index.shapes.get(id))
            && !self.copied(base.writer)
        {
            return Some(ShapeRec::Base(base));
        }
        self.shapes
            .get(self.txn, id)
            .and_then(|value| value.cast::<MapRef>().ok())
            .map(ShapeRec::Live)
    }

    pub fn story(&self, id: &str) -> Option<StoryRec<'a>> {
        if let Some(base) = self.index.and_then(|index| index.stories.get(id))
            && !self.copied(base.writer)
        {
            return Some(StoryRec::Base(base));
        }
        self.stories
            .get(self.txn, id)
            .and_then(|value| value.cast::<TextRef>().ok())
            .map(StoryRec::Live)
    }

    /// The live shape ids a slide lists, first occurrence of each.
    pub fn slide_shape_ids(&self, slide: &SlideRec<'_>) -> EditResult<Vec<String>> {
        let ids = match slide {
            SlideRec::Base(base) => base.shape_ids.clone(),
            SlideRec::Live(map) => crate::deck::string_array_ref(
                &crate::deck::slide_shape_order(map, self.txn)?,
                self.txn,
            ),
        };
        let mut seen = HashSet::new();
        Ok(ids
            .into_iter()
            .filter(|id| self.shape(id).is_some() && seen.insert(id.clone()))
            .collect())
    }

    pub fn shape_lists(&self, shape: &ShapeRec<'_>) -> EditResult<(Vec<String>, Vec<String>)> {
        Ok(match shape {
            ShapeRec::Base(base) => (base.story_ids.clone(), base.child_ids.clone()),
            ShapeRec::Live(map) => (
                map_string_array(map, self.txn, "textStories")?,
                map_string_array(map, self.txn, "children")?,
            ),
        })
    }

    /// Whether a source shape is still on a slide of the deck.
    pub fn base_shape_live(&self, shape_id: &str) -> EditResult<bool> {
        let Some(index) = self.index else {
            return Ok(true);
        };
        let mut child = shape_id.to_owned();
        let Some(mut parent) = index.shapes.get(shape_id).map(|shape| shape.parent.clone()) else {
            return Ok(false);
        };
        for _ in 0..crate::deck::MAX_SHAPE_DEPTH {
            if let Some(slide) = self
                .slide(&parent)
                .filter(|_| index.slides.contains_key(&parent))
            {
                return Ok(self.slide_shape_ids(&slide)?.contains(&child)
                    && self.order()?.contains(&parent));
            }
            let Some(shape) = self.shape(&parent) else {
                return Ok(false);
            };
            if !self.shape_lists(&shape)?.1.contains(&child) {
                return Ok(false);
            }
            child = parent.clone();
            parent = match index.shapes.get(&parent) {
                Some(shape) => shape.parent.clone(),
                None => return Ok(false),
            };
        }
        Ok(false)
    }
}

/// The source slide a seeded slide id names.
fn source_slide<'p>(package: &'p PptxPackage, index: &BaseIndex, id: &str) -> Option<&'p Slide> {
    package.slides.get(index.slides.get(id)?.index)
}

/// The source node a seeded shape id names (`<slide>:shape:<i>[.<j>…]`).
fn source_shape<'p>(
    package: &'p PptxPackage,
    index: &BaseIndex,
    id: &str,
) -> Option<&'p ShapeNode> {
    let slide = package.slides.get(index.shapes.get(id)?.slide_index)?;
    let (_, path) = id.rsplit_once(":shape:")?;
    let mut nodes = slide.shapes.as_slice();
    let mut node = None;
    for step in path.split('.') {
        let found = nodes.get(step.parse::<usize>().ok()?)?;
        nodes = match found {
            ShapeNode::Group(group) => &group.children,
            _ => &[],
        };
        node = Some(found);
    }
    node
}

/// The update that copies `entry` of the source into a document: what the
/// seed writes for it, under the entry's own writer.
pub(crate) fn copy_update(
    package: &PptxPackage,
    index: &BaseIndex,
    entry: Entry<'_>,
) -> EditResult<Option<Vec<u8>>> {
    let Some(client) = index.entry_writer(entry) else {
        return Ok(None);
    };
    let scratch = crate::doc_with_client_id(client);
    {
        let mut txn = scratch.transact_mut_with(COPY_ORIGIN);
        match entry {
            Entry::Order => {
                let order = txn.get_or_insert_array(SLIDE_ORDER);
                for id in &index.order {
                    order.push_back(&mut txn, id.as_str());
                }
            }
            Entry::Slide(id) => {
                let slide = source_slide(package, index, id)
                    .ok_or_else(|| EditError::SlideNotFound(id.to_owned()))?;
                let base = &index.slides[id];
                let slides = txn.get_or_insert_map(SLIDES);
                let map = slides.insert(&mut txn, id, MapPrelim::default());
                map.insert(&mut txn, "id", id);
                map.insert(&mut txn, "sourcePartPath", slide.part_path.as_str());
                if let Some(layout) = &slide.layout_part_path {
                    map.insert(&mut txn, "layoutPartPath", layout.as_str());
                }
                if let Some(name) = &slide.name {
                    map.insert(&mut txn, "name", name.as_str());
                }
                if !slide.notes.is_empty() {
                    map.insert(&mut txn, "notes", slide.notes.as_str());
                }
                let shapes = map.insert(&mut txn, "shapes", ArrayPrelim::default());
                for shape_id in &base.shape_ids {
                    shapes.push_back(&mut txn, shape_id.as_str());
                }
            }
            Entry::Shape(id) => {
                let node = source_shape(package, index, id)
                    .ok_or_else(|| EditError::ShapeNotFound(id.to_owned()))?;
                let base = &index.shapes[id];
                let shapes = txn.get_or_insert_map(SHAPES);
                crate::deck::write_shape_record(
                    &shapes,
                    &mut txn,
                    id,
                    node,
                    &base.story_ids,
                    &base.child_ids,
                )?;
            }
            Entry::Story(id) => {
                let (slide, body) = crate::deck::source_story_body(package, id)
                    .ok_or_else(|| EditError::StoryNotFound(id.to_owned()))?;
                let theme = pptx_parse::slide_theme(package, Some(&slide.part_path), None);
                let stories = txn.get_or_insert_map(STORIES);
                crate::story::seed_story(&stories, &mut txn, id, body, Some(&theme))?;
            }
            Entry::Comments => {
                let slide_id_by_part: HashMap<&str, &str> = index
                    .slides
                    .iter()
                    .filter_map(|(id, slide)| {
                        Some((
                            package.slides.get(slide.index)?.part_path.as_str(),
                            id.as_str(),
                        ))
                    })
                    .collect();
                crate::comments::seed_comments(&mut txn, package, &|part| {
                    slide_id_by_part.get(part).map(|id| (*id).to_owned())
                })?;
            }
        }
    }
    Ok(Some(
        scratch
            .transact()
            .encode_state_as_update_v1(&StateVector::default()),
    ))
}

/// The schema-6 seed: deck metadata only.
pub(crate) fn seed_meta(
    txn: &mut yrs::TransactionMut<'_>,
    package: &PptxPackage,
    fingerprint: &str,
) {
    let meta = txn.get_or_insert_map(META);
    meta.insert(txn, "schemaVersion", SCHEMA_OVERRIDES);
    meta.insert(txn, "fingerprint", fingerprint);
    meta.insert(txn, "widthEmu", package.presentation.width_emu as f64);
    meta.insert(txn, "heightEmu", package.presentation.height_emu as f64);
    meta.insert(
        txn,
        "commentFlavor",
        crate::comments::flavor_key(package.comment_flavor.unwrap_or_default()),
    );
}
