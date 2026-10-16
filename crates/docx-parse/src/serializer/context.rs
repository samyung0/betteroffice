//! Deterministic recursive serializer state.

use std::collections::{BTreeSet, HashMap, HashSet};

use crate::paragraph::HexIdAllocator;
use crate::paragraph_identity::{allocate_paragraph_id, format_paragraph_id};
use crate::xml::ParseError;

use super::s10::SerializerDeterminism;

/// Per-serialization state. Generated identities are taken only from the
/// injected seed; serializers never consult ambient randomness or a clock.
#[derive(Clone, Debug)]
pub struct SerializerContext {
    ids: HexIdAllocator,
    reserved_ids: HashSet<String>,
    now: String,
    rendered_page_breaks: Vec<bool>,
    /// Ids of the comments the package keeps; `None` keeps every comment marker.
    comments: Option<HashSet<u64>>,
    /// Comments whose reference mark is written, so none is generated for them.
    references: HashSet<u64>,
    warnings: Vec<String>,
    paragraph_ids: BTreeSet<u32>,
    splice: Option<SpliceRecorder>,
}

/// The XML written for each model paragraph a spliced part addresses by `sourceOrdinal`.
#[derive(Clone, Debug, Default)]
struct SpliceRecorder {
    expected: HashSet<u32>,
    recorded: RecordedParagraphs,
    conflict: bool,
}

/// The paragraphs written for a spliced part, outside any other paragraph.
#[derive(Clone, Debug, Default)]
pub(crate) struct RecordedParagraphs {
    /// The XML of each paragraph written from a source paragraph, by `sourceOrdinal`.
    pub(crate) written: HashMap<u32, String>,
    /// The XML of each paragraph without a `sourceOrdinal`, in writing order.
    pub(crate) inserted: Vec<String>,
    /// Every recorded paragraph in writing order: its `sourceOrdinal`, or `None` for the next
    /// of `inserted`.
    pub(crate) order: Vec<Option<u32>>,
    /// The empty paragraph written for each table cell without content, as the number of
    /// `order` entries written before it.
    pub(crate) placeholders: Vec<usize>,
}

impl SerializerContext {
    pub fn new(determinism: &SerializerDeterminism) -> Result<Self, ParseError> {
        determinism.validate()?;
        Ok(Self {
            ids: HexIdAllocator::from_sha256(&determinism.seed)?,
            reserved_ids: HashSet::new(),
            now: determinism.now.clone(),
            rendered_page_breaks: Vec::new(),
            comments: None,
            references: HashSet::new(),
            warnings: Vec::new(),
            paragraph_ids: BTreeSet::new(),
            splice: None,
        })
    }

    /// Limits comment anchors and references to these comments.
    pub(crate) fn keep_comments(&mut self, ids: impl IntoIterator<Item = f64>) {
        self.comments = Some(ids.into_iter().map(f64::to_bits).collect());
    }

    /// Whether a comment anchor or reference with this id is written.
    pub(crate) fn keeps_comment(&self, id: f64) -> bool {
        self.comments
            .as_ref()
            .is_none_or(|ids| ids.contains(&id.to_bits()))
    }

    pub(crate) fn note_comment_reference(&mut self, id: f64) {
        self.references.insert(id.to_bits());
    }

    pub(crate) fn references_comment(&self, id: f64) -> bool {
        self.references.contains(&id.to_bits())
    }

    /// Paragraph IDs the output already uses, which [`Self::allocate_paragraph_id`] avoids.
    pub fn reserve_paragraph_ids(&mut self, ids: impl IntoIterator<Item = u32>) {
        self.paragraph_ids.extend(ids);
    }

    /// A paragraph ID for `owner` through the shared deterministic allocator.
    pub fn allocate_paragraph_id(&mut self, owner: &str) -> String {
        let id = allocate_paragraph_id(owner, &self.paragraph_ids)
            .expect("parse budgets keep packages far below 2^31 paragraphs");
        self.paragraph_ids.insert(id);
        format_paragraph_id(id)
    }

    /// Records a non-fatal save diagnostic for the caller to report.
    pub fn warn(&mut self, warning: impl Into<String>) {
        self.warnings.push(warning.into());
    }

    /// Drains the diagnostics recorded so far.
    pub fn take_warnings(&mut self) -> Vec<String> {
        std::mem::take(&mut self.warnings)
    }

    pub fn allocate_hex_id(&mut self) -> String {
        loop {
            let id = self.ids.allocate();
            if self.reserved_ids.insert(id.clone()) {
                return id;
            }
        }
    }

    pub(crate) fn reserve_hex_id(&mut self, id: &str) {
        self.reserved_ids.insert(id.to_ascii_uppercase());
    }

    /// Injected clock for timestamp-bearing part writers.
    pub fn now(&self) -> &str {
        &self.now
    }

    /// `wp:docPr/@id` is an unsigned decimal integer. Reuse the canonical
    /// xorshift stream, converting its valid long-hex value to decimal.
    pub fn allocate_drawing_id(&mut self) -> String {
        let hex = self.allocate_hex_id();
        u32::from_str_radix(&hex, 16)
            .expect("HexIdAllocator always returns eight hexadecimal digits")
            .to_string()
    }

    /// Records the XML written for the model paragraphs whose `sourceOrdinal` is in `expected`
    /// until [`Self::end_splice`].
    pub(crate) fn begin_splice(&mut self, expected: impl IntoIterator<Item = u32>) {
        self.splice = Some(SpliceRecorder {
            expected: expected.into_iter().collect(),
            ..SpliceRecorder::default()
        });
    }

    /// The recorded paragraphs; `None` when one source paragraph was written twice differently,
    /// or one with a `sourceOrdinal` the part does not address was written.
    pub(crate) fn end_splice(&mut self) -> Option<RecordedParagraphs> {
        self.splice
            .take()
            .filter(|recorder| !recorder.conflict)
            .map(|recorder| recorder.recorded)
    }

    /// Records a paragraph written outside any other paragraph: by its `sourceOrdinal`, which
    /// must be one the part addresses, or as a paragraph the source lacks.
    pub(crate) fn record_paragraph(&mut self, ordinal: Option<u32>, xml: &str) {
        if !self.rendered_page_breaks.is_empty() {
            return;
        }
        let Some(recorder) = self.splice.as_mut() else {
            return;
        };
        let recorded = &mut recorder.recorded;
        match ordinal {
            None => recorded.inserted.push(xml.to_owned()),
            Some(ordinal) if recorder.expected.contains(&ordinal) => {
                if let Some(previous) = recorded.written.get(&ordinal) {
                    recorder.conflict |= previous != xml;
                    return;
                }
                recorded.written.insert(ordinal, xml.to_owned());
            }
            Some(_) => {
                recorder.conflict = true;
                return;
            }
        }
        recorded.order.push(ordinal);
    }

    /// Records the empty paragraph written outside any other paragraph for a table cell
    /// without content.
    pub(crate) fn record_placeholder(&mut self) {
        if !self.rendered_page_breaks.is_empty() {
            return;
        }
        if let Some(recorder) = self.splice.as_mut() {
            let recorded = &mut recorder.recorded;
            recorded.placeholders.push(recorded.order.len());
        }
    }

    pub(crate) fn enter_paragraph(&mut self, rendered_page_break_before: bool) {
        self.rendered_page_breaks.push(rendered_page_break_before);
    }

    pub(crate) fn leave_paragraph(&mut self) {
        let _ = self.rendered_page_breaks.pop();
    }

    /// Consumes every active rendered-page-break marker.
    pub(crate) fn take_rendered_page_breaks(&mut self) -> usize {
        let mut count = 0;
        for pending in &mut self.rendered_page_breaks {
            if *pending {
                *pending = false;
                count += 1;
            }
        }
        count
    }
}
