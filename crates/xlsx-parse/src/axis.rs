//! Source-to-current row and column addresses for a preserved sheet.

use std::ops::Range;

use xlsx_model::addr::{MAX_COLS, MAX_ROWS};
use xlsx_model::{AnchorCell, AnchorEditAs, ChartAnchor};

/// A monotone partial map from source index to current index on one axis.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct AxisMap {
    segments: Vec<Segment>,
    limit: u32,
}

/// `len` consecutive source indices from `source` that now start at `current`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct Segment {
    source: u32,
    len: u32,
    current: u32,
}

impl AxisMap {
    pub fn identity(limit: u32) -> Self {
        Self {
            segments: vec![Segment {
                source: 0,
                len: limit,
                current: 0,
            }],
            limit,
        }
    }

    /// A map from `(source, len, current)` runs, in increasing order on both.
    pub fn from_runs(limit: u32, runs: impl IntoIterator<Item = (u32, u32, u32)>) -> Self {
        let mut map = Self {
            segments: Vec::new(),
            limit,
        };
        let mut segments = Vec::new();
        for (source, len, current) in runs {
            map.push(
                &mut segments,
                Segment {
                    source,
                    len,
                    current,
                },
            );
        }
        map.segments = segments;
        map
    }

    pub fn is_identity(&self) -> bool {
        self.segments
            == [Segment {
                source: 0,
                len: self.limit,
                current: 0,
            }]
    }

    /// Current index of a source index; `None` when deleted or past the limit.
    pub fn current(&self, source: u32) -> Option<u32> {
        let index = self
            .segments
            .partition_point(|segment| segment.source + segment.len <= source);
        let segment = self.segments.get(index)?;
        (source >= segment.source).then(|| segment.current + (source - segment.source))
    }

    /// The source index of a current index; `None` for an inserted one.
    pub fn source(&self, current: u32) -> Option<u32> {
        let index = self
            .segments
            .partition_point(|segment| segment.current + segment.len <= current);
        let segment = self.segments.get(index)?;
        (current >= segment.current).then(|| segment.source + (current - segment.current))
    }

    /// Current spans of a source span, in order; gaps stay split.
    pub fn current_ranges(&self, source: Range<u32>) -> Vec<Range<u32>> {
        let mut ranges: Vec<Range<u32>> = Vec::new();
        let mut cursor = source.start;
        for segment in &self.segments {
            let start = segment.source.max(source.start);
            let end = (segment.source + segment.len).min(source.end);
            if start >= end {
                continue;
            }
            let current = segment.current + (start - segment.source);
            let range = current..current + (end - start);
            match ranges.last_mut() {
                Some(last) if last.end == range.start && cursor == start => {
                    last.end = range.end;
                }
                _ => ranges.push(range),
            }
            cursor = end;
        }
        ranges
    }

    /// The current span of the surviving indices in `first..=last`, gaps and
    /// insertions inside included; `None` when every one was deleted.
    pub fn span(&self, first: u32, last: u32) -> Option<(u32, u32)> {
        let ranges = self.current_ranges(first..last.saturating_add(1));
        Some((ranges.first()?.start, ranges.last()?.end - 1))
    }

    /// Where a drawing marker on source index `source` lands, as a chart's
    /// marker does through each edit: its current index, or the point its row
    /// or column was deleted at. One pushed past the grid stays on its edge.
    pub fn marker(&self, source: u32) -> u32 {
        let landing = self.current(source).unwrap_or_else(|| {
            match self.current_ranges(source..self.limit).first() {
                Some(next) => next.start,
                None => self
                    .current_ranges(0..source)
                    .last()
                    .map_or(0, |previous| previous.end),
            }
        });
        landing.min(self.limit.saturating_sub(1))
    }

    /// Records `count` indices inserted at current index `at`.
    pub fn insert(&mut self, at: u32, count: u32) {
        let mut segments = Vec::with_capacity(self.segments.len() + 1);
        for segment in &self.segments {
            if segment.current + segment.len <= at {
                self.push(&mut segments, *segment);
            } else if segment.current >= at {
                self.push(
                    &mut segments,
                    Segment {
                        current: segment.current.saturating_add(count),
                        ..*segment
                    },
                );
            } else {
                let head = at - segment.current;
                self.push(
                    &mut segments,
                    Segment {
                        len: head,
                        ..*segment
                    },
                );
                self.push(
                    &mut segments,
                    Segment {
                        source: segment.source + head,
                        len: segment.len - head,
                        current: at.saturating_add(count),
                    },
                );
            }
        }
        self.segments = segments;
    }

    /// Records `count` indices deleted from current index `at`.
    pub fn delete(&mut self, at: u32, count: u32) {
        let end = at.saturating_add(count);
        let mut segments = Vec::with_capacity(self.segments.len() + 1);
        for segment in &self.segments {
            let segment_end = segment.current + segment.len;
            if segment_end <= at {
                self.push(&mut segments, *segment);
                continue;
            }
            if segment.current >= end {
                self.push(
                    &mut segments,
                    Segment {
                        current: segment.current - count,
                        ..*segment
                    },
                );
                continue;
            }
            if segment.current < at {
                let head = at - segment.current;
                self.push(
                    &mut segments,
                    Segment {
                        len: head,
                        ..*segment
                    },
                );
            }
            if segment_end > end {
                let skipped = end - segment.current;
                self.push(
                    &mut segments,
                    Segment {
                        source: segment.source + skipped,
                        len: segment.len - skipped,
                        current: at,
                    },
                );
            }
        }
        self.segments = segments;
    }

    /// Appends a segment clipped to the limit, merging gapless neighbors.
    fn push(&self, segments: &mut Vec<Segment>, mut segment: Segment) {
        if segment.current >= self.limit {
            return;
        }
        segment.len = segment.len.min(self.limit - segment.current);
        if segment.len == 0 {
            return;
        }
        if let Some(last) = segments.last_mut()
            && last.source + last.len == segment.source
            && last.current + last.len == segment.current
        {
            last.len += segment.len;
            return;
        }
        segments.push(segment);
    }
}

/// Both axes of one preserved sheet.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SheetAxes {
    pub rows: AxisMap,
    pub cols: AxisMap,
}

impl Default for SheetAxes {
    fn default() -> Self {
        Self {
            rows: AxisMap::identity(MAX_ROWS),
            cols: AxisMap::identity(MAX_COLS),
        }
    }
}

impl SheetAxes {
    pub fn is_identity(&self) -> bool {
        self.rows.is_identity() && self.cols.is_identity()
    }

    /// A drawing anchor moved along the axes, honouring `editAs` as a chart
    /// anchor does (xlsx-ops `remap_anchor`): a two-cell anchor moves and
    /// resizes, a one-cell one moves and keeps its size, an absolute one stays.
    pub(crate) fn move_anchor(&self, anchor: ChartAnchor) -> ChartAnchor {
        let mark = |cell: AnchorCell| AnchorCell {
            row: self.rows.marker(cell.row),
            col: self.cols.marker(cell.col),
            ..cell
        };
        match anchor {
            ChartAnchor::Absolute { .. }
            | ChartAnchor::TwoCell {
                edit_as: AnchorEditAs::Absolute,
                ..
            } => anchor,
            ChartAnchor::OneCell { from, extent } => ChartAnchor::OneCell {
                from: mark(from),
                extent,
            },
            ChartAnchor::TwoCell {
                from,
                to,
                edit_as: AnchorEditAs::TwoCell,
            } => ChartAnchor::TwoCell {
                from: mark(from),
                to: mark(to),
                edit_as: AnchorEditAs::TwoCell,
            },
            ChartAnchor::TwoCell { from, to, edit_as } => {
                let moved = mark(from);
                let shift = |index: u32, delta: i64, limit: u32| {
                    (i64::from(index) + delta).clamp(0, i64::from(limit) - 1) as u32
                };
                ChartAnchor::TwoCell {
                    from: moved,
                    to: AnchorCell {
                        row: shift(to.row, i64::from(moved.row) - i64::from(from.row), MAX_ROWS),
                        col: shift(to.col, i64::from(moved.col) - i64::from(from.col), MAX_COLS),
                        ..to
                    },
                    edit_as,
                }
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn insert_shifts_everything_at_and_after_the_point() {
        let mut map = AxisMap::identity(100);
        map.insert(3, 2);
        assert_eq!(map.current(2), Some(2));
        assert_eq!(map.current(3), Some(5));
        assert_eq!(map.current(99), None);
        assert_eq!(map.source(3), None);
        assert_eq!(map.source(4), None);
        assert_eq!(map.source(5), Some(3));
        assert!(!map.is_identity());
    }

    #[test]
    fn delete_drops_the_span_and_pulls_later_indices_back() {
        let mut map = AxisMap::identity(100);
        map.delete(3, 2);
        assert_eq!(map.current(2), Some(2));
        assert_eq!(map.current(3), None);
        assert_eq!(map.current(4), None);
        assert_eq!(map.current(5), Some(3));
        assert_eq!(map.source(3), Some(5));
        assert_eq!(map.current_ranges(0..10), vec![0..3, 3..8]);
    }

    #[test]
    fn inverse_edits_restore_identity() {
        let mut map = AxisMap::identity(100);
        map.insert(3, 2);
        map.delete(3, 2);
        for index in 0..98 {
            assert_eq!(map.current(index), Some(index));
            assert_eq!(map.source(index), Some(index));
        }
        assert_eq!(map.current(98), None);
        assert_eq!(map.current(99), None);
        map.delete(10, 4);
        map.insert(10, 4);
        assert_eq!(map.current(9), Some(9));
        assert_eq!(map.current(10), None);
        assert_eq!(map.current(14), Some(14));
        assert_eq!(map.source(12), None);
    }

    #[test]
    fn runs_build_the_map_an_edit_sequence_would() {
        let mut edited = AxisMap::identity(100);
        edited.delete(3, 2);
        edited.insert(10, 4);
        let runs = AxisMap::from_runs(100, [(0, 3, 0), (5, 7, 3), (12, 88, 14)]);
        assert_eq!(runs, edited);
    }

    #[test]
    fn current_ranges_split_around_inserted_indices() {
        let mut map = AxisMap::identity(100);
        map.insert(4, 3);
        assert_eq!(map.current_ranges(0..10), vec![0..4, 7..13]);
        assert_eq!(map.current_ranges(20..20), Vec::<Range<u32>>::new());
    }
}
