use std::collections::{BTreeMap, BTreeSet};

use serde::{Deserialize, Serialize};

#[derive(Clone, Debug, PartialEq, Eq, PartialOrd, Ord, Serialize, Deserialize)]
pub(super) struct Point {
    pub run: String,
    pub offset: u64,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub(super) struct Span {
    pub run: String,
    pub start: u64,
    pub len: u64,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase")]
pub(super) enum Change {
    Insert {
        before: Point,
        after: Option<Point>,
        count: u32,
    },
    Delete {
        spans: Vec<Span>,
    },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) struct Axis {
    pub spans: Vec<Span>,
    limit: u32,
    gaps: Vec<(Span, u32)>,
    known: BTreeMap<String, u64>,
    insertions: BTreeMap<String, (Point, Option<Point>, u32)>,
}

impl Axis {
    #[cfg(test)]
    pub fn base(limit: u32) -> Self {
        Self {
            spans: vec![Span {
                run: "base".into(),
                start: 0,
                len: limit.into(),
            }],
            limit,
            gaps: Vec::new(),
            known: BTreeMap::from([("base".into(), u64::from(limit))]),
            insertions: BTreeMap::new(),
        }
    }

    pub fn project(
        limit: u32,
        changes: &BTreeMap<String, Change>,
        active: &BTreeSet<String>,
    ) -> Result<Self, String> {
        let mut insertions = BTreeMap::<String, BTreeMap<u64, Vec<String>>>::new();
        let mut removed = BTreeMap::<String, Vec<(u64, u64)>>::new();
        let mut base_length = u64::from(limit);
        for (id, change) in changes {
            match change {
                Change::Insert { before, count, .. } => {
                    if *count == 0 || *count > limit || id == "base" {
                        return Err("invalid inserted axis run".into());
                    }
                    insertions
                        .entry(before.run.clone())
                        .or_default()
                        .entry(before.offset)
                        .or_default()
                        .push(id.clone());
                }
                Change::Delete { spans } if active.contains(id) => {
                    if spans.is_empty() {
                        return Err("empty axis deletion".into());
                    }
                    for span in spans {
                        if span.len == 0 {
                            return Err("empty deleted axis span".into());
                        }
                        let end = span
                            .start
                            .checked_add(span.len)
                            .ok_or("axis span overflow")?;
                        base_length = base_length
                            .checked_add(span.len)
                            .ok_or("axis length overflow")?;
                        removed
                            .entry(span.run.clone())
                            .or_default()
                            .push((span.start, end));
                    }
                }
                _ => {}
            }
        }
        let maximum_base = changes
            .values()
            .try_fold(u64::from(limit), |size, change| {
                if let Change::Delete { spans } = change {
                    spans.iter().try_fold(size, |size, span| {
                        size.checked_add(span.len).ok_or("axis length overflow")
                    })
                } else {
                    Ok(size)
                }
            })?;
        for change in changes.values() {
            if let Change::Delete { spans } = change {
                let mut count = 0_u64;
                for span in spans {
                    count = count
                        .checked_add(span.len)
                        .ok_or("axis deletion length overflow")?;
                    let run_len = if span.run == "base" {
                        maximum_base
                    } else {
                        match changes.get(&span.run) {
                            Some(Change::Insert { count, .. }) => u64::from(*count),
                            _ => return Err("deleted axis run is absent".into()),
                        }
                    };
                    if span.len == 0
                        || span
                            .start
                            .checked_add(span.len)
                            .is_none_or(|end| end > run_len)
                    {
                        return Err("deleted axis span out of bounds".into());
                    }
                }
                if count == 0 || count > u64::from(limit) {
                    return Err("axis deletion exceeds grid".into());
                }
            }
        }
        for points in insertions.values_mut() {
            for ids in points.values_mut() {
                ids.sort();
            }
        }
        for ranges in removed.values_mut() {
            ranges.sort();
        }
        let mut spans = Vec::new();
        let mut gaps = Vec::new();
        let mut visited = BTreeSet::new();
        let mut length = 0;
        enum Task {
            Run(String, u64, bool),
            Span(String, u64, u64, bool),
        }
        let mut pending = vec![Task::Run("base".into(), base_length, true)];
        while let Some(task) = pending.pop() {
            match task {
                Task::Run(run, end, visible) => {
                    if !visited.insert(run.clone()) {
                        return Err("cyclic axis insertion".into());
                    }
                    let mut pieces = Vec::new();
                    let mut cursor = 0;
                    if let Some(boundaries) = insertions.get(&run) {
                        for (&offset, children) in boundaries {
                            if offset >= end {
                                return Err("axis insertion anchor out of bounds".into());
                            }
                            pieces.push(Task::Span(run.clone(), cursor, offset, visible));
                            for child in children {
                                let Some(Change::Insert { count, .. }) = changes.get(child) else {
                                    return Err("missing axis insertion".into());
                                };
                                pieces.push(Task::Run(
                                    child.clone(),
                                    u64::from(*count),
                                    active.contains(child),
                                ));
                            }
                            cursor = offset;
                        }
                    }
                    pieces.push(Task::Span(run, cursor, end, visible));
                    pending.extend(pieces.into_iter().rev());
                }
                Task::Span(run, start, end, visible) => {
                    if !visible {
                        if end > start {
                            gaps.push((
                                Span {
                                    run,
                                    start,
                                    len: end - start,
                                },
                                (length as u32).min(limit - 1),
                            ));
                        }
                        continue;
                    }
                    let mut cursor = start;
                    for &(low, high) in removed.get(&run).into_iter().flatten() {
                        if high <= cursor || low >= end {
                            continue;
                        }
                        if low > cursor {
                            push_span(&mut spans, &run, cursor, low.min(end), limit, &mut length);
                        }
                        let gap_start = cursor.max(low);
                        let gap_end = high.min(end);
                        if gap_end > gap_start {
                            gaps.push((
                                Span {
                                    run: run.clone(),
                                    start: gap_start,
                                    len: gap_end - gap_start,
                                },
                                (length as u32).min(limit - 1),
                            ));
                        }
                        cursor = cursor.max(high);
                        if cursor >= end {
                            break;
                        }
                    }
                    if cursor < end {
                        push_span(&mut spans, &run, cursor, end, limit, &mut length);
                    }
                }
            }
        }
        if changes
            .iter()
            .any(|(id, change)| matches!(change, Change::Insert { .. }) && !visited.contains(id))
        {
            return Err("axis insertion does not attach to the base".into());
        }
        if length != u64::from(limit) {
            return Err("axis does not cover the grid".into());
        }
        let mut known = BTreeMap::from([("base".into(), maximum_base)]);
        for (id, change) in changes {
            if let Change::Insert { count, .. } = change {
                known.insert(id.clone(), u64::from(*count));
            }
        }
        let insertions = changes
            .iter()
            .filter_map(|(id, change)| match change {
                Change::Insert {
                    before,
                    after,
                    count,
                } => Some((id.clone(), (before.clone(), after.clone(), *count))),
                _ => None,
            })
            .collect();
        let axis = Self {
            spans,
            limit,
            gaps,
            known,
            insertions,
        };
        for (_, after, _) in axis.insertions.values() {
            if let Some(after) = after {
                axis.validate_point(after)?;
            }
        }
        Ok(axis)
    }

    pub fn first_changed(&self) -> Option<u32> {
        let mut index = 0_u64;
        for span in &self.spans {
            if span.run != "base" || span.start != index {
                return Some(index as u32);
            }
            index += span.len;
        }
        None
    }

    pub fn at(&self, index: u32) -> Result<Point, String> {
        if index >= self.limit {
            return Err("axis index out of bounds".into());
        }
        let mut remaining = u64::from(index);
        for span in &self.spans {
            if remaining < span.len {
                return Ok(Point {
                    run: span.run.clone(),
                    offset: span.start + remaining,
                });
            }
            remaining -= span.len;
        }
        Err("axis index does not resolve".into())
    }

    pub fn index(&self, point: &Point) -> Option<u32> {
        self.index_in(&point.run, point.offset)
    }

    pub fn index_in(&self, run: &str, offset: u64) -> Option<u32> {
        let mut index = 0;
        for span in &self.spans {
            if span.run == run && offset >= span.start && offset < span.start + span.len {
                return u32::try_from(index + offset - span.start).ok();
            }
            index += span.len;
        }
        None
    }

    pub fn validate_point(&self, point: &Point) -> Result<(), String> {
        let end = self
            .known
            .get(&point.run)
            .ok_or_else(|| format!("pending axis run {}", point.run))?;
        if point.offset >= *end {
            return Err("axis point is outside its source run".into());
        }
        Ok(())
    }

    pub fn validate_spans(&self, spans: &[Span]) -> Result<(), String> {
        let mut count = 0_u64;
        for span in spans {
            let end = self
                .known
                .get(&span.run)
                .ok_or_else(|| format!("pending axis run {}", span.run))?;
            if span.len == 0
                || span
                    .start
                    .checked_add(span.len)
                    .is_none_or(|last| last > *end)
            {
                return Err("range span is outside its source run".into());
            }
            count = count.checked_add(span.len).ok_or("range span overflow")?;
        }
        if count > u64::from(self.limit) {
            return Err("range exceeds axis bounds".into());
        }
        Ok(())
    }

    pub fn clip(&self, point: &Point) -> Option<u32> {
        self.index(point)
            .or_else(|| {
                self.gaps.iter().find_map(|(span, index)| {
                    (span.run == point.run
                        && point.offset >= span.start
                        && point.offset < span.start + span.len)
                        .then_some(*index)
                })
            })
            .or_else(|| self.validate_point(point).ok().map(|_| self.limit - 1))
    }

    pub fn range(&self, start: u32, end: u32) -> Result<Vec<Span>, String> {
        if start > end || end >= self.limit {
            return Err("axis range out of bounds".into());
        }
        let mut index = 0;
        let mut result = Vec::new();
        for span in &self.spans {
            let low = u64::from(start).max(index);
            let high = (u64::from(end) + 1).min(index + span.len);
            if high > low {
                result.push(Span {
                    run: span.run.clone(),
                    start: span.start + low - index,
                    len: high - low,
                });
            }
            index += span.len;
            if index > u64::from(end) {
                break;
            }
        }
        Ok(result)
    }

    pub fn resolve_range(&self, target: &[Span]) -> Option<(u32, u32)> {
        // An insertion joins a range only when both neighbors observed by its
        // author already belonged to it. Keep that membership when those
        // neighbors are later deleted or their insertion is undone.
        let mut members = target.to_vec();
        let mut admitted = BTreeSet::new();
        loop {
            let mut changed = false;
            for (id, (before, after, count)) in &self.insertions {
                let contains = |point: &Point| {
                    members.iter().any(|span| {
                        span.run == point.run
                            && point.offset >= span.start
                            && point.offset < span.start + span.len
                    })
                };
                if !admitted.contains(id)
                    && after.as_ref().is_some_and(contains)
                    && contains(before)
                {
                    members.push(Span {
                        run: id.clone(),
                        start: 0,
                        len: u64::from(*count),
                    });
                    admitted.insert(id.clone());
                    changed = true;
                }
            }
            if !changed {
                break;
            }
        }
        let mut first = None;
        let mut last = 0;
        let mut index = 0;
        for span in &self.spans {
            for wanted in &members {
                if span.run != wanted.run {
                    continue;
                }
                let low = span.start.max(wanted.start);
                let high = (span.start + span.len).min(wanted.start + wanted.len);
                if high > low {
                    let from = (index + low - span.start) as u32;
                    let to = (index + high - span.start - 1) as u32;
                    first = Some(first.map_or(from, |old: u32| old.min(from)));
                    last = last.max(to);
                }
            }
            index += span.len;
        }
        first.map(|first| (first, last))
    }
}

fn push_span(spans: &mut Vec<Span>, run: &str, start: u64, end: u64, limit: u32, length: &mut u64) {
    let len = (end - start).min(u64::from(limit) - *length);
    if len == 0 {
        return;
    }
    if let Some(last) = spans.last_mut()
        && last.run == run
        && last.start + last.len == start
    {
        last.len += len;
    } else {
        spans.push(Span {
            run: run.to_owned(),
            start,
            len,
        });
    }
    *length += len;
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn sparse_axis_insert_delete_and_undo_preserve_identity() {
        let base = Axis::base(1_048_576);
        let mut changes = BTreeMap::from([
            (
                "a".into(),
                Change::Insert {
                    before: base.at(1).unwrap(),
                    after: Some(base.at(0).unwrap()),
                    count: 1,
                },
            ),
            (
                "b".into(),
                Change::Insert {
                    before: base.at(1).unwrap(),
                    after: Some(base.at(0).unwrap()),
                    count: 1,
                },
            ),
        ]);
        let mut active = BTreeSet::from(["a".into(), "b".into()]);
        let axis = Axis::project(1_048_576, &changes, &active).unwrap();
        assert_eq!(axis.spans.len(), 4);
        assert_eq!(axis.index(&base.at(1).unwrap()), Some(3));
        assert_eq!(axis.at(1).unwrap().run, "a");
        changes.insert(
            "delete".into(),
            Change::Delete {
                spans: axis.range(3, 3).unwrap(),
            },
        );
        active.insert("delete".into());
        assert_eq!(
            Axis::project(1_048_576, &changes, &active)
                .unwrap()
                .index(&base.at(1).unwrap()),
            None
        );
        active.remove("delete");
        assert_eq!(
            Axis::project(1_048_576, &changes, &active)
                .unwrap()
                .index(&base.at(1).unwrap()),
            Some(3)
        );
        changes.insert(
            "child".into(),
            Change::Insert {
                before: axis.at(1).unwrap(),
                after: Some(axis.at(0).unwrap()),
                count: 1,
            },
        );
        active.insert("child".into());
        active.remove("a");
        let undone = Axis::project(1_048_576, &changes, &active).unwrap();
        assert_eq!(undone.at(1).unwrap().run, "child");
        assert_eq!(undone.index(&base.at(1).unwrap()), Some(3));
    }
}
