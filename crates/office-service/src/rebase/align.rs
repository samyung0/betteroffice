//! A text's units and their alignment with another text (office-rebase.ts
//! `units`, `myers`, `align`, `pointAt`, `past`).

use yrs::block::ItemContent;
use yrs::iter::BlockIter;
use yrs::types::text::YChange;
use yrs::{Any, Map, Out, ReadTxn, Text, TextRef};

use super::fail;
use crate::error::Result;

/// One position of a text: a UTF-16 code unit, or an embed by its kind.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub(crate) enum Key {
    Unit(u16),
    /// `\0${kind}`: a map embed's `_kind` (as `String(...)`), or `embed`.
    Embed(String),
    /// A comment's reference field (`COMMENT_REFERENCE`).
    Reference,
}

pub(crate) const PILCROW: &str = "pilcrow";

/// A text's units, one per position, with the Yjs id of each.
#[derive(Clone, Debug, Default)]
pub(crate) struct Units {
    pub keys: Vec<Key>,
    pub ids: Vec<(u64, u32)>,
}

/// `String(value)` of a map entry as JS writes it.
pub(crate) fn any_string(value: Option<&Out>) -> String {
    match value {
        None => "undefined".into(),
        Some(Out::Any(any)) => match any {
            Any::Null => "null".into(),
            Any::Undefined => "undefined".into(),
            Any::Bool(value) => value.to_string(),
            Any::Number(value) => crate::js::number_to_string(*value),
            Any::BigInt(value) => value.to_string(),
            Any::String(text) => text.to_string(),
            Any::Buffer(_) | Any::Array(_) | Any::Map(_) => match any {
                Any::Array(items) => items
                    .iter()
                    .map(|item| match item {
                        Any::Null | Any::Undefined => String::new(),
                        item => any_string(Some(&Out::Any(item.clone()))),
                    })
                    .collect::<Vec<_>>()
                    .join(","),
                _ => "[object Object]".into(),
            },
        },
        Some(_) => "[object Object]".into(),
    }
}

/// The key an embedded value occupies.
pub(crate) fn embed_key<T: ReadTxn>(txn: &T, value: &Out) -> Key {
    match value {
        Out::YMap(map) => {
            let model = map.get(txn, "modelKind");
            if matches!(&model, Some(Out::Any(Any::String(kind))) if &**kind == "commentReference")
            {
                Key::Reference
            } else {
                Key::Embed(any_string(map.get(txn, "_kind").as_ref()))
            }
        }
        _ => Key::Embed("embed".into()),
    }
}

/// Visible items of a text: (offset, item units as values, first id).
pub(crate) fn items<T: ReadTxn>(txn: &T, text: &TextRef) -> Vec<(u32, Vec<Out>, (u64, u32))> {
    let _ = txn;
    let mut found = Vec::new();
    let mut offset = 0;
    for item in BlockIter::new(AsRef::<yrs::branch::Branch>::as_ref(text).start()) {
        if item.is_deleted() || !item.is_countable() {
            continue;
        }
        let id = (item.id().client.get(), item.id().clock);
        let units: Vec<Out> = match item.content() {
            ItemContent::String(_) => Vec::new(),
            content => content.get_content(),
        };
        let length = item.len();
        found.push((offset, units, id));
        offset += length;
    }
    found
}

pub(crate) fn units<T: ReadTxn>(txn: &T, text: &TextRef) -> Units {
    let mut found = Units::default();
    for item in BlockIter::new(AsRef::<yrs::branch::Branch>::as_ref(text).start()) {
        if item.is_deleted() || !item.is_countable() {
            continue;
        }
        let (client, clock) = (item.id().client.get(), item.id().clock);
        match item.content() {
            ItemContent::String(text) => {
                for (index, unit) in text.as_str().encode_utf16().enumerate() {
                    found.keys.push(Key::Unit(unit));
                    found.ids.push((client, clock + index as u32));
                }
            }
            content => {
                for (index, value) in content.get_content().iter().enumerate() {
                    found.keys.push(embed_key(txn, value));
                    found.ids.push((client, clock + index as u32));
                }
            }
        }
    }
    found
}

/// Embedded maps of a text with their positions.
pub(crate) fn embeds<T: ReadTxn>(txn: &T, text: &TextRef) -> Vec<(u32, yrs::MapRef)> {
    items(txn, text)
        .into_iter()
        .filter_map(|(offset, units, _)| match units.into_iter().next() {
            Some(Out::YMap(map)) => Some((offset, map)),
            _ => None,
        })
        .collect()
}

/// The embed at `offset` of a text.
pub(crate) fn child_at_offset<T: ReadTxn>(txn: &T, text: &TextRef, offset: u32) -> Option<Out> {
    let mut start = 0;
    for item in BlockIter::new(AsRef::<yrs::branch::Branch>::as_ref(text).start()) {
        if item.is_deleted() || !item.is_countable() {
            continue;
        }
        let length = item.len();
        if offset < start + length {
            return match item.content() {
                ItemContent::String(_) => None,
                content => content
                    .get_content()
                    .into_iter()
                    .nth((offset - start) as usize),
            };
        }
        start += length;
    }
    let _ = txn;
    None
}

/// Each unit's attributes (one map per run of them), as `unitAttributes`.
pub(crate) fn unit_attributes<T: ReadTxn>(
    txn: &T,
    text: &TextRef,
) -> Vec<std::rc::Rc<std::collections::HashMap<std::sync::Arc<str>, Any>>> {
    let mut out = Vec::new();
    for diff in text.diff(txn, YChange::identity) {
        let attributes = std::rc::Rc::new(diff.attributes.map(|attrs| *attrs).unwrap_or_default());
        let length = match &diff.insert {
            Out::Any(Any::String(text)) => text.encode_utf16().count(),
            _ => 1,
        };
        for _ in 0..length {
            out.push(attributes.clone());
        }
    }
    out
}

/// An edit script longer than this means the export did not reproduce the text.
const MAX_DIFFERENCES: i64 = 1000;

/// Where each captured unit stands in the seed's text, or -1.
#[derive(Clone, Debug)]
pub(crate) struct Alignment {
    pub from: Units,
    pub to: Units,
    pub map: Vec<i64>,
    /// Seed units some captured unit maps to.
    pub matched: Vec<bool>,
}

pub(crate) fn align(from: Units, to: Units) -> Result<Alignment> {
    let map = myers(&from.keys, &to.keys)?;
    let mut matched = vec![false; to.keys.len()];
    for &target in &map {
        if target >= 0 {
            matched[target as usize] = true;
        }
    }
    Ok(Alignment {
        from,
        to,
        map,
        matched,
    })
}

fn myers(a: &[Key], b: &[Key]) -> Result<Vec<i64>> {
    let mut map = vec![-1_i64; a.len()];
    let mut start = 0;
    while start < a.len() && start < b.len() && a[start] == b[start] {
        map[start] = start as i64;
        start += 1;
    }
    let (mut end_a, mut end_b) = (a.len(), b.len());
    while end_a > start && end_b > start && a[end_a - 1] == b[end_b - 1] {
        end_a -= 1;
        end_b -= 1;
        map[end_a] = end_b as i64;
    }
    let (n, m) = ((end_a - start) as i64, (end_b - start) as i64);
    if n == 0 || m == 0 {
        return Ok(map);
    }
    let max = (n + m).min(MAX_DIFFERENCES);
    let offset = max + 1;
    let mut v = vec![0_i64; (2 * max + 3) as usize];
    let mut trace: Vec<Vec<i64>> = Vec::new();
    let at = |v: &Vec<i64>, k: i64| v[(offset + k) as usize];
    for d in 0..=max {
        let mut k = -d;
        while k <= d {
            let mut x = if k == -d || (k != d && at(&v, k - 1) < at(&v, k + 1)) {
                at(&v, k + 1)
            } else {
                at(&v, k - 1) + 1
            };
            let mut y = x - k;
            while x < n && y < m && a[start + x as usize] == b[start + y as usize] {
                x += 1;
                y += 1;
            }
            v[(offset + k) as usize] = x;
            if x < n || y < m {
                k += 2;
                continue;
            }
            trace.push(v.clone());
            // Walk the script back, matching every diagonal step.
            let mut e = d;
            while e > 0 {
                let previous = &trace[(e - 1) as usize];
                let key = x - y;
                let back = if key == -e
                    || (key != e
                        && previous[(offset + key - 1) as usize]
                            < previous[(offset + key + 1) as usize])
                {
                    key + 1
                } else {
                    key - 1
                };
                let px = previous[(offset + back) as usize];
                let py = px - back;
                while x > px && y > py {
                    x -= 1;
                    y -= 1;
                    map[start + x as usize] = start as i64 + y;
                }
                x = px;
                y = py;
                e -= 1;
            }
            while x > 0 && y > 0 {
                x -= 1;
                y -= 1;
                map[start + x as usize] = start as i64 + y;
            }
            return Ok(map);
        }
        trace.push(v.clone());
    }
    Err(fail("the export wrote a text too differently"))
}

/// The seed unit at or after `s` that is not a comment reference only the seed holds.
pub(crate) fn past(f: &Alignment, mut s: i64) -> i64 {
    while (s as usize) < f.to.keys.len()
        && f.to.keys[s as usize] == Key::Reference
        && !f.matched[s as usize]
    {
        s += 1;
    }
    s
}

/// The seed position of an insertion before captured position `c`.
pub(crate) fn point_at(f: &Alignment, c: usize) -> Option<i64> {
    let left = if c == 0 { -1 } else { f.map[c - 1] };
    let right = if c == f.map.len() {
        f.to.keys.len() as i64
    } else {
        f.map[c]
    };
    if left >= 0 || c == 0 {
        if right >= 0 || c == f.map.len() {
            return (right == past(f, left + 1)).then_some(right);
        }
        return Some(left + 1);
    }
    (right >= 0).then_some(right)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn keys(text: &str) -> Vec<Key> {
        text.encode_utf16().map(Key::Unit).collect()
    }

    #[test]
    fn myers_maps_the_kept_units() {
        assert_eq!(
            myers(&keys("abcd"), &keys("abxd")).unwrap(),
            vec![0, 1, -1, 3]
        );
        assert_eq!(myers(&keys("ab"), &keys("xaby")).unwrap(), vec![1, 2]);
    }
}
