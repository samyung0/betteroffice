use crate::common::sha256_hex;
use crate::js::{J, canonical};
use crate::types::{Asset, BaselineEntry, Command, EffectKind, NetEffect, ObjectRef, Operation};

/// office-checkpoint.ts `Entry`: a baseline entry with its asset.
#[derive(Clone, Debug)]
pub(crate) struct Item {
    pub id: String,
    pub kind: EffectKind,
    pub label: String,
    pub value: String,
    pub position: String,
    pub asset_ref: Option<ObjectRef>,
    pub asset: Option<Asset>,
}

impl Item {
    pub fn text(
        id: impl Into<String>,
        label: impl Into<String>,
        value: impl Into<String>,
        position: impl Into<String>,
    ) -> Self {
        Self {
            id: id.into(),
            kind: EffectKind::Text,
            label: label.into(),
            value: value.into(),
            position: position.into(),
            asset_ref: None,
            asset: None,
        }
    }

    /// office-checkpoint.ts `visual`: the value is the canonical JSON.
    pub fn visual(
        id: impl Into<String>,
        label: impl Into<String>,
        value: &J,
        position: impl Into<String>,
    ) -> Self {
        Self {
            kind: EffectKind::Visual,
            ..Self::text(id, label, canonical(value), position)
        }
    }

    /// [`Item::visual`] of a JS value.
    pub fn visual_v(
        id: impl Into<String>,
        label: impl Into<String>,
        value: &crate::jsv::V,
        position: impl Into<String>,
    ) -> Self {
        Self {
            kind: EffectKind::Visual,
            ..Self::text(id, label, crate::jsv::canonical(value), position)
        }
    }

    /// The `officeBaseline` entry: visual values hashed, assets by hash.
    pub fn baseline(self) -> BaselineEntry {
        BaselineEntry {
            value: if self.kind == EffectKind::Visual {
                sha256_hex(self.value.as_bytes())
            } else {
                self.value
            },
            image_sha256: self.asset.map(|asset| asset.sha256),
            id: self.id,
            kind: self.kind,
            label: self.label,
            position: self.position,
            asset_ref: self.asset_ref,
        }
    }
}

/// One applied command: the target id and the command that undoes it.
pub(crate) struct Applied {
    pub id: String,
    pub inverse: Command,
}

/// office-checkpoint.ts `compareBaselines`.
pub fn compare_baselines(from: &[BaselineEntry], to: &[BaselineEntry]) -> Vec<NetEffect> {
    use std::collections::HashMap;
    fn index(entries: &[BaselineEntry]) -> HashMap<&str, &BaselineEntry> {
        let mut map = HashMap::with_capacity(entries.len());
        for entry in entries {
            map.insert(entry.id.as_str(), entry);
        }
        map
    }
    let (before, after) = (index(from), index(to));
    let mut ids: Vec<&str> = before.keys().chain(after.keys()).copied().collect();
    ids.sort_by(|a, b| crate::js::utf16_compare(a, b));
    ids.dedup();
    let mut effects = Vec::new();
    for id in ids {
        let (old, next) = (before.get(id).copied(), after.get(id).copied());
        if let (Some(old), Some(next)) = (old, next)
            && old.value == next.value
            && old.position == next.position
        {
            continue;
        }
        let Some(entry) = next.or(old) else { continue };
        let operation = match (old, next) {
            (None, _) => Operation::Add,
            (_, None) => Operation::Remove,
            (Some(old), Some(next)) if old.value == next.value => Operation::Move,
            _ => Operation::Replace,
        };
        let text = entry.kind == EffectKind::Text;
        effects.push(NetEffect {
            id: id.to_owned(),
            kind: entry.kind,
            operation,
            label: entry.label.clone(),
            before: old.filter(|_| text).map(|old| old.value.clone()),
            after: next.filter(|_| text).map(|next| next.value.clone()),
            asset_ref: next.and_then(|next| next.asset_ref.clone()),
            image_sha256: next.and_then(|next| next.image_sha256.clone()),
        });
    }
    effects
}

/// `compare_baselines` without the `move` of an entry an edit merely
/// shifted (#6): its position is an index in a story, slide or shape
/// (`<container>:<n>`), the same container in both, and its order among the
/// entries both baselines hold in that container is unchanged. A paragraph
/// typed into or inserted before others moves every later one's index or
/// offset; that is no change of its own.
pub fn unshifted_effects(from: &[BaselineEntry], to: &[BaselineEntry]) -> Vec<NetEffect> {
    use std::collections::HashMap;
    fn container(position: &str) -> Option<&str> {
        let (container, index) = position.rsplit_once(':')?;
        (!index.is_empty() && index.bytes().all(|byte| byte.is_ascii_digit())).then_some(container)
    }
    let (before, after): (HashMap<&str, &BaselineEntry>, HashMap<&str, &BaselineEntry>) = (
        from.iter()
            .map(|entry| (entry.id.as_str(), entry))
            .collect(),
        to.iter().map(|entry| (entry.id.as_str(), entry)).collect(),
    );
    // Each entry's rank among the entries of its container both hold.
    let ranks = |entries: &[BaselineEntry], other: &HashMap<&str, &BaselineEntry>| {
        let mut counts: HashMap<&str, usize> = HashMap::new();
        let mut ranks: HashMap<String, usize> = HashMap::new();
        for entry in entries {
            let Some(held) = other.get(entry.id.as_str()) else {
                continue;
            };
            let Some(place) = container(&entry.position) else {
                continue;
            };
            if container(&held.position) != Some(place) {
                continue;
            }
            let count = counts.entry(place).or_default();
            ranks.insert(entry.id.clone(), *count);
            *count += 1;
        }
        ranks
    };
    let (old, new) = (ranks(from, &after), ranks(to, &before));
    compare_baselines(from, to)
        .into_iter()
        .filter(|effect| {
            effect.operation != Operation::Move
                || !old
                    .get(&effect.id)
                    .is_some_and(|rank| new.get(&effect.id) == Some(rank))
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn entry(id: &str, value: &str, position: &str) -> BaselineEntry {
        BaselineEntry {
            id: id.into(),
            kind: EffectKind::Text,
            label: id.into(),
            value: value.into(),
            position: position.into(),
            asset_ref: None,
            image_sha256: None,
        }
    }

    fn operations(effects: Vec<NetEffect>) -> Vec<(String, Operation)> {
        effects
            .into_iter()
            .map(|effect| (effect.id, effect.operation))
            .collect()
    }

    /// A paragraph inserted before others and text typed before a field
    /// shift their indexes and offsets: only the insertion is an effect.
    #[test]
    fn shifted_entries_are_no_moves() {
        let from = [
            entry("p1", "a", "body:0"),
            entry("field", "2026", "body:5"),
            entry("p2", "b", "body:1"),
        ];
        let to = [
            entry("p0", "new", "body:0"),
            entry("p1", "a", "body:1"),
            entry("field", "2026", "body:9"),
            entry("p2", "b", "body:2"),
        ];
        assert_eq!(
            operations(unshifted_effects(&from, &to)),
            [("p0".into(), Operation::Add)]
        );
        assert_eq!(compare_baselines(&from, &to).len(), 4);
    }

    /// Entries whose order changed, or that left their story, still move.
    #[test]
    fn reordered_entries_move() {
        let from = [
            entry("p1", "a", "body:0"),
            entry("p2", "b", "body:1"),
            entry("p3", "c", "body:2"),
        ];
        let to = [
            entry("p2", "b", "body:0"),
            entry("p1", "a", "body:1"),
            entry("p3", "c", "hf:rId1:0"),
        ];
        assert_eq!(
            operations(unshifted_effects(&from, &to)),
            operations(compare_baselines(&from, &to))
        );
    }
}
