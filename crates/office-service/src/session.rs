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
