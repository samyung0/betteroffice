//! yrs/comments.ts: shared comment keys to the numeric ids a save writes.

use std::collections::{HashMap, HashSet};

use crate::error::{Error, Result};
use crate::js::utf16_compare;
use crate::jsv::{Obj, V};

pub(crate) fn comment_shared_id(comment: &V) -> String {
    comment
        .get("sharedId")
        .or_else(|| V::from(comment.get("id").to_js_string()))
        .to_js_string()
}

/// A decimal key within int32 keeps its number.
fn numeric_key(key: &str) -> Option<u32> {
    if key.is_empty() || !key.bytes().all(|byte| byte.is_ascii_digit()) {
        return None;
    }
    key.parse::<f64>()
        .ok()
        .filter(|value| *value <= 2_147_483_647.0)
        .map(|value| value as u32)
}

pub(crate) fn comment_numeric_id(id: &str) -> u32 {
    if let Some(number) = numeric_key(id) {
        return number;
    }
    let mut hash: u32 = 2_166_136_261;
    for byte in id.bytes() {
        hash = (hash ^ byte as u32).wrapping_mul(16_777_619);
    }
    hash >> 1
}

/// The collision-free OOXML numeric id a save writes for each shared comment key.
pub(crate) fn comment_ooxml_ids(keys: &[String]) -> HashMap<String, u32> {
    let mut sorted = keys.to_vec();
    sorted.sort_by(|a, b| utf16_compare(a, b));
    let mut ids = HashMap::new();
    let mut used = HashSet::new();
    for key in &sorted {
        if let Some(number) = numeric_key(key) {
            ids.insert(key.clone(), number);
            used.insert(number);
        }
    }
    for key in &sorted {
        if ids.contains_key(key) {
            continue;
        }
        let mut id = comment_numeric_id(key);
        while used.contains(&id) {
            id = (id + 1) & 0x7fff_ffff;
        }
        ids.insert(key.clone(), id);
        used.insert(id);
    }
    ids
}

/// Projects stable shared keys into collision-free OOXML numeric ids
/// (`projectYrsComments`).
pub(crate) fn project_comments(mut records: Vec<V>, base: &[V]) -> Result<Vec<V>> {
    let id_of = |record: &V| record.get("id").to_js_string();
    records.sort_by(|a, b| utf16_compare(&id_of(a), &id_of(b)));
    let ids = comment_ooxml_ids(&records.iter().map(id_of).collect::<Vec<_>>());
    let mut originals: HashMap<String, V> = HashMap::new();
    for comment in base {
        originals.insert(comment_shared_id(comment), comment.clone());
    }
    let mut projected = Vec::new();
    for record in records {
        let parent = record.get("parentId");
        if !parent.same(&V::Null) && !ids.contains_key(&parent.to_js_string()) {
            continue;
        }
        let key = id_of(&record);
        let original = originals.get(&key);
        let body = record.get("body");
        if !body.same(&V::Null) && body.arr().is_none() {
            return Err(Error::engine("Invalid shared comment body"));
        }
        let content = if body.same(&V::Null) {
            original
                .map(|original| original.get("content"))
                .unwrap_or_default()
                .or_else(|| V::array(Vec::new()))
        } else {
            body.clone()
        };
        let from_original = body.same(&V::Null) && original.is_some();
        let comment = original
            .and_then(V::obj)
            .map(Obj::spread)
            .unwrap_or_default();
        comment.set("id", V::Num(ids[&key] as f64));
        comment.set("sharedId", V::from(key));
        comment.set(
            "author",
            if from_original {
                original.map(|o| o.get("author")).unwrap_or_default()
            } else {
                record.get("author")
            },
        );
        comment.set(
            "date",
            if from_original {
                original.map(|o| o.get("date")).unwrap_or_default()
            } else {
                record.get("date")
            },
        );
        comment.set("content", content);
        comment.set("blockContent", V::Undef);
        comment.set(
            "parentId",
            if parent.same(&V::Null) {
                V::Undef
            } else {
                ids.get(&parent.to_js_string())
                    .map(|id| V::Num(*id as f64))
                    .unwrap_or_default()
            },
        );
        comment.set("done", record.get("done"));
        comment.set(
            "status",
            V::str(if record.get("done").truthy() {
                "resolved"
            } else {
                "active"
            }),
        );
        projected.push(V::Obj(comment));
    }
    Ok(projected)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn keys_keep_or_hash_their_numbers() {
        let ids = comment_ooxml_ids(&["3".into(), "client:7".into(), "12".into()]);
        assert_eq!(ids["3"], 3);
        assert_eq!(ids["12"], 12);
        assert_eq!(ids["client:7"], comment_numeric_id("client:7"));
    }
}
