//! The paragraph ID plan a save carries: IDs for source paragraphs outside
//! the edited stories, applied by occurrence to serialized parts, and parts
//! written as their source bytes with only paragraph IDs patched in place.

use std::collections::{BTreeMap, HashMap};

use serde::{Deserialize, Serialize};
use serde_json::Value;

use crate::paragraph_identity::{paragraph_occurrences, parse_paragraph_id, patch_paragraph_ids};
use crate::xml::ParseError;

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct S13ParagraphIds {
    /// IDs of source paragraphs by part and occurrence, applied to the parts
    /// this save serializes.
    #[serde(default)]
    pub assignments: Vec<S13ParagraphId>,
    /// Parts written as their source bytes with only paragraph IDs patched.
    #[serde(default)]
    pub patched_parts: Vec<S13PatchedPart>,
    /// Story parts written as their source bytes with only some paragraphs re-serialized.
    #[serde(default)]
    pub spliced_parts: Vec<S13SplicedPart>,
}

/// A story part whose model paragraphs each come from one source `w:p`, in source order, or are
/// new: it is written as its source XML with only the paragraphs that changed, or whose written
/// XML no longer agrees with their source on comments, revisions, notes or relationships,
/// replaced, the removed ones dropped and the new ones inserted.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct S13SplicedPart {
    pub part: String,
    /// Lowercase hex SHA-256 of the source part the ordinals address.
    pub sha256: String,
    /// The `sourceOrdinal` of every model paragraph written from the part.
    #[serde(default)]
    pub paragraphs: Vec<u32>,
    /// Of those, the ones whose content changed since the source was opened.
    #[serde(default)]
    pub changed: Vec<u32>,
    /// Where each model paragraph written without a `sourceOrdinal` goes, in writing order.
    #[serde(default)]
    pub inserted: Vec<S13SpliceAnchor>,
    /// The source paragraphs the model no longer holds.
    #[serde(default)]
    pub removed: Vec<u32>,
}

/// Where a spliced part takes a paragraph its source lacks: right before or right after the
/// source paragraph with this ordinal.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum S13SpliceAnchor {
    Before(u32),
    After(u32),
}

impl S13SpliceAnchor {
    pub(crate) fn ordinal(self) -> u32 {
        match self {
            Self::Before(ordinal) | Self::After(ordinal) => ordinal,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct S13ParagraphId {
    pub part: String,
    pub ordinal: u32,
    pub para_id: String,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct S13PatchedPart {
    pub part: String,
    /// `[ordinal, paraId]` pairs.
    #[serde(default)]
    pub para_ids: Vec<(u32, String)>,
}

pub(crate) const COMMENTS_PART: &str = "word/comments.xml";

impl S13ParagraphIds {
    pub(crate) fn validate(&self) -> Result<(), ParseError> {
        let values = self.assignments.iter().map(|id| id.para_id.as_str()).chain(
            self.patched_parts
                .iter()
                .flat_map(|part| part.para_ids.iter().map(|(_, id)| id.as_str())),
        );
        for value in values {
            if parse_paragraph_id(value).is_none() {
                return Err(ParseError::Canonical(format!(
                    "S13 package save: {value:?} is not a paragraph ID"
                )));
            }
        }
        Ok(())
    }

    /// Assignments grouped by part.
    pub(crate) fn assignments_by_part(&self) -> HashMap<&str, BTreeMap<u32, String>> {
        let mut parts: HashMap<&str, BTreeMap<u32, String>> = HashMap::new();
        for id in &self.assignments {
            parts
                .entry(id.part.as_str())
                .or_default()
                .insert(id.ordinal, id.para_id.clone());
        }
        parts
    }
}

/// Sets the assigned ID on every model paragraph whose `sourceOrdinal` it
/// names, and patches raw XML carried from the source part at its occurrences.
pub(crate) fn apply_assignments<T: Serialize + for<'de> Deserialize<'de>>(
    node: &mut T,
    ids: &BTreeMap<u32, String>,
) -> Result<(), ParseError> {
    let mut value =
        serde_json::to_value(&*node).map_err(|error| ParseError::Canonical(error.to_string()))?;
    assign(&mut value, ids);
    *node =
        serde_json::from_value(value).map_err(|error| ParseError::Canonical(error.to_string()))?;
    Ok(())
}

fn assign(value: &mut Value, ids: &BTreeMap<u32, String>) {
    match value {
        Value::Object(object) => {
            let ordinal = object
                .get("sourceOrdinal")
                .and_then(Value::as_u64)
                .and_then(|ordinal| u32::try_from(ordinal).ok());
            if let Some(ordinal) = ordinal {
                if object.get("type").and_then(Value::as_str) == Some("paragraph") {
                    if let Some(id) = ids.get(&ordinal) {
                        object.insert("paraId".to_owned(), Value::String(id.clone()));
                        object.remove("repeatedParaId");
                    }
                } else {
                    for key in ["xml", "verbatimXml"] {
                        if let Some(Value::String(xml)) = object.get_mut(key) {
                            patch_fragment(xml, ordinal, ids);
                        }
                    }
                }
            }
            for child in object.values_mut() {
                assign(child, ids);
            }
        }
        Value::Array(items) => {
            for item in items {
                assign(item, ids);
            }
        }
        _ => {}
    }
}

fn patch_fragment(xml: &mut String, base: u32, ids: &BTreeMap<u32, String>) {
    let Some(count) = paragraph_occurrences(xml).map(|occurrences| occurrences.len() as u32) else {
        return;
    };
    let local: BTreeMap<u32, String> = ids
        .range(base..base.saturating_add(count))
        .map(|(ordinal, id)| (ordinal - base, id.clone()))
        .collect();
    if !local.is_empty()
        && let Some(patched) = patch_paragraph_ids(xml, &local, false)
    {
        *xml = patched;
    }
}

/// The source part with only the planned paragraph IDs patched in place.
pub(crate) fn patch_part(original: &[u8], ids: &[(u32, String)]) -> Option<Vec<u8>> {
    let xml = std::str::from_utf8(original).ok()?;
    let patches: BTreeMap<u32, String> = ids.iter().cloned().collect();
    patch_paragraph_ids(xml, &patches, true).map(String::into_bytes)
}
