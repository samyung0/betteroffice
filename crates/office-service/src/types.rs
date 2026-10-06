use serde::{Deserialize, Serialize};

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum Format {
    Docx,
    Xlsx,
    Pptx,
}

impl Format {
    pub fn as_str(self) -> &'static str {
        match self {
            Format::Docx => "docx",
            Format::Xlsx => "xlsx",
            Format::Pptx => "pptx",
        }
    }
}

/// What a checkpoint call reads: the state over `base_sha256`'s package.
#[derive(Clone, Copy, Debug)]
pub struct Checkpoint<'a> {
    pub format: Format,
    pub schema_version: u32,
    pub base_sha256: &'a str,
    pub state: &'a [u8],
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum EffectKind {
    Text,
    Image,
    Visual,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum Operation {
    Add,
    Replace,
    Remove,
    Move,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum ObjectKind {
    Image,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ObjectRef {
    pub format: Format,
    pub kind: ObjectKind,
    pub id: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub story_id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub slide_id: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NetEffect {
    pub id: String,
    pub kind: EffectKind,
    pub operation: Operation,
    pub label: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub before: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub after: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub asset_ref: Option<ObjectRef>,
    #[serde(
        default,
        rename = "imageSHA256",
        skip_serializing_if = "Option::is_none"
    )]
    pub image_sha256: Option<String>,
}

/// One editable text entry: a DOCX or PPTX paragraph or an XLSX cell.
#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub struct Entry {
    pub id: String,
    pub label: String,
    pub value: String,
    pub position: String,
}

/// office-checkpoint.ts `OfficeBaselineEntry`.
#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct BaselineEntry {
    pub id: String,
    pub kind: EffectKind,
    pub label: String,
    pub value: String,
    pub position: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub asset_ref: Option<ObjectRef>,
    #[serde(
        default,
        rename = "imageSHA256",
        skip_serializing_if = "Option::is_none"
    )]
    pub image_sha256: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum Command {
    #[serde(rename_all = "camelCase")]
    ReplaceText {
        target_id: String,
        expected_text: String,
        text: String,
    },
    #[serde(rename_all = "camelCase")]
    SetCell {
        sheet: String,
        cell: String,
        expected_value: String,
        value: String,
    },
}

/// Yrs location of an edited target: root map name, nested keys, and the
/// UTF-16 range in a story.
#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub struct Target {
    pub id: String,
    pub path: Vec<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub range: Option<[u32; 2]>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CommandResult {
    pub state: Vec<u8>,
    pub inverse: Vec<Command>,
    pub targets: Vec<Target>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Asset {
    pub bytes: Vec<u8>,
    pub mime_type: String,
    pub sha256: String,
}

#[derive(Clone, Copy, Debug)]
pub struct Determinism<'a> {
    pub seed: &'a str,
    pub now: &'a str,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Rebased {
    pub state: Vec<u8>,
    pub effects: Vec<NetEffect>,
}
