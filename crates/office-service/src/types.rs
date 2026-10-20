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

/// An agent edit. XLSX `sheet` is a sheet id from inspect (`sheet:0`) or a
/// sheet name; cell values are inputs as typed in the editor (`=` formulas,
/// `""` empty).
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
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
    /// `range` is `A1` or `A1:C3`; both grids have its shape, row by row.
    #[serde(rename_all = "camelCase")]
    SetRange {
        sheet: String,
        range: String,
        expected: Vec<Vec<String>>,
        values: Vec<Vec<String>>,
    },
    #[serde(rename_all = "camelCase")]
    FormatCells {
        sheet: String,
        range: String,
        style: CellStyle,
    },
    /// Placed after the sheet `after` names, else last.
    #[serde(rename_all = "camelCase")]
    AddSheet {
        name: String,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        after: Option<String>,
    },
    #[serde(rename_all = "camelCase")]
    RenameSheet {
        sheet: String,
        expected_name: String,
        name: String,
    },
    #[serde(rename_all = "camelCase")]
    DeleteSheet {
        sheet: String,
        expected_name: String,
    },
    /// `at` is the 1-based row the first new row takes.
    #[serde(rename_all = "camelCase")]
    InsertRows { sheet: String, at: u32, count: u32 },
    /// `expected` is the band's inputs as TSV (rows from `at`, cells from
    /// column A), required; `""` expects an empty band.
    #[serde(rename_all = "camelCase")]
    DeleteRows {
        sheet: String,
        at: u32,
        count: u32,
        expected: String,
    },
    /// `at` is the column letter the first new column takes.
    #[serde(rename_all = "camelCase")]
    InsertColumns {
        sheet: String,
        at: String,
        count: u32,
    },
    /// `expected` is the band's inputs as TSV (rows from row 1, cells from
    /// column `at`), required; `""` expects an empty band.
    #[serde(rename_all = "camelCase")]
    DeleteColumns {
        sheet: String,
        at: String,
        count: u32,
        expected: String,
    },
}

/// What `format_cells` sets: the editor's style patch and a number format.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CellStyle {
    #[serde(flatten)]
    pub patch: betteroffice_xlsx::StylePatch,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub number_format: Option<betteroffice_xlsx::NumberFormatMutation>,
}

/// A scoped read. XLSX: no `range` lists the sheets; `sheet` and `range`
/// read that range's cells.
#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct InspectQuery {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub range: Option<String>,
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

#[derive(Clone, Debug, PartialEq)]
pub struct CommandResult {
    pub state: Vec<u8>,
    /// In undo order; `None` when a command has no inverse, so the edit has
    /// no Undo.
    pub inverse: Option<Vec<Command>>,
    pub targets: Vec<Target>,
    /// For the edit's tool result: the characters an Office file cannot store
    /// that the agent's text lost before the edit applied (Epo 2026-10-15).
    pub notice: Option<String>,
}

/// A file [`crate::create`] built from a starter.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Created {
    pub bytes: Vec<u8>,
    pub notice: Option<String>,
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
