//! Native checkpoint service over the Office engines: the port of
//! `shared/office-checkpoint.ts` and `shared/office-rebase.ts` with the same
//! inputs and outputs. Every call opens its own engine session on the calling
//! thread; only [`XlsxReplicas`] keeps state between calls.
//!
//! Calls recurse over document trees: run them on a thread with at least
//! 8 MiB of stack (the deepest golden and matrix call commits 0.45 MiB in a
//! release build). A stack overflow aborts the process; it is not a panic.

mod common;
mod docx;
#[cfg(test)]
mod effects_tests;
pub mod env;
mod error;
mod formula;
pub mod js;
pub mod jsv;
mod pptx;
mod rebase;
mod replicas;
mod session;
mod types;
pub mod wire;
mod xlsx;

use std::collections::{BTreeMap, HashSet};
use std::panic::{AssertUnwindSafe, catch_unwind};

pub use common::sha256_hex;
pub use error::{EditCode, Error, Result};
pub use replicas::{HEAP_PER_UNZIPPED_BYTE, REPLICA_IDLE, ReplicaStats, XlsxReplicas};
pub use session::{compare_baselines, unshifted_effects};
pub use types::*;

use common::{assert_checkpoint, removed_notice, storable_command};
use docx::DocxSession;
use pptx::PptxSession;
use session::{Applied, Item};
use xlsx::XlsxSession;

/// Top-level Yjs roots each engine's state may hold, Capy's contributor map
/// included. The PPTX and XLSX engines reject any other root.
pub fn document_roots(format: Format) -> &'static [&'static str] {
    match format {
        Format::Docx => &[
            "stories",
            "comments",
            "bookmarks",
            "__capy_pending_contributors",
        ],
        Format::Xlsx => &[
            "xlsx",
            "xlsx:cell-formats",
            "xlsx:sheet-order",
            "xlsx:sheets",
            "xlsx:axis-catalog",
            "xlsx:defined-names",
            "__capy_pending_contributors",
        ],
        Format::Pptx => &[
            "pptx:meta",
            "pptx:slide-order",
            "pptx:slides",
            "pptx:shapes",
            "pptx:stories",
            "pptx:comments",
            "__capy_pending_contributors",
        ],
    }
}

/// Runs an engine call, turning a panic into [`Error::Panic`].
fn guard<T>(call: impl FnOnce() -> Result<T>) -> Result<T> {
    let result = catch_unwind(AssertUnwindSafe(call));
    jsv::forget_interned();
    result.unwrap_or_else(|panic| {
        let message = panic
            .downcast_ref::<&str>()
            .map(|text| (*text).to_owned())
            .or_else(|| panic.downcast_ref::<String>().cloned())
            .unwrap_or_else(|| "unknown panic".into());
        Err(Error::Panic(message))
    })
}

/// One engine session of any format (office-checkpoint.ts `Session`).
enum Session {
    Docx(Box<DocxSession>),
    Xlsx(Box<XlsxSession>),
    Pptx(Box<PptxSession>),
}

impl Session {
    fn open(format: Format, base: &[u8], checkpoint: Option<Checkpoint>) -> Result<Self> {
        assert_checkpoint(format, base, checkpoint)?;
        let state = checkpoint.map(|checkpoint| checkpoint.state);
        Ok(match format {
            Format::Docx => Session::Docx(Box::new(DocxSession::open(base, state)?)),
            Format::Xlsx => Session::Xlsx(Box::new(XlsxSession::open(base, state)?)),
            Format::Pptx => Session::Pptx(Box::new(PptxSession::open(base, state)?)),
        })
    }

    fn state(&self) -> Vec<u8> {
        match self {
            Session::Docx(session) => session.state(),
            Session::Xlsx(session) => session.state(),
            Session::Pptx(session) => session.state(),
        }
    }

    fn entries(&mut self) -> Result<Vec<Item>> {
        match self {
            Session::Docx(session) => session.entries(),
            Session::Xlsx(session) => session.entries(),
            Session::Pptx(session) => session.entries(),
        }
    }

    fn editable(&mut self) -> Result<Vec<Entry>> {
        match self {
            Session::Docx(session) => session.editable(),
            Session::Xlsx(session) => session.editable(),
            Session::Pptx(session) => session.editable(),
        }
    }

    fn apply(&mut self, command: &Command) -> Result<Applied> {
        match self {
            Session::Docx(session) => session.apply(command),
            Session::Xlsx(session) => session.apply(command),
            Session::Pptx(session) => session.apply(command),
        }
    }

    /// After a call's commands (DOCX: the empty paragraphs its tables no
    /// longer need).
    fn finish(&mut self, commands: &[Command]) -> Result<()> {
        match self {
            Session::Docx(session) => session.finish_agent(commands),
            Session::Xlsx(_) | Session::Pptx(_) => Ok(()),
        }
    }

    fn locate(&mut self, ids: &[String]) -> Result<Vec<Target>> {
        match self {
            Session::Docx(session) => ids.iter().map(|id| session.locate(id)).collect(),
            Session::Xlsx(session) => session.locate_all(ids),
            Session::Pptx(session) => ids.iter().map(|id| session.locate(id)).collect(),
        }
    }

    fn inspect_scoped(&mut self, query: &InspectQuery) -> Result<serde_json::Value> {
        match self {
            Session::Xlsx(session) => session.inspect_scoped(query),
            Session::Docx(session) => session.inspect_scoped(query),
            Session::Pptx(_) => Err(Error::edit(
                EditCode::UnsupportedOperation,
                "scoped reads cover XLSX and DOCX files for now",
            )),
        }
    }

    /// Ends the session (a DOCX export frees it as it goes).
    fn export(self, determinism: Determinism) -> Result<Vec<u8>> {
        match self {
            Session::Docx(session) => session.export(determinism),
            Session::Xlsx(mut session) => session.export(determinism.now),
            Session::Pptx(session) => session.export(),
        }
    }
}

/// seed(base): the state a NULL row stands for. Byte-identical to the TS.
pub fn seed(format: Format, base: &[u8]) -> Result<Vec<u8>> {
    guard(|| {
        assert_checkpoint(format, base, None)?;
        match format {
            Format::Docx => docx::seed(base),
            Format::Xlsx => xlsx::seed(base),
            Format::Pptx => pptx::seed(base),
        }
    })
}

/// Semantic comparison data only; media is represented by its content hash.
pub fn baseline(base: &[u8], checkpoint: Checkpoint) -> Result<Vec<BaselineEntry>> {
    guard(|| baseline_unguarded(base, checkpoint))
}

pub(crate) fn baseline_unguarded(
    base: &[u8],
    checkpoint: Checkpoint,
) -> Result<Vec<BaselineEntry>> {
    let mut session = Session::open(checkpoint.format, base, Some(checkpoint))?;
    Ok(session.entries()?.into_iter().map(Item::baseline).collect())
}

/// DOCX and PPTX save effects (#6): the checkpoint against the indexed
/// baseline, without the `move` of entries an edit merely shifted
/// ([`unshifted_effects`]). The Capy save path compares the same baselines
/// with `compareBaselines`.
pub fn save_effects(
    base: &[u8],
    indexed: &[BaselineEntry],
    checkpoint: Checkpoint,
) -> Result<Vec<NetEffect>> {
    Ok(unshifted_effects(indexed, &baseline(base, checkpoint)?))
}

pub fn compare(base: &[u8], from: Checkpoint, to: Checkpoint) -> Result<Vec<NetEffect>> {
    if from.format != to.format {
        return Err(Error::engine("Cannot compare different Office formats"));
    }
    Ok(compare_baselines(
        &baseline(base, from)?,
        &baseline(base, to)?,
    ))
}

/// Pending XLSX effects against the base, read off the checkpoint's overrides
/// with a reader opened for this call ([`XlsxReplicas`] keeps one per room).
pub fn xlsx_pending_effects(base: &[u8], checkpoint: Checkpoint) -> Result<Vec<NetEffect>> {
    guard(|| {
        if checkpoint.format != Format::Xlsx {
            return Err(Error::Invalid("Expected an XLSX checkpoint".into()));
        }
        xlsx::check(base, checkpoint)?;
        xlsx::pending_effects(base, checkpoint, None)
    })
}

impl XlsxReplicas {
    /// [`xlsx_pending_effects`] off `room`'s replica.
    pub fn effects(
        &self,
        base: &[u8],
        checkpoint: Checkpoint,
        room: &str,
    ) -> Result<Vec<NetEffect>> {
        guard(|| {
            if checkpoint.format != Format::Xlsx {
                return Err(Error::Invalid("Expected an XLSX checkpoint".into()));
            }
            self.pending_effects(base, checkpoint, room)
        })
    }
}

/// Lands the edits saved after the capture on seed(export) and returns that
/// state with its effects against the export's derived baseline. A refusal is
/// [`Error::Rebase`]; the same inputs are always refused again.
pub fn rebase(
    base: &[u8],
    captured: Checkpoint,
    latest: Checkpoint,
    exported: &[u8],
) -> Result<Rebased> {
    guard(|| rebase::rebase(base, captured, latest, exported))
}

pub fn resolve_asset(base: &[u8], checkpoint: Checkpoint, object_ref: &ObjectRef) -> Result<Asset> {
    guard(|| {
        if object_ref.format != checkpoint.format {
            return Err(Error::engine("Image reference format mismatch"));
        }
        let mut session = Session::open(checkpoint.format, base, Some(checkpoint))?;
        let wanted = js::canonical(&ref_json(object_ref));
        session
            .entries()?
            .into_iter()
            .find(|entry| {
                entry
                    .asset_ref
                    .as_ref()
                    .is_some_and(|found| js::canonical(&ref_json(found)) == wanted)
            })
            .and_then(|entry| entry.asset)
            .ok_or_else(|| Error::engine("Image object is absent from the checkpoint"))
    })
}

fn ref_json(object_ref: &ObjectRef) -> js::J {
    serde_json::to_value(object_ref)
        .map(js::J::from)
        .unwrap_or_default()
}

pub fn export(base: &[u8], checkpoint: Checkpoint, determinism: Determinism) -> Result<Vec<u8>> {
    guard(|| {
        check_determinism(determinism)?;
        Session::open(checkpoint.format, base, Some(checkpoint))?.export(determinism)
    })
}

fn check_determinism(determinism: Determinism) -> Result<()> {
    let hex = determinism.seed.len() == 64
        && determinism
            .seed
            .bytes()
            .all(|byte| byte.is_ascii_digit() || (b'a'..=b'f').contains(&byte));
    let now = determinism.now.as_bytes();
    let iso = now.len() == 24
        && now.iter().enumerate().all(|(index, byte)| match index {
            4 | 7 => *byte == b'-',
            10 => *byte == b'T',
            13 | 16 => *byte == b':',
            19 => *byte == b'.',
            23 => *byte == b'Z',
            _ => byte.is_ascii_digit(),
        });
    if !hex || !iso {
        return Err(Error::Invalid(
            "Expected deterministic SHA-256 seed and ISO UTC millisecond time".into(),
        ));
    }
    Ok(())
}

pub fn inspect(base: &[u8], checkpoint: Checkpoint) -> Result<Vec<Entry>> {
    guard(|| Session::open(checkpoint.format, base, Some(checkpoint))?.editable())
}

/// What the agent reads of an Office file without projecting all of it
/// (XLSX: the sheet list, or a range's non-empty cells; DOCX: a window of
/// paragraphs).
pub fn inspect_scoped(
    base: &[u8],
    checkpoint: Checkpoint,
    query: &InspectQuery,
) -> Result<serde_json::Value> {
    guard(|| Session::open(checkpoint.format, base, Some(checkpoint))?.inspect_scoped(query))
}

/// The blank file the agent's `create` and the Create tab start from; a
/// format without one is refused as `unsupported_operation`.
pub fn starter(format: Format) -> Result<&'static [u8]> {
    match format {
        Format::Xlsx => Ok(include_bytes!("../starters/blank.xlsx")),
        Format::Docx => Ok(include_bytes!("../starters/blank.docx")),
        Format::Pptx => Err(Error::edit(
            EditCode::UnsupportedOperation,
            format!("{} files cannot be created yet", format.as_str()),
        )),
    }
}

/// Commands one call holds (Epo 2026-10-21). The pipeline sends a call's
/// commands only up to this and DOCX's inserted-paragraph limit, so these
/// refusals guard other callers.
const MAX_CALL_COMMANDS: usize = 50;

fn check_call(commands: &[Command]) -> Result<()> {
    if commands.len() > MAX_CALL_COMMANDS {
        return Err(Error::edit(
            EditCode::InvalidInput,
            format!(
                "the call holds {} commands; a call holds at most {MAX_CALL_COMMANDS}",
                commands.len()
            ),
        ));
    }
    docx::agent::check_call(commands)
}

/// A new file: `starter` seeded, `commands` applied in one session, exported.
pub fn create(
    format: Format,
    starter: &[u8],
    commands: &[Command],
    determinism: Determinism,
) -> Result<Created> {
    guard(|| {
        check_determinism(determinism)?;
        assert_checkpoint(format, starter, None)?;
        check_call(commands)?;
        let state = match format {
            Format::Docx => docx::seed(starter),
            Format::Xlsx => xlsx::seed(starter),
            Format::Pptx => pptx::seed(starter),
        }?;
        let sha = sha256_hex(starter);
        let checkpoint = Checkpoint {
            format,
            schema_version: 1,
            base_sha256: &sha,
            state: &state,
        };
        let mut session = Session::open(format, starter, Some(checkpoint))?;
        let mut removed = Vec::new();
        for command in commands {
            session.apply(&storable_command(command, &mut removed))?;
        }
        session.finish(commands)?;
        Ok(Created {
            bytes: session.export(determinism)?,
            notice: removed_notice(&removed),
        })
    })
}

/// Applies content commands in order; the inverse list is in undo order (none
/// when a command has none) and the targets describe the post-edit state.
/// Characters an Office file cannot store leave the commands' text first and
/// the result's notice names them.
pub fn apply_commands(
    base: &[u8],
    checkpoint: Checkpoint,
    commands: &[Command],
) -> Result<CommandResult> {
    guard(|| {
        if commands.is_empty() {
            return Err(Error::edit(EditCode::InvalidInput, "no commands"));
        }
        check_call(commands)?;
        let mut session = Session::open(checkpoint.format, base, Some(checkpoint))?;
        let mut inverse = Some(Vec::with_capacity(commands.len()));
        let mut ids: Vec<String> = Vec::new();
        let mut seen = HashSet::new();
        let mut removed = Vec::new();
        for command in commands {
            let applied = session.apply(&storable_command(command, &mut removed))?;
            inverse = inverse.zip(applied.inverse).map(|(mut list, command)| {
                list.insert(0, command);
                list
            });
            for id in applied.ids {
                if seen.insert(id.clone()) {
                    ids.push(id);
                }
            }
        }
        session.finish(commands)?;
        let state = session.state();
        let targets = session.locate(&ids)?;
        Ok(CommandResult {
            state,
            inverse,
            targets,
            notice: removed_notice(&removed),
        })
    })
}

pub fn locate_targets(base: &[u8], checkpoint: Checkpoint, ids: &[String]) -> Result<Vec<Target>> {
    guard(|| Session::open(checkpoint.format, base, Some(checkpoint))?.locate(ids))
}

/// The build's identity, in place of the WASM asset hashes.
pub fn runtime_manifest() -> BTreeMap<&'static str, String> {
    BTreeMap::from([("office-service", env!("CARGO_PKG_VERSION").to_owned())])
}
