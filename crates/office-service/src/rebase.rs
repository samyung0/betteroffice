//! office-checkpoint.ts `rebaseOffice`.

use crate::common::{CHECKPOINT_MISMATCH, sha256_hex};
use crate::error::{Error, Result};
use crate::types::{Checkpoint, Format, Rebased};
use crate::xlsx;

pub(crate) fn rebase(
    base: &[u8],
    captured: Checkpoint,
    latest: Checkpoint,
    exported: &[u8],
) -> Result<Rebased> {
    let format = captured.format;
    if base.is_empty() || exported.is_empty() {
        return Err(Error::Invalid(
            "Expected supported Office format and nonempty source bytes".into(),
        ));
    }
    let base_sha256 = sha256_hex(base);
    for checkpoint in [captured, latest] {
        if checkpoint.format != format
            || checkpoint.schema_version != 1
            || checkpoint.base_sha256 != base_sha256
            || checkpoint.state.is_empty()
        {
            return Err(Error::Engine(CHECKPOINT_MISMATCH.into()));
        }
    }
    match format {
        Format::Xlsx => {
            let state = xlsx::rebase(base, captured.state, latest.state, exported)?;
            let exported_sha256 = sha256_hex(exported);
            let checkpoint = Checkpoint {
                format,
                schema_version: 1,
                base_sha256: &exported_sha256,
                state: &state,
            };
            xlsx::check(exported, checkpoint)?;
            let effects = xlsx::pending_effects(exported, checkpoint, None)?;
            Ok(Rebased { state, effects })
        }
        Format::Docx | Format::Pptx => Err(Error::engine("rebase is not ported yet")),
    }
}
