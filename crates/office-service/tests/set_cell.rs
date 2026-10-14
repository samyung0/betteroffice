//! A refused agent `set_cell` carries the cell as inspect lists it, as a
//! refused paragraph edit does (Epo 2026-10-14).

use office_service::{
    Checkpoint, Command, EditCode, Error, Format, apply_commands, inspect, seed, sha256_hex,
};
use serde_json::json;

const SAMPLE: &[u8] = include_bytes!("../../../apps/demo/public/sample.xlsx");

#[test]
fn a_stale_set_cell_carries_the_cell_as_inspect_lists_it() {
    let sha = sha256_hex(SAMPLE);
    let checkpoint = |state| Checkpoint {
        format: Format::Xlsx,
        schema_version: 1,
        base_sha256: &sha,
        state,
    };
    let seeded = seed(Format::Xlsx, SAMPLE).unwrap();
    let first = inspect(SAMPLE, checkpoint(&seeded)).unwrap().remove(0);
    let sheet = first.label.split('!').next().unwrap().to_owned();
    let set = |cell: &str, expected: &str, value: &str| Command::SetCell {
        sheet: sheet.clone(),
        cell: cell.into(),
        expected_value: expected.into(),
        value: value.into(),
    };
    let written = apply_commands(
        SAMPLE,
        checkpoint(&seeded),
        &[set("ZZ200", "", "cell 😀 value")],
    )
    .unwrap();
    let added = inspect(SAMPLE, checkpoint(&written.state))
        .unwrap()
        .into_iter()
        .find(|entry| entry.label == format!("{sheet}!ZZ200"))
        .unwrap();
    let refused = |cell: &str| {
        apply_commands(
            SAMPLE,
            checkpoint(&written.state),
            &[set(cell, "old", "new")],
        )
        .unwrap_err()
    };
    assert_eq!(
        refused("ZZ200"),
        Error::Edit {
            code: EditCode::StaleTarget,
            message: "the cell differs from expected_value".into(),
            details: Some(json!({
                "block": { "id": added.id, "type": "cell", "text": "cell 😀 value" }
            })),
        }
    );
    // An empty cell is not listed, so its refusal carries nothing.
    assert_eq!(
        refused("ZZ201"),
        Error::Edit {
            code: EditCode::StaleTarget,
            message: "the cell differs from expected_value".into(),
            details: None,
        }
    );
}
