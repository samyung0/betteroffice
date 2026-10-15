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

/// A value holding characters an Office file cannot store applies without
/// them and the result names them for the tool result (Epo 2026-10-15); a
/// cell outside the grid and text over Excel's limit are refused as
/// `invalid_input` naming the rule, not as an engine failure.
#[test]
fn set_cell_drops_what_an_office_file_cannot_store_and_refuses_what_the_workbook_cannot_hold() {
    let sha = sha256_hex(SAMPLE);
    let seeded = seed(Format::Xlsx, SAMPLE).unwrap();
    let checkpoint = |state| Checkpoint {
        format: Format::Xlsx,
        schema_version: 1,
        base_sha256: &sha,
        state,
    };
    let first = inspect(SAMPLE, checkpoint(&seeded)).unwrap().remove(0);
    let sheet = first.label.split('!').next().unwrap().to_owned();
    let set = |cell: &str, value: &str| {
        apply_commands(
            SAMPLE,
            checkpoint(&seeded),
            &[Command::SetCell {
                sheet: sheet.clone(),
                cell: cell.into(),
                expected_value: String::new(),
                value: value.into(),
            }],
        )
    };
    let applied = set("ZZ300", "a\u{1}b\u{fffe}").unwrap();
    let written = inspect(SAMPLE, checkpoint(&applied.state))
        .unwrap()
        .into_iter()
        .find(|entry| entry.label == format!("{sheet}!ZZ300"))
        .unwrap();
    assert_eq!(written.value, "ab");
    assert_eq!(
        applied.notice.as_deref(),
        Some("Removed 2 characters Office files cannot store (U+0001, U+FFFE).")
    );
    let invalid = |message: &str| Error::Edit {
        code: EditCode::InvalidInput,
        message: message.into(),
        details: None,
    };
    assert_eq!(
        set("ZZZ9999999", "x").unwrap_err(),
        invalid("cell row 9999999, column 18278 is out of range")
    );
    assert_eq!(
        set("ZZ300", &"x".repeat(32_768)).unwrap_err(),
        invalid("cell text exceeds Excel's length limit")
    );
}

/// The Undo of an agent edit restores the cell as the file can store it: a
/// character the file cannot hold, there before the edit (a paste or an
/// import), is dropped instead of refusing the Undo.
#[test]
fn undo_restores_a_cell_without_a_character_the_file_cannot_store() {
    let parts: Vec<(String, Vec<u8>)> = ooxml_opc::unzip_parts(SAMPLE)
        .unwrap()
        .into_iter()
        .map(|(name, bytes)| match name.as_str() {
            "xl/sharedStrings.xml" => {
                let xml = String::from_utf8(bytes).unwrap();
                assert_eq!(xml.matches(">Item</t>").count(), 1);
                (
                    name,
                    xml.replace(">Item</t>", ">It\u{1}em</t>").into_bytes(),
                )
            }
            _ => (name, bytes),
        })
        .collect();
    let base = ooxml_opc::rezip_parts(&parts).unwrap();
    let sha = sha256_hex(&base);
    let checkpoint = |state| Checkpoint {
        format: Format::Xlsx,
        schema_version: 1,
        base_sha256: &sha,
        state,
    };
    let seeded = seed(Format::Xlsx, &base).unwrap();
    let cell = inspect(&base, checkpoint(&seeded))
        .unwrap()
        .into_iter()
        .find(|entry| entry.value == "It\u{1}em")
        .unwrap();
    let (sheet, address) = cell.position.rsplit_once(':').unwrap();
    let edit = apply_commands(
        &base,
        checkpoint(&seeded),
        &[Command::SetCell {
            sheet: sheet.into(),
            cell: address.into(),
            expected_value: cell.value.clone(),
            value: "Line".into(),
        }],
    )
    .unwrap();
    let undone = apply_commands(&base, checkpoint(&edit.state), &edit.inverse).unwrap();
    let restored = inspect(&base, checkpoint(&undone.state))
        .unwrap()
        .into_iter()
        .find(|entry| entry.id == cell.id)
        .unwrap();
    assert_eq!(restored.value, "Item");
}
