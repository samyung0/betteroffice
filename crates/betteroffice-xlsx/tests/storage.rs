//! Golden schema 8 seeds and state-size budgets. A changed seed fails here and
//! ships in a maintenance window; a byte regression fails its budget.
use betteroffice_xlsx::{CalculationOptions, CellInput, CellRef, SheetId, Workbook};
use sha2::{Digest, Sha256};

/// Fixture, seed SHA-256 and seed byte budget.
const SEEDS: [(&str, &[u8], &str, usize); 4] = [
    (
        "course-guide.xlsx",
        include_bytes!("fixtures/storage/course-guide.xlsx"),
        "96b4640f904e841efb251e1740fd85f07f4d33b271e75302272b18aee9595f70",
        24_800,
    ),
    (
        "cells-1k.xlsx",
        include_bytes!("fixtures/storage/cells-1k.xlsx"),
        "4967af69daa06a01b510058b0bd20906821344e2bc34c7ecb8ac012509549b80",
        4_300,
    ),
    (
        "cells-10k.xlsx",
        include_bytes!("fixtures/storage/cells-10k.xlsx"),
        "97aa866997b1a6baa55a2a0c1569d3adb3ae43827c116c749ce20b40a8809d07",
        4_300,
    ),
    (
        "cells-100k.xlsx",
        include_bytes!("fixtures/storage/cells-100k.xlsx"),
        "877e06a90ce431fc4b08dcaa824cbb84c816b0d7c0de3a635b80668a2c0d3368",
        4_300,
    ),
];

#[test]
fn seeds_are_golden_deterministic_and_within_budget() {
    for (name, bytes, hash, budget) in SEEDS {
        let seed = Workbook::open_collaborative(bytes, 11)
            .unwrap()
            .encode_state_as_update_v1();
        let again = Workbook::open_collaborative(bytes, 12)
            .unwrap()
            .encode_state_as_update_v1();
        assert_eq!(seed, again, "{name} seeds differently per replica");
        assert_eq!(format!("{:x}", Sha256::digest(&seed)), hash, "{name}");
        assert!(seed.len() <= budget, "{name} seed is {} bytes", seed.len());
    }
}

#[test]
fn a_hundred_edited_cells_stay_within_budget() {
    let mut book = Workbook::open_collaborative(SEEDS[2].1, 13).unwrap();
    let edits = (0..100)
        .map(|index| CellInput {
            cell: CellRef::new(1 + index, 0),
            input: (500_000 + index).to_string(),
        })
        .collect::<Vec<_>>();
    book.edit_cells(SheetId(0), &edits, CalculationOptions::default())
        .unwrap();
    let state = book.encode_state_as_update_v1().len();
    assert!(state <= 17_400, "100 edits take {state} bytes");
}
