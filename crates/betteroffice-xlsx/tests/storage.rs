//! Golden schema 8 seeds and state-size budgets. A changed seed fails here and
//! ships in a maintenance window; a byte regression fails its budget.
use betteroffice_xlsx::{CalculationOptions, CellInput, CellRef, SheetId, Workbook};
use sha2::{Digest, Sha256};

/// Fixture, seed SHA-256 and seed byte budget.
const SEEDS: [(&str, &[u8], &str, usize); 4] = [
    (
        "course-guide.xlsx",
        include_bytes!("fixtures/storage/course-guide.xlsx"),
        "3f2d3e7929a7ef4c4b44256bf065554952e9f678ad7eb3fe70fda21276cea352",
        31_000,
    ),
    (
        "cells-1k.xlsx",
        include_bytes!("fixtures/storage/cells-1k.xlsx"),
        "355584b350f7d1493e893dfa03afcfb805eb4ecc3b20988f06f22e0fc5ac4036",
        4_300,
    ),
    (
        "cells-10k.xlsx",
        include_bytes!("fixtures/storage/cells-10k.xlsx"),
        "c688181f3b29891f3a20e65b76716050909b6652954f70914d4d088ef724cf15",
        4_300,
    ),
    (
        "cells-100k.xlsx",
        include_bytes!("fixtures/storage/cells-100k.xlsx"),
        "e4fc0c407d83fc4c7ee922885d29d4e8359f75e30632af9b3fd0c9a45a12a85c",
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
        assert_eq!(
            Workbook::seed_collaborative(bytes).unwrap(),
            seed,
            "{name} seeds differently without a replica"
        );
        assert_eq!(format!("{:x}", Sha256::digest(&seed)), hash, "{name}");
        assert!(seed.len() <= budget, "{name} seed is {} bytes", seed.len());
    }
}

#[test]
fn seeds_without_a_replica_match_every_fixture() {
    for (name, bytes) in [
        (
            "charts",
            &include_bytes!("../../../packages/xlsx/test-fixtures/charts.xlsx")[..],
        ),
        (
            "defined-names",
            include_bytes!("../../../packages/xlsx/test-fixtures/defined-names.xlsx"),
        ),
        (
            "unsupported-charts",
            include_bytes!("../../../packages/xlsx/test-fixtures/unsupported-charts.xlsx"),
        ),
        (
            "sample",
            include_bytes!("../../../apps/demo/public/sample.xlsx"),
        ),
        (
            "showcase",
            include_bytes!("../../../apps/demo/public/showcase.xlsx"),
        ),
        (
            "feature-rich",
            include_bytes!("../../../poc/fixtures/feature-rich.xlsx"),
        ),
        (
            "column-style-undo",
            include_bytes!("fixtures/column-style-undo.xlsx"),
        ),
        (
            "hidden-dimensions",
            include_bytes!("fixtures/hidden-dimensions.xlsx"),
        ),
    ] {
        assert_eq!(
            Workbook::seed_collaborative(bytes).unwrap(),
            Workbook::open_collaborative(bytes, 11)
                .unwrap()
                .encode_state_as_update_v1(),
            "{name}"
        );
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
