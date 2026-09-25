//! Legacy positional snapshots cannot establish the exact source and stable
//! topology identity the stable schema requires. New live sessions reject them.
use betteroffice_xlsx::{CalculationOptions, CellRef, SheetId, Workbook};

const SAMPLE: &[u8] = include_bytes!("../../../apps/demo/public/sample.xlsx");
const SHOWCASE: &[u8] = include_bytes!("../../../apps/demo/public/showcase.xlsx");
const HIDDEN: &[u8] = include_bytes!("fixtures/hidden-dimensions.xlsx");
const SAMPLE_V5: &[u8] = include_bytes!("fixtures/workbook-schema-v5-sample.update.bin");
const SHOWCASE_V5: &[u8] = include_bytes!("fixtures/workbook-schema-v5-showcase.update.bin");
const HIDDEN_V5: &[u8] = include_bytes!("fixtures/workbook-schema-v5-hidden.update.bin");

#[test]
fn positional_checkpoints_are_refused_without_mutating_the_exact_source_session() {
    for (source, legacy) in [
        (SAMPLE, SAMPLE_V5),
        (SHOWCASE, SHOWCASE_V5),
        (HIDDEN, HIDDEN_V5),
        (SAMPLE, SHOWCASE_V5),
    ] {
        for edited in [false, true] {
            let mut workbook = Workbook::open_collaborative(source, 5001).unwrap();
            if edited {
                workbook
                    .edit_cell(
                        SheetId(0),
                        CellRef::parse_a1("A1").unwrap(),
                        "local",
                        CalculationOptions {
                            now_serial: Some(46000.0),
                        },
                    )
                    .unwrap();
            }
            workbook.recalculate_all(CalculationOptions {
                now_serial: Some(46000.0),
            });
            let before = workbook.model().clone();
            let state = workbook.encode_state_as_update_v1();
            assert!(
                workbook
                    .apply_update_v1(
                        legacy,
                        CalculationOptions {
                            now_serial: Some(46000.0)
                        }
                    )
                    .is_err()
            );
            assert_eq!(workbook.model(), &before);
            assert_eq!(workbook.encode_state_as_update_v1(), state);
            let mut peer = Workbook::open_collaborative(source, 5002).unwrap();
            peer.apply_update_v1(
                &state,
                CalculationOptions {
                    now_serial: Some(46000.0),
                },
            )
            .unwrap();
            assert_eq!(peer.model(), workbook.model());
            workbook
                .edit_cell(
                    SheetId(0),
                    CellRef::parse_a1("A2").unwrap(),
                    "after rejection",
                    CalculationOptions {
                        now_serial: Some(46000.0),
                    },
                )
                .unwrap();
            let update = workbook
                .encode_diff_v1(&peer.encode_state_vector_v1())
                .unwrap();
            peer.apply_update_v1(
                &update,
                CalculationOptions {
                    now_serial: Some(46000.0),
                },
            )
            .unwrap();
            assert_eq!(peer.model(), workbook.model());
            assert_eq!(peer.save().unwrap(), workbook.save().unwrap());
        }
    }
}
