use betteroffice_xlsx::{
    CalculationOptions, Cell, CellRange, CellRef, CellState, CellValue, ColStyle, DefinedName, Op,
    Sheet, SheetId, StylePatch, Workbook, WorkbookModel,
};
fn at(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}
fn base() -> WorkbookModel {
    let mut data = Sheet::new("Data");
    data.set_cell(
        at("A1"),
        Cell {
            value: CellValue::Number { value: 10.0 },
            ..Cell::default()
        },
    );
    data.set_cell(
        at("A2"),
        Cell {
            value: CellValue::Number { value: 20.0 },
            ..Cell::default()
        },
    );
    let mut summary = Sheet::new("Summary");
    summary.set_cell(
        at("A1"),
        Cell {
            formula: Some("Data!A2".into()),
            ..Cell::default()
        },
    );
    WorkbookModel {
        sheets: vec![data, summary],
        defined_names: vec![DefinedName {
            name: "Range".into(),
            formula: "Data!$A$1:$A$2".into(),
            local_sheet: None,
            hidden: false,
        }],
        ..WorkbookModel::default()
    }
}
fn peer(id: u64) -> Workbook {
    Workbook::from_model_collaborative(base(), id).unwrap()
}
fn apply(book: &mut Workbook, op: Op) {
    book.apply_ops(vec![op], CalculationOptions::default())
        .unwrap();
}
fn sync(left: &mut Workbook, right: &mut Workbook) {
    let a = left.encode_state_as_update_v1();
    let b = right.encode_state_as_update_v1();
    left.apply_update_v1(&b, CalculationOptions::default())
        .unwrap();
    right
        .apply_update_v1(&a, CalculationOptions::default())
        .unwrap();
    assert_eq!(left.model(), right.model());
}
#[test]
fn concurrent_insert_and_formula_bind_to_the_authors_rows() {
    let mut a = peer(2001);
    let mut b = peer(2002);
    apply(
        &mut a,
        Op::InsertRows {
            sheet: SheetId(0),
            at: 1,
            count: 1,
        },
    );
    b.edit_cell(SheetId(0), at("B1"), "=A2", CalculationOptions::default())
        .unwrap();
    sync(&mut a, &mut b);
    assert_eq!(
        a.model().sheets[0]
            .cell(at("B1"))
            .unwrap()
            .formula
            .as_deref(),
        Some("A3")
    );
    assert_eq!(
        a.model().sheets[0].cell(at("A3")).unwrap().value,
        CellValue::Number { value: 20.0 }
    );
    let mut fresh = peer(2003);
    fresh
        .apply_update_v1(
            &a.encode_state_as_update_v1(),
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(fresh.model(), a.model());
    assert_eq!(fresh.save().unwrap(), a.save().unwrap());
}
#[test]
fn independent_inserts_survive_and_delete_undo_resurrects_remote_edits() {
    let mut a = peer(2011);
    let mut b = peer(2012);
    apply(
        &mut a,
        Op::InsertRows {
            sheet: SheetId(0),
            at: 1,
            count: 1,
        },
    );
    apply(
        &mut b,
        Op::InsertRows {
            sheet: SheetId(0),
            at: 1,
            count: 1,
        },
    );
    a.edit_cell(SheetId(0), at("A2"), "left", CalculationOptions::default())
        .unwrap();
    b.edit_cell(SheetId(0), at("A2"), "right", CalculationOptions::default())
        .unwrap();
    sync(&mut a, &mut b);
    assert_eq!(
        a.model().sheets[0].cell(at("A4")).unwrap().value,
        CellValue::Number { value: 20.0 }
    );
    let mut a = peer(2021);
    let mut b = peer(2022);
    apply(
        &mut a,
        Op::DeleteRows {
            sheet: SheetId(0),
            at: 1,
            count: 1,
        },
    );
    b.edit_cell(
        SheetId(0),
        at("A2"),
        "remote edit",
        CalculationOptions::default(),
    )
    .unwrap();
    sync(&mut a, &mut b);
    assert_eq!(
        a.model().sheets[1]
            .cell(at("A1"))
            .unwrap()
            .formula
            .as_deref(),
        Some("#REF!")
    );
    a.undo(CalculationOptions::default()).unwrap();
    sync(&mut a, &mut b);
    assert_eq!(
        a.model().sheets[0].cell(at("A2")).unwrap().value,
        CellValue::Text {
            value: "remote edit".into()
        }
    );
}
#[test]
fn sheet_structure_and_undo_keep_original_part_ownership_on_a_fresh_replica() {
    let source = Workbook::from_model(base()).unwrap().save().unwrap();
    let mut a = Workbook::open_collaborative(&source, 2031).unwrap();
    apply(&mut a, Op::RemoveSheet { index: 0 });
    a.edit_cell(
        SheetId(0),
        at("B1"),
        "remaining",
        CalculationOptions::default(),
    )
    .unwrap();
    let mut fresh = Workbook::open_collaborative(&source, 2032).unwrap();
    fresh
        .apply_update_v1(
            &a.encode_state_as_update_v1(),
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(a.save().unwrap(), fresh.save().unwrap());
    assert_eq!(
        Workbook::open(&fresh.save().unwrap())
            .unwrap()
            .model()
            .sheets[0]
            .name,
        "Summary"
    );
}

#[test]
fn structural_metadata_survives_peer_restore_and_local_undo() {
    use betteroffice_xlsx::{CellRange, FreezePane, Hyperlink};
    let source = Workbook::from_model(base()).unwrap().save().unwrap();
    let mut a = Workbook::open_collaborative(&source, 2041).unwrap();
    let mut b = Workbook::open_collaborative(&source, 2042).unwrap();
    a.recalculate_all(CalculationOptions::default());
    b.recalculate_all(CalculationOptions::default());
    let range = CellRange::new(at("B2"), at("C3"));
    let ops = vec![
        Op::SetFreezePane {
            sheet: SheetId(0),
            pane: Some(FreezePane::new(1, 1, at("B2"))),
        },
        Op::SetHyperlinks {
            sheet: SheetId(0),
            hyperlinks: vec![Hyperlink {
                range,
                external_target: None,
                location: Some("#Data!A2".into()),
                tooltip: Some("Target".into()),
                display: Some("Jump".into()),
            }],
        },
        Op::MergeCells {
            sheet: SheetId(0),
            range,
        },
        Op::SetRowHeight {
            sheet: SheetId(0),
            row: 1,
            height: Some(25.0),
        },
        Op::SetColWidth {
            sheet: SheetId(0),
            col: 1,
            width: Some(20.0),
        },
        Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
        Op::InsertCols {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
        Op::RenameSheet {
            sheet: SheetId(0),
            name: "Renamed".into(),
        },
        Op::AddSheet {
            index: 1,
            name: "Added".into(),
        },
        Op::RemoveSheet { index: 1 },
        Op::DeleteRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
        Op::DeleteCols {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
        Op::UnmergeCells {
            sheet: SheetId(0),
            range,
        },
    ];
    for op in ops {
        let before = a.model().clone();
        apply(&mut a, op);
        let after = a.model().clone();
        sync(&mut a, &mut b);
        a.undo(CalculationOptions::default()).unwrap();
        sync(&mut a, &mut b);
        assert_eq!(a.model(), &before);
        a.redo(CalculationOptions::default()).unwrap();
        sync(&mut a, &mut b);
        assert_eq!(a.model(), &after);
    }
    a.set_active_sheet(SheetId(1)).unwrap();
    let mut fresh = Workbook::open_collaborative(&source, 2043).unwrap();
    fresh
        .apply_update_v1(
            &a.encode_state_as_update_v1(),
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(a.model(), fresh.model());
    assert_eq!(
        a.save().unwrap(),
        fresh.save().unwrap(),
        "local tab selection must not alter checkpoint exports"
    );
}

#[test]
fn structural_updates_buffer_out_of_order_and_duplicate_delivery() {
    let mut a = peer(2051);
    let mut fresh = peer(2052);
    let before = a.encode_state_vector_v1();
    apply(
        &mut a,
        Op::InsertCols {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
    );
    let topology = a.encode_diff_v1(&before).unwrap();
    let after_insert = a.encode_state_vector_v1();
    a.edit_cell(
        SheetId(0),
        at("A1"),
        "inserted cell",
        CalculationOptions::default(),
    )
    .unwrap();
    let edit = a.encode_diff_v1(&after_insert).unwrap();
    fresh
        .apply_update_v1(&edit, CalculationOptions::default())
        .unwrap();
    fresh
        .apply_update_v1(&topology, CalculationOptions::default())
        .unwrap();
    fresh
        .apply_update_v1(&edit, CalculationOptions::default())
        .unwrap();
    fresh
        .apply_update_v1(&topology, CalculationOptions::default())
        .unwrap();
    assert_eq!(a.model(), fresh.model());
    assert_eq!(
        a.cell_identity(SheetId(0), at("A1")).unwrap(),
        fresh.cell_identity(SheetId(0), at("A1")).unwrap()
    );
}

#[test]
fn formula_membership_clips_spaced_ranges_and_preserves_hyperlink_prefix() {
    use betteroffice_xlsx::{CellRange, Hyperlink};
    let mut model = base();
    model.defined_names = vec![DefinedName {
        name: "Spaced".into(),
        formula: "Data!A1: Data!A2".into(),
        local_sheet: None,
        hidden: false,
    }];
    let mut a = Workbook::from_model_collaborative(model.clone(), 2061).unwrap();
    apply(
        &mut a,
        Op::SetHyperlinks {
            sheet: SheetId(1),
            hyperlinks: vec![Hyperlink {
                range: CellRange::new(at("A1"), at("A1")),
                external_target: None,
                location: Some("#Data!A2".into()),
                tooltip: None,
                display: None,
            }],
        },
    );
    apply(
        &mut a,
        Op::DeleteRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
    );
    assert!(!a.model().defined_names[0].formula.contains("#REF!"));
    assert_eq!(
        a.model().sheets[1].hyperlinks[0].location.as_deref(),
        Some("#'Data'!A1")
    );
    let mut fresh = Workbook::from_model_collaborative(model, 2062).unwrap();
    fresh
        .apply_update_v1(
            &a.encode_state_as_update_v1(),
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(a.model(), fresh.model());
}

#[test]
fn concurrent_names_merges_and_last_sheet_deletions_resolve_in_any_order() {
    use betteroffice_xlsx::CellRange;
    for scenario in 0..3 {
        let mut a = peer(2071);
        let mut b = peer(2072);
        match scenario {
            0 => {
                apply(
                    &mut a,
                    Op::RenameSheet {
                        sheet: SheetId(0),
                        name: "Shared".into(),
                    },
                );
                apply(
                    &mut b,
                    Op::RenameSheet {
                        sheet: SheetId(1),
                        name: "Shared".into(),
                    },
                );
            }
            1 => {
                apply(
                    &mut a,
                    Op::MergeCells {
                        sheet: SheetId(0),
                        range: CellRange::new(at("A1"), at("B2")),
                    },
                );
                apply(
                    &mut b,
                    Op::MergeCells {
                        sheet: SheetId(0),
                        range: CellRange::new(at("B2"), at("C3")),
                    },
                );
            }
            _ => {
                apply(&mut a, Op::RemoveSheet { index: 0 });
                apply(&mut b, Op::RemoveSheet { index: 1 });
            }
        }
        let left = a.encode_state_as_update_v1();
        let right = b.encode_state_as_update_v1();
        sync(&mut a, &mut b);
        let mut forward = peer(2073);
        let mut reverse = peer(2074);
        for update in [&left, &right, &left] {
            forward
                .apply_update_v1(update, CalculationOptions::default())
                .unwrap();
        }
        for update in [&right, &left, &right] {
            reverse
                .apply_update_v1(update, CalculationOptions::default())
                .unwrap();
        }
        assert_eq!(a.model(), forward.model());
        assert_eq!(a.model(), reverse.model());
        assert_eq!(forward.save().unwrap(), reverse.save().unwrap());
        match scenario {
            0 => {
                assert_eq!(a.model().sheets[0].name, "Shared");
                assert!(a.model().sheets[1].name.starts_with("Shared ("));
            }
            1 => {
                assert_eq!(a.model().sheets[0].merges.len(), 1);
                assert!(a.model().sheets[0].cell(at("A2")).is_some());
            }
            _ => {
                assert_eq!(a.model().sheets.len(), 1);
                assert_eq!(a.model().sheets[0].name, "Data");
                assert!(a.model().sheets[0].cell(at("A2")).is_some());
            }
        }
        let converged = a.model().clone();
        a.undo(CalculationOptions::default()).unwrap();
        sync(&mut a, &mut b);
        a.redo(CalculationOptions::default()).unwrap();
        sync(&mut a, &mut b);
        assert_eq!(a.model(), &converged);
        if scenario == 2 {
            apply(
                &mut a,
                Op::AddSheet {
                    index: 1,
                    name: "New".into(),
                },
            );
            sync(&mut a, &mut b);
            assert_eq!(a.model().sheets[0].name, "Data");
            assert!(a.model().sheets[0].cell(at("A2")).is_some());
        }
    }
}

#[test]
fn blank_axis_edits_keep_concurrent_cell_identity_and_undo() {
    let model = WorkbookModel {
        sheets: vec![Sheet::new("Blank")],
        ..WorkbookModel::default()
    };
    for rows in [true, false] {
        for insert in [true, false] {
            for index in [0, 100] {
                let mut left = Workbook::from_model_collaborative(model.clone(), 2101).unwrap();
                let mut right = Workbook::from_model_collaborative(model.clone(), 2102).unwrap();
                let point = if rows {
                    CellRef::new(index, 0)
                } else {
                    CellRef::new(0, index)
                };
                let following = if rows {
                    CellRef::new(index + 1, 0)
                } else {
                    CellRef::new(0, index + 1)
                };
                let original_id = left.cell_identity(SheetId(0), point).unwrap();
                let operation = match (rows, insert) {
                    (true, true) => Op::InsertRows {
                        sheet: SheetId(0),
                        at: index,
                        count: 1,
                    },
                    (true, false) => Op::DeleteRows {
                        sheet: SheetId(0),
                        at: index,
                        count: 1,
                    },
                    (false, true) => Op::InsertCols {
                        sheet: SheetId(0),
                        at: index,
                        count: 1,
                    },
                    (false, false) => Op::DeleteCols {
                        sheet: SheetId(0),
                        at: index,
                        count: 1,
                    },
                };
                assert!(
                    left.apply_ops(vec![operation], CalculationOptions::default())
                        .unwrap()
                        .applied
                );
                assert_ne!(left.cell_identity(SheetId(0), point).unwrap(), original_id);
                assert!(left.can_undo());
                right
                    .edit_cell(
                        SheetId(0),
                        point,
                        "remote cell",
                        CalculationOptions::default(),
                    )
                    .unwrap();
                if !insert {
                    right
                        .edit_cell(
                            SheetId(0),
                            following,
                            "following cell",
                            CalculationOptions::default(),
                        )
                        .unwrap();
                }
                sync(&mut left, &mut right);
                let expected = if insert { following } else { point };
                let expected_value = if insert {
                    "remote cell"
                } else {
                    "following cell"
                };
                assert_eq!(
                    left.cell(SheetId(0), expected).unwrap().input,
                    expected_value
                );
                let converged = left.model().clone();
                let mut fresh = Workbook::from_model_collaborative(model.clone(), 2103).unwrap();
                fresh
                    .apply_update_v1(
                        &left.encode_state_as_update_v1(),
                        CalculationOptions::default(),
                    )
                    .unwrap();
                assert_eq!(fresh.model(), &converged);
                assert_eq!(fresh.save().unwrap(), left.save().unwrap());
                assert!(left.undo(CalculationOptions::default()).unwrap().applied);
                sync(&mut left, &mut right);
                assert_eq!(left.cell_identity(SheetId(0), point).unwrap(), original_id);
                assert_eq!(left.cell(SheetId(0), point).unwrap().input, "remote cell");
                assert!(left.redo(CalculationOptions::default()).unwrap().applied);
                sync(&mut left, &mut right);
                assert_eq!(left.model(), &converged);
            }
        }
    }
}

#[test]
fn inserted_range_members_survive_endpoint_deletion_and_peer_undo() {
    for rows in [true, false] {
        let model = WorkbookModel {
            sheets: vec![Sheet::new("Data")],
            ..WorkbookModel::default()
        };
        let mut left = Workbook::from_model_collaborative(model.clone(), 2111).unwrap();
        let mut right = Workbook::from_model_collaborative(model.clone(), 2112).unwrap();
        let formula_at = if rows { at("F1") } else { at("A6") };
        let formula = if rows { "=SUM(A2:A3)" } else { "=SUM(B1:C1)" };
        let insert = if rows {
            Op::InsertRows {
                sheet: SheetId(0),
                at: 2,
                count: 1,
            }
        } else {
            Op::InsertCols {
                sheet: SheetId(0),
                at: 2,
                count: 1,
            }
        };
        apply(&mut left, insert);
        right
            .edit_cell(
                SheetId(0),
                formula_at,
                formula,
                CalculationOptions::default(),
            )
            .unwrap();
        sync(&mut left, &mut right);
        let inserted = if rows { at("A3") } else { at("C1") };
        right
            .edit_cell(SheetId(0), inserted, "7", CalculationOptions::default())
            .unwrap();
        sync(&mut left, &mut right);
        let vector = left.encode_state_vector_v1();
        let delete = |index| {
            if rows {
                Op::DeleteRows {
                    sheet: SheetId(0),
                    at: index,
                    count: 1,
                }
            } else {
                Op::DeleteCols {
                    sheet: SheetId(0),
                    at: index,
                    count: 1,
                }
            }
        };
        apply(&mut left, delete(3));
        let first_delete = left.encode_diff_v1(&vector).unwrap();
        let vector = left.encode_state_vector_v1();
        apply(&mut left, delete(1));
        let second_delete = left.encode_diff_v1(&vector).unwrap();
        right
            .apply_update_v1(&second_delete, CalculationOptions::default())
            .unwrap();
        right
            .apply_update_v1(&first_delete, CalculationOptions::default())
            .unwrap();
        assert_eq!(left.model(), right.model());
        assert_eq!(
            left.cell(SheetId(0), formula_at).unwrap().input,
            if rows { "=SUM(A2)" } else { "=SUM(B1)" }
        );
        assert_eq!(
            left.model().sheets[0].cell(formula_at).unwrap().value,
            CellValue::Number { value: 7.0 }
        );
        let converged = left.model().clone();
        let saved = left.save().unwrap();
        let reopened = Workbook::open(&saved).unwrap();
        assert_eq!(
            reopened.cell(SheetId(0), formula_at).unwrap().input,
            left.cell(SheetId(0), formula_at).unwrap().input
        );
        let mut fresh = Workbook::from_model_collaborative(model, 2113).unwrap();
        fresh
            .apply_update_v1(
                &left.encode_state_as_update_v1(),
                CalculationOptions::default(),
            )
            .unwrap();
        assert_eq!(fresh.model(), &converged);
        assert_eq!(fresh.save().unwrap(), saved);
        for _ in 0..3 {
            left.undo(CalculationOptions::default()).unwrap();
            sync(&mut left, &mut right);
        }
        assert_eq!(left.cell(SheetId(0), formula_at).unwrap().input, formula);
        for _ in 0..3 {
            left.redo(CalculationOptions::default()).unwrap();
            sync(&mut left, &mut right);
        }
        assert_eq!(left.model(), &converged);
    }
}

#[test]
fn insertion_before_a_clipped_range_start_does_not_join_it() {
    let mut workbook = peer(2121);
    workbook
        .edit_cell(
            SheetId(0),
            at("F1"),
            "=SUM(A2:A3)",
            CalculationOptions::default(),
        )
        .unwrap();
    apply(
        &mut workbook,
        Op::DeleteRows {
            sheet: SheetId(0),
            at: 1,
            count: 1,
        },
    );
    apply(
        &mut workbook,
        Op::InsertRows {
            sheet: SheetId(0),
            at: 1,
            count: 1,
        },
    );
    assert_eq!(
        workbook.cell(SheetId(0), at("F1")).unwrap().input,
        "=SUM(A3)"
    );
}

#[test]
fn reordered_structural_update_after_undo_returns_and_restores() {
    let (send, receive) = std::sync::mpsc::channel();
    std::thread::spawn(move || {
        let source = include_bytes!("../../../apps/demo/public/sample.xlsx");
        let mut a = Workbook::open_collaborative(source, 50100).unwrap();
        let mut b = Workbook::open_collaborative(source, 50101).unwrap();
        let options = CalculationOptions::default();
        let vector = a.encode_state_vector_v1();
        apply(
            &mut a,
            Op::DeleteRows {
                sheet: SheetId(0),
                at: 4,
                count: 1,
            },
        );
        b.apply_update_v1(&a.encode_diff_v1(&vector).unwrap(), options)
            .unwrap();
        let vector = a.encode_state_vector_v1();
        a.undo(options).unwrap();
        b.undo(options).unwrap();
        b.apply_update_v1(&a.encode_diff_v1(&vector).unwrap(), options)
            .unwrap();
        let vector = b.encode_state_vector_v1();
        apply(
            &mut b,
            Op::InsertCols {
                sheet: SheetId(0),
                at: 1,
                count: 1,
            },
        );
        let column_insert = b.encode_diff_v1(&vector).unwrap();
        let vector = b.encode_state_vector_v1();
        apply(
            &mut b,
            Op::DeleteCols {
                sheet: SheetId(0),
                at: 3,
                count: 1,
            },
        );
        let column_delete = b.encode_diff_v1(&vector).unwrap();
        a.apply_update_v1(&column_insert, options).unwrap();
        let vector = b.encode_state_vector_v1();
        apply(
            &mut b,
            Op::InsertRows {
                sheet: SheetId(0),
                at: 1,
                count: 1,
            },
        );
        let row_insert = b.encode_diff_v1(&vector).unwrap();
        a.undo(options).unwrap();
        apply(
            &mut a,
            Op::DeleteRows {
                sheet: SheetId(0),
                at: 4,
                count: 1,
            },
        );
        apply(
            &mut b,
            Op::DeleteRows {
                sheet: SheetId(0),
                at: 3,
                count: 1,
            },
        );
        a.apply_update_v1(&row_insert, options).unwrap();
        let mut fresh = Workbook::open_collaborative(source, 90000).unwrap();
        fresh
            .apply_update_v1(&a.encode_state_as_update_v1(), options)
            .unwrap();
        assert_eq!(a.model(), fresh.model());
        assert_eq!(a.save().unwrap(), fresh.save().unwrap());
        a.apply_update_v1(&column_delete, options).unwrap();
        sync(&mut a, &mut b);
        send.send(()).unwrap();
    });
    receive
        .recv_timeout(std::time::Duration::from_secs(10))
        .expect("reordered update must complete");
}

fn reopen_rebased(source: &[u8], state: &[u8], client: u64) -> Workbook {
    let mut book = Workbook::open_collaborative(source, client).unwrap();
    book.apply_update_v1(state, CalculationOptions::default())
        .unwrap();
    book
}

#[test]
fn rebase_preserves_saved_structure_and_indexed_cell_identity_without_old_source() {
    let source = Workbook::from_model(base()).unwrap().save().unwrap();
    let mut book = Workbook::open_collaborative(&source, 7011).unwrap();
    apply(
        &mut book,
        Op::InsertRows {
            sheet: SheetId(0),
            at: 1,
            count: 2,
        },
    );
    apply(
        &mut book,
        Op::InsertCols {
            sheet: SheetId(0),
            at: 1,
            count: 1,
        },
    );
    book.edit_cell(SheetId(0), at("A2"), "same", CalculationOptions::default())
        .unwrap();
    book.edit_cell(SheetId(0), at("A3"), "same", CalculationOptions::default())
        .unwrap();
    let captured = book.encode_state_as_update_v1();
    let published = book.save().unwrap();
    apply(
        &mut book,
        Op::DeleteRows {
            sheet: SheetId(0),
            at: 1,
            count: 1,
        },
    );
    apply(
        &mut book,
        Op::InsertCols {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
    );
    apply(
        &mut book,
        Op::RenameSheet {
            sheet: SheetId(0),
            name: "Renamed".into(),
        },
    );
    apply(
        &mut book,
        Op::AddSheet {
            index: 0,
            name: "New".into(),
        },
    );
    book.edit_cell(SheetId(1), at("C2"), "=B3", CalculationOptions::default())
        .unwrap();
    let expected = Workbook::open(&book.save().unwrap()).unwrap().into_model();
    let result = Workbook::rebase_checkpoint(
        &source,
        &captured,
        &book.encode_state_as_update_v1(),
        &published,
        7012,
    )
    .unwrap();
    drop(source);
    let mut reopened = reopen_rebased(&published, &result.state, 7013);
    let indexed = reopen_rebased(&published, &result.indexed_state, 7014);
    assert_eq!(reopened.model(), &expected);
    assert_eq!(
        indexed.checkpoint_sheet_ids().unwrap(),
        vec!["sheet:1", "sheet:2"]
    );
    assert_eq!(
        indexed
            .checkpoint_cell_identities([(SheetId(0), at("A3"))])
            .unwrap()[0],
        reopened.cell_identity(SheetId(1), at("B2")).unwrap()
    );
    assert_ne!(
        indexed
            .checkpoint_cell_identities([(SheetId(0), at("A2"))])
            .unwrap()[0],
        reopened.cell_identity(SheetId(1), at("B2")).unwrap()
    );
    assert_eq!(
        reopened.model().sheets[2]
            .cell(at("A1"))
            .unwrap()
            .formula
            .as_deref(),
        Some("'Renamed'!B3")
    );
    apply(
        &mut reopened,
        Op::InsertRows {
            sheet: SheetId(1),
            at: 0,
            count: 1,
        },
    );
    assert_eq!(
        reopened.model().sheets[2]
            .cell(at("A1"))
            .unwrap()
            .formula
            .as_deref(),
        Some("'Renamed'!B4")
    );
    let peer = reopen_rebased(&published, &reopened.encode_state_as_update_v1(), 7015);
    assert_eq!(peer.model(), reopened.model());
    assert_eq!(
        Workbook::open(&peer.save().unwrap()).unwrap().model(),
        reopened.model()
    );
}

#[test]
fn second_rebase_releases_previous_overlay_and_preserves_later_edits() {
    let source = Workbook::from_model(base()).unwrap().save().unwrap();
    let mut book = Workbook::open_collaborative(&source, 7021).unwrap();
    let captured = book.encode_state_as_update_v1();
    let published = book.save().unwrap();
    book.edit_cell(SheetId(0), at("B1"), "first", CalculationOptions::default())
        .unwrap();
    let first = Workbook::rebase_checkpoint(
        &source,
        &captured,
        &book.encode_state_as_update_v1(),
        &published,
        7022,
    )
    .unwrap();
    let mut book = reopen_rebased(&published, &first.state, 7023);
    let captured = book.encode_state_as_update_v1();
    let next_published = book.save().unwrap();
    apply(
        &mut book,
        Op::DeleteCols {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
    );
    book.edit_cell(
        SheetId(0),
        at("A2"),
        "second",
        CalculationOptions::default(),
    )
    .unwrap();
    let second = Workbook::rebase_checkpoint(
        &published,
        &captured,
        &book.encode_state_as_update_v1(),
        &next_published,
        7024,
    )
    .unwrap();
    let mut rebased = reopen_rebased(&next_published, &second.state, 7025);
    assert_eq!(rebased.model(), book.model());
    let mut wrong_source = Workbook::open_collaborative(&source, 7026).unwrap();
    assert!(
        wrong_source
            .apply_update_v1(&second.state, CalculationOptions::default())
            .is_err()
    );
    let mut peer = reopen_rebased(&next_published, &second.state, 7027);
    rebased
        .edit_cell(SheetId(0), at("C3"), "left", CalculationOptions::default())
        .unwrap();
    peer.edit_cell(SheetId(0), at("D3"), "right", CalculationOptions::default())
        .unwrap();
    sync(&mut rebased, &mut peer);
}

#[test]
fn rebase_overlay_is_sparse_binary_and_immutable_after_bootstrap() {
    use yrs::updates::decoder::Decode;
    use yrs::{Doc, Map, Out, ReadTxn, StateVector, Transact, WriteTxn};
    let source = Workbook::from_model(base()).unwrap().save().unwrap();
    let mut book = Workbook::open_collaborative(&source, 7031).unwrap();
    let captured = book.encode_state_as_update_v1();
    let published = book.save().unwrap();
    book.edit_cell(SheetId(0), at("C5"), "saved", CalculationOptions::default())
        .unwrap();
    book.recalculate_all(CalculationOptions::default());
    let current = book.save().unwrap();
    let result = Workbook::rebase_checkpoint(
        &source,
        &captured,
        &book.encode_state_as_update_v1(),
        &published,
        7032,
    )
    .unwrap();
    let snapshot = Doc::with_client_id(7033);
    snapshot
        .transact_mut()
        .apply_update(yrs::Update::decode_v1(&result.state).unwrap())
        .unwrap();
    let before = ooxml_opc::unzip_parts(&published)
        .unwrap()
        .into_iter()
        .collect::<std::collections::BTreeMap<_, _>>();
    let after = ooxml_opc::unzip_parts(&current)
        .unwrap()
        .into_iter()
        .collect::<std::collections::BTreeMap<_, _>>();
    {
        let txn = snapshot.transact();
        let root = txn.get_map("xlsx:rebase").unwrap();
        let mut copied = 0;
        for (key, value) in root.iter(&txn) {
            if key == "$manifest" {
                continue;
            }
            let path = key.strip_prefix("part:").unwrap();
            let Out::Any(yrs::Any::Buffer(bytes)) = value else {
                panic!("part was not stored as bytes")
            };
            assert_eq!(bytes.as_ref(), after[path]);
            assert!(
                before.get(path).is_none_or(|old| old != &after[path]),
                "copied an unchanged part"
            );
            copied += 1;
        }
        assert_eq!(
            copied,
            after
                .iter()
                .filter(|(path, bytes)| before.get(*path).is_none_or(|old| old != *bytes))
                .count()
        );
    }
    let mut target = reopen_rebased(&published, &result.state, 7034);
    let previous = target.encode_state_as_update_v1();
    {
        let mut txn = snapshot.transact_mut();
        let root = txn.get_or_insert_map("xlsx:rebase");
        root.insert(
            &mut txn,
            "part:xl/worksheets/sheet1.xml",
            yrs::Any::Buffer(vec![0].into()),
        );
    }
    let corrupt = snapshot
        .transact()
        .encode_state_as_update_v1(&StateVector::default());
    assert!(
        target
            .apply_update_v1(&corrupt, CalculationOptions::default())
            .is_err()
    );
    assert_eq!(target.encode_state_as_update_v1(), previous);
    assert!(
        target
            .apply_update_v1(
                &book.encode_state_as_update_v1(),
                CalculationOptions::default()
            )
            .is_err()
    );
}

#[test]
fn capy_contributor_metadata_survives_native_live_and_durable_updates() {
    use yrs::updates::decoder::Decode;
    use yrs::{Any, Doc, Map, ReadTxn, StateVector, Transact, WriteTxn};
    const ROOT: &str = "__capy_pending_contributors";
    let source = Workbook::from_model(base()).unwrap().save().unwrap();
    let mut native = Workbook::open_collaborative(&source, 7041).unwrap();
    let server = Doc::with_client_id(7042);
    server
        .transact_mut()
        .apply_update(yrs::Update::decode_v1(&native.encode_state_as_update_v1()).unwrap())
        .unwrap();
    let marker = Any::Map(std::sync::Arc::new(std::collections::HashMap::from([
        ("access".into(), Any::String("write".into())),
        ("nonce".into(), Any::String("n".into())),
        ("userId".into(), Any::String("u".into())),
    ])));
    {
        let mut txn = server.transact_mut();
        txn.get_or_insert_map(ROOT)
            .insert(&mut txn, "actor", marker.clone());
    }
    let live = server
        .transact()
        .encode_state_as_update_v1(&StateVector::default());
    native
        .apply_update_v1(&live, CalculationOptions::default())
        .unwrap();
    native
        .edit_cell(SheetId(0), at("C1"), "kept", CalculationOptions::default())
        .unwrap();
    let retained = Doc::new();
    retained
        .transact_mut()
        .apply_update(yrs::Update::decode_v1(&native.encode_state_as_update_v1()).unwrap())
        .unwrap();
    {
        let txn = retained.transact();
        assert_eq!(
            txn.get_map(ROOT).unwrap().get(&txn, "actor"),
            Some(yrs::Out::Any(marker))
        );
    }
    server
        .transact_mut()
        .apply_update(yrs::Update::decode_v1(&native.encode_state_as_update_v1()).unwrap())
        .unwrap();
    let vector = server.transact().state_vector();
    {
        let mut txn = server.transact_mut();
        txn.get_map(ROOT).unwrap().remove(&mut txn, "actor");
    }
    let removal = server.transact().encode_diff_v1(&vector);
    native
        .apply_update_v1(&removal, CalculationOptions::default())
        .unwrap();
    retained
        .transact_mut()
        .apply_update(yrs::Update::decode_v1(&native.encode_state_as_update_v1()).unwrap())
        .unwrap();
    {
        let txn = retained.transact();
        assert!(txn.get_map(ROOT).unwrap().get(&txn, "actor").is_none());
        assert_eq!(txn.snapshot(), server.transact().snapshot());
    }
    let saved = native.encode_state_as_update_v1();
    let mut restored = reopen_rebased(&source, &saved, 7043);
    assert_eq!(restored.model(), native.model());
    let published = restored.save().unwrap();
    restored
        .edit_cell(SheetId(0), at("D1"), "newer", CalculationOptions::default())
        .unwrap();
    let rebased = Workbook::rebase_checkpoint(
        &source,
        &saved,
        &restored.encode_state_as_update_v1(),
        &published,
        7044,
    )
    .unwrap();
    assert_eq!(
        reopen_rebased(&published, &rebased.state, 7045).model(),
        restored.model()
    );
    let previous = native.encode_state_as_update_v1();
    {
        let mut txn = server.transact_mut();
        txn.get_or_insert_map("__capy_unrelated_root")
            .insert(&mut txn, "actor", true);
    }
    let unrelated = server
        .transact()
        .encode_state_as_update_v1(&StateVector::default());
    assert!(
        native
            .apply_update_v1(&unrelated, CalculationOptions::default())
            .is_err()
    );
    assert_eq!(native.encode_state_as_update_v1(), previous);
}

#[test]
fn source_column_styles_tables_and_array_formulas_follow_structural_edits() {
    let range = |address: &str| CellRange::parse_a1(address).unwrap();
    let mut model = base();
    let data = &mut model.sheets[0];
    data.col_styles = vec![ColStyle {
        first: 0,
        last: 1,
        xf: 0,
    }];
    data.format.default_row_height_pt = Some(20.0);
    data.set_cell(
        at("B1"),
        Cell {
            value: CellValue::Number { value: 20.0 },
            formula: Some("A1:A2*2".into()),
            ..Cell::default()
        },
    );
    data.set_array_formula(at("B1"), range("B1:B2"));
    model.tables = vec![xlsx_model::Table {
        name: "Values".into(),
        sheet: SheetId(0),
        range: range("A1:A2"),
        header_rows: 1,
        totals_rows: 0,
        columns: vec!["Value".into()],
    }];
    let mut a = Workbook::from_model_collaborative(model.clone(), 2101).unwrap();
    let mut b = Workbook::from_model_collaborative(model, 2102).unwrap();
    apply(
        &mut a,
        Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
    );
    apply(
        &mut a,
        Op::InsertCols {
            sheet: SheetId(0),
            at: 1,
            count: 1,
        },
    );
    sync(&mut a, &mut b);
    let sheet = &b.model().sheets[0];
    assert_eq!(sheet.format.default_row_height_pt, Some(20.0));
    assert_eq!(
        sheet.col_styles,
        [
            ColStyle {
                first: 0,
                last: 0,
                xf: 0
            },
            ColStyle {
                first: 2,
                last: 2,
                xf: 0
            }
        ]
    );
    assert_eq!(
        sheet.array_formulas().collect::<Vec<_>>(),
        [(at("C2"), range("C2:C3"))]
    );
    assert_eq!(b.model().tables[0].range, range("A2:A3"));

    b.edit_cell(SheetId(0), at("C2"), "5", CalculationOptions::default())
        .unwrap();
    assert_eq!(b.model().sheets[0].array_formulas().count(), 0);
    apply(&mut b, Op::RemoveSheet { index: 0 });
    assert!(b.model().tables.is_empty());
}

#[test]
fn a_batch_sets_a_cell_to_a_style_an_earlier_op_of_the_batch_interned() {
    let mut a = peer(1);
    let mut b = peer(2);
    // C1 first, so the batch interns bold before the styles a row-major read meets first.
    let mut ops = [
        (
            "C1",
            StylePatch {
                bold: Some(true),
                ..StylePatch::default()
            },
        ),
        (
            "A1",
            StylePatch {
                italic: Some(true),
                ..StylePatch::default()
            },
        ),
        (
            "B1",
            StylePatch {
                strikethrough: Some(true),
                ..StylePatch::default()
            },
        ),
    ]
    .into_iter()
    .map(|(cell, patch)| Op::PatchRangeStyle {
        sheet: SheetId(0),
        range: CellRange::new(at(cell), at(cell)),
        patch,
    })
    .collect::<Vec<_>>();
    let mut preview = a.model().clone();
    for op in &ops {
        xlsx_ops::apply_in_place(&mut preview, op).unwrap();
    }
    let bold = preview.sheets[0].cell(at("C1")).unwrap().style;
    ops.push(Op::SetCell {
        sheet: SheetId(0),
        at: at("D1"),
        cell: CellState {
            value: CellValue::Number { value: 1.0 },
            formula: None,
            style: bold,
        },
    });
    a.apply_ops(ops, CalculationOptions::default()).unwrap();
    sync(&mut a, &mut b);
    for book in [&a, &b] {
        let model = book.model();
        let format = |cell: &str| {
            model
                .styles
                .cell_format(model.sheets[0].cell(at(cell)).unwrap().style)
        };
        assert_eq!(format("D1"), format("C1"));
        assert_ne!(format("D1"), format("A1"));
    }
}

fn overrides(book: &Workbook, sheet: &str) -> Vec<(String, String)> {
    use yrs::updates::decoder::Decode;
    use yrs::{Doc, Map, MapRef, ReadTxn, Transact};
    let doc = Doc::new();
    doc.transact_mut()
        .apply_update(yrs::Update::decode_v1(&book.encode_state_as_update_v1()).unwrap())
        .unwrap();
    let txn = doc.transact();
    let sheet = txn
        .get_map("xlsx:sheets")
        .unwrap()
        .get(&txn, sheet)
        .unwrap()
        .cast::<MapRef>()
        .unwrap();
    let contents = sheet
        .get(&txn, "contents")
        .unwrap()
        .cast::<MapRef>()
        .unwrap();
    let mut entries = contents
        .iter(&txn)
        .map(|(key, value)| (key.to_owned(), value.to_string(&txn)))
        .collect::<Vec<_>>();
    entries.sort();
    entries
}

#[test]
fn source_cells_are_overridden_cleared_reverted_and_undone() {
    let options = CalculationOptions::default();
    let mut book = peer(3001);
    assert!(overrides(&book, "sheet:0").is_empty());
    book.edit_cell(SheetId(0), at("A1"), "", options).unwrap();
    assert!(book.model().sheets[0].cell(at("A1")).is_none());
    assert_eq!(
        overrides(&book, "sheet:0"),
        [(
            r#"[{"run":"base","offset":0},{"run":"base","offset":0}]"#.to_owned(),
            r#"{"value":{"kind":"empty"},"formula":null}"#.to_owned()
        )]
    );
    book.edit_cell(SheetId(0), at("A1"), "10", options).unwrap();
    assert!(
        overrides(&book, "sheet:0").is_empty(),
        "the source value reverts"
    );
    book.edit_cell(SheetId(0), at("A2"), "25", options).unwrap();
    assert_eq!(overrides(&book, "sheet:0").len(), 1);
    book.undo(options).unwrap();
    assert!(overrides(&book, "sheet:0").is_empty());
    assert_eq!(
        book.model().sheets[0].cell(at("A2")).unwrap().value,
        CellValue::Number { value: 20.0 }
    );
    let mut fresh = peer(3002);
    fresh
        .apply_update_v1(&book.encode_state_as_update_v1(), options)
        .unwrap();
    assert_eq!(fresh.model(), book.model());
}

#[test]
fn a_concurrent_clear_and_set_of_a_source_cell_resolve_by_client_order() {
    let options = CalculationOptions::default();
    for (clearing, setting) in [(3011, 3012), (3022, 3021)] {
        let mut clear = peer(clearing);
        let mut set = peer(setting);
        clear.edit_cell(SheetId(0), at("A1"), "", options).unwrap();
        set.edit_cell(SheetId(0), at("A1"), "set", options).unwrap();
        sync(&mut clear, &mut set);
        let expected = (setting > clearing).then(|| CellValue::Text {
            value: "set".into(),
        });
        assert_eq!(
            clear.model().sheets[0]
                .cell(at("A1"))
                .map(|cell| cell.value.clone()),
            expected
        );
    }
}

#[test]
fn a_hidden_source_sheet_still_validates_its_formulas() {
    let mut model = base();
    model.sheets[1].set_cell(
        at("B1"),
        Cell {
            formula: Some("SUM(Values[Value])".into()),
            ..Cell::default()
        },
    );
    let mut book = Workbook::from_model_collaborative(model, 3031).unwrap();
    let state = book.encode_state_as_update_v1();
    assert!(
        book.apply_ops(
            vec![Op::RemoveSheet { index: 1 }],
            CalculationOptions::default()
        )
        .is_err()
    );
    assert_eq!(book.encode_state_as_update_v1(), state);
}

#[test]
fn pending_effects_come_from_the_overrides() {
    let options = CalculationOptions::default();
    let source = Workbook::from_model(base()).unwrap().save().unwrap();
    let mut book = Workbook::open_collaborative(&source, 3041).unwrap();
    let effects = |book: &Workbook| {
        let effects: Vec<serde_json::Value> =
            serde_json::from_str(&book.pending_effects_json().unwrap()).unwrap();
        effects
            .iter()
            .map(|effect| {
                let field = |name: &str| effect[name].as_str().unwrap_or_default().to_owned();
                [
                    field("kind"),
                    field("operation"),
                    field("label"),
                    field("before"),
                    field("after"),
                ]
                .join("|")
            })
            .collect::<Vec<_>>()
    };
    assert!(effects(&book).is_empty());
    book.edit_cell(SheetId(0), at("A1"), "15", options).unwrap();
    book.edit_cell(SheetId(0), at("C1"), "new", options)
        .unwrap();
    apply(
        &mut book,
        Op::PatchRangeStyle {
            sheet: SheetId(0),
            range: CellRange::new(at("A1"), at("B2")),
            patch: StylePatch {
                bold: Some(true),
                ..StylePatch::default()
            },
        },
    );
    apply(
        &mut book,
        Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 2,
        },
    );
    apply(
        &mut book,
        Op::DeleteRows {
            sheet: SheetId(0),
            at: 3,
            count: 1,
        },
    );
    apply(
        &mut book,
        Op::RenameSheet {
            sheet: SheetId(1),
            name: "Totals".into(),
        },
    );
    let mut listed = effects(&book);
    listed.sort();
    assert_eq!(
        listed,
        [
            r#"text|add|Data!C3||{"kind":"text","value":"new"}"#,
            "text|add|Data: 2 rows inserted||",
            r#"text|remove|Data: 1 rows deleted|{"kind":"number","value":20.0}|"#,
            r#"text|replace|Data!A3|{"kind":"number","value":10.0}|{"kind":"number","value":15.0}"#,
            "text|replace|Sheet Totals|Summary|Totals",
            "visual|replace|Data!A3:B3 formatting||",
        ]
    );
    let json: Vec<serde_json::Value> =
        serde_json::from_str(&book.pending_effects_json().unwrap()).unwrap();
    assert!(
        json.iter().any(|effect| effect["id"]
            == format!("{}", book.cell_identity(SheetId(0), at("A3")).unwrap()))
    );
    book.edit_cell(SheetId(0), at("C3"), "", options).unwrap();
    assert_eq!(effects(&book).len(), 5);
    for _ in 0..7 {
        book.undo(options).unwrap();
    }
    assert!(effects(&book).is_empty());
}
