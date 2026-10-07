//! A remote cell edit applied incrementally must leave the replica exactly as
//! the whole projection and recalculation leave its twin: same model, result,
//! emitted update, shared state and history. The twin also commits its own
//! cell edits through the whole projection, so the replica's cell-only commits
//! must emit the same update bytes and leave the same model and history.
use super::*;
use xlsx_model::Cell;

struct Rng(u64);

impl Rng {
    fn next(&mut self) -> u64 {
        self.0 ^= self.0 << 13;
        self.0 ^= self.0 >> 7;
        self.0 ^= self.0 << 17;
        self.0
    }

    fn below(&mut self, n: u64) -> u64 {
        self.next() % n
    }

    fn pick<'a, T>(&mut self, items: &'a [T]) -> &'a T {
        &items[self.below(items.len() as u64) as usize]
    }
}

fn cell(value: CellValue, formula: Option<&str>) -> Cell {
    Cell {
        value,
        formula: formula.map(str::to_owned),
        style: None,
    }
}

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

/// Values, cross-sheet formulas, a defined name, a function the engine leaves
/// at its cached value, and a third sheet reading the first, through an array
/// formula when `spill`. A spill keeps local edits on the whole projection.
fn base(spill: bool) -> WorkbookModel {
    let at = |a1: &str| CellRef::parse_a1(a1).unwrap();
    let mut data = Sheet::new("Data");
    for row in 0..4 {
        data.set_cell(CellRef::new(row, 0), cell(number(f64::from(row + 1)), None));
        data.set_cell(
            CellRef::new(row, 1),
            cell(CellValue::Empty, Some(&format!("A{}*2", row + 1))),
        );
    }
    data.set_cell(at("C1"), cell(CellValue::Empty, Some("SUM(A1:A4)")));
    data.set_cell(at("E1"), cell(CellValue::Empty, Some("Summary!A1+1")));
    data.set_cell(at("F1"), cell(number(7.0), Some("FOOBAR(A1)")));
    let mut summary = Sheet::new("Summary");
    summary.set_cell(at("A1"), cell(CellValue::Empty, Some("SUM(Data!B1:B4)")));
    summary.set_cell(at("A2"), cell(CellValue::Empty, Some("Data!C1*2")));
    summary.set_cell(at("B1"), cell(CellValue::Empty, Some("SUM(Range)")));
    let mut arrays = Sheet::new("Arrays");
    if spill {
        // Cached results under the array, as a saved file holds them.
        arrays.set_cell(at("A1"), cell(number(10.0), Some("Data!A1:A3*10")));
        arrays.set_cell(at("A2"), cell(number(20.0), None));
        arrays.set_cell(at("A3"), cell(number(30.0), None));
        arrays.set_array_formula(at("A1"), CellRange::parse_a1("A1:A3").unwrap());
    } else {
        arrays.set_cell(at("A1"), cell(CellValue::Empty, Some("SUM(Data!A1:A3)*10")));
    }
    arrays.set_cell(at("B1"), cell(CellValue::Empty, Some("SUM(A1:A4)")));
    WorkbookModel {
        sheets: vec![data, summary, arrays],
        defined_names: vec![xlsx_model::DefinedName {
            name: "Range".into(),
            formula: "Data!$A$1:$A$4".into(),
            local_sheet: None,
            hidden: false,
        }],
        ..WorkbookModel::default()
    }
}

const INPUTS: [&str; 20] = [
    "",
    "5",
    "42",
    "-3.5",
    "text",
    "TRUE",
    "=A1+1",
    "=B2*2",
    "=SUM(A1:A5)",
    "=C3",
    "=Summary!A1+Data!A2",
    "=A1&\"x\"",
    "=IF(A1>2,1,0)",
    "=Arrays!A2+1",
    "=SUM(Arrays!A1:A4)",
    "=A2",
    "=B1+C1",
    "=FOOBAR(B1)",
    "=Summary!B1",
    "=1/0",
];

fn random_input(rng: &mut Rng) -> String {
    if rng.below(4) == 0 {
        rng.below(100).to_string()
    } else {
        (*rng.pick(&INPUTS)).to_owned()
    }
}

fn random_cell(rng: &mut Rng) -> (SheetId, CellRef) {
    (
        {
            let sheets = if rng.below(6) == 0 { 3 } else { 2 };
            SheetId(rng.below(sheets) as u32)
        },
        CellRef::new(rng.below(6) as u32, rng.below(6) as u32),
    )
}

/// A local step. Only the author steps through history: Yrs restores a
/// multi-item step in hash order, so twins would number its items differently.
fn local_step(book: &mut Workbook, rng: &mut Rng, history: bool) {
    let options = CalculationOptions::default();
    let (sheet, at) = random_cell(rng);
    let _ = match rng.below(if history { 20 } else { 16 }) {
        0..=12 => book.edit_cell(sheet, at, &random_input(rng), options),
        13 => book.edit_cells(
            sheet,
            &[
                CellInput {
                    cell: at,
                    input: random_input(rng),
                },
                CellInput {
                    cell: CellRef::new(at.row + 1, at.col),
                    input: random_input(rng),
                },
            ],
            options,
        ),
        14 => book.apply_ops(
            vec![Op::PatchRangeStyle {
                sheet,
                range: CellRange::new(at, CellRef::new(at.row + 1, at.col)),
                patch: xlsx_ops::StylePatch {
                    bold: Some(rng.below(2) == 0),
                    italic: (rng.below(3) == 0).then_some(true),
                    ..xlsx_ops::StylePatch::default()
                },
            }],
            options,
        ),
        15 => book.apply_ops(
            vec![if rng.below(2) == 0 {
                Op::InsertRows {
                    sheet,
                    at: at.row,
                    count: 1,
                }
            } else {
                Op::DeleteRows {
                    sheet,
                    at: at.row,
                    count: 1,
                }
            }],
            options,
        ),
        16..=18 => book.undo(options),
        _ => book.redo(options),
    };
}

fn observe(book: &Workbook) -> (UpdateSubscription, Arc<Mutex<Vec<UpdateEvent>>>) {
    let sink = Arc::new(Mutex::new(Vec::new()));
    let events = sink.clone();
    let subscription = book
        .observe_update_v1(move |event| events.lock().unwrap().push(event))
        .unwrap();
    (subscription, sink)
}

fn drain(sink: &Arc<Mutex<Vec<UpdateEvent>>>) -> Vec<UpdateEvent> {
    std::mem::take(&mut *sink.lock().unwrap())
}

fn replica(client_id: u64, spill: bool) -> Workbook {
    let mut book = Workbook::from_model_collaborative(base(spill), client_id).unwrap();
    book.recalculate_all(CalculationOptions::default());
    book.mark_exact();
    book
}

fn assert_twins(fast: &Workbook, whole: &Workbook, context: &str) {
    assert_eq!(fast.model, whole.model, "{context}: model");
    assert_eq!(
        fast.encode_state_as_update_v1(),
        whole.encode_state_as_update_v1(),
        "{context}: state"
    );
    assert_eq!(
        fast.history_state(),
        whole.history_state(),
        "{context}: history"
    );
    assert_eq!(
        fast.last_calculation, whole.last_calculation,
        "{context}: calculation"
    );
    assert_eq!(
        fast.sheet_info().unwrap(),
        whole.sheet_info().unwrap(),
        "{context}: sheet info"
    );
}

#[test]
fn incremental_remote_cell_edits_equal_the_whole_projection() {
    let options = CalculationOptions::default();
    let mut incremental = 0;
    let mut whole_path = 0;
    for seed in 1..=6_u64 {
        let mut rng = Rng(seed.wrapping_mul(0x9e37_79b9_7f4a_7c15));
        let spill = seed % 2 == 1;
        let mut author = replica(7001, spill);
        let mut fast = replica(7002, spill);
        let mut whole = replica(7002, spill);
        whole.whole_local_commits = true;
        let (_a, author_events) = observe(&author);
        let (_f, fast_events) = observe(&fast);
        let (_w, whole_events) = observe(&whole);
        for step in 0..200 {
            let context = format!("seed {seed} step {step}");
            if rng.below(5) == 0 {
                let mut twin = Rng(rng.next());
                let mut copy = Rng(twin.0);
                local_step(&mut fast, &mut twin, false);
                local_step(&mut whole, &mut copy, false);
                let sent = drain(&fast_events);
                assert_eq!(sent, drain(&whole_events), "{context}: local update");
                assert_twins(&fast, &whole, &context);
                for event in sent {
                    author.apply_update_v1(&event.update, options).unwrap();
                }
                drain(&author_events);
                continue;
            }
            local_step(&mut author, &mut rng, true);
            for event in drain(&author_events) {
                let before = fast.model.clone();
                let result = match fast.apply_cell_update(&event.update, options) {
                    Some(result) => {
                        incremental += 1;
                        Ok(result)
                    }
                    None => {
                        whole_path += 1;
                        fast.apply_update_v1(&event.update, options)
                    }
                };
                whole.exact_epoch = None;
                let expected = whole.apply_update_v1(&event.update, options);
                assert_eq!(
                    format!("{result:?}"),
                    format!("{expected:?}"),
                    "{context}: result"
                );
                assert_eq!(
                    drain(&fast_events),
                    drain(&whole_events),
                    "{context}: emitted"
                );
                assert_eq!(
                    result.unwrap().changed,
                    changed_cells_between(&before, &fast.model),
                    "{context}: changed"
                );
                assert_twins(&fast, &whole, &context);
            }
        }
        // The author recalculated its own edits from the edited cells alone,
        // the peers each update whole: both must give the same workbook.
        assert_eq!(
            author.encode_state_vector_v1(),
            fast.encode_state_vector_v1(),
            "seed {seed}: shared state converged"
        );
        if author.model != fast.model {
            for at in changed_cells_between(&author.model, &fast.model) {
                eprintln!(
                    "seed {seed} {}!{}: author {:?} peer {:?}",
                    at.sheet.0,
                    at.cell.to_a1(),
                    author.model.sheet(at.sheet).and_then(|s| s.cell(at.cell)),
                    fast.model.sheet(at.sheet).and_then(|s| s.cell(at.cell))
                );
            }
        }
        assert_eq!(author.model, fast.model, "seed {seed}: converged");
    }
    // A cycle anywhere sends every update down the whole path, and the random
    // formulas make many.
    assert!(
        5 * incremental > whole_path,
        "incremental {incremental}, whole {whole_path}"
    );
}

#[test]
fn cell_commits_on_source_files_emit_the_whole_projections_update() {
    let options = CalculationOptions::default();
    for bytes in [
        &include_bytes!("../../tests/fixtures/storage/course-guide.xlsx")[..],
        &include_bytes!("../../tests/fixtures/storage/cells-1k.xlsx")[..],
    ] {
        let open = || Workbook::open_collaborative_recalculated(bytes, 7003, options).unwrap();
        let (mut fast, mut whole) = (open(), open());
        whole.whole_local_commits = true;
        let (_f, fast_events) = observe(&fast);
        let (_w, whole_events) = observe(&whole);
        let mut rng = Rng(0x5eed);
        for step in 0..40 {
            let context = format!("step {step}");
            let at = CellRef::new(rng.below(12) as u32, rng.below(6) as u32);
            let input = random_input(&mut rng);
            if rng.below(4) == 0 {
                let edits = [
                    CellInput {
                        cell: at,
                        input: input.clone(),
                    },
                    CellInput {
                        cell: CellRef::new(at.row + 1, at.col),
                        input: random_input(&mut rng),
                    },
                ];
                assert_eq!(
                    format!("{:?}", fast.edit_cells(SheetId(0), &edits, options)),
                    format!("{:?}", whole.edit_cells(SheetId(0), &edits, options)),
                    "{context}: result"
                );
            } else {
                assert_eq!(
                    format!("{:?}", fast.edit_cell(SheetId(0), at, &input, options)),
                    format!("{:?}", whole.edit_cell(SheetId(0), at, &input, options)),
                    "{context}: result"
                );
            }
            assert_eq!(
                drain(&fast_events),
                drain(&whole_events),
                "{context}: emitted"
            );
            assert_twins(&fast, &whole, &context);
        }
    }
}

/// A local edit keeps what array formulas spill, through either commit, and
/// an edit under a spill recomputes its anchor: a replica that projects the
/// same state whole agrees after every edit.
#[test]
fn local_edits_keep_spilled_values_as_a_whole_recalculation_does() {
    let options = CalculationOptions::default();
    let at = |a1: &str| CellRef::parse_a1(a1).unwrap();
    for whole_local_commits in [false, true] {
        let mut book = replica(7004, true);
        book.whole_local_commits = whole_local_commits;
        let arrays = SheetId(2);
        for (sheet, cell, input) in [
            (SheetId(0), "A2", "7"),
            (SheetId(0), "F6", "1"),
            (arrays, "A3", "x"),
            (arrays, "D4", "=A2+1"),
        ] {
            book.edit_cell(sheet, at(cell), input, options).unwrap();
            let mut fresh = replica(7005, true);
            fresh
                .apply_update_v1(&book.encode_state_as_update_v1(), options)
                .unwrap();
            assert_eq!(book.model, fresh.model, "{cell} {input}");
        }
        let value = |cell: &str| {
            book.model
                .sheet(arrays)
                .and_then(|sheet| sheet.cell(at(cell)))
                .map(|cell| cell.value.clone())
        };
        assert_eq!(value("A2"), Some(number(70.0)));
        assert_eq!(value("A3"), Some(number(30.0)));
        assert_eq!(value("D4"), Some(number(71.0)));
    }
}
