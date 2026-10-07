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

/// Column A holds values only and every formula reads column A or a name over
/// it, so no edit makes a cycle; with `spill`, a third sheet holds a source
/// array over `Data!A1:A3` with its cached results, as a saved file does.
fn cycle_free_base(spill: bool) -> WorkbookModel {
    let at = |a1: &str| CellRef::parse_a1(a1).unwrap();
    let mut data = Sheet::new("Data");
    for row in 0..6 {
        data.set_cell(CellRef::new(row, 0), cell(number(f64::from(row + 1)), None));
        data.set_cell(
            CellRef::new(row, 1),
            cell(CellValue::Empty, Some(&format!("A{}*2", row + 1))),
        );
    }
    data.set_cell(at("C1"), cell(CellValue::Empty, Some("SUM(A1:A6)")));
    data.set_cell(at("D1"), cell(CellValue::Empty, Some("SUM(Range)+NOW()*0")));
    let mut summary = Sheet::new("Summary");
    summary.set_cell(at("A1"), cell(number(3.0), None));
    summary.set_cell(at("B1"), cell(CellValue::Empty, Some("Data!A2+A1")));
    summary.set_cell(at("C1"), cell(CellValue::Empty, Some("SUM(Data!A1:A4)")));
    let mut arrays = Sheet::new("Arrays");
    if spill {
        arrays.set_cell(at("A1"), cell(number(10.0), Some("Data!A1:A3*10")));
        arrays.set_cell(at("A2"), cell(number(20.0), None));
        arrays.set_cell(at("A3"), cell(number(30.0), None));
        arrays.set_array_formula(at("A1"), CellRange::parse_a1("A1:A3").unwrap());
    }
    arrays.set_cell(at("C1"), cell(CellValue::Empty, Some("SUM(Data!A1:A6)")));
    WorkbookModel {
        sheets: vec![data, summary, arrays],
        defined_names: vec![xlsx_model::DefinedName {
            name: "Range".into(),
            formula: "Data!$A$1:$A$6".into(),
            local_sheet: None,
            hidden: false,
        }],
        ..WorkbookModel::default()
    }
}

const VALUES: [&str; 7] = ["", "5", "-2", "text", "TRUE", "17", "0.5"];
const READS_OF_A: [&str; 8] = [
    "=A1+1",
    "=SUM(A1:A6)",
    "=Data!A3*3",
    "=SUM(Range)",
    "=IF(A4>2,A5,A6)",
    "=COUNT(Data!A1:A6)",
    "=A2&\"x\"",
    "=SUMPRODUCT(A1:A3,A4:A6)",
];

/// A local step that cannot make a cycle; `history` adds Undo and Redo.
fn cycle_free_step(book: &mut Workbook, rng: &mut Rng, history: bool) {
    let options = CalculationOptions {
        now_serial: Some(45_000.25),
    };
    let row = rng.below(7) as u32;
    let kind = rng.below(if history { 24 } else { 20 });
    let sheet = SheetId(rng.below(3) as u32);
    let _ = match kind {
        // values into column A, the array's anchor and cells under it included
        0..=8 | 19 => book.edit_cell(
            sheet,
            CellRef::new(row.min(5), 0),
            rng.pick(&VALUES),
            options,
        ),
        9..=15 if sheet.0 < 2 => book.edit_cell(
            sheet,
            CellRef::new(row, 1 + rng.below(4) as u32),
            rng.pick(&READS_OF_A),
            options,
        ),
        9..=15 => book.edit_cell(
            sheet,
            CellRef::new(row, 1 + rng.below(4) as u32),
            rng.pick(&VALUES),
            options,
        ),
        16 => book.edit_cells(
            sheet,
            &[
                CellInput {
                    cell: CellRef::new(row, 0),
                    input: (*rng.pick(&VALUES)).to_owned(),
                },
                CellInput {
                    cell: CellRef::new(row + 1, 0),
                    input: (*rng.pick(&VALUES)).to_owned(),
                },
            ],
            options,
        ),
        17 => book.apply_ops(
            vec![Op::PatchRangeStyle {
                sheet,
                range: CellRange::new(CellRef::new(row, 0), CellRef::new(row + 1, 1)),
                patch: xlsx_ops::StylePatch {
                    bold: Some(rng.below(2) == 0),
                    ..xlsx_ops::StylePatch::default()
                },
            }],
            options,
        ),
        18 => book.apply_ops(
            vec![Op::InsertRows {
                sheet,
                at: row,
                count: 1,
            }],
            options,
        ),
        20..=22 => book.undo(options),
        _ => book.redo(options),
    };
}

type Sink = (UpdateSubscription, Arc<Mutex<Vec<UpdateEvent>>>);

/// Queues each peer's local updates for every other peer.
fn forward(sinks: &[Sink], inbox: &mut [Vec<(usize, Vec<u8>)>]) {
    for (from, (_, sink)) in sinks.iter().enumerate() {
        for update in local_events(sink) {
            for (to, queue) in inbox.iter_mut().enumerate() {
                if to != from {
                    queue.push((from, update.clone()));
                }
            }
        }
    }
}

fn local_events(sink: &Arc<Mutex<Vec<UpdateEvent>>>) -> Vec<Vec<u8>> {
    drain(sink)
        .into_iter()
        .filter(|event| matches!(event.origin, UpdateOrigin::Local))
        .map(|event| event.update)
        .collect()
}

/// Three peers edit concurrently and hear each other in any order a room can
/// deliver. Peer 1 takes remote cell edits incrementally and its twin takes
/// every one whole; they must agree after every step, and at the end every
/// peer must equal a replica that projects the final state whole. Two seeds
/// in three add a source array, and one of those shuffles delivery through
/// the pending queue; the rest deliver each sender in order without arrays,
/// the shape the shortcut is for, and must take it at least four times as
/// often as the whole path.
#[test]
fn concurrent_cycle_free_peers_take_the_shortcut_and_converge() {
    let options = CalculationOptions {
        now_serial: Some(45_000.25),
    };
    let fresh = |client_id, spill| {
        let mut book =
            Workbook::from_model_collaborative(cycle_free_base(spill), client_id).unwrap();
        book.recalculate_all(options);
        book.mark_exact();
        book
    };
    let (mut incremental, mut whole_path) = (0, 0);
    for seed in 1..=9_u64 {
        let in_order = seed % 3 != 2;
        let spill = seed % 3 != 0;
        let counted = !spill;
        let mut rng = Rng(seed.wrapping_mul(0x9e37_79b9_7f4a_7c15) | 1);
        let mut peers = vec![fresh(8000, spill), fresh(8001, spill), fresh(8002, spill)];
        let mut twin = fresh(8001, spill);
        twin.whole_local_commits = true;
        let sinks = peers.iter().map(observe).collect::<Vec<_>>();
        let (_t, twin_sink) = observe(&twin);
        let mut inbox: Vec<Vec<(usize, Vec<u8>)>> = vec![Vec::new(); 3];
        let mut deliver = |peers: &mut [Workbook],
                           twin: &mut Workbook,
                           to: usize,
                           update: &[u8],
                           context: &str| {
            let before = peers[to].model.clone();
            let result = if to == 1 {
                let result = match peers[1].apply_cell_update(update, options) {
                    Some(result) => {
                        incremental += usize::from(counted);
                        Ok(result)
                    }
                    None => {
                        whole_path += usize::from(counted);
                        peers[1].apply_update_v1(update, options)
                    }
                };
                twin.exact_epoch = None;
                let expected = twin.apply_update_v1(update, options);
                assert_eq!(
                    format!("{result:?}"),
                    format!("{expected:?}"),
                    "{context}: result"
                );
                assert_twins(&peers[1], twin, context);
                result
            } else {
                peers[to].apply_update_v1(update, options)
            };
            assert_eq!(
                result
                    .unwrap_or_else(|error| panic!("{context}: {error:?}"))
                    .changed,
                changed_cells_between(&before, &peers[to].model),
                "{context}: changed"
            );
        };
        for step in 0..300 {
            let context = format!("seed {seed} step {step}");
            if rng.below(3) == 0 {
                let who = rng.below(3) as usize;
                if who == 1 {
                    let mut copy = Rng(rng.next());
                    let mut same = Rng(copy.0);
                    cycle_free_step(&mut peers[1], &mut copy, false);
                    cycle_free_step(&mut twin, &mut same, false);
                    assert_twins(&peers[1], &twin, &context);
                    let sent = drain(&sinks[1].1);
                    assert_eq!(sent, drain(&twin_sink), "{context}: local update");
                    sinks[1].1.lock().unwrap().extend(sent);
                } else {
                    cycle_free_step(&mut peers[who], &mut rng, true);
                }
            } else {
                let to = rng.below(3) as usize;
                if !inbox[to].is_empty() {
                    let mut index = rng.below(inbox[to].len() as u64) as usize;
                    if in_order {
                        let from = inbox[to][index].0;
                        index = inbox[to]
                            .iter()
                            .position(|(sender, _)| *sender == from)
                            .unwrap();
                    }
                    let (_, update) = inbox[to].remove(index);
                    deliver(&mut peers, &mut twin, to, &update, &context);
                }
            }
            drain(&twin_sink);
            forward(&sinks, &mut inbox);
        }
        loop {
            forward(&sinks, &mut inbox);
            if inbox.iter().all(Vec::is_empty) {
                break;
            }
            for (to, queue) in inbox.iter_mut().enumerate() {
                for (_, update) in std::mem::take(queue) {
                    deliver(
                        &mut peers,
                        &mut twin,
                        to,
                        &update,
                        &format!("seed {seed} drain"),
                    );
                }
            }
            drain(&twin_sink);
        }
        let state = peers[0].encode_state_as_update_v1();
        let mut oracle = fresh(8999, spill);
        oracle.apply_update_v1(&state, options).unwrap();
        for (index, peer) in peers.iter().enumerate() {
            assert_eq!(
                peer.encode_state_vector_v1(),
                oracle.encode_state_vector_v1(),
                "seed {seed}: peer {index} state"
            );
            assert_eq!(
                peer.model, oracle.model,
                "seed {seed}: peer {index} converged"
            );
        }
    }
    assert!(
        incremental >= 4 * whole_path,
        "incremental {incremental}, whole {whole_path}"
    );
}
