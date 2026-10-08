//! The slide order under concurrent slide edits and Undo: readers show the
//! first entry of each slide with a record, and a write deletes only entries
//! of the slide it acts on, so no peer's cleanup can drop a slide another
//! peer restored.

use std::collections::HashSet;

use pptx_edit::{DeckSession, EditCtx, UndoCaptureMode};
use yrs::{Array, Map, ReadTxn, Transact};

const DEMO: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

fn ctx() -> EditCtx {
    EditCtx::local("test")
}

fn sync(from: &DeckSession, to: &DeckSession) {
    let diff = from.encode_diff_v1(&to.encode_state_vector_v1()).unwrap();
    to.apply_update_v1(&diff).unwrap();
}

/// A plain yrs delivery, without the engine's staged validation, for the
/// random schedules: the order rules alone decide what each peer shows.
fn deliver(from: &DeckSession, to: &DeckSession) {
    use yrs::Update;
    use yrs::updates::decoder::Decode;
    let known = to.yrs_doc().transact().state_vector();
    let diff = from.yrs_doc().transact().encode_diff_v1(&known);
    to.yrs_doc()
        .transact_mut_with("remote")
        .apply_update(Update::decode_v1(&diff).unwrap())
        .unwrap();
}

fn sync_all(peers: &[&DeckSession]) {
    for from in peers {
        for to in peers {
            if !std::ptr::eq(*from, *to) {
                sync(from, to);
            }
        }
    }
}

fn peer(client: u64) -> DeckSession {
    let session = DeckSession::open(DEMO, client).unwrap();
    session.set_undo_capture_mode(UndoCaptureMode::Manual);
    session
}

/// Slide records that no order entry lists: slides no peer can show.
fn hidden_slides(session: &DeckSession) -> Vec<String> {
    let txn = session.yrs_doc().transact();
    let order: HashSet<String> = txn
        .get_array("pptx:slide-order")
        .unwrap()
        .iter(&txn)
        .map(|value| value.to_string(&txn))
        .collect();
    let slides = txn.get_map("pptx:slides").unwrap();
    let mut hidden: Vec<String> = slides
        .keys(&txn)
        .filter(|id| !order.contains(*id))
        .map(str::to_owned)
        .collect();
    hidden.sort();
    hidden
}

/// REVIEW2's case: A deletes a slide while B moves it, and A undoes the
/// delete. A then moves one slide while C, which has not seen the Undo,
/// moves another. Compacting writers dropped the restored slide's two
/// entries from both sides, one as a repeat and one as stale.
#[test]
fn a_slide_undo_restored_survives_concurrent_moves() {
    let (a, b, c) = (peer(901), peer(902), peer(903));
    let before = a.slide_ids().unwrap();
    let restored = before[2].clone();
    a.delete_slide(&ctx(), &restored).unwrap();
    a.add_undo_barrier();
    b.move_slide(&ctx(), &restored, 0).unwrap();
    b.add_undo_barrier();
    sync_all(&[&a, &b, &c]);
    assert!(!a.slide_ids().unwrap().contains(&restored));

    assert!(a.undo());
    assert!(a.slide_ids().unwrap().contains(&restored));
    let shown = a.slide_ids().unwrap();
    a.move_slide(&ctx(), &shown[1], 2).unwrap();
    let shown = c.slide_ids().unwrap();
    c.move_slide(&ctx(), &shown[0], 1).unwrap();
    sync_all(&[&a, &b, &c]);

    let ids = a.slide_ids().unwrap();
    assert!(ids.contains(&restored), "{ids:?}");
    assert_eq!(ids.len(), before.len());
    for session in [&b, &c] {
        assert_eq!(session.snapshot().unwrap(), a.snapshot().unwrap());
    }
    assert_eq!(hidden_slides(&a), Vec::<String>::new());
}

struct Rng(u64);

impl Rng {
    fn next(&mut self) -> u64 {
        self.0 ^= self.0 << 13;
        self.0 ^= self.0 >> 7;
        self.0 ^= self.0 << 17;
        self.0
    }

    fn below(&mut self, n: usize) -> usize {
        if n == 0 {
            0
        } else {
            (self.next() % n as u64) as usize
        }
    }
}

/// One slide insert, delete or move, Undo or Redo, on the slides `session`
/// shows.
fn random_step(rng: &mut Rng, session: &DeckSession) {
    let ids = session.slide_ids().unwrap();
    let pick = |rng: &mut Rng| ids[rng.below(ids.len())].clone();
    match rng.below(10) {
        0 | 1 => {
            let at = rng.below(ids.len() + 1) as u32;
            session.insert_slide(&ctx(), at, None).unwrap();
        }
        2 if ids.len() > 1 => {
            session.delete_slide(&ctx(), &pick(rng)).unwrap();
        }
        3 | 4 if !ids.is_empty() => {
            let to = rng.below(ids.len()) as u32;
            session.move_slide(&ctx(), &pick(rng), to).unwrap();
        }
        5..=7 => {
            session.undo();
        }
        _ => {
            session.redo();
        }
    }
    session.add_undo_barrier();
}

/// Random slide schedules on three peers with Undo and Redo, delivered
/// pairwise in random orders. After every delivery no peer holds a slide
/// record that no order entry lists; at the end the peers take each other's
/// updates through the engine and converge. `SLIDE_SCHEDULES=2000` runs the
/// longer scan.
#[test]
fn random_slide_schedules_lose_no_slide() {
    let schedules: u64 = std::env::var("SLIDE_SCHEDULES").map_or(60, |n| n.parse().unwrap());
    let state = peer(700).encode_state_as_update_v1();
    for schedule in 0..schedules {
        let mut rng = Rng(schedule.wrapping_mul(0x9E37_79B9_7F4A_7C15) | 1);
        let peers: Vec<DeckSession> = (701..704)
            .map(|client| {
                let session =
                    DeckSession::open_from_update_with_source(&state, DEMO, client).unwrap();
                session.set_undo_capture_mode(UndoCaptureMode::Manual);
                session
            })
            .collect();
        for round in 0..6 {
            for _ in 0..2 {
                for session in &peers {
                    random_step(&mut rng, session);
                }
            }
            for _ in 0..4 {
                let (from, to) = (rng.below(3), rng.below(3));
                if from != to {
                    deliver(&peers[from], &peers[to]);
                    for session in &peers {
                        let hidden = hidden_slides(session);
                        assert!(
                            hidden.is_empty(),
                            "schedule {schedule} round {round}: {hidden:?}"
                        );
                    }
                }
            }
        }
        sync_all(&peers.iter().collect::<Vec<_>>());
        for session in &peers[1..] {
            assert_eq!(
                session.slide_ids().unwrap(),
                peers[0].slide_ids().unwrap(),
                "schedule {schedule}"
            );
        }
        assert_eq!(hidden_slides(&peers[0]), Vec::<String>::new());
    }
}
