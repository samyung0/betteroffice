//! XLSX replicas: a room's source opened for reading pending effects
//! ([`EffectsReader`]), kept between saves so a save does not reopen the
//! workbook. The rules are office-checkpoint.ts's: a replica never applies a
//! state, so it serves any state of its base; a call on a replica that fails
//! drops it; under the budget a new replica pushes out least recently used
//! ones only once they have been idle for [`REPLICA_IDLE`], else it is not
//! kept. Calls may now run on several threads: a replica in use is checked
//! out, and another call for its room opens a reader of its own.

use std::sync::{Mutex, MutexGuard};
use std::time::Duration;

use indexmap::IndexMap;

use crate::error::Result;
use crate::types::{Checkpoint, NetEffect};
use crate::xlsx::{self, EffectsReader};

/// Estimated memory a replica holds per byte of its unzipped package. The TS
/// measured 16 as WASM linear memory (`XLSX_HEAP_PER_UNZIPPED_BYTE`); a
/// native replica keeps about 21.6 (8.1 MiB gradebook, process working set
/// after the miss, Windows; API.md "XLSX replicas").
pub const HEAP_PER_UNZIPPED_BYTE: u64 = 22;
/// How long a replica must have gone unused before a new one may push it out.
pub const REPLICA_IDLE: Duration = Duration::from_secs(120);

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct ReplicaStats {
    pub replicas: usize,
    pub replica_bytes: u64,
    pub hits: u64,
    pub misses: u64,
    pub evictions: u64,
}

struct Slot {
    base_sha256: String,
    /// None while a call holds it.
    reader: Option<EffectsReader>,
    bytes: u64,
    /// `env::perf_ms()` of its last use.
    used_at: f64,
    generation: u64,
}

#[derive(Default)]
struct State {
    budget: u64,
    bytes: u64,
    /// Recency order: the first entry is the least recently used.
    slots: IndexMap<String, Slot>,
    generation: u64,
    stats: ReplicaStats,
}

impl State {
    fn remove(&mut self, room: &str) -> Option<Slot> {
        let slot = self.slots.shift_remove(room)?;
        self.bytes -= slot.bytes;
        Some(slot)
    }

    /// Drops least recently used replicas until `needed` more bytes fit; with
    /// `idle`, only those unused for that long. Whether they fit.
    fn evict(&mut self, needed: u64, idle: Option<Duration>) -> bool {
        let now = crate::env::perf_ms();
        while self.bytes + needed > self.budget {
            let Some((room, slot)) = self.slots.first() else {
                break;
            };
            if idle.is_some_and(|idle| now - slot.used_at < idle.as_secs_f64() * 1000.0) {
                break;
            }
            let room = room.clone();
            self.remove(&room);
            self.stats.evictions += 1;
        }
        self.bytes + needed <= self.budget
    }
}

/// The replica cache. One per process; `Send + Sync`.
pub struct XlsxReplicas {
    state: Mutex<State>,
}

/// A replica taken out of the cache for one call; dropping it without
/// [`Checkout::restore`] drops the replica.
struct Checkout<'a> {
    replicas: &'a XlsxReplicas,
    room: &'a str,
    generation: u64,
}

impl Checkout<'_> {
    /// A new replica counts as used once its reader is open.
    fn opened(&self) {
        let mut state = self.replicas.lock();
        if let Some(slot) = state.slots.get_mut(self.room)
            && slot.generation == self.generation
        {
            slot.used_at = crate::env::perf_ms();
        }
    }

    fn restore(self, reader: EffectsReader) {
        let mut state = self.replicas.lock();
        if let Some(slot) = state.slots.get_mut(self.room)
            && slot.generation == self.generation
        {
            slot.reader = Some(reader);
        }
        std::mem::forget(self);
    }
}

impl Drop for Checkout<'_> {
    fn drop(&mut self) {
        let mut state = self.replicas.lock();
        if state
            .slots
            .get(self.room)
            .is_some_and(|slot| slot.generation == self.generation)
        {
            state.remove(self.room);
        }
    }
}

enum Lease<'a> {
    /// A held replica.
    Hit(Checkout<'a>, Box<EffectsReader>),
    /// A new replica to open and keep.
    Opening(Checkout<'a>),
    /// A reader for this call only.
    Transient,
}

impl XlsxReplicas {
    /// Replicas within `budget_bytes` of estimated heap; 0 keeps none.
    pub fn new(budget_bytes: u64) -> Self {
        Self {
            state: Mutex::new(State {
                budget: budget_bytes,
                ..State::default()
            }),
        }
    }

    fn lock(&self) -> MutexGuard<'_, State> {
        self.state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
    }

    /// Sets the budget and drops replicas down to it (0 drops every one).
    pub fn configure(&self, budget_bytes: u64) {
        let mut state = self.lock();
        state.budget = budget_bytes;
        state.evict(0, None);
    }

    /// Drops a room's replica (the room unloaded).
    pub fn drop_room(&self, room: &str) {
        self.lock().remove(room);
    }

    pub fn stats(&self) -> ReplicaStats {
        let state = self.lock();
        ReplicaStats {
            replicas: state.slots.len(),
            replica_bytes: state.bytes,
            ..state.stats
        }
    }

    fn lease<'a>(&'a self, room: &'a str, base: &[u8], base_sha256: &str) -> Lease<'a> {
        let mut state = self.lock();
        let now = crate::env::perf_ms();
        if let Some(held) = state.slots.get_mut(room)
            && held.base_sha256 == base_sha256
            && held.reader.is_some()
        {
            held.used_at = now;
            let reader = held.reader.take();
            let generation = held.generation;
            let index = state.slots.get_index_of(room).unwrap_or_default();
            let last = state.slots.len() - 1;
            state.slots.move_index(index, last);
            state.stats.hits += 1;
            if let Some(reader) = reader {
                let checkout = Checkout {
                    replicas: self,
                    room,
                    generation,
                };
                return Lease::Hit(checkout, Box::new(reader));
            }
        }
        if state
            .slots
            .get(room)
            .is_some_and(|held| held.reader.is_some() || held.base_sha256 != base_sha256)
        {
            state.remove(room);
        }
        state.stats.misses += 1;
        if state.slots.contains_key(room) {
            return Lease::Transient;
        }
        let Some(unzipped) = crate::common::unzipped_bytes(base) else {
            return Lease::Transient;
        };
        let bytes = unzipped * HEAP_PER_UNZIPPED_BYTE;
        if bytes > state.budget || !state.evict(bytes, Some(REPLICA_IDLE)) {
            return Lease::Transient;
        }
        state.generation += 1;
        let generation = state.generation;
        state.slots.insert(
            room.to_owned(),
            Slot {
                base_sha256: base_sha256.to_owned(),
                reader: None,
                bytes,
                used_at: now,
                generation,
            },
        );
        state.bytes += bytes;
        Lease::Opening(Checkout {
            replicas: self,
            room,
            generation,
        })
    }

    /// office-checkpoint.ts `xlsxPendingEffects` with a room: the effects off
    /// the room's replica, or off a reader opened for this call. A state the
    /// reader would not adopt whole goes to a fresh session.
    pub fn pending_effects(
        &self,
        base: &[u8],
        checkpoint: Checkpoint,
        room: &str,
    ) -> Result<Vec<NetEffect>> {
        xlsx::check(base, checkpoint)?;
        if self.lock().budget == 0 || room.is_empty() {
            return xlsx::pending_effects(base, checkpoint, None);
        }
        let read = |reader: &EffectsReader| reader.effects(checkpoint.state);
        let effects = match self.lease(room, base, checkpoint.base_sha256) {
            Lease::Hit(checkout, reader) => {
                let effects = read(&reader)?;
                checkout.restore(*reader);
                effects
            }
            Lease::Opening(checkout) => {
                let reader = EffectsReader::open(base)?;
                checkout.opened();
                let effects = read(&reader)?;
                checkout.restore(reader);
                effects
            }
            Lease::Transient => read(&EffectsReader::open(base)?)?,
        };
        match effects {
            Some(effects) => Ok(effects),
            None => xlsx::XlsxSession::open(base, Some(checkpoint.state))?.effects(),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::env::{Env, with_env};
    use crate::types::Format;

    const SAMPLE: &[u8] = include_bytes!("../../../apps/demo/public/sample.xlsx");
    const SHOWCASE: &[u8] = include_bytes!("../../../apps/demo/public/showcase.xlsx");

    struct Source {
        base: &'static [u8],
        sha: String,
        state: Vec<u8>,
    }

    impl Source {
        fn new(base: &'static [u8]) -> Self {
            Self {
                base,
                sha: crate::common::sha256_hex(base),
                state: crate::seed(Format::Xlsx, base).expect("seeds"),
            }
        }

        fn checkpoint(&self) -> Checkpoint<'_> {
            self.checkpoint_with(&self.state)
        }

        fn checkpoint_with<'a>(&'a self, state: &'a [u8]) -> Checkpoint<'a> {
            Checkpoint {
                format: Format::Xlsx,
                schema_version: 1,
                base_sha256: &self.sha,
                state,
            }
        }

        /// Security regression (office-native REVIEW1 D3): the estimate a
        /// replica is admitted by is the measured native heap, 22 bytes per
        /// unzipped byte (Epo 65a9ae87); at 16 the budget let the engine
        /// process keep about 1.4 times the memory it was given.
        fn bytes(&self) -> u64 {
            crate::common::unzipped_bytes(self.base).expect("a package") * 22
        }
    }

    /// `performance.now()` at `ms` for one call.
    fn at<T>(ms: f64, call: impl FnOnce() -> T) -> T {
        let env = Env {
            perf_ms: Some(ms),
            ..Env::default()
        };
        with_env(env, call)
    }

    fn effects(replicas: &XlsxReplicas, source: &Source, room: &str, ms: f64) {
        at(ms, || {
            replicas.pending_effects(source.base, source.checkpoint(), room)
        })
        .expect("effects");
    }

    #[test]
    fn a_replica_serves_its_room_until_it_does_not_fit() {
        let sample = Source::new(SAMPLE);
        let replicas = XlsxReplicas::new(1 << 30);
        effects(&replicas, &sample, "a", 0.0);
        effects(&replicas, &sample, "a", 1.0);
        let stats = replicas.stats();
        assert_eq!((stats.hits, stats.misses, stats.replicas), (1, 1, 1));
        assert_eq!(stats.replica_bytes, sample.bytes());
        // Another base for the room replaces its replica.
        let showcase = Source::new(SHOWCASE);
        effects(&replicas, &showcase, "a", 2.0);
        assert_eq!(replicas.stats().replica_bytes, showcase.bytes());
        // A replica larger than the budget is not kept.
        replicas.configure(sample.bytes() - 1);
        assert_eq!(replicas.stats().replicas, 0);
        effects(&replicas, &sample, "a", 3.0);
        assert_eq!(replicas.stats().replicas, 0);
    }

    #[test]
    fn a_new_replica_waits_for_the_least_recently_used_to_go_idle() {
        let sample = Source::new(SAMPLE);
        let replicas = XlsxReplicas::new(sample.bytes());
        effects(&replicas, &sample, "a", 0.0);
        // "a" was used within REPLICA_IDLE: "b" reads without keeping a replica.
        effects(&replicas, &sample, "b", 60_000.0);
        let stats = replicas.stats();
        assert_eq!((stats.misses, stats.evictions, stats.replicas), (2, 0, 1));
        effects(&replicas, &sample, "a", 61_000.0);
        assert_eq!(replicas.stats().hits, 1);
        // Idle long enough, "a" makes room for "b".
        effects(&replicas, &sample, "b", 61_000.0 + 120_000.0);
        effects(&replicas, &sample, "b", 182_000.0);
        let stats = replicas.stats();
        assert_eq!((stats.evictions, stats.replicas, stats.hits), (1, 1, 2));
    }

    #[test]
    fn a_room_dropped_or_a_budget_cut_during_a_call_loses_its_replica() {
        let sample = Source::new(SAMPLE);
        let replicas = XlsxReplicas::new(1 << 30);
        for change in [
            (|replicas: &XlsxReplicas| replicas.drop_room("a")) as fn(&XlsxReplicas),
            |replicas| replicas.configure(0),
        ] {
            replicas.configure(1 << 30);
            let Lease::Opening(checkout) = replicas.lease("a", sample.base, &sample.sha) else {
                panic!("a new replica opens");
            };
            change(&replicas);
            checkout.restore(EffectsReader::open(sample.base).expect("opens"));
            assert_eq!(replicas.stats().replicas, 0);
        }
    }

    #[test]
    fn a_replica_in_use_leaves_other_calls_a_reader_of_their_own() {
        let sample = Source::new(SAMPLE);
        let replicas = XlsxReplicas::new(1 << 30);
        effects(&replicas, &sample, "a", 0.0);
        let Lease::Hit(checkout, reader) = replicas.lease("a", sample.base, &sample.sha) else {
            panic!("the replica is held");
        };
        // Another thread's call for the room meanwhile: a transient reader.
        std::thread::scope(|scope| {
            scope.spawn(|| effects(&replicas, &sample, "a", 1.0));
        });
        assert_eq!(replicas.stats().misses, 2);
        checkout.restore(*reader);
        effects(&replicas, &sample, "a", 2.0);
        let stats = replicas.stats();
        assert_eq!((stats.hits, stats.replicas), (2, 1));
    }

    #[test]
    fn a_call_that_fails_or_panics_on_its_replica_drops_it() {
        let sample = Source::new(SAMPLE);
        let replicas = XlsxReplicas::new(1 << 30);
        effects(&replicas, &sample, "a", 0.0);
        // A state the reader does not adopt fails in the fresh session; the
        // replica served the call and stays (office-checkpoint.ts too).
        let broken = [0_u8, 1, 2, 3];
        let read = at(1.0, || {
            replicas.pending_effects(sample.base, sample.checkpoint_with(&broken), "a")
        });
        assert!(read.is_err());
        assert_eq!(replicas.stats().replicas, 1);
        // A read that errs or panics returns no reader: the replica goes.
        let Lease::Hit(checkout, _reader) = replicas.lease("a", sample.base, &sample.sha) else {
            panic!("the replica is held");
        };
        drop(checkout);
        assert_eq!(replicas.stats().replicas, 0);
        effects(&replicas, &sample, "a", 2.0);
        let held = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            let Lease::Hit(_checkout, _reader) = replicas.lease("a", sample.base, &sample.sha)
            else {
                panic!("the replica is held");
            };
            panic!("the engine panics");
        }));
        assert!(held.is_err());
        assert_eq!(replicas.stats().replicas, 0);
        effects(&replicas, &sample, "a", 3.0);
        assert_eq!(replicas.stats().replicas, 1);
    }
}
