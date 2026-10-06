//! The two inputs the TS read from its runtime: `randomInt` for Yjs client
//! ids and `Date.now()` for XLSX volatile cells. Tests fix both per call.

use std::cell::RefCell;
use std::collections::hash_map::RandomState;
use std::hash::{BuildHasher, Hasher};
use std::sync::OnceLock;
use std::time::{Instant, SystemTime, UNIX_EPOCH};

/// `randomInt(1, 0x1fffffffffff)`'s upper bound (exclusive).
const CLIENT_LIMIT: u64 = 0x1fff_ffff_ffff;

#[derive(Clone, Debug, Default)]
pub struct Env {
    /// Client ids handed out in order; random when empty.
    pub clients: Vec<u64>,
    /// Milliseconds since the Unix epoch; the system clock when unset.
    pub now_ms: Option<f64>,
    /// `performance.now()` for the replica cache; a monotonic clock when unset.
    pub perf_ms: Option<f64>,
}

thread_local! {
    static ENV: RefCell<Option<Env>> = const { RefCell::new(None) };
}

/// Runs `call` with fixed client ids and clock on this thread.
pub fn with_env<T>(env: Env, call: impl FnOnce() -> T) -> T {
    struct Reset(Option<Env>);
    impl Drop for Reset {
        fn drop(&mut self) {
            ENV.with(|slot| *slot.borrow_mut() = self.0.take());
        }
    }
    let previous = ENV.with(|slot| slot.borrow_mut().replace(env));
    let _reset = Reset(previous);
    call()
}

pub(crate) fn next_client() -> u64 {
    let fixed = ENV.with(|slot| {
        let mut slot = slot.borrow_mut();
        let env = slot.as_mut()?;
        (!env.clients.is_empty()).then(|| env.clients.remove(0))
    });
    fixed.unwrap_or_else(|| {
        let mut hasher = RandomState::new().build_hasher();
        hasher.write_u128(
            SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .map(|elapsed| elapsed.as_nanos())
                .unwrap_or_default(),
        );
        1 + hasher.finish() % (CLIENT_LIMIT - 1)
    })
}

/// `Date.now()`.
pub(crate) fn now_ms() -> f64 {
    ENV.with(|slot| slot.borrow().as_ref().and_then(|env| env.now_ms))
        .unwrap_or_else(|| {
            SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .map(|elapsed| elapsed.as_millis() as f64)
                .unwrap_or_default()
        })
}

/// `performance.now()`: milliseconds on a monotonic clock.
pub(crate) fn perf_ms() -> f64 {
    static START: OnceLock<Instant> = OnceLock::new();
    ENV.with(|slot| slot.borrow().as_ref().and_then(|env| env.perf_ms))
        .unwrap_or_else(|| START.get_or_init(Instant::now).elapsed().as_secs_f64() * 1000.0)
}
