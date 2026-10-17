use std::collections::HashSet;
use std::sync::{Arc, Mutex};

use yrs::sync::time::Clock;
use yrs::{Doc, Origin, ReadTxn, Transact};

use crate::{COMMENTS, EditError, EditResult, META, SHAPES, SLIDE_ORDER, SLIDES, STORIES};

const CAPTURE_TIMEOUT_MS: u64 = 500;

/// Policy for grouping tracked local transactions.
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub enum UndoCaptureMode {
    #[default]
    Auto,
    Manual,
}

struct CaptureClock {
    source: Arc<dyn Clock>,
    state: Mutex<CaptureClockState>,
}

struct CaptureClockState {
    mode: UndoCaptureMode,
    ticks: u64,
    last_source: u64,
}

impl CaptureClock {
    fn new(source: Arc<dyn Clock>) -> Self {
        Self {
            state: Mutex::new(CaptureClockState {
                mode: UndoCaptureMode::Auto,
                ticks: 1,
                last_source: source.now(),
            }),
            source,
        }
    }

    fn mode(&self) -> UndoCaptureMode {
        self.state.lock().unwrap().mode
    }

    fn set_mode(&self, mode: UndoCaptureMode) {
        self.state.lock().unwrap().mode = mode;
    }
}

impl Clock for CaptureClock {
    fn now(&self) -> u64 {
        let mut state = self.state.lock().unwrap();
        let now = self.source.now();
        let elapsed = match state.mode {
            UndoCaptureMode::Auto => now.saturating_sub(state.last_source),
            UndoCaptureMode::Manual => 0,
        };
        state.last_source = now;
        state.ticks = state.ticks.saturating_add(elapsed);
        state.ticks
    }
}

pub struct DeckUndoManager {
    clock: Arc<CaptureClock>,
    inner: yrs::undo::UndoManager<()>,
    /// The story the current typing run edits; any other local edit ends the run.
    typing: Option<String>,
}

impl DeckUndoManager {
    pub(crate) fn new(doc: &Doc, client_id: u64) -> EditResult<Self> {
        let txn = doc.transact();
        let order = txn
            .get_array(SLIDE_ORDER)
            .ok_or_else(|| EditError::InvalidState("missing slide order".to_owned()))?;
        let slides = txn
            .get_map(SLIDES)
            .ok_or_else(|| EditError::InvalidState("missing slides map".to_owned()))?;
        let shapes = txn
            .get_map(SHAPES)
            .ok_or_else(|| EditError::InvalidState("missing shapes map".to_owned()))?;
        let stories = txn
            .get_map(STORIES)
            .ok_or_else(|| EditError::InvalidState("missing stories map".to_owned()))?;
        let comments = txn
            .get_map(COMMENTS)
            .ok_or_else(|| EditError::InvalidState("missing comments map".to_owned()))?;
        let meta = txn
            .get_map(META)
            .ok_or_else(|| EditError::InvalidState("missing meta map".to_owned()))?;
        drop(txn);
        let clock = Arc::new(CaptureClock::new(default_clock()));
        let options = yrs::undo::Options {
            capture_timeout_millis: CAPTURE_TIMEOUT_MS,
            tracked_origins: HashSet::from([Origin::from(client_id)]),
            capture_transaction: None,
            timestamp: clock.clone(),
            init_undo_stack: Vec::new(),
            init_redo_stack: Vec::new(),
        };
        let mut inner = yrs::undo::UndoManager::with_options(options);
        inner.expand_scope(doc, &order);
        inner.expand_scope(doc, &slides);
        inner.expand_scope(doc, &shapes);
        inner.expand_scope(doc, &stories);
        inner.expand_scope(doc, &comments);
        inner.expand_scope(doc, &meta);
        Ok(Self {
            inner,
            clock,
            typing: None,
        })
    }

    pub fn capture_mode(&self) -> UndoCaptureMode {
        self.clock.mode()
    }

    pub fn set_capture_mode(&mut self, mode: UndoCaptureMode) {
        if self.capture_mode() != mode {
            self.add_undo_barrier();
            self.clock.set_mode(mode);
        }
    }

    pub fn undo(&mut self) -> bool {
        let applied = self.inner.undo_blocking();
        self.add_undo_barrier();
        applied
    }

    pub fn redo(&mut self) -> bool {
        let applied = self.inner.redo_blocking();
        self.add_undo_barrier();
        applied
    }

    /// In automatic capture a local edit other than typing is its own step.
    pub(crate) fn begin_edit(&mut self) {
        if self.capture_mode() == UndoCaptureMode::Auto {
            self.add_undo_barrier();
        }
    }

    /// Typing in one story inside the capture window is one step.
    pub(crate) fn begin_typing(&mut self, story_id: &str) {
        if self.capture_mode() == UndoCaptureMode::Auto && self.typing.as_deref() != Some(story_id)
        {
            self.add_undo_barrier();
            self.typing = Some(story_id.to_owned());
        }
    }

    pub fn can_undo(&self) -> bool {
        self.inner.can_undo()
    }

    pub fn can_redo(&self) -> bool {
        self.inner.can_redo()
    }

    pub fn add_undo_barrier(&mut self) {
        self.inner.reset();
        self.typing = None;
    }
}

/// System clock on native targets and `Date.now` in the browser build, so keystrokes inside
/// the capture window form one Undo step. A wasm build without the `wasm` feature has no
/// clock: its counter advances a full window per reading, so every transaction is a step.
fn default_clock() -> Arc<dyn Clock> {
    #[cfg(not(all(target_family = "wasm", target_os = "unknown")))]
    {
        Arc::new(|| {
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .map(|elapsed| elapsed.as_millis() as u64)
                .unwrap_or_default()
        })
    }
    #[cfg(all(target_family = "wasm", target_os = "unknown", feature = "wasm"))]
    {
        Arc::new(|| js_sys::Date::now() as u64)
    }
    #[cfg(all(target_family = "wasm", target_os = "unknown", not(feature = "wasm")))]
    {
        use std::sync::atomic::{AtomicU64, Ordering};
        let ticks = AtomicU64::new(0);
        Arc::new(move || ticks.fetch_add(CAPTURE_TIMEOUT_MS + 1, Ordering::Relaxed))
    }
}
