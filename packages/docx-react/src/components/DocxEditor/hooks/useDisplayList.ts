import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  buildRustDisplayList,
  buildRustDisplayFrame,
  applyFrameDeltaOwned,
  createCanvasImageResolver,
  createDisplayListQueries,
  decodeFrameDelta,
  demoDisplayList,
  encodeDisplayListFrameExtras,
  isDisplayListQuerySourceDead,
  type DisplayList,
  type DisplayListQueries,
  type DisplayPage,
  type GlyphOutlineProvider,
  type ImageResolver,
  type RustDisplayListEngine,
  type RetainedFrame,
  type ResidentDisplayListQueryEngine,
} from '@betteroffice/docx/layout/render';
import { getLayoutKernelInputs, isRecordedResidentLayout } from '@betteroffice/docx/editor';
import {
  canUseResidentEngineWorker,
  residentCaretSnapshotForFrame,
  ResidentCommittedEditError,
  ResidentEngineWorkerClient,
  ResidentWorkerFailureError,
  sameYrsSelection,
  type ResidentCaretPaintStyle,
  type ResidentEngineOffscreenPage,
  type ResidentEngineWorkerFrame,
  type ResidentFirstPagesViewport,
  type SlicedPaint,
  type YrsResidentCaretSnapshot,
  type YrsSelection,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import type { RustFontChainsProvider } from './useRustMeasurement';
import { displayListNeedsHostImages } from '../canvasPresentation';
import { CARET_PAINT_IDLE_MS, PaintedCaretMachine } from '../paintedCaret';
import {
  DisplayListQueryEpochGate,
  type ResolveDisplayListQueries,
} from './displayListQueryEpochGate';

// provider for the canvas renderer's display list: returns the injected value
// when the host supplies one, otherwise the demo fixture. consumers only ever
// see a DisplayList, never where it came from.
export function useDisplayList(injected?: DisplayList | null): DisplayList {
  return injected ?? demoDisplayList;
}

export interface UseRustDisplayListResult {
  /** the latest successfully built display list (kept across rebuilds so the canvas never blanks mid-compute) */
  displayList: DisplayList | null;
  /** fatal error from resolving inputs or building the display list */
  error: Error | null;
  /** true until the first display list for the current document is ready */
  loading: boolean;
  /** Binary retained-frame state; null on the compatibility JSON path. */
  frame: RetainedFrame | null;
  /** Query facade built from the same display list as `frame`. */
  queries: DisplayListQueries | null;
  /** Resolve the newest query facade after pending document/frame changes. */
  resolveQueries: ResolveDisplayListQueries;
  /** Worker-computed caret tagged to `frame`. */
  caret: YrsResidentCaretSnapshot | null;
  /** Apply a plain-text edit through the resident engine and publish its frame. */
  applyInput(text: string): Promise<ResidentFrameApplyResult | null>;
  /** Apply one collapsed deletion/paragraph merge through the resident engine. */
  applyDelete(direction: 'backward' | 'forward'): Promise<ResidentFrameApplyResult | null>;
  /**
   * True while the worker owns the visible page surfaces. Sticky across
   * invalidation (remote/structural updates) so the canvas keeps its last
   * pixels instead of remounting; drops only on genuine fallback or reset.
   */
  workerSurfacesActive: boolean;
  workerPresentationActive: boolean;
  setWorkerPresentationActive(active: boolean): void;
  attachOffscreenCanvases(
    pages: ResidentEngineOffscreenPage[],
    activePageIds: string[],
    devicePixelRatio: number,
    zoom: number,
    caretStyle: ResidentCaretPaintStyle
  ): Promise<boolean>;
  /** True while the worker-painted caret line owns the caret (DOM caret hidden). */
  paintedCaretActive: boolean;
  /** Starts the session's worker as it opens, seeded from `update` (see YrsSession.residentWorkerSeed). */
  startWorker(engine: YrsSession, update: Uint8Array | null): void;
  /** Local text input: keeps painted-caret mode alive for follow-up frames. */
  notifyCaretInput(): void;
  /** Text input dispatched: hide the DOM caret before the worker round-trip. */
  notifyCaretInputDispatched(): void;
  /** Selection move / blur / IME start / mode change: immediate swap to the DOM caret. */
  notifyCaretInterrupt(): void;
  /** Resolves once these pages are built (Q815 C: a first frame leaves the pages out of view unbuilt). */
  ensurePages(pageIndices: number[]): Promise<void>;
  /** The pages in view, which build first. */
  setVisiblePages(first: number, last: number): void;
  /** Resolves once the pages in view are built (every page before the view is known). */
  ensureVisiblePages(): Promise<void>;
  /** Unbuilt pages may hold pictures only the page can draw. */
  unbuiltMayHoldImages: boolean;
}

export interface ResidentFrameApplyResult {
  frameEpoch: number | null;
  caretSynchronized: boolean;
}

/** test seam: unit tests inject a fake engine/inputs-resolver instead of the wasm module */
export interface RustDisplayListHookOverrides {
  build?: typeof buildRustDisplayList;
  getInputs?: typeof getLayoutKernelInputs;
}

type ResidentInputOperation =
  | { kind: 'insert'; text: string }
  | { kind: 'delete'; direction: 'backward' | 'forward' };

interface RustDisplayListSnapshot {
  displayList: DisplayList | null;
  frame: RetainedFrame | null;
  queries: DisplayListQueries | null;
  caret: YrsResidentCaretSnapshot | null;
}

interface MainThreadBuild {
  displayList: DisplayList;
  frame: RetainedFrame | null;
  queryEngine: RustDisplayListEngine | null | undefined;
}

/** The newest frame the worker should produce. A queued run reads it when it
 * starts, so layouts that land while the worker is busy cost one request. */
interface WorkerFrameRequest {
  hostEngine: YrsSession;
  extras: string;
  /** The page recorded this pass without laying it out: the worker's reply
   * carries its layout for the page's readers. */
  wantLayout: boolean;
  /** Leave the worker for the main-thread engine and build there instead. */
  fallback(cause: unknown): void;
}

/** Unbuilt pages one request builds (as DOCX View's batches). */
const PAGE_BUILD_BATCH = 8;

/** The next unbuilt pages to build: those in view, those asked for, then
 * (with `fill`) outward from the view (from the top without one). */
export function pageBuildBatch(
  pages: readonly DisplayPage[],
  visible: { first: number; last: number } | null,
  demanded: Iterable<number>,
  fill = true
): number[] {
  const batch: number[] = [];
  const add = (index: number): void => {
    if (batch.length < PAGE_BUILD_BATCH && pages[index]?.unbuilt && !batch.includes(index)) {
      batch.push(index);
    }
  };
  const first = visible?.first ?? 0;
  const last = visible?.last ?? -1;
  for (let index = first; index <= last; index += 1) add(index);
  for (const index of demanded) add(index);
  for (
    let step = 1;
    fill && batch.length < PAGE_BUILD_BATCH && (last + step < pages.length || first - step >= 0);
    step += 1
  ) {
    add(last + step);
    add(first - step);
  }
  return batch;
}

const EMPTY_DISPLAY_LIST_SNAPSHOT: RustDisplayListSnapshot = {
  displayList: null,
  frame: null,
  queries: null,
  caret: null,
};

// rebuilds the display list through the rust wasm engine after every layout
// pass. dumb replay glue: the `{ measured, options, layout }` triple (plus the
// kernel-recorded `headersFooters` payload when the document has HF parts) is
// serialized as-is (the same envelope the golden fixtures pin) and every
// geometry/paint decision happens in rust. Every adapter interaction is
// display-list-backed. A generation counter drops stale async results so only
// the newest layout wins.
export function useRustDisplayList(
  layout: Layout | null,
  overrides?: RustDisplayListHookOverrides,
  // Host slot the Rust measure source fills with the merged doc-wide font
  // chains. Read at build time and passed through to the builder — its
  // presence (a non-empty map) is what activates GlyphRun emission. Null /
  // undefined ⇒ the builder emits browser-shaped TextRunPrimitives (unchanged).
  fontChainsProviderRef?: React.RefObject<RustFontChainsProvider | null>,
  // Resolved comment threads whose range wash the canvas must hide (the crate
  // drops their comment-range decorations and stamps status="resolved").
  // Changing the set rebuilds the display list so resolve/reopen — and the
  // "expanded resolved card re-tints its range" flow — repaint immediately.
  resolvedCommentIds?: ReadonlySet<number>,
  engine?: RustDisplayListEngine | null,
  // The viewport the worker's first frame fills; the other pages build later.
  firstPagesViewport?: React.RefObject<(() => ResidentFirstPagesViewport | null) | null>
): UseRustDisplayListResult {
  const [snapshot, setSnapshot] = useState<RustDisplayListSnapshot>(EMPTY_DISPLAY_LIST_SNAPSHOT);
  const snapshotRef = useRef<RustDisplayListSnapshot>(EMPTY_DISPLAY_LIST_SNAPSHOT);
  const layoutRef = useRef(layout);
  layoutRef.current = layout;
  const queryEpochGateRef = useRef<DisplayListQueryEpochGate | null>(null);
  if (!queryEpochGateRef.current) queryEpochGateRef.current = new DisplayListQueryEpochGate();
  const queryEpochGate = queryEpochGateRef.current;
  const contentEpochRef = useRef(0);
  const [error, setError] = useState<Error | null>(null);
  const [loading, setLoading] = useState(true);
  const generationRef = useRef(0);
  const workerRef = useRef<{
    engine: YrsSession;
    client: ResidentEngineWorkerClient;
  } | null>(null);
  const workerFallbackEngineRef = useRef<YrsSession | null>(null);
  const workerInputQueueRef = useRef<Promise<void>>(Promise.resolve());
  const workerFrameRequestRef = useRef<WorkerFrameRequest | null>(null);
  const workerFrameQueuedRef = useRef(false);
  const suppressWorkerInvalidationRef = useRef(0);
  const [workerSurfacesActive, setWorkerSurfacesActive] = useState(false);
  const workerPresentationActiveRef = useRef(false);
  const [workerPresentationActive, setWorkerPresentationActiveState] = useState(false);

  // Pages the worker's frames leave unbuilt (Q815 C) build on the worker's
  // queue: those in view or asked for at once, the rest a batch at a time.
  // Keys, frames and those builds go ahead of a queued batch (Q838); a batch
  // the worker is building finishes first.
  const visiblePagesRef = useRef<{ first: number; last: number } | null>(null);
  const demandedPagesRef = useRef(new Set<number>());
  const pageBuildQueuedRef = useRef(false);
  const demandBuildQueuedRef = useRef(false);
  const workerWaitingRef = useRef(0);
  const runPageBuildRef = useRef<(onDemand: boolean) => Promise<void>>(async () => {});
  const pageWaitersRef = useRef(new Set<{ pages: number[]; resolve: () => void }>());
  const workerMediaRef = useRef(false);
  /** Queues a worker request; a page batch queued before it waits for it. */
  const enqueueWorkerRequest = useCallback(<T>(run: () => Promise<T>): Promise<T> => {
    workerWaitingRef.current += 1;
    const start = (): Promise<T> => {
      workerWaitingRef.current -= 1;
      return run();
    };
    const pending = workerInputQueueRef.current.then(start, start);
    workerInputQueueRef.current = pending.then(
      () => undefined,
      () => undefined
    );
    return pending;
  }, []);
  const schedulePageBuilds = useCallback((): void => {
    if (pageBuildQueuedRef.current || !workerRef.current) return;
    pageBuildQueuedRef.current = true;
    const run = (): Promise<void> => {
      pageBuildQueuedRef.current = false;
      // Requests queued behind this batch go first: it queues again after them.
      if (workerWaitingRef.current > 0) {
        schedulePageBuilds();
        return Promise.resolve();
      }
      return runPageBuildRef.current(false);
    };
    workerInputQueueRef.current = workerInputQueueRef.current
      .then(run, run)
      .catch((error) => console.error('[CanvasRenderer] Page build failed', error));
  }, []);
  const scheduleDemandBuild = useCallback((): void => {
    if (demandBuildQueuedRef.current || !workerRef.current) return;
    demandBuildQueuedRef.current = true;
    enqueueWorkerRequest(() => {
      demandBuildQueuedRef.current = false;
      return runPageBuildRef.current(true);
    }).catch((error) => console.error('[CanvasRenderer] Page build failed', error));
  }, [enqueueWorkerRequest]);
  const demandPages = useCallback(
    (pageIndices: number[]): void => {
      for (const index of pageIndices) demandedPagesRef.current.add(index);
      scheduleDemandBuild();
    },
    [scheduleDemandBuild]
  );

  const paintedCaretMachineRef = useRef<PaintedCaretMachine | null>(null);
  if (!paintedCaretMachineRef.current) paintedCaretMachineRef.current = new PaintedCaretMachine();
  const paintedCaretMachine = paintedCaretMachineRef.current;
  const [paintedCaretActive, setPaintedCaretActiveState] = useState(false);
  const paintedCaretIdleTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const dispatchHoldTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const residentPaintInflightRef = useRef(0);

  // Idle/interrupt swap: the DOM caret renders first (state flip), then the
  // erase posts two rAFs later — an overlap is possible, a caret gap is not.
  // The post re-checks the machine so a paint that landed meanwhile survives.
  const requestPaintedCaretErase = useCallback((): void => {
    setPaintedCaretActiveState(false);
    const post = (): void => {
      if (paintedCaretMachine.isActive() || residentPaintInflightRef.current > 0) return;
      workerRef.current?.client.eraseCaret();
    };
    if (typeof requestAnimationFrame === 'function') {
      requestAnimationFrame(() => requestAnimationFrame(post));
    } else {
      setTimeout(post, 0);
    }
  }, [paintedCaretMachine]);

  const schedulePaintedCaretIdle = useCallback((): void => {
    if (paintedCaretIdleTimerRef.current !== null) clearTimeout(paintedCaretIdleTimerRef.current);
    const fire = (): void => {
      paintedCaretIdleTimerRef.current = null;
      const now = performance.now();
      if (paintedCaretMachine.idleTimeout(now)) {
        requestPaintedCaretErase();
        return;
      }
      if (paintedCaretMachine.isActive()) {
        paintedCaretIdleTimerRef.current = setTimeout(
          fire,
          Math.max(1, paintedCaretMachine.msUntilIdle(now))
        );
      }
    };
    paintedCaretIdleTimerRef.current = setTimeout(fire, CARET_PAINT_IDLE_MS);
  }, [paintedCaretMachine, requestPaintedCaretErase]);

  const applyPaintedCaretReply = useCallback(
    (painted: boolean, token: number): void => {
      if (painted && paintedCaretMachine.framePainted(token)) {
        setPaintedCaretActiveState(true);
        schedulePaintedCaretIdle();
        return;
      }
      // Painted but interrupted mid-flight: the DOM caret is already showing,
      // so the stale line can be erased immediately.
      if (painted) workerRef.current?.client.eraseCaret();
      // An in-flight resident input owns the next verdict; deciding here would
      // remount the DOM caret for a frame in the middle of a burst.
      if (residentPaintInflightRef.current > 0) return;
      paintedCaretMachine.frameUnpainted();
      setPaintedCaretActiveState(false);
    },
    [paintedCaretMachine, schedulePaintedCaretIdle]
  );

  const notifyCaretInput = useCallback((): void => {
    paintedCaretMachine.noteInput(performance.now());
  }, [paintedCaretMachine]);

  // Dispatch-time hide: called synchronously from the input event, BEFORE the
  // worker round-trip. The worker presents glyphs+painted caret atomically off
  // the main thread, so a caret left mounted until the reply commits shows the
  // new character with a stale caret one position behind. Hiding for the
  // in-flight window is invisible; the dispatch hold self-expires and every
  // resolution path (painted, unpainted, interrupt) reconciles it.
  const notifyCaretInputDispatched = useCallback((): void => {
    if (!workerPresentationActiveRef.current) return;
    paintedCaretMachine.noteDispatch(performance.now());
    setPaintedCaretActiveState(true);
    if (dispatchHoldTimerRef.current !== null) clearTimeout(dispatchHoldTimerRef.current);
    dispatchHoldTimerRef.current = setTimeout(() => {
      dispatchHoldTimerRef.current = null;
      const now = performance.now();
      if (!paintedCaretMachine.isActive() && !paintedCaretMachine.isHolding(now)) {
        setPaintedCaretActiveState(false);
      }
    }, CARET_PAINT_IDLE_MS + 16);
  }, [paintedCaretMachine]);

  const notifyCaretInterrupt = useCallback((): void => {
    if (paintedCaretMachine.interrupt()) {
      requestPaintedCaretErase();
      return;
    }
    // A pending dispatch hold must not survive a selection change.
    setPaintedCaretActiveState(false);
  }, [paintedCaretMachine, requestPaintedCaretErase]);

  useEffect(() => {
    if (typeof window === 'undefined') return;
    window.addEventListener('blur', notifyCaretInterrupt);
    return () => window.removeEventListener('blur', notifyCaretInterrupt);
  }, [notifyCaretInterrupt]);

  const setWorkerPresentationActive = useCallback(
    (active: boolean): void => {
      workerPresentationActiveRef.current = active;
      if (!active) notifyCaretInterrupt();
      setWorkerPresentationActiveState((current) => (current === active ? current : active));
    },
    [notifyCaretInterrupt]
  );

  const residentEngine = isWorkerHostEngine(engine) ? engine : null;

  // The host asks again for the font slices that failed: the engine worker does too.
  useEffect(
    () => residentEngine?.onFontSlicesRetry(() => workerRef.current?.client.retryFontSlices()),
    [residentEngine]
  );

  useEffect(() => {
    if (!residentEngine) return;
    return residentEngine.onUpdate((update) => {
      if (suppressWorkerInvalidationRef.current > 0) return;
      contentEpochRef.current += 1;
      queryEpochGate.invalidate();
      // Update observers fire from inside the wasm transaction. Calling any
      // other EditSession method here would re-enter the borrowed wasm object
      // (wasm-bindgen correctly rejects that unsafe alias). Selection is sent
      // with the next input request after the transaction has returned.
      if (canUseResidentEngineWorker()) workerRef.current?.client.invalidate(update, null);
      // Only frame freshness drops here. The worker keeps the page surfaces
      // (workerSurfacesActive) so the canvas retains its pixels until the
      // post-sync frame lands — flipping surfaces would remount every page.
    });
  }, [queryEpochGate, residentEngine]);

  useEffect(
    () => () => {
      if (paintedCaretIdleTimerRef.current !== null) {
        clearTimeout(paintedCaretIdleTimerRef.current);
      }
      if (dispatchHoldTimerRef.current !== null) {
        clearTimeout(dispatchHoldTimerRef.current);
      }
      workerRef.current?.client.destroy();
      workerRef.current = null;
      queryEpochGate.clear();
      for (const waiter of pageWaitersRef.current) waiter.resolve();
      pageWaitersRef.current.clear();
    },
    [queryEpochGate]
  );

  const publishQuerySnapshot = useCallback(
    (nextSnapshot: RustDisplayListSnapshot, contentEpoch: number): void => {
      const queries = nextSnapshot.queries;
      if (!queries) {
        if (contentEpoch === contentEpochRef.current) queryEpochGate.clear();
        return;
      }
      if (contentEpoch !== contentEpochRef.current) return;
      queryEpochGate.invalidate();
      const publish = (): void => {
        if (
          snapshotRef.current !== nextSnapshot ||
          contentEpoch !== contentEpochRef.current ||
          !queries.isReady()
        ) {
          return;
        }
        queryEpochGate.publish({
          queries,
          frameEpoch: nextSnapshot.frame?.frameEpoch ?? null,
        });
      };
      if (queries.isReady()) {
        publish();
        return;
      }
      void queries.whenReady().then(publish, () => {
        if (
          snapshotRef.current === nextSnapshot &&
          contentEpoch === contentEpochRef.current
        ) {
          queryEpochGate.clear();
        }
      });
    },
    [queryEpochGate]
  );

  const resolveQueries = useCallback<ResolveDisplayListQueries>(
    (minimumFrameEpoch = null) => queryEpochGate.resolve(minimumFrameEpoch),
    [queryEpochGate]
  );

  // Handle acquisition (page-delta serialization into the Rust query store) is
  // deferred out of the keystroke path: prime the newest facade from idle time
  // so the first interaction after a typing burst pays one accumulated delta
  // at most. Superseded facades no-op their pending prime.
  useEffect(() => {
    const queries = snapshot.queries;
    if (!queries) return;
    if (typeof requestIdleCallback === 'function') {
      const id = requestIdleCallback(() => queries.prime());
      return () => cancelIdleCallback(id);
    }
    const id = setTimeout(() => queries.prime(), 200);
    return () => clearTimeout(id);
  }, [snapshot.queries]);

  const applyResidentInput = useCallback(
    (operation: ResidentInputOperation): Promise<ResidentFrameApplyResult | null> => {
      const replayInputOnMainThread = async (
        pending: ResidentInputOperation,
        hostEngine: YrsSession,
        frameEpoch: number,
        paintToken: number
      ): Promise<ResidentFrameApplyResult | null> => {
        let expectedFrameEpoch = frameEpoch;
        if (workerFallbackEngineRef.current !== hostEngine) {
          queryEpochGate.clear();
          workerFallbackEngineRef.current = hostEngine;
          // The retained frame carries the worker's epochs and the host engine
          // numbers its own: drop it and ask for a full frame, so the host's
          // frames are not refused as stale (the effect's fallback does the same).
          snapshotRef.current = { ...snapshotRef.current, frame: null, queries: null, caret: null };
          expectedFrameEpoch = 0;
        }
        if (workerRef.current?.engine === hostEngine) {
          workerRef.current.client.destroy();
          workerRef.current = null;
        }
        setWorkerSurfacesActive(false);
        setWorkerPresentationActive(false);
        // The page lays out from now on. A pass the worker laid out has no
        // current layout in this engine: the edit takes the structural path,
        // whose pass lays out here.
        hostEngine.publishResidentLayout?.({ kind: 'page' });
        if (isRecordedResidentLayout(layoutRef.current)) return null;
        let encoded: Uint8Array;
        suppressWorkerInvalidationRef.current += 1;
        try {
          encoded =
            pending.kind === 'insert'
              ? hostEngine.applyInput(pending.text, expectedFrameEpoch)
              : hostEngine.applyDelete(pending.direction, expectedFrameEpoch);
        } catch (error) {
          suppressWorkerInvalidationRef.current -= 1;
          if (
            error instanceof Error &&
            error.message.includes('resident input state is not ready')
          ) {
            return null;
          }
          throw error;
        }
        suppressWorkerInvalidationRef.current -= 1;
        const delta = decodeFrameDelta(encoded);
        const previous = snapshotRef.current;
        const nextFrame = delta.full
          ? applyFrameDeltaOwned(null, delta)
          : applyFrameDeltaOwned(previous.frame, delta);
        const hostSelection = hostEngine.selection();
        const caret = residentCaretForSelection(
          hostEngine.residentCaretSnapshot(),
          hostSelection,
          hostSelection,
          nextFrame
        );
        const nextSnapshot = createRustDisplayListSnapshot(
          nextFrame.displayList,
          nextFrame,
          caret,
          null,
          { ...previous, queries: null },
          demandPages
        );
        generationRef.current += 1;
        snapshotRef.current = nextSnapshot;
        publishQuerySnapshot(nextSnapshot, contentEpochRef.current);
        setSnapshot(nextSnapshot);
        setError(null);
        setLoading(false);
        applyPaintedCaretReply(false, paintToken);
        return { frameEpoch: nextFrame.frameEpoch, caretSynchronized: false };
      };
      const run = async (): Promise<ResidentFrameApplyResult | null> => {
        const worker = workerRef.current;
        const currentFrame = snapshotRef.current.frame;
        if (!worker || !worker.client.isReady() || !currentFrame) return null;
        const selection = worker.engine.selection();
        if (!selection) return null;
        paintedCaretMachine.noteInput(performance.now());
        const paintCaret = workerPresentationActiveRef.current;
        const paintToken = paintedCaretMachine.token();
        residentPaintInflightRef.current += 1;
        let result;
        let workerDelta: ReturnType<typeof decodeFrameDelta> | undefined;
        try {
          result =
            operation.kind === 'insert'
              ? await worker.client.applyInput(
                  operation.text,
                  selection,
                  currentFrame.frameEpoch,
                  false,
                  paintCaret
                )
              : await worker.client.applyDelete(
                  operation.direction,
                  selection,
                  currentFrame.frameEpoch,
                  false,
                  paintCaret
                );
          if (result.applied) {
            try {
              workerDelta = decodeFrameDelta(result.frame);
            } catch (error) {
              throw new ResidentWorkerFailureError(
                `Resident engine worker returned an undecodable FrameDelta: ${error instanceof Error ? error.message : String(error)}`
              );
            }
          }
        } catch (error) {
          if (error instanceof ResidentCommittedEditError) {
            // The worker committed the edit, then failed: the edit comes back
            // into the main replica (never replayed), and the main thread paints.
            const frameRequest = workerFrameRequestRef.current;
            if (frameRequest) suppressWorkerInvalidationRef.current += 1;
            try {
              for (const update of error.updates) worker.engine.applyLocalUpdate(update);
            } finally {
              if (frameRequest) suppressWorkerInvalidationRef.current -= 1;
            }
            frameRequest?.fallback(error);
            return { frameEpoch: null, caretSynchronized: false };
          }
          if (!(error instanceof ResidentWorkerFailureError)) throw error;
          console.error(
            '[CanvasRenderer] Resident engine worker unavailable; falling back to the main-thread engine',
            error
          );
          return replayInputOnMainThread(
            operation,
            worker.engine,
            currentFrame.frameEpoch,
            paintToken
          );
        } finally {
          residentPaintInflightRef.current -= 1;
        }
        if (!result.applied) return null;
        const delta = workerDelta ?? decodeFrameDelta(result.frame);
        suppressWorkerInvalidationRef.current += 1;
        try {
          for (const update of result.updates) worker.engine.applyLocalUpdate(update);
        } finally {
          suppressWorkerInvalidationRef.current -= 1;
        }
        if (result.fontErrors.length > 0) worker.engine.reportResidentFontErrors(result.fontErrors);
        const frameRequest = workerFrameRequestRef.current;
        if (result.paintError !== undefined && frameRequest) {
          // The edit is in the main replica now; only the worker's painting
          // failed, so the main thread takes over painting (no replay).
          frameRequest.fallback(new Error(`Resident engine worker could not paint: ${result.paintError}`));
          return { frameEpoch: null, caretSynchronized: false };
        }
        const previous = snapshotRef.current;
        if (previous.frame && delta.frameEpoch <= previous.frame.frameEpoch) {
          // Superseded: a newer frame's reply owns the painted-caret verdict.
          return { frameEpoch: null, caretSynchronized: false };
        }
        const nextFrame = applyFrameDeltaOwned(previous.frame, delta);
        const caret = residentCaretForSelection(
          result.caret,
          result.selection,
          worker.engine.selection(),
          nextFrame
        );
        const nextSnapshot = createRustDisplayListSnapshot(
          nextFrame.displayList,
          nextFrame,
          caret,
          null,
          previous,
          demandPages
        );
        // Supersede an older async compatibility build before publishing the
        // frame produced by the edit transaction.
        generationRef.current += 1;
        snapshotRef.current = nextSnapshot;
        publishQuerySnapshot(nextSnapshot, contentEpochRef.current);
        setSnapshot(nextSnapshot);
        setError(null);
        setLoading(false);
        applyPaintedCaretReply(Boolean(result.caretPainted && caret?.caretRect), paintToken);
        return {
          frameEpoch: nextFrame.frameEpoch,
          caretSynchronized: Boolean(
            caret?.caretRect &&
              workerPresentationActiveRef.current &&
              !displayListNeedsHostImages(nextFrame.displayList)
          ),
        };
      };
      return enqueueWorkerRequest(run).catch((error) => {
        const nextError =
          error instanceof Error ? error : new Error(`Resident input failed: ${String(error)}`);
        // This Rust rejection happens before the edit transaction. Let the
        // caller use the structural compatibility path for unsupported
        // paragraphs (inline media/float-dependent measurement, non-body
        // stories, or a frame that has not established resident state).
        if (nextError.message.includes('resident input state is not ready')) return null;
        console.error('[CanvasRenderer] Resident input failed', nextError);
        queryEpochGate.clear();
        setError(nextError);
        // Once invoked, never fall through to the legacy op: a worker failure
        // may have happened after committing the transaction.
        return { frameEpoch: null, caretSynchronized: false };
      });
    },
    [
      applyPaintedCaretReply,
      demandPages,
      enqueueWorkerRequest,
      paintedCaretMachine,
      publishQuerySnapshot,
      queryEpochGate,
    ]
  );

  const applyInput = useCallback(
    (text: string) => applyResidentInput({ kind: 'insert', text }),
    [applyResidentInput]
  );

  const applyDelete = useCallback(
    (direction: 'backward' | 'forward') =>
      applyResidentInput({ kind: 'delete', direction }),
    [applyResidentInput]
  );

  // Every frame-producing worker request runs on `workerInputQueueRef`, one at
  // a time: each names the epoch of the frame it diffs against after the
  // previous reply was applied, and every reply is applied, even one whose
  // content a newer update already passed. A dropped delta would leave the
  // next one without its base and cost a full frame.
  const runWorkerFrame = useCallback(async (): Promise<void> => {
    workerFrameQueuedRef.current = false;
    const request = workerFrameRequestRef.current;
    const worker = workerRef.current;
    if (!request || !worker || worker.engine !== request.hostEngine) return;
    const { hostEngine, extras } = request;
    const client = worker.client;
    const probe = hostEngine.residentWorkerProbe();
    if (!probe) return;
    // Every update so far reached the worker before this request does.
    const contentEpoch = contentEpochRef.current;
    const paintToken = paintedCaretMachine.token();
    const bootstrapping = client.layoutRevision() === 0;
    const previous = bootstrapping ? null : snapshotRef.current.frame;
    const expectedFrameEpoch = previous?.frameEpoch ?? 0;
    // Structural text input reaches the worker as a relayout; keep the painted
    // caret glued to those frames while the typing burst lasts.
    const paintCaret =
      !bootstrapping &&
      workerPresentationActiveRef.current &&
      paintedCaretMachine.shouldPaint(performance.now());
    // On a fresh client both hints are null, so a bootstrap snapshot is
    // complete; a sync snapshot ships a state diff and skips font bytes the
    // worker already holds.
    const snapshot = () => {
      const next = hostEngine.residentWorkerSnapshot({
        knownStateVector: client.remoteStateVector(),
        knownFontsRevision: client.syncedFontsRevision(),
      });
      if (!next) throw new Error('Resident worker snapshot was not available');
      return next;
    };
    let result: ResidentEngineWorkerFrame;
    let nextFrame: RetainedFrame;
    try {
      const { wantLayout } = request;
      result = bootstrapping
        ? await client.bootstrap(
            snapshot(),
            extras,
            wantLayout,
            firstPagesViewport?.current?.() ?? undefined
          )
        : client.layoutRevision() === probe.layoutRevision
          ? await client.buildFrame(extras, expectedFrameEpoch, paintCaret)
          : probe.regionLayoutInput !== null &&
              client.syncedFontsRevision() === probe.fontsRevision
            ? await client.relayout(
                probe.regionLayoutInput,
                extras,
                expectedFrameEpoch,
                hostEngine.selection(),
                probe.layoutRevision,
                paintCaret,
                wantLayout
              )
            : await client.sync(snapshot(), extras, expectedFrameEpoch, paintCaret, wantLayout);
      if (result.fontErrors.length > 0) hostEngine.reportResidentFontErrors(result.fontErrors);
      if (result.unbuiltPages > 0) workerMediaRef.current = result.media;
      // Torn down, or the frame this delta builds on was reset meanwhile (the
      // next request then asks for a full frame).
      if (workerRef.current?.client !== client) return;
      if (!bootstrapping && snapshotRef.current.frame !== previous) return;
      // Serialized: this delta applies once, so position-only pages shift in
      // place and the query store and mirrors keep them (as for local input).
      nextFrame = applyFrameDeltaOwned(previous, decodeFrameDelta(result.frame));
    } catch (cause) {
      if (workerRef.current?.client !== client) return;
      (workerFrameRequestRef.current ?? request).fallback(cause);
      return;
    }
    // The worker's layout for the page's readers, parsed after its frame
    // paints (the pages come from the frame). A pass recorded or laid out on
    // the page since then supersedes it.
    if (result.layoutJson !== undefined) {
      const json = result.layoutJson;
      const { layoutRevision } = result;
      setTimeout(() => {
        if (
          workerRef.current?.client !== client ||
          hostEngine.residentWorkerProbe()?.layoutRevision !== layoutRevision
        ) {
          return;
        }
        hostEngine.publishResidentLayout?.({ kind: 'layout', json });
      }, 0);
    }
    if (workerFrameQueuedRef.current) {
      // A newer frame is already queued: the next delta builds on this one
      // (the worker painted it), and only the newest is rendered.
      snapshotRef.current = { ...snapshotRef.current, frame: nextFrame };
      return;
    }
    const nextSnapshot = createRustDisplayListSnapshot(
      nextFrame.displayList,
      nextFrame,
      residentCaretForSelection(result.caret, result.selection, hostEngine.selection(), nextFrame),
      null,
      snapshotRef.current,
      demandPages
    );
    // Supersede an older async main-thread build.
    generationRef.current += 1;
    snapshotRef.current = nextSnapshot;
    publishQuerySnapshot(nextSnapshot, contentEpoch);
    setSnapshot(nextSnapshot);
    setError(null);
    setLoading(false);
    setWorkerSurfacesActive(true);
    applyPaintedCaretReply(
      Boolean(result.caretPainted && nextSnapshot.caret?.caretRect),
      paintToken
    );
  }, [
    applyPaintedCaretReply,
    demandPages,
    firstPagesViewport,
    paintedCaretMachine,
    publishQuerySnapshot,
  ]);

  // Unbuilt pages, applied as any worker frame: on demand those in view and
  // those asked for, else a batch.
  runPageBuildRef.current = async (onDemand: boolean): Promise<void> => {
    const worker = workerRef.current;
    const previous = snapshotRef.current.frame;
    if (!worker || !previous) return;
    for (const index of [...demandedPagesRef.current]) {
      if (!previous.displayList.pages[index]?.unbuilt) demandedPagesRef.current.delete(index);
    }
    const pages = pageBuildBatch(
      previous.displayList.pages,
      visiblePagesRef.current,
      demandedPagesRef.current,
      !onDemand
    );
    if (pages.length === 0) return;
    const contentEpoch = contentEpochRef.current;
    const paintToken = paintedCaretMachine.token();
    const paintCaret =
      workerPresentationActiveRef.current && paintedCaretMachine.shouldPaint(performance.now());
    let result: ResidentEngineWorkerFrame;
    let nextFrame: RetainedFrame;
    try {
      result = await worker.client.buildPages(pages, previous.frameEpoch, paintCaret);
      if (workerRef.current?.client !== worker.client || snapshotRef.current.frame !== previous) {
        return;
      }
      nextFrame = applyFrameDeltaOwned(previous, decodeFrameDelta(result.frame));
    } catch (cause) {
      if (workerRef.current?.client === worker.client) workerFrameRequestRef.current?.fallback(cause);
      return;
    }
    if (result.fontErrors.length > 0) worker.engine.reportResidentFontErrors(result.fontErrors);
    for (const index of pages) demandedPagesRef.current.delete(index);
    const nextSnapshot = createRustDisplayListSnapshot(
      nextFrame.displayList,
      nextFrame,
      residentCaretForSelection(result.caret, result.selection, worker.engine.selection(), nextFrame),
      null,
      snapshotRef.current,
      demandPages
    );
    generationRef.current += 1;
    snapshotRef.current = nextSnapshot;
    publishQuerySnapshot(nextSnapshot, contentEpoch);
    setSnapshot(nextSnapshot);
    applyPaintedCaretReply(Boolean(result.caretPainted && nextSnapshot.caret?.caretRect), paintToken);
  };

  // While a frame holds unbuilt pages the next batch follows; a caller
  // waiting for pages resumes once they are built.
  useEffect(() => {
    const pages = snapshot.frame?.displayList.pages;
    const unbuilt = pages?.some((page) => page.unbuilt) ?? false;
    for (const waiter of [...pageWaitersRef.current]) {
      if (!unbuilt || waiter.pages.every((index) => !pages?.[index]?.unbuilt)) {
        pageWaitersRef.current.delete(waiter);
        waiter.resolve();
      }
    }
    if (unbuilt) schedulePageBuilds();
  }, [schedulePageBuilds, snapshot.frame]);

  const ensurePages = useCallback(
    (pageIndices: number[]): Promise<void> => {
      const pages = snapshotRef.current.frame?.displayList.pages;
      const wanted = pageIndices.filter((index) => pages?.[index]?.unbuilt);
      if (wanted.length === 0) return Promise.resolve();
      return new Promise<void>((resolve) => {
        pageWaitersRef.current.add({ pages: wanted, resolve });
        demandPages(wanted);
      });
    },
    [demandPages]
  );

  const setVisiblePages = useCallback(
    (first: number, last: number): void => {
      const current = visiblePagesRef.current;
      if (current?.first === first && current.last === last) return;
      visiblePagesRef.current = { first, last };
      const pages = snapshotRef.current.frame?.displayList.pages;
      if (pages?.slice(first, last + 1).some((page) => page.unbuilt)) scheduleDemandBuild();
    },
    [scheduleDemandBuild]
  );

  const ensureVisiblePages = useCallback((): Promise<void> => {
    const count = snapshotRef.current.frame?.displayList.pages.length ?? 0;
    const { first, last } = visiblePagesRef.current ?? { first: 0, last: count - 1 };
    return ensurePages(Array.from({ length: last - first + 1 }, (_, index) => first + index));
  }, [ensurePages]);

  const unbuiltMayHoldImages = useMemo(
    () =>
      workerMediaRef.current &&
      Boolean(snapshot.frame?.displayList.pages.some((page) => page.unbuilt)),
    [snapshot.frame]
  );

  // The worker starts as the session opens (Q815): it loads the state the
  // page's replica loaded while the page renders, so the first frame request
  // ships only what changed since. A seed that fails leaves that request to
  // send the whole state.
  const startWorker = useCallback((engine: YrsSession, update: Uint8Array | null): void => {
    if (!canUseResidentEngineWorker() || workerRef.current?.engine === engine) return;
    const failed = (error: unknown) =>
      console.error('[CanvasRenderer] Resident engine worker seed failed', error);
    try {
      const seed = engine.residentWorkerSeed(update);
      workerRef.current?.client.destroy();
      const client = new ResidentEngineWorkerClient();
      workerRef.current = { engine, client };
      // A client destroyed before its seed replied (unmount or swap) is no failure.
      const seeded = client.seed(seed).catch((error: unknown) => {
        if (workerRef.current?.client === client) failed(error);
      });
      // The first frame request's snapshot waits for the seed's state vector.
      workerInputQueueRef.current = workerInputQueueRef.current.then(() => seeded);
    } catch (error) {
      failed(error);
    }
  }, []);

  const attachOffscreenCanvases = useCallback(
    async (
      pages: ResidentEngineOffscreenPage[],
      activePageIds: string[],
      devicePixelRatio: number,
      zoom: number,
      caretStyle: ResidentCaretPaintStyle
    ): Promise<boolean> => {
      const worker = workerRef.current?.client;
      // Queue even while the worker is mid-invalidation: requests are handled
      // FIFO, so an attach lands after the sync that follows and the worker
      // rasters the newly attached surfaces itself. Refusing here would strand
      // already-transferred canvases (they cannot be re-transferred).
      if (!worker) return false;
      await worker.attachCanvases(pages, activePageIds, devicePixelRatio, zoom, caretStyle);
      return true;
    },
    []
  );

  useEffect(() => {
    if (!layout) {
      // layout reset (document change) — drop the stale pages
      generationRef.current++;
      contentEpochRef.current += 1;
      queryEpochGate.clear();
      snapshotRef.current = EMPTY_DISPLAY_LIST_SNAPSHOT;
      setSnapshot(EMPTY_DISPLAY_LIST_SNAPSHOT);
      setError(null);
      setLoading(true);
      workerFrameRequestRef.current = null;
      setWorkerSurfacesActive(false);
      setWorkerPresentationActive(false);
      notifyCaretInterrupt();
      return;
    }
    queryEpochGate.invalidate();
    const contentEpoch = contentEpochRef.current;
    const inputs = (overrides?.getInputs ?? getLayoutKernelInputs)(layout);
    const generation = ++generationRef.current;
    if (!inputs) {
      queryEpochGate.clear();
      setError(new Error('No display-list inputs were recorded for the current layout.'));
      setLoading(false);
      return;
    }
    const build = overrides?.build ?? buildRustDisplayList;
    // Merged doc-wide font chains from the Rust measure source (when active).
    // A non-empty map activates GlyphRun emission; absent ⇒ TextRunPrimitive.
    const fontChains = fontChainsProviderRef?.current?.();
    // Getters so the worker-rendered path (extras only) never materializes
    // the retained measured arena; the main-thread fallback pays the fetch once.
    const buildInputs = {
      get measured() {
        return inputs.measured;
      },
      get options() {
        return inputs.options;
      },
      layout,
      ...(inputs.headersFooters ? { headersFooters: inputs.headersFooters } : {}),
      ...(fontChains ? { fontChains } : {}),
      ...(resolvedCommentIds && resolvedCommentIds.size > 0
        ? { resolvedCommentIds: [...resolvedCommentIds].sort((a, b) => a - b) }
        : {}),
    };
    const workerEligible =
      residentEngine !== null && workerFallbackEngineRef.current !== residentEngine;
    // Cheap probe only: the full snapshot (document state, font bytes) is
    // built lazily in the queued worker request, and only for bootstrap/sync —
    // steady-state frame builds never encode state or copy fonts.
    const probe = workerEligible ? residentEngine.residentWorkerProbe() : null;
    const buildOnMainThread = (): Promise<MainThreadBuild> =>
      overrides?.build
        ? build(buildInputs, engine ?? undefined).then((displayList) => ({
            displayList,
            frame: null as RetainedFrame | null,
            queryEngine: engine,
          }))
        : buildRustDisplayFrame(buildInputs, engine ?? undefined, snapshotRef.current.frame).then(
            (result) => ({ ...result, queryEngine: engine })
          );
    const paintToken = paintedCaretMachine.token();
    const finish = (pending: Promise<MainThreadBuild>, finishGeneration: number): void => {
      pending
        .then((result) => {
          if (
            finishGeneration !== generationRef.current ||
            contentEpoch !== contentEpochRef.current
          ) {
            return;
          }
          const nextSnapshot = createRustDisplayListSnapshot(
            result.displayList,
            result.frame,
            null,
            result.queryEngine,
            snapshotRef.current
          );
          snapshotRef.current = nextSnapshot;
          publishQuerySnapshot(nextSnapshot, contentEpoch);
          setSnapshot(nextSnapshot);
          setError(null);
          setLoading(false);
          setWorkerSurfacesActive(false);
          applyPaintedCaretReply(false, paintToken);
        })
        .catch((error) => {
          if (
            finishGeneration !== generationRef.current ||
            contentEpoch !== contentEpochRef.current
          ) {
            return;
          }
          const nextError =
            error instanceof Error ? error : new Error(`Display-list build failed: ${String(error)}`);
          console.error('[CanvasRenderer] Rust display-list build failed', nextError);
          queryEpochGate.clear();
          setError(nextError);
          setLoading(false);
        });
    };
    if (!overrides?.build && probe && canUseResidentEngineWorker()) {
      const hostEngine = residentEngine;
      if (!hostEngine) throw new Error('Resident worker snapshot requires a host engine');
      const fallback = (cause: unknown) => {
        const nextError =
          cause instanceof Error
            ? cause
            : new Error(`Resident engine worker failed: ${String(cause)}`);
        console.error(
          '[CanvasRenderer] Resident engine worker unavailable; falling back to the main-thread engine',
          nextError
        );
        if (workerFallbackEngineRef.current !== hostEngine) {
          queryEpochGate.clear();
          const fallbackSnapshot = {
            ...snapshotRef.current,
            frame: null,
            queries: null,
            caret: null,
          };
          snapshotRef.current = fallbackSnapshot;
          setSnapshot(fallbackSnapshot);
          workerFallbackEngineRef.current = hostEngine;
        }
        if (workerRef.current?.engine === hostEngine) {
          workerRef.current.client.destroy();
          workerRef.current = null;
        }
        workerFrameRequestRef.current = null;
        setWorkerSurfacesActive(false);
        setWorkerPresentationActive(false);
        // The page lays out from now on; a recorded pass is laid out there
        // again, and that layout builds here.
        hostEngine.publishResidentLayout?.({ kind: 'page' });
        if (isRecordedResidentLayout(layout)) return;
        finish(buildOnMainThread(), ++generationRef.current);
      };
      try {
        if (workerRef.current?.engine !== hostEngine) {
          workerRef.current?.client.destroy();
          workerRef.current = {
            engine: hostEngine,
            client: new ResidentEngineWorkerClient(),
          };
        }
        // The newest layout wins: a request still queued picks these up.
        workerFrameRequestRef.current = {
          hostEngine,
          extras: encodeDisplayListFrameExtras(buildInputs),
          wantLayout: isRecordedResidentLayout(layout),
          fallback,
        };
        if (!workerFrameQueuedRef.current) {
          workerFrameQueuedRef.current = true;
          enqueueWorkerRequest(runWorkerFrame).catch((error) => {
            console.error('[CanvasRenderer] Resident frame publication failed', error);
          });
        }
      } catch (error) {
        fallback(error);
      }
    } else if (isRecordedResidentLayout(layout)) {
      // A recorded pass with no worker to lay it out: the page lays out.
      residentEngine?.publishResidentLayout?.({ kind: 'page' });
    } else {
      finish(buildOnMainThread(), generation);
    }
  }, [
    layout,
    overrides,
    fontChainsProviderRef,
    resolvedCommentIds,
    engine,
    residentEngine,
    setWorkerPresentationActive,
    paintedCaretMachine,
    applyPaintedCaretReply,
    notifyCaretInterrupt,
    publishQuerySnapshot,
    queryEpochGate,
    runWorkerFrame,
    enqueueWorkerRequest,
  ]);

  return {
    displayList: snapshot.displayList,
    error,
    loading,
    frame: snapshot.frame,
    queries: snapshot.queries,
    resolveQueries,
    caret: snapshot.caret,
    applyInput,
    applyDelete,
    workerSurfacesActive,
    workerPresentationActive,
    setWorkerPresentationActive,
    attachOffscreenCanvases,
    paintedCaretActive,
    startWorker,
    notifyCaretInput,
    notifyCaretInputDispatched,
    notifyCaretInterrupt,
    ensurePages,
    setVisiblePages,
    ensureVisiblePages,
    unbuiltMayHoldImages,
  };
}

function createRustDisplayListSnapshot(
  displayList: DisplayList,
  frame: RetainedFrame | null,
  caret: YrsResidentCaretSnapshot | null,
  engine: RustDisplayListEngine | null | undefined,
  previous: RustDisplayListSnapshot,
  demandPages?: (pageIndices: number[]) => void
): RustDisplayListSnapshot {
  const residentQueries = residentDisplayListQueryEngine(engine);
  return {
    displayList,
    frame,
    queries: createDisplayListQueries(displayList, residentQueries, previous.queries, demandPages),
    caret,
  };
}

function residentCaretForSelection(
  caret: YrsResidentCaretSnapshot,
  computedFor: YrsSelection | null,
  current: YrsSelection | null,
  frame: RetainedFrame
): YrsResidentCaretSnapshot | null {
  if (!computedFor || !sameYrsSelection(computedFor, current)) return null;
  const validated = residentCaretSnapshotForFrame(caret, frame);
  return validated ? { ...validated, selection: computedFor } : null;
}

/**
 * The editing wasm's own query surface, unless a panic already poisoned that
 * instance — then queries route through the layout module instead of dying on
 * the leaked borrow guard.
 */
function residentDisplayListQueryEngine(
  engine: RustDisplayListEngine | null | undefined
): ResidentDisplayListQueryEngine | undefined {
  const resident =
    engine?.displayHitTestRegionsJson &&
    engine.displayVerticalMoveJson &&
    engine.displayRangeRectsJson &&
    engine.displayRangeRectsRegionJson
      ? (engine as ResidentDisplayListQueryEngine)
      : undefined;
  return isDisplayListQuerySourceDead(resident) ? undefined : resident;
}

function isWorkerHostEngine(
  engine: RustDisplayListEngine | null | undefined
): engine is YrsSession {
  return Boolean(
    engine &&
    'residentWorkerSnapshot' in engine &&
    'onUpdate' in engine &&
    'selection' in engine &&
    'applyUpdate' in engine
  );
}

// memoized per display list: one query facade (and one JSON stringify inside
// it) per build, shared by pointer routing, selection overlay, and sidebar
// anchors. Null while the canvas is off or before the first build lands.
export function useDisplayListQueries(
  displayList: DisplayList | null,
  engine?: RustDisplayListEngine | null
): DisplayListQueries | null {
  const residentQueries = residentDisplayListQueryEngine(engine);
  // The previous facade seeds handle adoption: consecutive builds patch only
  // changed pages into the Rust query store instead of re-serializing the
  // whole display list per build.
  const previousRef = useRef<DisplayListQueries | null>(null);
  return useMemo(() => {
    const next = displayList
      ? createDisplayListQueries(displayList, residentQueries, previousRef.current)
      : null;
    previousRef.current = next;
    return next;
  }, [displayList, residentQueries]);
}

/** The engine a canvas renderer builds with: its outlines, and the sliced faces' outline installs. */
type CanvasDisplayEngine = RustDisplayListEngine & {
  outlineGlyphJson?: GlyphOutlineProvider;
  prepareSlicedPaint?: (page: unknown, fontChains: Readonly<Record<string, number[]>>, complete?: boolean) => Promise<SlicedPaint>;
  onGlyphOutlines?: (listener: (fontIds: Set<number>) => void) => () => void;
};

export interface UseCanvasRendererResult {
  /** latest display list built from the live document (null until the first pass lands) */
  displayList: DisplayList | null;
  /** Retained binary frame and its damaged page set. */
  frame: RetainedFrame | null;
  /** sole visible renderer lifecycle */
  status: 'loading' | 'ready' | 'error';
  /** fatal display-list error; non-null exactly while status is `error` */
  error: Error | null;
  /** feed PagedEditor's per-pass Layout into the interaction query source */
  onLayoutComputed: (
    layout: Layout | null,
    engine?: CanvasDisplayEngine | null
  ) => void;
  /** media resolver for CanvasPagesView */
  resolveImage: ImageResolver;
  /** Rust display-list query facade for adapter interactions. */
  queries: DisplayListQueries | null;
  /** Resolve the newest facade after pending edits and relayouts. */
  resolveQueries: ResolveDisplayListQueries;
  /** Worker caret from the same atomic renderer snapshot. */
  caret: YrsResidentCaretSnapshot | null;
  /** Whether worker-presented pixels make `caret` authoritative. */
  authoritativeCaretActive: boolean;
  /** host element of the canvas pages, so pointer routing can map client → page-local coords */
  canvasHostRef: React.RefObject<HTMLDivElement | null>;
  /** Glyph outlines sourced from the same resident font store as measurement. */
  glyphOutlineProvider: GlyphOutlineProvider | null;
  /** Readies a page's sliced faces before it paints: outlines, and FontFaces for its markers (see YrsSession); `complete` for print. */
  prepareSlicedPaint: ((page: unknown, complete?: boolean) => Promise<SlicedPaint>) | null;
  /** Font ids whose outlines installed after the page that needed them painted. */
  onGlyphOutlines: ((listener: (fontIds: Set<number>) => void) => () => void) | null;
  /** One-call ordinary text insertion; false until resident state is ready. */
  applyInput(text: string): Promise<ResidentFrameApplyResult | null>;
  /** One-call ordinary deletion/merge; false until resident state is ready. */
  applyDelete(
    direction: 'backward' | 'forward'
  ): Promise<ResidentFrameApplyResult | null>;
  setWorkerPresentationActive(active: boolean): void;
  /** OffscreenCanvas replay bridge; null keeps DOM-canvas replay. */
  offscreenReplay: {
    attach(
      pages: ResidentEngineOffscreenPage[],
      activePageIds: string[],
      devicePixelRatio: number,
      zoom: number,
      caretStyle: ResidentCaretPaintStyle
    ): Promise<boolean>;
  } | null;
  /** True while the worker-painted caret line owns the caret (DOM caret hidden). */
  paintedCaretActive: boolean;
  /** Starts the session's worker as it opens, seeded from the state it loaded. */
  startWorker(engine: YrsSession, update: Uint8Array | null): void;
  /** Local text input notification for the painted-caret mode machine. */
  notifyCaretInput(): void;
  /** Text input dispatched: hide the DOM caret before the worker round-trip. */
  notifyCaretInputDispatched(): void;
  /** Selection move / blur / IME / mode change: immediate swap to the DOM caret. */
  notifyCaretInterrupt(): void;
  /** Resolves once these pages are built (a first frame builds only the pages in view, Q815 C). */
  ensurePages(pageIndices: number[]): Promise<void>;
  /** The pages in view, which build first. */
  setVisiblePages(first: number, last: number): void;
  /** Resolves once the pages in view are built. */
  ensureVisiblePages(): Promise<void>;
}

// bundles the canvas-renderer host wiring for DocxEditor: collects each
// layout pass and rebuilds the display list through the rust engine. Canvas is
// the only visible renderer: the first build has an explicit loading state and
// a hard failure has an explicit error state. `fontChainsProviderRef`
// is the host slot the Rust measure source fills; reading it at build time is
// what activates GlyphRun emission when Rust measurement is on.
export function useCanvasRenderer(
  fontChainsProviderRef?: React.RefObject<RustFontChainsProvider | null>,
  // Resolved comment ids whose range wash the canvas hides; identity changes
  // rebuild the display list (resolve / reopen / expand-a-resolved-card).
  resolvedCommentIds?: ReadonlySet<number>,
  // The viewport the first frame fills (Q815 C): the other pages build after.
  firstPagesViewport?: React.RefObject<(() => ResidentFirstPagesViewport | null) | null>
): UseCanvasRendererResult {
  const [layout, setLayout] = useState<Layout | null>(null);
  const [engine, setEngine] = useState<
    CanvasDisplayEngine | null
  >(null);
  const onLayoutComputed = useCallback(
    (
      next: Layout | null,
      nextEngine?: CanvasDisplayEngine | null
    ) => {
      setLayout(next);
      setEngine(nextEngine ?? null);
    },
    []
  );
  const {
    displayList,
    error,
    loading,
    frame,
    queries: snapshotQueries,
    resolveQueries,
    caret,
    applyInput,
    applyDelete,
    workerSurfacesActive,
    workerPresentationActive,
    setWorkerPresentationActive,
    attachOffscreenCanvases,
    paintedCaretActive,
    startWorker,
    notifyCaretInput,
    notifyCaretInputDispatched,
    notifyCaretInterrupt,
    ensurePages,
    setVisiblePages,
    ensureVisiblePages,
    unbuiltMayHoldImages,
  } = useRustDisplayList(
    layout,
    undefined,
    fontChainsProviderRef,
    resolvedCommentIds,
    engine,
    firstPagesViewport
  );
  const resolveImage = useMemo(() => createCanvasImageResolver(), []);
  const status: UseCanvasRendererResult['status'] = error
    ? 'error'
    : loading || displayList == null
      ? 'loading'
      : 'ready';
  const [geometryReady, setGeometryReady] = useState(false);
  useEffect(() => {
    if (!snapshotQueries) {
      setGeometryReady(false);
      return;
    }
    if (snapshotQueries.isReady()) {
      setGeometryReady(true);
      return;
    }
    let cancelled = false;
    setGeometryReady(false);
    void snapshotQueries.whenReady().then(
      () => {
        if (!cancelled) setGeometryReady(true);
      },
      () => {
        if (!cancelled) setGeometryReady(false);
      }
    );
    return () => {
      cancelled = true;
    };
  }, [snapshotQueries]);
  const canvasHostRef = useRef<HTMLDivElement | null>(null);
  // Offscreen replay is the default when the browser supports transferable
  // canvas surfaces. `offscreenReplay=0` is a diagnostic escape hatch; pages
  // containing host-resolved media still select the DOM-canvas fallback in
  // CanvasPagesView.
  const offscreenAllowed = (() => {
    if (typeof window === 'undefined') return false;
    return new URLSearchParams(window.location.search).get('offscreenReplay') !== '0';
  })();
  // Keyed off surface ownership, not frame freshness: an invalidated worker
  // frame must not tear down the offscreen canvases (CanvasPagesView keys its
  // page surfaces on this), or every remote/structural edit remounts and
  // blanks the whole document. Pages not built yet may hold pictures, which
  // only the page draws: a document with media keeps the page's canvases
  // until every page is built, as when its pictures are built.
  const offscreenReplay = useMemo(
    () =>
      workerSurfacesActive &&
      !unbuiltMayHoldImages &&
      offscreenAllowed &&
      typeof OffscreenCanvas !== 'undefined' &&
      typeof HTMLCanvasElement !== 'undefined' &&
      'transferControlToOffscreen' in HTMLCanvasElement.prototype
        ? { attach: attachOffscreenCanvases }
        : null,
    [attachOffscreenCanvases, offscreenAllowed, unbuiltMayHoldImages, workerSurfacesActive]
  );
  const authoritativeCaretActive = Boolean(
    workerPresentationActive &&
      displayList &&
      caret?.caretRect &&
      !displayListNeedsHostImages(displayList)
  );
  // Markers and watermarks name their faces by family, through the measurement's chains.
  const prepareSlicedPaint = useMemo(() => {
    const prepare = engine?.prepareSlicedPaint;
    return prepare ? (page: unknown, complete?: boolean) => prepare(page, fontChainsProviderRef?.current?.() ?? {}, complete) : null;
  }, [engine, fontChainsProviderRef]);
  return {
    displayList,
    frame,
    // The display surface is already valid when a delta lands. Keep it mounted
    // while the replacement main-thread geometry cache warms; interactions are
    // briefly gated by a null query facade instead of throwing away every page
    // canvas and forcing a full replay.
    status,
    error,
    onLayoutComputed,
    resolveImage,
    queries: geometryReady ? snapshotQueries : null,
    resolveQueries,
    caret,
    authoritativeCaretActive,
    canvasHostRef,
    glyphOutlineProvider: engine?.outlineGlyphJson ?? null,
    prepareSlicedPaint,
    onGlyphOutlines: engine?.onGlyphOutlines ?? null,
    applyInput,
    applyDelete,
    setWorkerPresentationActive,
    offscreenReplay,
    paintedCaretActive,
    startWorker,
    notifyCaretInput,
    notifyCaretInputDispatched,
    notifyCaretInterrupt,
    ensurePages,
    setVisiblePages,
    ensureVisiblePages,
  };
}
