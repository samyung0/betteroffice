/** Resident Rust layout scheduling and React paint-state publication. */

import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';

import type { LayoutBlock, Layout } from '@betteroffice/docx/layout/pagination';
import {
  adoptResidentLayout,
  buildResidentRegionLayoutRequest,
  computeLayout,
  recordResidentLayout,
  type LayoutComputation,
} from '@betteroffice/docx/editor';
import { findVerticalScrollParentOrRoot } from '@betteroffice/docx/utils/findVerticalScrollParent';
import type { Document } from '@betteroffice/docx/types/document';
import type {
  ResidentFontRequirement,
  ResidentMeasurementConfig,
} from '@betteroffice/docx/layout';
import {
  canUseResidentEngineWorker,
  type YrsLoc,
  type YrsRenderEnv,
  type YrsSession,
  type YrsStickyPosition,
} from '@betteroffice/docx/yrs';

import type { LayoutSelectionGate } from '../internals/LayoutSelectionGate';
import {
  CANVAS_PAGE_GAP_PX,
  resolveDisplayPageClientRect,
  type DisplayListQueries,
} from '@betteroffice/docx/layout/render';
import { viewportMinHeightPx } from '../internals/scrollUtils';
import {
  captureDisplayListScrollAnchor,
  captureDisplayListViewportAnchor,
  restoreDisplayListScrollAnchor,
  restoreDisplayListViewportAnchor,
  restoreScrollSnapshot,
  type DisplayListScrollAnchor,
  type DisplayListViewportAnchor,
} from '../internals/scrollRestore';
import {
  layoutPassOrigin,
  mergeLayoutUpdateOrigin,
  PendingScrollRestoreController,
  rescaleScrollTop,
  type LayoutUpdateOrigin,
} from '../internals/viewportAnchoring';

interface SelectionScrollRestore {
  kind: 'selection';
  anchor: DisplayListScrollAnchor;
}

interface ViewportScrollRestore {
  kind: 'viewport';
  anchor: DisplayListViewportAnchor;
}

type PendingScrollRestore = SelectionScrollRestore | ViewportScrollRestore;

interface CurrentViewportAnchor {
  anchor: DisplayListViewportAnchor;
  navigationEpoch: number;
}

export interface UseLayoutPipelineOptions {
  onError?: (error: Error) => void;
  document: Document | null;
  session: YrsSession | null;
  renderEnv: YrsRenderEnv;
  pageGap: number;
  zoom: number;
  residentMeasurementConfig: (
    requirements: ResidentFontRequirement[]
  ) => ResidentMeasurementConfig | null;
  /**
   * Rust measurement readiness gate (`useRustMeasurement.deferLayoutPass`).
   * Checked before computing (engine may still be loading) and before
   * committing (a pass may have discovered unresolved font chains); a
   * deferred pass is skipped/discarded and the measurement hook re-runs the
   * pipeline once the engine/fonts settle.
   */
  deferLayoutPass: (blocks?: LayoutBlock[]) => boolean;
  /** Queries available while the canvas renderer is active. */
  displayListQueries?: DisplayListQueries | null;
  interactionPageHostRef?: React.RefObject<HTMLDivElement | null>;
  pagesContainerRef: React.RefObject<HTMLDivElement | null>;
  viewportLayoutRef: React.RefObject<HTMLDivElement | null>;
  /** Current display-list position used to preserve the scroll anchor. */
  getSelectionHead?: () => number;
  displayPositionToYrsLoc?: (position: number) => YrsLoc | null;
  yrsLocToDisplayPosition?: (loc: YrsLoc) => number | null;
  syncCoordinator: LayoutSelectionGate;
  getScrollContainer: () => HTMLDivElement | null;
  onTotalPagesChange?: (totalPages: number) => void;
  /** Receives each computed layout and resets with null. */
  onLayoutComputed?: (layout: Layout | null) => void;
  onAnchorPositionsChange?: (positions: Map<string, number>) => void;
  /**
   * The canvas renderer's resident worker lays out every pass: the page only
   * records it and adopts the worker's layout. A table of contents (which
   * reads its pages at once) and every pass after the worker failed lay out
   * on the page.
   */
  residentWorkerLayout?: boolean;
}

export interface LayoutPipelineOptions {
  /** Lay out on the page even when the worker lays out passes. */
  onPage?: boolean;
}

export interface UseLayoutPipelineReturn {
  layout: Layout | null;
  layoutUpdateOrigin: LayoutUpdateOrigin;
  runLayoutPipeline: (options?: LayoutPipelineOptions) => void;
  /** The layout the last pass computed, before React renders it. */
  getLatestLayout: () => Layout | null;
  scheduleLayout: (origin?: LayoutUpdateOrigin) => void;
  cancelPendingScrollRestore: () => void;
}

export function useLayoutPipeline(opts: UseLayoutPipelineOptions): UseLayoutPipelineReturn {
  const {
    document,
    session,
    renderEnv,
    pageGap,
    zoom,
    residentMeasurementConfig,
    deferLayoutPass,
    displayListQueries,
    interactionPageHostRef,
    pagesContainerRef,
    viewportLayoutRef,
    getSelectionHead,
    displayPositionToYrsLoc,
    yrsLocToDisplayPosition,
    syncCoordinator,
    getScrollContainer,
    onError,
    onTotalPagesChange,
    onLayoutComputed,
    onAnchorPositionsChange,
    residentWorkerLayout = false,
  } = opts;

  const [layout, setLayout] = useState<Layout | null>(null);
  const latestLayoutRef = useRef<Layout | null>(null);
  // The latest pass recorded for the resident worker: the canvas renderer
  // builds it in the worker, whose layout then becomes `layout` for the page's
  // readers. Null once a pass lays out on the page.
  const [recordedPass, setRecordedPassState] = useState<{
    layout: Layout;
    session: YrsSession;
  } | null>(null);
  const recordedPassRef = useRef(recordedPass);
  const setRecordedPass = useCallback(
    (recorded: Layout | null) => {
      const next = recorded && session ? { layout: recorded, session } : null;
      recordedPassRef.current = next;
      setRecordedPassState(next);
    },
    [session]
  );
  // The session whose worker failed: its passes lay out on the page.
  const workerFailedSessionRef = useRef<YrsSession | null>(null);
  // The last font check and what it read: a pass over the same request and
  // document (the one after the fonts it asked for load, a web-font load's)
  // reuses its answer instead of reading every story again (Q815).
  const fontCheckRef = useRef<{
    session: YrsSession;
    docEpoch: number;
    input: string;
    requirements: ResidentFontRequirement[];
  } | null>(null);

  // Callback refs — parent may hand in a fresh closure every render. Mirroring
  // these in refs keeps `runLayoutPipeline`'s dep array stable; otherwise
  // every parent re-render would invalidate the rAF-coalesced scheduler.
  const onErrorRef = useRef(onError);
  onErrorRef.current = onError;
  const onTotalPagesChangeRef = useRef(onTotalPagesChange);
  const onLayoutComputedRef = useRef(onLayoutComputed);
  const onAnchorPositionsChangeRef = useRef(onAnchorPositionsChange);
  // Query facades are immutable per display-list build. Reading the current
  // facade through a ref keeps runLayoutPipeline identity-stable when a build
  // lands; otherwise useLayoutTriggers sees a new callback, starts another
  // layout, and creates a display-list -> layout feedback loop.
  const displayListQueriesRef = useRef(displayListQueries);
  const deferLayoutPassRef = useRef(deferLayoutPass);
  const displayPositionToYrsLocRef = useRef(displayPositionToYrsLoc);
  const yrsLocToDisplayPositionRef = useRef(yrsLocToDisplayPosition);
  onTotalPagesChangeRef.current = onTotalPagesChange;
  onLayoutComputedRef.current = onLayoutComputed;
  onAnchorPositionsChangeRef.current = onAnchorPositionsChange;
  displayListQueriesRef.current = displayListQueries;
  deferLayoutPassRef.current = deferLayoutPass;
  displayPositionToYrsLocRef.current = displayPositionToYrsLoc;
  yrsLocToDisplayPositionRef.current = yrsLocToDisplayPosition;

  // Total-pages notifier — fires only when count changes (including N → 0).
  const lastTotalPagesRef = useRef<number>(0);
  useEffect(() => {
    // A pass recorded for an earlier session leaves this one nothing to build.
    const recorded = recordedPass?.session === session ? recordedPass.layout : null;
    onLayoutComputedRef.current?.(recorded ?? (recordedPass ? null : layout));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [recordedPass, layout]);
  useEffect(() => {
    const total = layout?.pages.length ?? 0;
    if (total === lastTotalPagesRef.current) return;
    lastTotalPagesRef.current = total;
    onTotalPagesChangeRef.current?.(total);
  }, [layout]);

  const scrollRestoreControllerRef =
    useRef<PendingScrollRestoreController<PendingScrollRestore> | null>(null);
  if (!scrollRestoreControllerRef.current) {
    scrollRestoreControllerRef.current =
      new PendingScrollRestoreController<PendingScrollRestore>();
  }
  const scrollRestoreController = scrollRestoreControllerRef.current;
  const navigationEpochRef = useRef(0);
  const cancelPendingScrollRestore = useCallback(() => {
    navigationEpochRef.current += 1;
    scrollRestoreController.cancel();
  }, [scrollRestoreController]);
  const currentViewportAnchorRef = useRef<CurrentViewportAnchor | null>(null);
  const viewportAnchorSessionRef = useRef(session);
  const viewportAnchorCaptureReadyRef = useRef(true);
  const expectedScrollTopRef = useRef<number | null>(null);
  if (viewportAnchorSessionRef.current !== session) {
    viewportAnchorSessionRef.current = session;
    currentViewportAnchorRef.current = null;
  }
  const pendingLayoutOriginRef = useRef<LayoutUpdateOrigin | null>(null);
  const layoutUpdateOriginRef = useRef<LayoutUpdateOrigin>('local');

  const captureViewportPosition = useCallback(
    (position: number) => {
      const loc = displayPositionToYrsLocRef.current?.(position);
      // Round-trip guard: a projection that cannot map the loc back to the same
      // display position would resolve the anchor into another story on restore.
      if (!session || !loc || yrsLocToDisplayPositionRef.current?.(loc) !== position) return null;
      try {
        return session.encodeStickyPosition(loc);
      } catch {
        return null;
      }
    },
    [session]
  );
  const resolveViewportPosition = useCallback(
    (position: YrsStickyPosition) => {
      const loc = session?.resolveStickyPosition(position);
      if (!loc) return null;
      return yrsLocToDisplayPositionRef.current?.(loc) ?? null;
    },
    [session]
  );
  const captureCurrentViewportAnchor = useCallback(() => {
    const queries = displayListQueriesRef.current;
    const pagesEl = pagesContainerRef.current;
    const host = interactionPageHostRef?.current ?? pagesEl;
    const scrollParent =
      getScrollContainer() ?? (pagesEl ? findVerticalScrollParentOrRoot(pagesEl) : null);
    const anchor =
      queries && host && scrollParent?.isConnected
        ? captureDisplayListViewportAnchor(queries, host, scrollParent, captureViewportPosition)
        : null;
    currentViewportAnchorRef.current = anchor
      ? { anchor, navigationEpoch: navigationEpochRef.current }
      : null;
  }, [
    captureViewportPosition,
    getScrollContainer,
    interactionPageHostRef,
    pagesContainerRef,
  ]);

  // =========================================================================
  // Layout Pipeline
  // =========================================================================

  const runLayoutPipeline = useCallback(
    (options?: LayoutPipelineOptions) => {
      const layoutUpdateOrigin = layoutPassOrigin(pendingLayoutOriginRef.current);
      pendingLayoutOriginRef.current = null;
      if (layoutUpdateOrigin === 'local') scrollRestoreController.cancel();
      const pipelineStart = performance.now();

      const currentEpoch = syncCoordinator.getStateSeq();
      syncCoordinator.onLayoutStart();

      if (deferLayoutPassRef.current() || !session) {
        pendingLayoutOriginRef.current = mergeLayoutUpdateOrigin(
          pendingLayoutOriginRef.current,
          layoutUpdateOrigin
        );
        syncCoordinator.onLayoutComplete(currentEpoch);
        return;
      }

      let measurement: ResidentMeasurementConfig | null = null;
      try {
        const input = JSON.stringify(
          buildResidentRegionLayoutRequest(document, pageGap, renderEnv)
        );
        const docEpoch = session.docEpoch();
        const known = fontCheckRef.current;
        const requirements =
          known?.session === session && known.docEpoch === docEpoch && known.input === input
            ? known.requirements
            : (JSON.parse(session.layoutFontRequirementsJson(input)) as ResidentFontRequirement[]);
        fontCheckRef.current = { session, docEpoch, input, requirements };
        measurement = residentMeasurementConfig(requirements);
      } catch (error) {
        console.error('[PagedEditor] Resident font preflight error:', error);
        onErrorRef.current?.(error instanceof Error ? error : new Error(String(error)));
        syncCoordinator.onLayoutComplete(currentEpoch);
        return;
      }
      if (!measurement) {
        pendingLayoutOriginRef.current = mergeLayoutUpdateOrigin(
          pendingLayoutOriginRef.current,
          layoutUpdateOrigin
        );
        syncCoordinator.onLayoutComplete(currentEpoch);
        return;
      }

      const computeInputs = { document, pageGap, session, renderEnv, measurement };

      // Step 4+: paint + scroll/events with the computed values.
      const captureAnchor = () => {
        const pagesEl = pagesContainerRef.current;
        const scrollParent =
          getScrollContainer() ?? (pagesEl ? findVerticalScrollParentOrRoot(pagesEl) : null);
        const interactionHost = interactionPageHostRef?.current ?? pagesEl;
        const queries = displayListQueriesRef.current;
        const currentViewportAnchor = currentViewportAnchorRef.current;
        const anchor =
          scrollParent?.isConnected && interactionHost && queries
            ? layoutUpdateOrigin === 'remote'
              ? currentViewportAnchor?.navigationEpoch === navigationEpochRef.current
                ? {
                    kind: 'viewport' as const,
                    anchor: currentViewportAnchor.anchor,
                  }
                : null
              : {
                  kind: 'selection' as const,
                  anchor: captureDisplayListScrollAnchor(
                    queries,
                    interactionHost,
                    scrollParent,
                    getSelectionHead?.() ?? 0
                  ),
                }
            : null;

        viewportAnchorCaptureReadyRef.current = false;
        layoutUpdateOriginRef.current = layoutUpdateOrigin;
        if (scrollParent?.isConnected && anchor) {
          scrollRestoreController.capture(anchor);
        } else {
          scrollRestoreController.cancel();
        }
      };
      const applyComputation = (computation: LayoutComputation) => {
        captureAnchor();
        commitLayoutRef.current(computation.layout);
        setRecordedPass(null);
        const totalTime = performance.now() - pipelineStart;
        if (totalTime > 2000) {
          console.warn(
            `[PagedEditor] Layout pipeline took ${Math.round(totalTime)}ms total ` +
              `(${computation.layout.pages.length} pages)`
          );
        }
      };

      // The resident worker lays the pass out: the page records it for the
      // canvas renderer and adopts the worker's layout when it arrives.
      if (
        residentWorkerLayout &&
        !options?.onPage &&
        workerFailedSessionRef.current !== session &&
        session.recordResidentLayout &&
        canUseResidentEngineWorker()
      ) {
        try {
          captureAnchor();
          setRecordedPass(recordResidentLayout(computeInputs));
        } catch (error) {
          console.error('[PagedEditor] Resident layout request error:', error);
          onErrorRef.current?.(error instanceof Error ? error : new Error(String(error)));
        }
        syncCoordinator.onLayoutComplete(currentEpoch);
        return;
      }

      // Every pagination pass performs a full relayout.
      try {
        applyComputation(computeLayout(computeInputs));
      } catch (error) {
        console.error('[PagedEditor] Layout pipeline error:', error);
        onErrorRef.current?.(error instanceof Error ? error : new Error(String(error)));
      }
      syncCoordinator.onLayoutComplete(currentEpoch);
    },
    [
      pageGap,
      zoom,
      syncCoordinator,
      document,
      session,
      renderEnv,
      residentMeasurementConfig,
      getScrollContainer,
      getSelectionHead,
      interactionPageHostRef,
      pagesContainerRef,
      viewportLayoutRef,
      scrollRestoreController,
      setRecordedPass,
      residentWorkerLayout,
    ]
  );

  // The layout the page's readers see, and the scroll height it gives.
  const commitLayout = useCallback(
    (newLayout: Layout) => {
      latestLayoutRef.current = newLayout;
      setLayout(newLayout);
      const vp = viewportLayoutRef.current;
      if (vp) {
        const mh = viewportMinHeightPx(newLayout, pageGap);
        vp.style.minHeight = `${mh}px`;
        vp.style.marginBottom = zoom !== 1 ? `${mh * (zoom - 1)}px` : '';
      }
    },
    [pageGap, zoom, viewportLayoutRef]
  );
  const commitLayoutRef = useRef(commitLayout);
  commitLayoutRef.current = commitLayout;

  // Hold the exact scrollTop while the next display-list commit is built.
  useLayoutEffect(() => {
    const ticket = scrollRestoreController.peek();
    if (!ticket) return;
    const pagesEl = pagesContainerRef.current;
    const scrollParent =
      getScrollContainer() ?? (pagesEl ? findVerticalScrollParentOrRoot(pagesEl) : null);
    if (scrollParent?.isConnected) {
      scrollRestoreController.run(ticket, () => {
        restoreScrollSnapshot(ticket.value.anchor, scrollParent);
        expectedScrollTopRef.current = scrollParent.scrollTop;
      });
    }
  }, [layout, getScrollContainer, pagesContainerRef, scrollRestoreController]);

  // A new immutable display-list/query facade is the geometry commit signal.
  useLayoutEffect(() => {
    const ticket = scrollRestoreController.peek();
    const pagesEl = pagesContainerRef.current;
    const host = interactionPageHostRef?.current ?? pagesEl;
    const scrollParent =
      getScrollContainer() ?? (pagesEl ? findVerticalScrollParentOrRoot(pagesEl) : null);
    if (!ticket || !displayListQueries || !host || !scrollParent?.isConnected) return;
    const pending = scrollRestoreController.take();
    if (!pending) return;
    const restore = (): void => {
      if (pending.value.kind === 'viewport') {
        restoreDisplayListViewportAnchor(
          pending.value.anchor,
          displayListQueries,
          host,
          scrollParent,
          resolveViewportPosition
        );
      } else {
        restoreDisplayListScrollAnchor(
          pending.value.anchor,
          displayListQueries,
          host,
          scrollParent
        );
      }
      expectedScrollTopRef.current = scrollParent.scrollTop;
    };
    if (!scrollRestoreController.run(pending, restore)) return;
    const rafId = requestAnimationFrame(() => {
      if (scrollParent.isConnected) scrollRestoreController.run(pending, restore);
    });
    return () => cancelAnimationFrame(rafId);
  }, [
    displayListQueries,
    getScrollContainer,
    interactionPageHostRef,
    pagesContainerRef,
    resolveViewportPosition,
    scrollRestoreController,
  ]);

  useLayoutEffect(() => {
    if (!displayListQueries) return;
    viewportAnchorCaptureReadyRef.current = true;
    captureCurrentViewportAnchor();
    const rafId = requestAnimationFrame(() => {
      captureCurrentViewportAnchor();
    });
    return () => cancelAnimationFrame(rafId);
  }, [captureCurrentViewportAnchor, displayListQueries]);

  // A zoom change (a re-fit on resize, View › Zoom, a pinch) keeps the point at
  // the pane's top where it was, as View's scroll anchoring does. It rescales
  // the last scroll position seen, so a fling or a scroll just before needs no
  // captured anchor, and a shrink that clamped the scroll does not count.
  const anchoredZoomRef = useRef(zoom);
  const lastScrollTopRef = useRef<number | null>(null);
  useLayoutEffect(() => {
    const from = anchoredZoomRef.current;
    if (from === zoom) return;
    anchoredZoomRef.current = zoom;
    const queries = displayListQueriesRef.current;
    const pagesEl = pagesContainerRef.current;
    const host = interactionPageHostRef?.current ?? pagesEl;
    const scrollParent =
      getScrollContainer() ?? (pagesEl ? findVerticalScrollParentOrRoot(pagesEl) : null);
    if (!queries || !host || !scrollParent?.isConnected) return;
    const first = resolveDisplayPageClientRect(host, queries, 0);
    if (!first) return;
    const columnTop =
      scrollParent.scrollTop + first.top - scrollParent.getBoundingClientRect().top;
    const heights = Array.from(
      { length: queries.pageCount() },
      (_, index) => queries.pageSize(index)?.height ?? 0
    );
    scrollParent.scrollTop = rescaleScrollTop(
      lastScrollTopRef.current ?? scrollParent.scrollTop,
      columnTop,
      heights,
      CANVAS_PAGE_GAP_PX,
      from,
      zoom
    );
    expectedScrollTopRef.current = scrollParent.scrollTop;
    lastScrollTopRef.current = scrollParent.scrollTop;
  }, [zoom, getScrollContainer, interactionPageHostRef, pagesContainerRef]);

  useEffect(() => {
    const pagesEl = pagesContainerRef.current;
    const scrollParent =
      getScrollContainer() ?? (pagesEl ? findVerticalScrollParentOrRoot(pagesEl) : null);
    if (!scrollParent) return;
    scrollParent.addEventListener('wheel', cancelPendingScrollRestore, {
      capture: true,
      passive: true,
    });
    scrollParent.addEventListener('touchstart', cancelPendingScrollRestore, {
      capture: true,
      passive: true,
    });
    scrollParent.addEventListener('pointerdown', cancelPendingScrollRestore, {
      capture: true,
      passive: true,
    });
    let captureTimer: ReturnType<typeof setTimeout> | null = null;
    const captureAfterScroll = (): void => {
      lastScrollTopRef.current = scrollParent.scrollTop;
      const expectedScrollTop = expectedScrollTopRef.current;
      expectedScrollTopRef.current = null;
      if (
        expectedScrollTop == null ||
        Math.abs(scrollParent.scrollTop - expectedScrollTop) > 0.5
      ) {
        cancelPendingScrollRestore();
      }
      if (captureTimer !== null) clearTimeout(captureTimer);
      captureTimer = setTimeout(() => {
        captureTimer = null;
        if (viewportAnchorCaptureReadyRef.current) captureCurrentViewportAnchor();
      }, 80);
    };
    scrollParent.addEventListener('scroll', captureAfterScroll, { passive: true });
    return () => {
      scrollParent.removeEventListener('wheel', cancelPendingScrollRestore, true);
      scrollParent.removeEventListener('touchstart', cancelPendingScrollRestore, true);
      scrollParent.removeEventListener('pointerdown', cancelPendingScrollRestore, true);
      scrollParent.removeEventListener('scroll', captureAfterScroll);
      if (captureTimer !== null) clearTimeout(captureTimer);
    };
  }, [
    cancelPendingScrollRestore,
    captureCurrentViewportAnchor,
    getScrollContainer,
    pagesContainerRef,
  ]);

  // =========================================================================
  // Coalesced Layout (rAF throttle)
  // =========================================================================

  /**
   * Multiple rapid transactions (e.g. typing "hello") within the same frame
   * are coalesced so only the final state triggers a full layout pass. The
   * coalescer lives in core (`createLayoutScheduler`) so React and Vue share
   * it; the `runRef` indirection lets the stable scheduler always call the
   * latest `runLayoutPipeline` without recreating itself.
   */
  const runRef = useRef(runLayoutPipeline);
  runRef.current = runLayoutPipeline;
  const schedulerRef = useRef<number | null>(null);
  const scheduleLayout = useCallback((origin: LayoutUpdateOrigin = 'local') => {
    if (origin === 'local') scrollRestoreController.cancel();
    pendingLayoutOriginRef.current = mergeLayoutUpdateOrigin(
      pendingLayoutOriginRef.current,
      origin
    );
    if (schedulerRef.current != null) return;
    schedulerRef.current = requestAnimationFrame(() => {
      schedulerRef.current = null;
      if (pendingLayoutOriginRef.current) runRef.current();
    });
  }, [scrollRestoreController]);

  // The resident worker's layout reaches the page's readers. A worker that
  // failed hands layout back: a pass it was to lay out is laid out here, as a
  // peer's is (the view stays).
  useEffect(() => {
    if (!session?.onResidentLayout) return;
    return session.onResidentLayout((reply) => {
      if (reply.kind === 'layout') {
        try {
          commitLayoutRef.current(adoptResidentLayout(reply.json));
        } catch (error) {
          console.error('[PagedEditor] Resident layout reply error:', error);
          onErrorRef.current?.(error instanceof Error ? error : new Error(String(error)));
        }
        return;
      }
      workerFailedSessionRef.current = session;
      if (recordedPassRef.current?.session === session) scheduleLayout('remote');
    });
  }, [session, scheduleLayout]);

  // Clean up pending rAF on unmount
  useEffect(() => {
    return () => {
      if (schedulerRef.current != null) cancelAnimationFrame(schedulerRef.current);
    };
  }, []);

  return {
    layout,
    layoutUpdateOrigin: layoutUpdateOriginRef.current,
    runLayoutPipeline,
    getLatestLayout: useCallback(() => latestLayoutRef.current, []),
    scheduleLayout,
    cancelPendingScrollRestore,
  };
}
