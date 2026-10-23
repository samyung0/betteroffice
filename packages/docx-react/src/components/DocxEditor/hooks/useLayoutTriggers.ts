/**
 * Layout-trigger effects for PagedEditor.
 *
 * Re-runs `runLayoutPipeline` for two state-shifts that the pipeline's
 * own dep array doesn't catch automatically:
 *
 *  1. Web-font loading completes — measurement is Rust-only (font bytes,
 *     not browser fonts), so the layout itself cannot change; the re-run
 *     re-rasterizes canvas text drawn through CSS-font fallback paths so
 *     late-loading embedded faces show up. It is no edit, so it lays out as
 *     a peer's update does: the viewport stays and the caret is not scrolled
 *     to (a wider toolbar after a rotation loads faces too).
 *
 *  2. Header / footer content or render-env changes — runLayoutPipeline does include
 *     these in its deps, but only re-runs when explicitly called. The
 *     first render already laid out when the Yrs session became ready, so this
 *     effect skips the initial render via a one-shot epoch counter. Only
 *     the hidden-text toggle is the reader's action here: it lays out as a
 *     local pass and keeps the caret in view. A peer's change (a comment
 *     anywhere renumbers the env, a peer's header edit) keeps the view; the
 *     reader's own header or footer save schedules its local pass itself
 *     (DocxEditor's header/footer push).
 */

import { useEffect, useRef } from 'react';

import type { HeaderFooter } from '@betteroffice/docx/types/document';
import type { YrsRenderEnv } from '@betteroffice/docx/yrs';
import type { LayoutUpdateOrigin } from '../internals/viewportAnchoring';
export interface UseLayoutTriggersOptions {
  runLayoutPipeline: () => void;
  scheduleLayout: (origin?: LayoutUpdateOrigin) => void;
  updateSelectionOverlay: () => void;
  headerContent?: HeaderFooter | null;
  footerContent?: HeaderFooter | null;
  firstPageHeaderContent?: HeaderFooter | null;
  firstPageFooterContent?: HeaderFooter | null;
  renderEnv?: YrsRenderEnv;
}

export function useLayoutTriggers(opts: UseLayoutTriggersOptions): void {
  const {
    runLayoutPipeline,
    scheduleLayout,
    updateSelectionOverlay,
    headerContent,
    footerContent,
    firstPageHeaderContent,
    firstPageFooterContent,
    renderEnv,
  } = opts;
  const runLayoutPipelineRef = useRef(runLayoutPipeline);
  runLayoutPipelineRef.current = runLayoutPipeline;
  const scheduleLayoutRef = useRef(scheduleLayout);
  scheduleLayoutRef.current = scheduleLayout;
  const updateSelectionOverlayRef = useRef(updateSelectionOverlay);
  updateSelectionOverlayRef.current = updateSelectionOverlay;

  // Re-layout on web-font load. FontFaceSet.onloadingdone catches new
  // fonts as they finish loading.
  useEffect(() => {
    const handleFontsLoaded = () => {
      scheduleLayoutRef.current('remote');
      updateSelectionOverlayRef.current();
    };
    window.document.fonts.addEventListener('loadingdone', handleFontsLoaded);
    return () => {
      window.document.fonts.removeEventListener('loadingdone', handleFontsLoaded);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Re-layout when H/F content or the render env changes (HF editor save,
  // showHiddenText toggle, a peer's edit or comment). Only the toggle is a
  // local pass here (Epo 2026-10-23).
  const contentEpochRef = useRef(0);
  const showHiddenTextRef = useRef(renderEnv?.showHiddenText);
  useEffect(() => {
    // Skip the initial render — session readiness already triggered the first layout.
    if (contentEpochRef.current === 0) {
      contentEpochRef.current = 1;
      return;
    }
    if (renderEnv?.showHiddenText !== showHiddenTextRef.current) {
      showHiddenTextRef.current = renderEnv?.showHiddenText;
      scheduleLayoutRef.current('local');
    }
    runLayoutPipelineRef.current();
  }, [
    headerContent,
    footerContent,
    firstPageHeaderContent,
    firstPageFooterContent,
    renderEnv,
  ]);
}
