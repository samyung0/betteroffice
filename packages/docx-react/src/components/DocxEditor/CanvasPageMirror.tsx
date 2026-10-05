/**
 * Mounts the accessibility mirror (core `buildMirrorPage`) 1:1 under one
 * canvas page: same origin, same page-local pixel space, so
 * `getBoundingClientRect` on mirror nodes returns the rects the canvas
 * painted. The mirror is invisible (opacity 0) and inert to the pointer
 * (pointer-events none) but deliberately NOT aria-hidden — it is the
 * accessible content of the canvas. Rebuilt whenever the page's display list
 * changes — the same trigger that re-rasters the canvas.
 *
 * Pages away from the viewport (`full` false) mount the plain-text form
 * (`buildMirrorTextPage`): the same accessible text without positioned runs,
 * so screen readers reach the whole document without every scroll frame
 * restyling every page's positioned mirror.
 *
 * Focus never lands here: the hidden input remains the editing surface.
 *
 * `selectable` (read-only viewers) turns the positioned mirror into the page's
 * text layer, as a PDF viewer's: its text stays transparent over the canvas and
 * only the selection shows (styles: `.canvas-page-mirror--selectable`). A page
 * holding a selection end keeps its positioned form when it leaves the
 * viewport, until the selection lets go of it, so scrolling away and back
 * keeps the selection.
 */

import { useEffect, useRef } from 'react';
import {
  buildMirrorPage,
  buildMirrorTextPage,
  displayPageRevision,
  type DisplayPage,
} from '@betteroffice/docx/layout/render';
import type { TFunction } from '@betteroffice/docx-i18n';
import { useTranslation } from '../../i18n';

export function CanvasPageMirror({
  page,
  zoom = 1,
  defer = false,
  full = true,
  selectable = false,
}: {
  page: DisplayPage;
  zoom?: number;
  /** Off-window pages build at idle time instead of inside the mount flush. */
  defer?: boolean;
  /** Positioned mirror near the viewport; plain accessible text elsewhere. */
  full?: boolean;
  /** Text layer of a read-only viewer: selectable, transparent text. */
  selectable?: boolean;
}) {
  const hostRef = useRef<HTMLDivElement>(null);
  // Position-shift deltas mutate primitives in place — identity alone is stale.
  const builtForRef = useRef<{
    page: DisplayPage;
    revision: number;
    t: TFunction;
    full: boolean;
  } | null>(null);
  const { t } = useTranslation();

  useEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    const built = builtForRef.current;
    if (
      built?.page === page &&
      built.revision === displayPageRevision(page) &&
      built.t === t &&
      built.full === full
    ) {
      return;
    }
    // A page crossing the viewport window already holds the other form of its
    // mirror, so the swap waits for idle time instead of a scroll frame.
    const swapOnly =
      built?.page === page && built.revision === displayPageRevision(page) && built.t === t;
    let releaseHold: (() => void) | undefined;
    const build = (): void => {
      // Going plain would drop the selection end this page holds: wait until
      // the selection lets go of it.
      if (swapOnly && selectable && !full && holdsSelectionEnd(host)) {
        const doc = host.ownerDocument;
        const retry = (): void => {
          if (holdsSelectionEnd(host)) return;
          doc.removeEventListener('selectionchange', retry);
          releaseHold = undefined;
          build();
        };
        doc.addEventListener('selectionchange', retry);
        releaseHold = () => doc.removeEventListener('selectionchange', retry);
        return;
      }
      const mirror = (full ? buildMirrorPage : buildMirrorTextPage)(page, {
        labels: {
          page: t('a11y.pageLabel', { number: page.pageIndex + 1 }),
          header: t('a11y.headerLabel'),
          footer: t('a11y.footerLabel'),
        },
      });
      // Keep the previous mirror connected until this replacement is ready.
      // Clearing in effect cleanup creates a detached-DOM window on every page
      // update; unmounting already removes the host and its complete subtree.
      host.replaceChildren(mirror);
      builtForRef.current = { page, revision: displayPageRevision(page), t, full };
    };
    if (!defer && !swapOnly) {
      build();
      return;
    }
    if (typeof requestIdleCallback === 'function') {
      const id = requestIdleCallback(build, { timeout: 1500 });
      return () => {
        cancelIdleCallback(id);
        releaseHold?.();
      };
    }
    const id = setTimeout(build, 150);
    return () => {
      clearTimeout(id);
      releaseHold?.();
    };
  }, [page, t, defer, full, selectable]);

  return (
    <div
      ref={hostRef}
      className={
        selectable ? 'canvas-page-mirror canvas-page-mirror--selectable' : 'canvas-page-mirror'
      }
      // The mirror content is built in page-local px; when the canvas is
      // enlarged for zoom (CSS size = page * zoom), CSS-scale the mirror by the
      // same factor from its top-left origin so its nodes' `getBoundingClientRect`
      // still lands on the painted glyphs. At zoom = 1 this is an identity scale.
      style={{
        position: 'absolute',
        left: 0,
        top: 0,
        pointerEvents: 'none',
        transform: zoom !== 1 ? `scale(${zoom})` : undefined,
        transformOrigin: '0 0',
      }}
    />
  );
}

/** Also a caret: a drag starts as one, before its first extension. */
function holdsSelectionEnd(host: HTMLElement): boolean {
  const selection = host.ownerDocument.getSelection();
  if (!selection) return false;
  return host.contains(selection.anchorNode) || host.contains(selection.focusNode);
}
