/**
 * Mounts the accessibility mirror (core `buildMirrorPage`) 1:1 under one
 * canvas page: same origin, same page-local pixel space, so
 * `getBoundingClientRect` on mirror nodes returns the rects the canvas
 * painted. The mirror is invisible (opacity 0) and inert to the pointer
 * (pointer-events none) but deliberately NOT aria-hidden — it is the
 * accessible content of the canvas. Rebuilt whenever the page's display list
 * changes — the same trigger that re-rasters the canvas — and patched into the
 * mounted mirror, so unchanged blocks keep their nodes. A position shift
 * (same page object, new revision) rewrites the existing nodes' positions
 * instead.
 *
 * Pages away from the viewport (`full` false) mount the plain-text form
 * (`buildMirrorTextPage`): the same accessible text without positioned runs,
 * so screen readers reach the whole document without every scroll frame
 * restyling every page's positioned mirror.
 *
 * Focus never lands here: the hidden input remains the editing surface.
 *
 * The positioned mirror joins a line's touching runs, so screen readers read
 * whole words. `selectable` (read-only viewers) turns it into the page's text
 * layer instead, as a PDF viewer's: one element per glyph cluster so selection
 * sits on the painted glyphs, its text transparent over the canvas and only the
 * selection showing (styles: `.canvas-page-mirror--selectable`). A page
 * holding a selection end keeps its positioned form when it leaves the
 * viewport, until the selection lets go of it, so scrolling away and back
 * keeps the selection.
 *
 * Screen readers would spell that text layer glyph by glyph, so while it is
 * positioned it is aria-hidden and the editor's whole-word mirror is mounted
 * beside it, built and patched with it, for screen readers only: never
 * selected, copied, hit or painted (`.canvas-page-mirror--reading`, Q753 C).
 * The plain-text form already reads whole words and stays readable.
 */

import { useEffect, useRef } from 'react';
import {
  buildMirrorPage,
  buildMirrorTextPage,
  patchMirrorPage,
  syncMirrorPositions,
  type DisplayPage,
} from '@betteroffice/docx/layout/render';
import type { TFunction } from '@betteroffice/docx-i18n';
import { useTranslation } from '../../i18n';

export function CanvasPageMirror({
  page,
  revision,
  zoom = 1,
  defer = false,
  full = true,
  selectable = false,
}: {
  page: DisplayPage;
  /** `displayPageRevision(page)`: a position shift keeps the page object. */
  revision: number;
  zoom?: number;
  /** Off-window pages build at idle time instead of inside the mount flush. */
  defer?: boolean;
  /** Positioned mirror near the viewport; plain accessible text elsewhere. */
  full?: boolean;
  /** Text layer of a read-only viewer: selectable, transparent text. */
  selectable?: boolean;
}) {
  const hostRef = useRef<HTMLDivElement>(null);
  const readingRef = useRef<HTMLDivElement>(null);
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
    const reading = readingRef.current;
    let built = builtForRef.current;
    // A shift moved the page's positions in place: the mirror keeps its nodes.
    if (built?.page === page && built.t === t && built.revision !== revision) {
      const mirror = host.firstElementChild;
      const readingMirror = reading?.firstElementChild;
      if (
        mirror instanceof HTMLElement &&
        syncMirrorPositions(mirror) &&
        (!(readingMirror instanceof HTMLElement) || syncMirrorPositions(readingMirror))
      ) {
        built = { ...built, revision };
        builtForRef.current = built;
      }
    }
    if (
      built?.page === page &&
      built.revision === revision &&
      built.t === t &&
      built.full === full
    ) {
      return;
    }
    // A page crossing the viewport window already holds the other form of its
    // mirror, so the swap waits for idle time instead of a scroll frame.
    const swapOnly =
      built?.page === page && built.revision === revision && built.t === t;
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
      const labels = {
        page: t('a11y.pageLabel', { number: page.pageIndex + 1 }),
        header: t('a11y.headerLabel'),
        footer: t('a11y.footerLabel'),
      };
      mount(host, (full ? buildMirrorPage : buildMirrorTextPage)(page, { labels, clusters: selectable }));
      if (reading) {
        if (full) {
          mount(reading, buildMirrorPage(page, { labels }));
          host.setAttribute('aria-hidden', 'true');
        } else {
          reading.replaceChildren();
          host.removeAttribute('aria-hidden');
        }
      }
      builtForRef.current = { page, revision, t, full };
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
  }, [page, revision, t, defer, full, selectable]);

  // The mirror content is built in page-local px; when the canvas is enlarged
  // for zoom (CSS size = page * zoom), CSS-scale the mirror by the same factor
  // from its top-left origin so its nodes' `getBoundingClientRect` still lands
  // on the painted glyphs. At zoom = 1 this is an identity scale.
  const style = {
    position: 'absolute',
    left: 0,
    top: 0,
    pointerEvents: 'none',
    transform: zoom !== 1 ? `scale(${zoom})` : undefined,
    transformOrigin: '0 0',
  } as const;
  return (
    <>
      {/* First, so the ids both layers carry (notes, note references, header
          cells) resolve to the layer screen readers follow links into. */}
      {selectable ? (
        <div ref={readingRef} className="canvas-page-mirror canvas-page-mirror--reading" style={style} />
      ) : null}
      <div
        ref={hostRef}
        className={
          selectable ? 'canvas-page-mirror canvas-page-mirror--selectable' : 'canvas-page-mirror'
        }
        style={style}
      />
    </>
  );
}

/**
 * Keeps the previous mirror connected until its replacement is ready (clearing
 * in effect cleanup would leave a detached-DOM window on every page update;
 * unmounting already removes the host and its subtree), patching it so a
 * keystroke restyles and lays out only the blocks it changed.
 */
function mount(host: HTMLElement, mirror: HTMLElement): void {
  const current = host.firstElementChild;
  if (!(current instanceof HTMLElement && patchMirrorPage(current, mirror))) {
    host.replaceChildren(mirror);
  }
}

/** Also a caret: a drag starts as one, before its first extension. */
function holdsSelectionEnd(host: HTMLElement): boolean {
  const selection = host.ownerDocument.getSelection();
  if (!selection) return false;
  return host.contains(selection.anchorNode) || host.contains(selection.focusNode);
}
