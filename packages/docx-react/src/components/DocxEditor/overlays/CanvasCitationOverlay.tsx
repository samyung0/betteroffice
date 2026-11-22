/**
 * A host's citation mark over the canvas pages (`DocxEditorRef.showCitation`).
 *
 * Display only: the paragraph is found by reading the session, and the mark is
 * a portaled overlay, so nothing reaches the document, the undo history, the
 * selection or the presence collaborators see.
 */

import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import type { DisplayListQueries, DisplayListRect } from '@betteroffice/docx/layout/render';
import type { YrsLoc, YrsSession } from '@betteroffice/docx/yrs';

/** What `DocxEditorRef.showCitation` takes. */
export interface DocxCitation {
  /** A literal piece every paragraph `match` could pick holds: the engine's
   * text search finds the candidates with it. */
  probe: string;
  /** Picks the cited paragraph by its text, or none: the host's citation rule. */
  match: <T extends { text: string }>(paragraphs: T[]) => T | null;
  /** The mark's CSS colour. */
  color: string;
}

/** The paragraph a citation picked, by its live identity. */
export interface CitedParagraph {
  story: string;
  paraId: string;
}

/**
 * The paragraph `citation` picks among those holding its probe, in every
 * story, so a quote that also appears in a header or a note is not unique;
 * only a body paragraph (table cells included) is marked.
 */
export function findCitedParagraph(
  session: YrsSession,
  citation: DocxCitation
): CitedParagraph | null {
  const candidates = new Map<string, CitedParagraph & { text: string }>();
  for (const { paraId, story } of session.searchText(citation.probe)) {
    const key = `${story}/${paraId}`;
    if (!candidates.has(key))
      candidates.set(key, { paraId, story, text: session.paragraphUnitText(story, paraId) });
  }
  const match = citation.match([...candidates.values()]);
  return match && (match.story === 'body' || match.story.startsWith('body:'))
    ? { paraId: match.paraId, story: match.story }
    : null;
}

/**
 * The cited paragraph's live display range while it still holds the quote
 * (checked on every change, with the same rule); null once it does not or is
 * gone. The mark never moves to text elsewhere.
 */
export function citedRange(
  session: YrsSession,
  cited: CitedParagraph,
  citation: DocxCitation,
  toDisplay: (loc: YrsLoc) => number | null
): { from: number; to: number } | null {
  try {
    if (!citation.match([{ text: session.paragraphUnitText(cited.story, cited.paraId) }]))
      return null;
    const span = session.locateParagraph(cited.story, cited.paraId);
    const from = toDisplay({ ...cited, offset: 0 });
    const to = toDisplay({ ...cited, offset: span.end - span.start });
    return from == null || to == null || to <= from ? null : { from, to };
  } catch {
    return null;
  }
}

export interface CanvasCitationOverlayProps {
  /** The cited paragraph's display range, or null for no mark. */
  range: { from: number; to: number } | null;
  color: string;
  /** A new key centres the mark once more, as each citation in the viewer does. */
  scrollKey: unknown;
  /** The pages have been presented: the first centring waits for it. */
  presented: boolean;
  /** Portal target — `editorContentRef.current`, sharing the canvas host's top-left. */
  overlayTarget: HTMLElement;
  /** `.canvas-pages` host — live per-page `<canvas>` rects are read from here. */
  canvasHostRef: React.RefObject<HTMLDivElement | null>;
  displayListQueries: DisplayListQueries;
  /** Sidebar open — recompute after its `translateX` transition settles. */
  sidebarOpen: boolean;
  /** Zoom — recompute when the canvas re-rasters larger. */
  zoom: number;
}

interface ProjectedRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** What the reader does to move: after it, a zoom change leaves the scroll alone. */
const READER_MOVES = ['wheel', 'pointerdown', 'keydown', 'touchstart'] as const;

export function CanvasCitationOverlay({
  range,
  color,
  scrollKey,
  presented,
  overlayTarget,
  canvasHostRef,
  displayListQueries,
  sidebarOpen,
  zoom,
}: CanvasCitationOverlayProps) {
  const [rects, setRects] = useState<ProjectedRect[]>([]);
  const firstRef = useRef<HTMLDivElement>(null);
  const recomputeRef = useRef(() => {});
  const from = range?.from ?? 0;
  const to = range?.to ?? 0;
  const marked = range != null;

  // The live range's rects; the observer below calls the latest through the
  // ref, so a key that shifts the paragraph only recomputes.
  useLayoutEffect(() => {
    recomputeRef.current = () => {
      const host = canvasHostRef.current;
      if (!marked || !host) {
        setRects([]);
        return;
      }
      const targetRect = overlayTarget.getBoundingClientRect();
      // The live `<canvas>` rect folds in centering, the sidebar shift and zoom.
      const project = (r: DisplayListRect): ProjectedRect | null => {
        const canvasEl = host.querySelector<HTMLCanvasElement>(
          `canvas[data-page-index="${r.pageIndex}"]`
        );
        const size = displayListQueries.pageSize(r.pageIndex);
        if (!canvasEl || !size) return null;
        const canvasRect = canvasEl.getBoundingClientRect();
        const scaleX = size.width > 0 ? canvasRect.width / size.width : 1;
        const scaleY = size.height > 0 ? canvasRect.height / size.height : 1;
        return {
          left: canvasRect.left - targetRect.left + r.x * scaleX,
          top: canvasRect.top - targetRect.top + r.y * scaleY,
          width: r.width * scaleX,
          height: r.height * scaleY,
        };
      };
      const next: ProjectedRect[] = [];
      for (const r of displayListQueries.rangeRects(from, to)) {
        const p = project(r);
        if (p) next.push(p);
      }
      setRects(next);
    };
    recomputeRef.current();
  }, [marked, from, to, overlayTarget, canvasHostRef, displayListQueries, sidebarOpen, zoom]);

  useEffect(() => {
    const host = canvasHostRef.current;
    if (!marked || !host) return;
    const recompute = () => recomputeRef.current();
    const ro = new ResizeObserver(recompute);
    ro.observe(host);
    ro.observe(overlayTarget);
    window.addEventListener('resize', recompute);
    host.addEventListener('transitionend', recompute);
    return () => {
      ro.disconnect();
      window.removeEventListener('resize', recompute);
      host.removeEventListener('transitionend', recompute);
    };
  }, [marked, overlayTarget, canvasHostRef]);

  // Centred once the pages are presented, once per citation, as the viewer
  // does; a zoom change before the reader moves (a fitted zoom settling)
  // centres it again.
  const scrolledRef = useRef<{ key: unknown; zoom: number; moved: boolean } | null>(null);
  useEffect(() => {
    const doc = overlayTarget.ownerDocument;
    const moved = () => {
      if (scrolledRef.current) scrolledRef.current.moved = true;
    };
    for (const type of READER_MOVES)
      doc.addEventListener(type, moved, { capture: true, passive: true });
    return () => {
      for (const type of READER_MOVES) doc.removeEventListener(type, moved, { capture: true });
    };
  }, [overlayTarget]);
  useLayoutEffect(() => {
    if (!presented || rects.length === 0) return;
    const last = scrolledRef.current;
    if (last && last.key === scrollKey && (last.zoom === zoom || last.moved)) return;
    scrolledRef.current = { key: scrollKey, moved: false, zoom };
    firstRef.current?.scrollIntoView({ block: 'center' });
  }, [presented, rects, scrollKey, zoom]);

  if (rects.length === 0) return null;

  return createPortal(
    <div
      data-testid="canvas-citation"
      style={{
        position: 'absolute',
        top: 0,
        left: 0,
        right: 0,
        bottom: 0,
        pointerEvents: 'none',
        overflow: 'hidden',
        zIndex: 7,
      }}
    >
      {rects.map((r, i) => (
        <div
          key={`citation-${i}-${r.left}-${r.top}`}
          ref={i === 0 ? firstRef : undefined}
          data-citation-highlight="true"
          aria-hidden="true"
          style={{
            position: 'absolute',
            left: r.left,
            top: r.top,
            width: r.width,
            height: r.height,
            background: color,
            pointerEvents: 'none',
          }}
        />
      ))}
    </div>,
    overlayTarget
  );
}
