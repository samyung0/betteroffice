/**
 * A host's citation mark over the canvas pages (`DocxEditorRef.showCitation`).
 *
 * Display only: the paragraph is found by reading the session, and the mark is
 * a portaled overlay, so nothing reaches the document, the undo history, the
 * selection or the presence collaborators see.
 */

import { useLayoutEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import type { DisplayListQueries, DisplayListRect } from '@betteroffice/docx/layout/render';
import type { YrsLoc, YrsSession } from '@betteroffice/docx/yrs';

/** What `DocxEditorRef.showCitation` takes. */
export interface DocxCitation {
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
 * The paragraph `citation` picks. Every story is a candidate, so a quote that
 * also appears in a header or a note is not unique; only a body paragraph
 * (table cells included) is marked.
 */
export function findCitedParagraph(
  session: YrsSession,
  citation: DocxCitation
): CitedParagraph | null {
  const paragraphs = session
    .storyIds()
    .flatMap((story) =>
      session.paragraphs(story).map(({ paraId, text }) => ({ paraId, story, text }))
    );
  const match = citation.match(paragraphs);
  return match && (match.story === 'body' || match.story.startsWith('body:'))
    ? { paraId: match.paraId, story: match.story }
    : null;
}

/** The cited paragraph's live display range; null once it is gone. */
export function citedRange(
  session: YrsSession,
  cited: CitedParagraph,
  toDisplay: (loc: YrsLoc) => number | null
): { from: number; to: number } | null {
  try {
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

export function CanvasCitationOverlay({
  range,
  color,
  scrollKey,
  overlayTarget,
  canvasHostRef,
  displayListQueries,
  sidebarOpen,
  zoom,
}: CanvasCitationOverlayProps) {
  const [rects, setRects] = useState<ProjectedRect[]>([]);
  const firstRef = useRef<HTMLDivElement>(null);
  const scrolledRef = useRef<unknown>(null);
  const from = range?.from ?? 0;
  const to = range?.to ?? 0;
  const marked = range != null;

  useLayoutEffect(() => {
    const host = canvasHostRef.current;
    if (!marked || !host) {
      setRects([]);
      return;
    }
    const recompute = () => {
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
    recompute();
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
  }, [marked, from, to, overlayTarget, canvasHostRef, displayListQueries, sidebarOpen, zoom]);

  // Centred once the pages hold it, once per citation, as the viewer does.
  useLayoutEffect(() => {
    if (rects.length === 0 || scrolledRef.current === scrollKey) return;
    scrolledRef.current = scrollKey;
    firstRef.current?.scrollIntoView({ block: 'center' });
  }, [rects, scrollKey]);

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
