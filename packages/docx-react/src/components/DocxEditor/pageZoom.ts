import { type RefObject, useEffect, useLayoutEffect, useRef, useState } from 'react';
import { CANVAS_PAGES_PADDING_PX } from '@betteroffice/docx/layout/render';

/** A zoom level (1 = 100%), or 'fit': the widest page fitted to the pane's width. */
export type DocxZoom = number | 'fit';

/** A pinch zooms within 25–200%. */
export const PINCH_ZOOM_MIN = 0.25;
export const PINCH_ZOOM_MAX = 2;

/**
 * The level that fits a `pageWidth` page (px at 100%) and the page column's
 * side padding into `paneWidth`, at most 100%; null while the pane has no room.
 */
export function fitZoom(paneWidth: number, pageWidth: number): number | null {
  const room = paneWidth - 2 * CANVAS_PAGES_PADDING_PX;
  return room > 0 && pageWidth > 0 ? Math.min(1, room / pageWidth) : null;
}

/** `pane`'s client width, followed through resizes while `enabled`; 0 until measured. */
export function usePaneWidth(pane: RefObject<HTMLElement | null>, enabled: boolean): number {
  const [width, setWidth] = useState(0);
  useLayoutEffect(() => {
    const node = pane.current;
    if (!enabled || !node) return;
    const measure = () => setWidth(node.clientWidth);
    measure();
    const observer = new ResizeObserver(measure);
    observer.observe(node);
    return () => observer.disconnect();
  }, [pane, enabled]);
  return width;
}

function scrollerOf(element: HTMLElement): HTMLElement {
  for (let node = element.parentElement; node; node = node.parentElement) {
    const { overflowX, overflowY } = getComputedStyle(node);
    if (/auto|scroll/.test(overflowX + overflowY)) return node;
  }
  return (document.scrollingElement as HTMLElement | null) ?? document.documentElement;
}

interface PinchAnchor {
  /** The page under the fingers when the pinch began, and the point on it (0–1). */
  page: HTMLElement | null;
  fx: number;
  fy: number;
  /** Where the fingers were when they lifted. */
  x: number;
  y: number;
}

/**
 * A two-finger pinch over `target` zooms the document instead of the browser
 * page (`touch-action: pan-x pan-y` there): `content` previews the scale as a
 * transform about the fingers, and the release commits the level (whole
 * percent, 25–200%) through `onZoom`, with the page point that was under the
 * fingers back under them. The host must apply the level it is given.
 */
export function usePinchZoom({
  target,
  content,
  enabled,
  zoom,
  onZoom,
}: {
  target: RefObject<HTMLElement | null>;
  content: RefObject<HTMLElement | null>;
  enabled: boolean;
  zoom: number;
  onZoom: ((zoom: number) => void) | undefined;
}): void {
  const zoomRef = useRef(zoom);
  zoomRef.current = zoom;
  const onZoomRef = useRef(onZoom);
  onZoomRef.current = onZoom;
  const anchorRef = useRef<PinchAnchor | null>(null);

  useEffect(() => {
    const surface = target.current;
    const scaled = content.current;
    if (!enabled || !onZoomRef.current || !surface || !scaled) return;
    surface.style.touchAction = 'pan-x pan-y';
    let gesture: (PinchAnchor & { spread: number; zoom: number; level: number }) | null = null;
    let frame = 0;
    const spread = (touches: TouchList) =>
      Math.hypot(touches[0].clientX - touches[1].clientX, touches[0].clientY - touches[1].clientY);
    const middle = (touches: TouchList) => ({
      x: (touches[0].clientX + touches[1].clientX) / 2,
      y: (touches[0].clientY + touches[1].clientY) / 2,
    });
    const clearPreview = () => {
      cancelAnimationFrame(frame);
      frame = 0;
      scaled.style.transform = '';
      scaled.style.transformOrigin = '';
    };
    const start = (event: TouchEvent) => {
      if (gesture || event.touches.length !== 2) return;
      const { x, y } = middle(event.touches);
      const box = scaled.getBoundingClientRect();
      let page: HTMLElement | null = null;
      let rect: DOMRect | undefined;
      let distance = Infinity;
      for (const candidate of scaled.querySelectorAll<HTMLElement>('[data-page-index]')) {
        const bounds = candidate.getBoundingClientRect();
        const off = Math.max(bounds.top - y, y - bounds.bottom, 0);
        if (off < distance) {
          page = candidate;
          rect = bounds;
          distance = off;
        }
      }
      gesture = {
        page,
        fx: rect && rect.width ? (x - rect.left) / rect.width : 0,
        fy: rect && rect.height ? (y - rect.top) / rect.height : 0,
        x,
        y,
        spread: spread(event.touches),
        zoom: zoomRef.current,
        level: zoomRef.current,
      };
      scaled.style.transformOrigin = `${x - box.left}px ${y - box.top}px`;
    };
    const move = (event: TouchEvent) => {
      if (!gesture || event.touches.length < 2 || !gesture.spread) return;
      const level = (gesture.zoom * spread(event.touches)) / gesture.spread;
      gesture.level = Math.min(PINCH_ZOOM_MAX, Math.max(PINCH_ZOOM_MIN, level));
      Object.assign(gesture, middle(event.touches));
      if (frame) return;
      frame = requestAnimationFrame(() => {
        frame = 0;
        if (gesture) scaled.style.transform = `scale(${gesture.level / gesture.zoom})`;
      });
    };
    const end = (event: TouchEvent) => {
      if (!gesture || event.touches.length >= 2) return;
      const done = gesture;
      gesture = null;
      const level = Math.round(done.level * 100) / 100;
      // Fingers that did not spread or close (a two-finger tap) leave the zoom.
      if (Math.round(done.zoom * 100) / 100 === level) {
        clearPreview();
        return;
      }
      cancelAnimationFrame(frame);
      frame = 0;
      anchorRef.current = { page: done.page, fx: done.fx, fy: done.fy, x: done.x, y: done.y };
      onZoomRef.current?.(level);
    };
    surface.addEventListener('touchstart', start, { passive: true });
    surface.addEventListener('touchmove', move, { passive: true });
    surface.addEventListener('touchend', end, { passive: true });
    surface.addEventListener('touchcancel', end, { passive: true });
    return () => {
      surface.removeEventListener('touchstart', start);
      surface.removeEventListener('touchmove', move);
      surface.removeEventListener('touchend', end);
      surface.removeEventListener('touchcancel', end);
      surface.style.touchAction = '';
      clearPreview();
    };
  }, [target, content, enabled]);

  // The committed level is laid out: drop the preview in the same frame and
  // scroll the anchored page point back under the fingers.
  useLayoutEffect(() => {
    const anchor = anchorRef.current;
    const scaled = content.current;
    if (!anchor || !scaled) return;
    anchorRef.current = null;
    scaled.style.transform = '';
    scaled.style.transformOrigin = '';
    if (!anchor.page?.isConnected) return;
    const rect = anchor.page.getBoundingClientRect();
    const scroller = scrollerOf(scaled);
    scroller.scrollLeft += rect.left + anchor.fx * rect.width - anchor.x;
    scroller.scrollTop += rect.top + anchor.fy * rect.height - anchor.y;
  }, [content, zoom]);
}
