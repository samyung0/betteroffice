import { describe, expect, it } from 'bun:test';
import { clampScroll, uniformOffsets, visibleRange, visibleCells, zoomedViewport } from './index';
import type { ViewportState } from './index';

describe('visibleRange', () => {
  const offsets = uniformOffsets(20, 80); // 20 tracks, 80px each, total 1600

  it('returns leading tracks at scroll 0', () => {
    expect(visibleRange(offsets, 0, 0, 320)).toEqual({ first: 0, last: 4 });
  });

  it('clips leading tracks when scrolled', () => {
    // scroll 250 lands inside track 3 (240..320); window 250..570 covers 3..7.
    expect(visibleRange(offsets, 0, 250, 320)).toEqual({ first: 3, last: 7 });
  });

  it('excludes frozen tracks and shortens the window by their width', () => {
    // 2 frozen (160px) pinned; body window starts at 160 + scroll.
    const r = visibleRange(offsets, 2, 200, 320);
    expect(r.first).toBeGreaterThanOrEqual(2);
    expect(r.last).toBeGreaterThanOrEqual(r.first);
  });

  it('is empty for a zero-track axis', () => {
    expect(visibleRange([0], 0, 0, 100)).toEqual({ first: 0, last: -1 });
  });
});

describe('clampScroll', () => {
  const colOffsets = uniformOffsets(10, 80); // total 800
  const rowOffsets = uniformOffsets(10, 20); // total 200
  const base: ViewportState = {
    scrollX: 0,
    scrollY: 0,
    width: 400,
    height: 100,
    dpr: 1,
    frozenRows: 0,
    frozenCols: 0,
  };

  it('returns the same object when already in range', () => {
    expect(clampScroll(base, colOffsets, rowOffsets)).toBe(base);
  });

  it('clamps overscroll to the max offset', () => {
    const state = { ...base, scrollX: 10_000, scrollY: 10_000 };
    const clamped = clampScroll(state, colOffsets, rowOffsets);
    expect(clamped.scrollX).toBe(400); // 800 content - 400 viewport
    expect(clamped.scrollY).toBe(100); // 200 content - 100 viewport
  });

  it('clamps negative scroll to zero', () => {
    const clamped = clampScroll({ ...base, scrollX: -50 }, colOffsets, rowOffsets);
    expect(clamped.scrollX).toBe(0);
  });
});

describe('visibleCells', () => {
  it('computes both axes together', () => {
    const state: ViewportState = {
      scrollX: 0,
      scrollY: 0,
      width: 320,
      height: 60,
      dpr: 1,
      frozenRows: 0,
      frozenCols: 0,
    };
    const result = visibleCells(state, uniformOffsets(20, 80), uniformOffsets(20, 20));
    expect(result.cols).toEqual({ first: 0, last: 4 });
    expect(result.rows).toEqual({ first: 0, last: 3 });
  });
});

describe('zoomedViewport', () => {
  const box = { scrollLeft: 300, scrollTop: 120, clientWidth: 800, clientHeight: 600 };

  it('shows a smaller part of the sheet zoomed in and a larger one zoomed out', () => {
    expect(zoomedViewport(box, 1)).toEqual({ x: 300, y: 120, width: 800, height: 600 });
    expect(zoomedViewport(box, 2)).toEqual({ x: 150, y: 60, width: 400, height: 300 });
    expect(zoomedViewport(box, 0.5)).toEqual({ x: 600, y: 240, width: 1600, height: 1200 });
  });

  it('keeps the top-left row and column when the box scrolls with the zoom', () => {
    // 24px rows and 96px columns with two frozen: what sits under the panes at
    // 100% sits there at every preset once the scroll is scaled by the zoom.
    const rows = uniformOffsets(200, 24);
    const cols = uniformOffsets(60, 96);
    const at100 = zoomedViewport(box, 1);
    const first = (view: typeof at100) => ({
      row: visibleRange(rows, 2, view.y, view.height).first,
      col: visibleRange(cols, 2, view.x, view.width).first,
    });
    for (const zoom of [0.5, 0.75, 0.9, 1.25, 1.5, 2]) {
      const scaled = {
        ...box,
        scrollLeft: at100.x * zoom,
        scrollTop: at100.y * zoom,
      };
      expect(first(zoomedViewport(scaled, zoom))).toEqual(first(at100));
    }
  });
});
