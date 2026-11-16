import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import type { DisplayList, GlyphRunPrimitive } from '@betteroffice/docx/layout/render';
import type { SlicedPaint } from '@betteroffice/docx/yrs';
import { CanvasPagesView } from './CanvasPagesView';

const { act, cleanup, render } = await import('@testing-library/react');

// A 2D context and Path2D that accept every call: what is drawn shows through
// the outline provider the page's glyph cache asks.
const realm = globalThis as { Path2D?: unknown };
const getContext = HTMLCanvasElement.prototype.getContext;
const context = new Proxy({} as Record<PropertyKey, unknown>, {
  get: (state, key) => (key in state ? state[key] : key === 'measureText' ? () => ({ width: 0 }) : () => {}),
  set: (state, key, value) => {
    state[key] = value;
    return true;
  },
});
beforeAll(() => {
  HTMLCanvasElement.prototype.getContext = (() => context) as unknown as typeof getContext;
  realm.Path2D = class {
    moveTo() {}
    lineTo() {}
    quadraticCurveTo() {}
    bezierCurveTo() {}
    closePath() {}
  };
});
afterEach(() => cleanup());
afterAll(async () => {
  HTMLCanvasElement.prototype.getContext = getContext;
  delete realm.Path2D;
  if (ownsDom) await GlobalRegistrator.unregister();
});

const settle = () => act(() => new Promise((resolve) => setTimeout(resolve, 50)));

/** One page drawing glyph 3000 of the sliced font 7. */
function pages(): DisplayList {
  const run: GlyphRunPrimitive = {
    kind: 'glyphRun',
    fontId: 7,
    size: 16,
    color: '#000000',
    text: '鰯',
    fallbackFont: '16px "MS Gothic"',
    glyphs: [{ id: 3000, x: 10, y: 20, cluster: 0 }],
  };
  return { pages: [{ pageIndex: 0, width: 100, height: 100, primitives: [run] }] } as unknown as DisplayList;
}

test('a page readies its sliced faces before it rasters, drops the outlines that changed, and repaints when late outlines land', async () => {
  const asked: string[] = [];
  const provider = (fontId: number, glyphId: number) => {
    asked.push(`${fontId}:${glyphId}`);
    return JSON.stringify({ upem: 1000, cmds: [{ t: 'M', x: 0, y: 0 }, { t: 'L', x: 1, y: 1 }, { t: 'Z' }] });
  };
  let changed = new Set<number>();
  let gate: Promise<void> | null = null;
  const prepareSlicedPaint = async (): Promise<SlicedPaint> => {
    await gate;
    return { changed, outlineOnlyFontIds: new Set([7]), fontFamilies: new Map() };
  };
  let late: ((fontIds: Set<number>) => void) | null = null;
  const onGlyphOutlines = (listener: (fontIds: Set<number>) => void) => {
    late = listener;
    return () => {
      late = null;
    };
  };
  let release!: () => void;
  gate = new Promise((resolve) => {
    release = resolve;
  });
  const view = (list: DisplayList) => (
    <CanvasPagesView
      displayList={list}
      glyphOutlineProvider={provider}
      prepareSlicedPaint={prepareSlicedPaint}
      onGlyphOutlines={onGlyphOutlines}
    />
  );
  const shown = render(view(pages()));
  await settle();
  // Nothing rasters while the page's outlines are still being readied.
  expect(asked).toEqual([]);
  gate = null;
  release();
  await settle();
  expect(asked).toEqual(['7:3000']);

  // A repaint whose faces did not change draws from the cache.
  shown.rerender(view(pages()));
  await settle();
  expect(asked).toEqual(['7:3000']);
  // Outlines of font 7 that changed are dropped and read again.
  changed = new Set([7]);
  shown.rerender(view(pages()));
  await settle();
  expect(asked).toEqual(['7:3000', '7:3000']);

  // Outlines that land after the page painted repaint it from a fresh cache.
  changed = new Set();
  await act(async () => late!(new Set([7])));
  await settle();
  expect(asked).toEqual(['7:3000', '7:3000', '7:3000']);
});
