/**
 * A double click on a header or footer opens that band, on any page (upstream
 * #1163), also from another open band (#1165). The press must not move the
 * body caret: scrolling the caret into view would move the page before the
 * click, which hit-tests at click time.
 */

import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, test } from 'bun:test';
import type { DisplayListQueries, DisplayListRegionHit } from '@betteroffice/docx/layout/render';
import type { YrsSession } from '@betteroffice/docx/yrs';
import { usePagesPointer, type UsePagesPointerOptions } from './usePagesPointer';
import type { PartEdit } from '../partEdit';
import type { YrsInputRef } from '../YrsInput';
import type { YrsPositionProjection } from '../internals/yrsPositionProjection';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');

const PAGE = { width: 800, height: 1000 };
const HEADER_BOTTOM = 80;
const FOOTER_TOP = 920;
/** How far the content moves under the pointer once a press placed the body caret (its scroll). */
const SCROLL = 300;

let host: HTMLDivElement;

beforeEach(() => {
  host = document.createElement('div');
  host.className = 'canvas-pages';
  const canvas = document.createElement('canvas');
  canvas.dataset.pageIndex = '0';
  canvas.getBoundingClientRect = () =>
    ({ left: 0, top: 0, right: PAGE.width, bottom: PAGE.height, ...PAGE }) as DOMRect;
  host.append(canvas);
  document.body.append(host);
});

afterEach(() => {
  cleanup();
  host.remove();
});

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function setup(partEdit: PartEdit | null = null) {
  const placed: number[] = [];
  const words: number[] = [];
  const opened: Array<[string, number | undefined]> = [];
  const bodyClicks: number[] = [];
  const queries = {
    displayList: { pages: [{ pageIndex: 0, ...PAGE, primitives: [] }] },
    pageCount: () => 1,
    pageSize: () => PAGE,
    hitTestRegions: (_pageIndex: number, x: number, y: number): DisplayListRegionHit => {
      const at = y + (placed.length > 0 ? SCROLL : 0);
      if (at < HEADER_BOTTOM) return { region: 'header', rId: 'rId7', pos: 1, target: 'text' };
      if (at >= FOOTER_TOP) return { region: 'footer', rId: 'rId8', pos: 1, target: 'text' };
      return { region: 'body', pos: Math.floor(x / 10), target: 'text' };
    },
    imageAtPoint: () => null,
  } as unknown as DisplayListQueries;
  const input = {
    focus: () => {},
    beginPointerSelection: () => true,
    setSelectionFromDisplay: (anchor: number) => placed.push(anchor),
    selectWordAtDisplay: (position: number) => words.push(position),
    displaySelection: () => null,
  } as unknown as YrsInputRef;
  const projection = {
    size: 100,
    targetAt: (position: number) =>
      position >= 0 && position < 100 ? { story: 'body', displayPosition: position } : null,
    tableAtPosition: () => null,
    cellPosition: () => null,
  } as unknown as YrsPositionProjection;
  const options: UsePagesPointerOptions = {
    pagesContainerRef: { current: null },
    yrsInputRef: { current: input },
    yrsSession: { cellSelection: () => null } as unknown as YrsSession,
    yrsRootStory: 'body',
    getYrsPositionProjection: () => projection,
    applyYrsCommand: () => false,
    syncYrsInputState: () => false,
    readOnly: false,
    displayListQueries: queries,
    canvasHostRef: { current: host },
    onHeaderFooterDoubleClick: (region, page) => opened.push([region, page]),
    partEdit,
    onBodyClick: () => bodyClicks.push(1),
    setSelectionRects: () => {},
    setCaretPosition: () => {},
    setIsFocused: () => {},
    scrollToPositionImpl: () => {},
  };
  renderHook(() => usePagesPointer(options));
  return { placed, words, opened, bodyClicks };
}

function doubleClick(x: number, y: number): void {
  const canvas = host.firstElementChild as HTMLCanvasElement;
  for (const detail of [1, 2]) {
    for (const type of ['mousedown', 'mouseup', 'click']) {
      const init = { bubbles: true, cancelable: true, clientX: x, clientY: y, button: 0, detail };
      act(() => {
        canvas.dispatchEvent(new MouseEvent(type, init));
      });
    }
  }
}

test('a double click on a header opens it and leaves the body caret alone', () => {
  const { placed, words, opened } = setup();
  doubleClick(400, 40);
  expect(opened).toEqual([['header', 1]]);
  expect(placed).toEqual([]);
  expect(words).toEqual([]);
});

test('a double click on a footer opens it and leaves the body caret alone', () => {
  const { placed, words, opened } = setup();
  doubleClick(400, 960);
  expect(opened).toEqual([['footer', 1]]);
  expect(placed).toEqual([]);
  expect(words).toEqual([]);
});

test('a double click on a header with the footer open opens the header without closing to the body', () => {
  const { placed, opened, bodyClicks } = setup({ kind: 'footer', rId: 'rId8' });
  doubleClick(400, 40);
  expect(opened.at(-1)).toEqual(['header', 1]);
  expect(bodyClicks).toEqual([]);
  expect(placed).toEqual([]);
});
