import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, mock, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { useRef, useState } from 'react';
import { applyFrameDeltaOwned, decodeFrameDelta, type DisplayList } from '@betteroffice/docx/layout/render';
import { createEditSession, preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { DocxDisplayListViewer } from '../DocxViewer';
import { type DocxZoom, fitZoom, usePinchZoom } from './pageZoom';

const { act, cleanup, render } = await import('@testing-library/react');

// Observers the component keeps connected, fired by hand.
const observers = new Set<{ callback: () => void }>();
class FakeResizeObserver {
  constructor(readonly callback: () => void) {}
  observe() {
    observers.add(this);
  }
  disconnect() {
    observers.delete(this);
  }
  unobserve() {}
}
const RealResizeObserver = globalThis.ResizeObserver;
let paneWidth = 0;
const clientWidth = Object.getOwnPropertyDescriptor(window.HTMLElement.prototype, 'clientWidth');

let displayList: DisplayList;
beforeAll(async () => {
  await preloadEditWasm(
    new Uint8Array(readFileSync(resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')))
  );
  const native = createEditSession(9311);
  native.load_json(JSON.stringify([{ storyId: 'body', paragraphs: [{ text: 'One page' }] }]));
  const request = JSON.stringify({
    bodyStory: 'body',
    regions: { sections: [{ sectionId: 'main', properties: {} }] },
    measurement: { defaults: { fontSize: 11, fontFamily: 'Calibri' } },
    renderEnv: {},
  });
  displayList = applyFrameDeltaOwned(null, decodeFrameDelta(native.relayout_frame(request, '{}', 0))).displayList;
  native.free();
  globalThis.ResizeObserver = FakeResizeObserver as unknown as typeof ResizeObserver;
  Object.defineProperty(window.HTMLElement.prototype, 'clientWidth', {
    configurable: true,
    get(this: HTMLElement) {
      return this.classList.contains('canvas-pages') ? paneWidth : 0;
    },
  });
});
afterEach(() => cleanup());
afterAll(async () => {
  globalThis.ResizeObserver = RealResizeObserver;
  if (clientWidth) Object.defineProperty(window.HTMLElement.prototype, 'clientWidth', clientWidth);
  if (ownsDom) await GlobalRegistrator.unregister();
});

const nextFrame = () => act(() => new Promise((done) => requestAnimationFrame(() => done(undefined))));

/** A touch event with these fingers (clientX, clientY). */
function touch(type: string, fingers: [number, number][]) {
  const event = new Event(type, { bubbles: true });
  Object.defineProperty(event, 'touches', { value: fingers.map(([clientX, clientY]) => ({ clientX, clientY })) });
  return event;
}

/** Two fingers `from` px apart, spread to `to` px, lifted. */
async function pinch(target: Element, from: number, to: number) {
  await act(async () => {
    target.dispatchEvent(touch('touchstart', [[100, 100], [100 + from, 100]]));
    target.dispatchEvent(touch('touchmove', [[100, 100], [100 + to, 100]]));
  });
  await nextFrame();
  await act(async () => {
    target.dispatchEvent(touch('touchend', [[100, 100]]));
  });
}

test('fit zoom fits the page and the column padding into the pane, at most 100%', () => {
  // 24 px of column padding each side.
  expect(fitZoom(348, 816)).toBeCloseTo(300 / 816, 6);
  expect(fitZoom(390, 816)).toBeCloseTo(342 / 816, 6);
  expect(fitZoom(348, 794)).toBeCloseTo(300 / 794, 6);
  expect(fitZoom(864, 816)).toBe(1);
  expect(fitZoom(1280, 816)).toBe(1);
  expect(fitZoom(48, 816)).toBeNull();
  expect(fitZoom(0, 816)).toBeNull();
  expect(fitZoom(348, 0)).toBeNull();
});

test('a fitted viewer re-fits on resize until a pinch picks a level, then never', async () => {
  const levels: number[] = [];
  function Host() {
    const [zoom, setZoom] = useState<DocxZoom>('fit');
    return (
      <DocxDisplayListViewer
        displayList={displayList}
        zoom={zoom}
        pinchZoom
        onZoomChange={(level) => {
          levels.push(level);
          setZoom(level);
        }}
      />
    );
  }
  const page = displayList.pages[0]!;
  paneWidth = 348;
  const { container } = render(<Host />);
  const canvas = () => container.querySelector<HTMLCanvasElement>('canvas[data-page-index="0"]')!;
  const width = () => Number.parseFloat(canvas().style.width);
  expect(width()).toBeCloseTo(page.width * fitZoom(348, page.width)!, 3);
  const host = container.querySelector('.canvas-pages')!;
  expect((host as HTMLElement).style.touchAction).toBe('pan-x pan-y');

  paneWidth = 390;
  await act(async () => {
    for (const observer of observers) observer.callback();
  });
  const fitted = fitZoom(390, page.width)!;
  expect(width()).toBeCloseTo(page.width * fitted, 3);
  expect(observers.size).toBe(1);

  await pinch(host, 100, 150);
  expect(levels).toEqual([Math.round(fitted * 1.5 * 100) / 100]);
  expect(width()).toBeCloseTo(page.width * levels[0]!, 3);
  // The preview is gone once the level is laid out.
  expect((host as HTMLElement).style.transform).toBe('');
  // Nothing follows the pane any more.
  expect(observers.size).toBe(0);
  paneWidth = 700;
  await act(async () => {
    for (const observer of observers) observer.callback();
  });
  expect(width()).toBeCloseTo(page.width * levels[0]!, 3);
});

test('a pinch previews its scale, stays within 25–200%, and a two-finger tap keeps the level', async () => {
  const onZoom = mock((_level: number) => {});
  let preview = '';
  function Pinchable({ zoom, enabled }: { zoom: number; enabled: boolean }) {
    const ref = useRef<HTMLDivElement>(null);
    usePinchZoom({ target: ref, content: ref, enabled, zoom, onZoom });
    return <div ref={ref} />;
  }
  const { container, rerender } = render(<Pinchable zoom={1} enabled />);
  const surface = container.firstElementChild as HTMLElement;
  expect(surface.style.touchAction).toBe('pan-x pan-y');

  await act(async () => {
    surface.dispatchEvent(touch('touchstart', [[0, 0], [100, 0]]));
    surface.dispatchEvent(touch('touchmove', [[0, 0], [130, 0]]));
  });
  await nextFrame();
  preview = surface.style.transform;
  expect(preview).toBe('scale(1.3)');
  await act(async () => {
    surface.dispatchEvent(touch('touchmove', [[0, 0], [1000, 0]]));
    surface.dispatchEvent(touch('touchend', [[0, 0]]));
  });
  expect(onZoom.mock.calls.map(([level]) => level)).toEqual([2]);
  rerender(<Pinchable zoom={2} enabled />);
  expect(surface.style.transform).toBe('');

  await pinch(surface, 100, 1);
  expect(onZoom.mock.calls.at(-1)?.[0]).toBe(0.25);
  rerender(<Pinchable zoom={0.25} enabled />);

  await pinch(surface, 100, 100);
  expect(onZoom).toHaveBeenCalledTimes(2);
  expect(surface.style.transform).toBe('');

  rerender(<Pinchable zoom={0.25} enabled={false} />);
  expect(surface.style.touchAction).toBe('');
  await pinch(surface, 100, 200);
  expect(onZoom).toHaveBeenCalledTimes(2);
});

test('the page point between the fingers stays under them after the level commits', async () => {
  // A page 100 px below the scroller's top, at 24 px from its left, sized by zoom.
  let scrollTop = 0;
  let scrollLeft = 0;
  function Scroller({ zoom, onZoom }: { zoom: number; onZoom: (level: number) => void }) {
    const ref = useRef<HTMLDivElement>(null);
    usePinchZoom({ target: ref, content: ref, enabled: true, zoom, onZoom });
    return (
      <div style={{ overflowX: 'auto', overflowY: 'auto' }} data-scroller="">
        <div ref={ref}>
          <div
            data-page-index="0"
            ref={(node) => {
              if (node)
                node.getBoundingClientRect = () =>
                  ({ top: 100 - scrollTop, left: 24 - scrollLeft, width: 816 * zoom, height: 1056 * zoom, bottom: 100 - scrollTop + 1056 * zoom }) as DOMRect;
            }}
          />
        </div>
      </div>
    );
  }
  let level = 0.5;
  const { container, rerender } = render(<Scroller zoom={level} onZoom={(next) => (level = next)} />);
  const scroller = container.querySelector<HTMLElement>('[data-scroller]')!;
  Object.defineProperty(scroller, 'scrollTop', { get: () => scrollTop, set: (value: number) => (scrollTop = value) });
  Object.defineProperty(scroller, 'scrollLeft', { get: () => scrollLeft, set: (value: number) => (scrollLeft = value) });
  const surface = scroller.firstElementChild!;
  // Fingers around (150, 300): page point ((150-24)/408, (300-100)/528) at 50%.
  await act(async () => {
    surface.dispatchEvent(touch('touchstart', [[100, 300], [200, 300]]));
    surface.dispatchEvent(touch('touchmove', [[50, 300], [250, 300]]));
    surface.dispatchEvent(touch('touchend', [[50, 300]]));
  });
  expect(level).toBe(1);
  rerender(<Scroller zoom={level} onZoom={(next) => (level = next)} />);
  // At 100% that point is at (24 + 252, 100 + 400): scrolled back to (150, 300).
  expect([scrollLeft, scrollTop]).toEqual([126, 200]);
});
