/**
 * A press before the page geometry is ready, as right after opening, takes the
 * caret once the geometry arrives and typing lands there (Google Docs is the
 * reference). A press that places no caret leaves typing where the caret was.
 * The real input and session, with a fake geometry.
 */

import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, beforeEach, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRef, type RefObject } from 'react';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import {
  createYrsInputPositionMap,
  createYrsSession,
  displayPositionToYrsLoc,
  yrsLocToDisplayPosition,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { HELD_PRESS_MS, usePagesPointer } from './usePagesPointer';
import type { YrsPositionProjection } from '../internals/yrsPositionProjection';
import { YrsInput, type YrsInputRef } from '../YrsInput';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render, renderHook } = await import('@testing-library/react');

const PAGE = { width: 800, height: 1000 };
const IN_BODY = 200;
const IN_HEADER = 40;
const IN_NOTE_AREA = 950;

let host: HTMLDivElement;
const sessions: YrsSession[] = [];

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(
        resolve(import.meta.dir, '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
      )
    )
  )
);

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
  for (const session of sessions.splice(0)) session.destroy();
});

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

/** The body's position is x / 10: x = 30 is the caret after "Se" in "Seed". */
function geometry(): DisplayListQueries {
  return {
    displayList: { pages: [{ pageIndex: 0, ...PAGE, primitives: [] }] },
    pageCount: () => 1,
    pageSize: () => PAGE,
    hitTestRegions: (_pageIndex: number, x: number, y: number) => {
      if (y < 80) return { region: 'header', rId: 'rId7', pos: 1, target: 'text' };
      if (y >= 900) return { region: 'footnote', noteId: undefined, pos: null, target: 'none' };
      return { region: 'body', pos: Math.floor(x / 10), target: 'text' };
    },
    imageAtPoint: () => null,
  } as unknown as DisplayListQueries;
}

async function setup(queries: DisplayListQueries | null = null) {
  const session = await createYrsSession();
  sessions.push(session);
  const { paraId } = session.createStory('body', 'Seed');
  session.setSelection({ story: 'body', paraId, offset: 4 });
  const map = () => createYrsInputPositionMap('body', session.paragraphSpans('body'));
  const input = createRef<YrsInputRef>();
  const pending: boolean[] = [];
  const view = render(
    <YrsInput
      ref={input}
      enabled
      readOnly={false}
      session={session}
      inputPositionMap={map}
      displayPositionToLoc={(position) => displayPositionToYrsLoc(map(), position)}
      locToDisplayPosition={(loc) => yrsLocToDisplayPosition(map(), loc)}
      onStateChange={() => {}}
      onDirectInput={() => {}}
      onPendingChange={(value) => pending.push(value)}
      paragraphStyleValues={() => ({})}
    />
  );
  const projection = {
    size: 6,
    targetAt: (position: number) => ({ story: 'body', displayPosition: position }),
    tableAtPosition: () => null,
    cellPosition: () => null,
  } as unknown as YrsPositionProjection;
  const pointer = renderHook(
    (props: { queries: DisplayListQueries | null; readOnly: boolean }) =>
      usePagesPointer({
        pagesContainerRef: { current: null },
        yrsInputRef: input as RefObject<YrsInputRef | null>,
        yrsSession: session,
        yrsRootStory: 'body',
        getYrsPositionProjection: () => projection,
        applyYrsCommand: () => false,
        syncYrsInputState: () => false,
        readOnly: props.readOnly,
        displayListQueries: props.queries,
        canvasHostRef: { current: host },
        setSelectionRects: () => {},
        setCaretPosition: () => {},
        setIsFocused: () => {},
        scrollToPositionImpl: () => {},
      }),
    { initialProps: { queries, readOnly: false } }
  );
  const flush = () =>
    act(async () => {
      await input.current!.flushPendingInput();
    });
  return {
    arrive: () => pointer.rerender({ queries: geometry(), readOnly: false }),
    flush,
    goReadOnly: () => pointer.rerender({ queries: null, readOnly: true }),
    input,
    pending,
    session,
    text: () => session.paragraphs('body')[0].text,
    textarea: view.getByTestId('yrs-input') as HTMLTextAreaElement,
  };
}

function mouse(type: string, x: number, y: number, target: EventTarget = host.firstElementChild!) {
  act(() => {
    target.dispatchEvent(
      new MouseEvent(type, {
        bubbles: true,
        button: 0,
        cancelable: true,
        clientX: x,
        clientY: y,
        detail: 1,
      })
    );
  });
}

const settle = () =>
  act(async () => {
    await new Promise((done) => setTimeout(done, 0));
  });

test('a press before the geometry takes the caret when it arrives, and typing lands there', async () => {
  const { arrive, flush, input, pending, text, textarea } = await setup();
  mouse('mousedown', 30, IN_BODY);
  // Nothing is typed yet, and the input takes what comes.
  expect(pending).toEqual([]);
  expect(textarea.readOnly).toBe(false);
  expect(textarea.dataset.pointerPlacement).toBe('ready');
  act(() => input.current!.insertText('X'));
  await settle();
  expect(text()).toBe('Seed');
  arrive();
  await flush();
  expect(text()).toBe('SeXed');
});

test('the last press before the geometry takes the caret', async () => {
  const { arrive, flush, input, text } = await setup();
  mouse('mousedown', 30, IN_BODY);
  mouse('mousedown', 20, IN_BODY);
  arrive();
  act(() => input.current!.insertText('X'));
  await flush();
  expect(text()).toBe('SXeed');
});

test('a press placed when the geometry arrives starts no drag', async () => {
  const { arrive, session } = await setup();
  mouse('mousedown', 30, IN_BODY);
  arrive();
  await settle();
  mouse('mousemove', 50, IN_BODY, window);
  await act(async () => {
    await new Promise((done) => requestAnimationFrame(done));
  });
  const selection = session.selection();
  expect(selection?.head.offset).toBe(2);
  expect(selection?.anchor.offset).toBe(2);
});

test('a flush waits for a held press, and read-only drops the press', async () => {
  const { goReadOnly, input, session } = await setup();
  mouse('mousedown', 30, IN_BODY);
  let flushed = false;
  const flush = input.current!.flushPendingInput().then(() => {
    flushed = true;
  });
  await settle();
  expect(flushed).toBe(false);
  goReadOnly();
  await act(async () => {
    await flush;
  });
  expect(flushed).toBe(true);
  expect(session.selection()?.head.offset).toBe(4);
});

test('a press whose geometry never comes gives the input back', async () => {
  const realSetTimeout = globalThis.setTimeout;
  let expire = () => {};
  globalThis.setTimeout = ((run: () => void, ms?: number) => {
    if (ms !== HELD_PRESS_MS) return realSetTimeout(run, ms);
    expire = run;
    return 0;
  }) as typeof setTimeout;
  try {
    const { flush, input, text } = await setup();
    mouse('mousedown', 30, IN_BODY);
    act(() => input.current!.insertText('X'));
    await settle();
    expect(text()).toBe('Seed');
    act(() => expire());
    await flush();
    expect(text()).toBe('SeedX');
  } finally {
    globalThis.setTimeout = realSetTimeout;
  }
});

for (const [where, y] of [
  ['a header band', IN_HEADER],
  ['a note area no note claims', IN_NOTE_AREA],
] as const) {
  test(`a press on ${where} leaves typing where the caret was`, async () => {
    const { flush, input, text, textarea } = await setup(geometry());
    mouse('mousedown', 30, y);
    expect(textarea.dataset.pointerPlacement).toBe('ready');
    expect(textarea.readOnly).toBe(false);
    act(() => input.current!.insertText('X'));
    await flush();
    expect(text()).toBe('SeedX');
  });
}
