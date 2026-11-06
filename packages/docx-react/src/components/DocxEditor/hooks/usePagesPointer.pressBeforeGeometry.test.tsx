/**
 * A press that cannot be placed yet, before the page geometry is ready or while
 * earlier input is still landing, waits in the input's queue: what came before
 * it lands at the old caret, what comes after lands where it was, and typing is
 * never dropped (Google Docs is the reference). The real input and session
 * through `applyResidentInput`, with a fake geometry.
 */

import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, beforeEach, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { Component, createRef, type ReactNode, type RefObject, useState } from 'react';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import {
  createYrsInputPositionMap,
  createYrsSession,
  displayPositionToYrsLoc,
  yrsLocToDisplayPosition,
  type YrsLoc,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { HELD_PRESS_MS, usePagesPointer } from './usePagesPointer';
import { useYrsCoreSession } from './useYrsCoreSession';
import type { YrsPositionProjection } from '../internals/yrsPositionProjection';
import { partEditStory, type PartEdit } from '../partEdit';
import { YrsInput, type YrsInputProps, type YrsInputRef } from '../YrsInput';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');

const PAGE = { width: 800, height: 1000 };
const IN_BODY = 200;
const IN_HEADER = 40;
const IN_ENDNOTE = 870;
const IN_FOOTNOTE = 950;
/** A note area on this side of the page belongs to no note. */
const UNCLAIMED_X = 500;
/** x past this hits a position the projection has no target for. */
const NO_TARGET_X = 700;
/** x here hits a position the projection targets and the story does not hold. */
const NO_LOC_X = 650;

let host: HTMLDivElement;
let pageRect: { left: number; top: number; width: number; height: number };
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
  pageRect = { left: 0, top: 0, ...PAGE };
  host = document.createElement('div');
  host.className = 'canvas-pages';
  const canvas = document.createElement('canvas');
  canvas.dataset.pageIndex = '0';
  canvas.getBoundingClientRect = () =>
    ({
      ...pageRect,
      bottom: pageRect.top + pageRect.height,
      right: pageRect.left + pageRect.width,
    }) as DOMRect;
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

/** Display position x / 10 in every region: x = 30 is the caret after two characters. */
function geometry(): DisplayListQueries {
  return {
    displayList: { pages: [{ pageIndex: 0, ...PAGE, primitives: [] }] },
    pageCount: () => 1,
    pageSize: () => PAGE,
    hitTestRegions: (_pageIndex: number, x: number, y: number) => {
      const pos = Math.floor(x / 10);
      if (y < 80) return { region: 'header', rId: 'rId7', pos, target: 'text' };
      if (y >= 900 && x >= UNCLAIMED_X) {
        return { region: 'footnote', noteId: undefined, pos: null, target: 'none' };
      }
      if (y >= 900) return { region: 'footnote', noteId: 2, pos, target: 'text' };
      if (y >= 850) return { region: 'endnote', noteId: 3, pos, target: 'text' };
      return { region: 'body', pos, target: 'text' };
    },
    imageAtPoint: () => null,
  } as unknown as DisplayListQueries;
}

class Boundary extends Component<{ children: ReactNode }, { failed: boolean }> {
  state = { failed: false };
  static getDerivedStateFromError() {
    return { failed: true };
  }
  render() {
    return this.state.failed ? null : this.props.children;
  }
}

type PointerProps = {
  input: RefObject<YrsInputRef | null>;
  partEdit: PartEdit | null;
  queries: DisplayListQueries | null;
  readOnly: boolean;
  session: YrsSession;
  setPartEdit: (part: PartEdit | null) => void;
  throwOnPlace: boolean;
};

function Pointer(props: PointerProps) {
  const { input, partEdit, queries, readOnly, session, setPartEdit, throwOnPlace } = props;
  usePagesPointer({
    pagesContainerRef: { current: null },
    yrsInputRef: input,
    yrsSession: session,
    yrsRootStory: partEditStory(partEdit),
    getYrsPositionProjection: (rootStory) => {
      if (throwOnPlace && queries) throw new Error('no projection');
      return {
        size: 100,
        targetAt: (position: number) =>
          position < NO_TARGET_X / 10 ? { story: rootStory, displayPosition: position } : null,
        tableAtPosition: () => null,
        cellPosition: () => null,
      } as unknown as YrsPositionProjection;
    },
    applyYrsCommand: () => false,
    syncYrsInputState: () => false,
    readOnly,
    partEdit,
    displayListQueries: queries,
    canvasHostRef: { current: host },
    onBodyClick: () => setPartEdit(null),
    onHeaderFooterDoubleClick: (region) => setPartEdit({ kind: region, rId: 'rId7' }),
    onNoteClick: setPartEdit,
    setSelectionRects: () => {},
    setCaretPosition: () => {},
    setIsFocused: () => {},
    scrollToPositionImpl: () => {},
  });
  return null;
}

type EditorProps = {
  applyResidentInput: YrsInputProps['applyResidentInput'];
  initialPart: PartEdit | null;
  input: RefObject<YrsInputRef | null>;
  onPendingChange?: YrsInputProps['onPendingChange'];
  queries: DisplayListQueries | null;
  readOnly: boolean;
  session: YrsSession;
  throwOnPlace?: boolean;
};

/** The input and the pointer as PagedEditor wires them: the root story follows the open part. */
function Editor(props: EditorProps) {
  const { applyResidentInput, input, onPendingChange, queries, readOnly, session } = props;
  const [partEdit, setPartEdit] = useState(props.initialPart);
  const root = partEditStory(partEdit);
  const map = (story = root) => createYrsInputPositionMap(story, session.paragraphSpans(story));
  return (
    <>
      <YrsInput
        ref={input}
        enabled
        readOnly={readOnly}
        session={session}
        story={root}
        inputPositionMap={map}
        displayPositionToLoc={(position, story) => displayPositionToYrsLoc(map(story), position)}
        // Bound to the render's root story, as PagedEditor's is.
        locToDisplayPosition={(loc: YrsLoc) =>
          loc.story === root ? yrsLocToDisplayPosition(map(loc.story), loc) : null
        }
        onStateChange={() => {}}
        onDirectInput={() => {}}
        onPendingChange={onPendingChange}
        applyResidentInput={applyResidentInput}
        paragraphStyleValues={() => ({})}
      />
      <Boundary>
        <Pointer
          input={input}
          partEdit={partEdit}
          queries={queries}
          readOnly={readOnly}
          session={session}
          setPartEdit={setPartEdit}
          throwOnPlace={props.throwOnPlace ?? false}
        />
      </Boundary>
    </>
  );
}

/** Resident input that waits for `open()` when `gated`, then falls to the session. */
function residentInput(gated: boolean) {
  const calls: string[] = [];
  let open = () => {};
  const gate = gated
    ? new Promise<void>((done) => {
        open = done;
      })
    : Promise.resolve();
  const apply: YrsInputProps['applyResidentInput'] = async (text) => {
    calls.push(text);
    await gate;
    return null;
  };
  return { apply, calls, open };
}

const nextFrame = () =>
  act(async () => {
    await new Promise((done) => requestAnimationFrame(done));
  });

async function setup(
  options: {
    gated?: boolean;
    part?: PartEdit | null;
    queries?: DisplayListQueries | null;
    secondParagraph?: boolean;
    throwOnPlace?: boolean;
  } = {}
) {
  const session = await createYrsSession();
  sessions.push(session);
  const { paraId } = session.createStory('body', 'Seed');
  if (options.secondParagraph) session.splitParagraph({ story: 'body', paraId, offset: 4 });
  session.createStory('fn:2', 'Note');
  session.createStory('en:3', 'Endn');
  session.createStory('hf:rId7', 'Head');
  session.createStory('hf:rId8', 'Foot');
  session.setSelection({ story: 'body', paraId, offset: 4 });
  const input = createRef<YrsInputRef>() as RefObject<YrsInputRef | null>;
  const pending: boolean[] = [];
  const resident = residentInput(options.gated ?? false);
  const editor = (queries: DisplayListQueries | null, readOnly = false, throwOnPlace = false) => (
    <Editor
      applyResidentInput={resident.apply}
      initialPart={options.part ?? null}
      input={input}
      onPendingChange={(value) => pending.push(value)}
      queries={queries}
      readOnly={readOnly}
      session={session}
      throwOnPlace={throwOnPlace}
    />
  );
  const view = render(editor(options.queries ?? null, false, options.throwOnPlace));
  // A fresh session takes the focus on its next frame.
  await nextFrame();
  return {
    arrive: () => view.rerender(editor(geometry(), false, options.throwOnPlace)),
    flush: () =>
      act(async () => {
        await input.current!.flushPendingInput();
      }),
    goReadOnly: () => view.rerender(editor(null, true)),
    input,
    openResident: () => act(() => resident.open()),
    pending,
    resident,
    session,
    text: (story = 'body') => session.paragraphs(story)[0].text,
    textarea: view.getByTestId('yrs-input') as HTMLTextAreaElement,
    view,
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

/** Whether a promise settles soon: a press must not make a flush wait. */
async function settlesSoon(promise: Promise<unknown>): Promise<boolean> {
  let settled = false;
  void promise.finally(() => {
    settled = true;
  });
  await act(async () => {
    await new Promise((done) => setTimeout(done, 30));
  });
  return settled;
}

test('a press before the geometry takes the caret when it arrives, and typing lands there', async () => {
  const { arrive, flush, input, pending, resident, text, textarea } = await setup();
  mouse('mousedown', 30, IN_BODY);
  // Nothing is typed yet, and the input takes what comes.
  expect(pending).toEqual([]);
  expect(textarea.readOnly).toBe(false);
  expect(textarea.dataset.pointerPlacement).toBe('pending');
  act(() => input.current!.insertText('X'));
  await settle();
  expect(text()).toBe('Seed');
  arrive();
  await flush();
  expect(text()).toBe('SeXed');
  expect(resident.calls).toEqual(['X']);
  expect(textarea.dataset.pointerPlacement).toBe('ready');
});

test('a held press focuses the input at once, so a drag off the pages still types', async () => {
  const { textarea } = await setup();
  mouse('mousedown', 30, IN_BODY);
  expect(document.activeElement === textarea).toBe(true);
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
  await nextFrame();
  const selection = session.selection();
  expect(selection?.head.offset).toBe(2);
  expect(selection?.anchor.offset).toBe(2);
});

test('a press keeps its place on the page when scrolling or zoom moves the page', async () => {
  const { arrive, flush, input, text } = await setup();
  mouse('mousedown', 30, IN_BODY);
  pageRect = { height: 2000, left: 100, top: -300, width: 1600 };
  act(() => input.current!.insertText('X'));
  arrive();
  await flush();
  expect(text()).toBe('SeXed');
});

test('a flush ends a held press: what was typed lands at the old caret and nothing waits', async () => {
  const { arrive, flush, input, session, text, textarea } = await setup();
  mouse('mousedown', 30, IN_BODY);
  act(() => input.current!.insertText('X'));
  expect(await settlesSoon(input.current!.flushPendingInput())).toBe(true);
  expect(text()).toBe('SeedX');
  expect(textarea.dataset.pointerPlacement).toBe('ready');
  arrive();
  await flush();
  expect(session.selection()?.head.offset).toBe(5);
});

test('going read-only drops a held press and lands what was typed at the old caret', async () => {
  const { goReadOnly, input, session, text } = await setup();
  mouse('mousedown', 30, IN_BODY);
  act(() => input.current!.insertText('X'));
  goReadOnly();
  await settle();
  expect(text()).toBe('SeedX');
  expect(session.selection()?.head.offset).toBe(5);
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

test('a throw while placing a held press leaves no hold', async () => {
  const consoleError = console.error;
  console.error = () => {};
  try {
    const { arrive, input, text } = await setup({ throwOnPlace: true });
    mouse('mousedown', 30, IN_BODY);
    act(() => input.current!.insertText('X'));
    await settle();
    arrive();
    expect(await settlesSoon(input.current!.flushPendingInput())).toBe(true);
    expect(text()).toBe('SeedX');
  } finally {
    console.error = consoleError;
  }
});

for (const [where, x, y] of [
  ['a header band', 30, IN_HEADER],
  ['a note area no note claims', UNCLAIMED_X, IN_FOOTNOTE],
  ['a position with no target', NO_TARGET_X, IN_BODY],
  ['a position its story does not hold', NO_LOC_X, IN_BODY],
] as const) {
  test(`a press on ${where} leaves typing where the caret was`, async () => {
    const { flush, input, text, textarea } = await setup({ queries: geometry() });
    mouse('mousedown', x, y);
    expect(textarea.dataset.pointerPlacement).toBe('ready');
    expect(textarea.readOnly).toBe(false);
    act(() => input.current!.insertText('X'));
    await flush();
    expect(text()).toBe('SeedX');
  });
}

for (const inFlight of [false, true]) {
  test(`a press during ${inFlight ? 'in-flight' : 'queued'} input takes the caret behind it, and nothing is dropped`, async () => {
    const { flush, input, openResident, session, text, textarea } = await setup({
      gated: true,
      queries: geometry(),
    });
    act(() => input.current!.insertText(' accepted'));
    if (inFlight) await settle();
    mouse('mousedown', 30, IN_BODY);
    act(() => input.current!.insertText('Right '));
    fireEvent.keyDown(textarea, { key: 'Enter' });
    act(() => input.current!.insertText('Next'));
    await settle();
    expect(text()).toBe('Seed');
    openResident();
    await flush();
    // What came before the press at the old caret, what came after at the press.
    expect(session.paragraphs('body').map((paragraph) => paragraph.text)).toEqual([
      'SeRight ',
      'Nexted accepted',
    ]);
    expect(textarea.dataset.pointerPlacement).toBe('ready');
  });
}

test('keys typed after a press never join the open text batch ahead of it', async () => {
  const { arrive, flush, input, openResident, text } = await setup({ gated: true });
  act(() => input.current!.insertText('A'));
  await settle();
  // A second batch waits behind the first, still open to more keys.
  act(() => input.current!.insertText('B'));
  mouse('mousedown', 30, IN_BODY);
  act(() => input.current!.insertText('C'));
  openResident();
  await settle();
  arrive();
  await flush();
  expect(text()).toBe('SeCedAB');
});

test('a composition begun before a press lands at the old caret, then the press takes the caret', async () => {
  const { flush, input, text, textarea } = await setup({ queries: geometry() });
  fireEvent.compositionStart(textarea);
  textarea.value = '日本';
  mouse('mousedown', 30, IN_BODY);
  fireEvent.compositionEnd(textarea, { data: '日本' });
  await settle();
  act(() => input.current!.insertText('X'));
  await flush();
  expect(text()).toBe('SeXed日本');
});

test('a composition during a hold lands where the press was', async () => {
  const { arrive, flush, text, textarea } = await setup();
  mouse('mousedown', 30, IN_BODY);
  fireEvent.compositionStart(textarea);
  textarea.value = '日本';
  arrive();
  fireEvent.compositionEnd(textarea, { data: '日本' });
  await settle();
  await flush();
  expect(text()).toBe('Se日本ed');
});

test('a second press while held text drains waits behind it too', async () => {
  const { arrive, flush, input, openResident, text } = await setup({ gated: true });
  mouse('mousedown', 30, IN_BODY);
  act(() => input.current!.insertText('A'));
  arrive();
  await settle();
  mouse('mousedown', 20, IN_BODY);
  act(() => input.current!.insertText('B'));
  openResident();
  await settle();
  await flush();
  expect(text()).toBe('SBeAed');
});

test('Down during a hold moves from where the press was', async () => {
  const { arrive, flush, session, textarea } = await setup({ secondParagraph: true });
  const [first, second] = session.paragraphs('body');
  mouse('mousedown', 30, IN_BODY);
  fireEvent.keyDown(textarea, { key: 'ArrowDown' });
  arrive();
  await flush();
  expect(session.selection()?.head.paraId).toBe(second.paraId);
  expect(session.selection()?.head.paraId).not.toBe(first.paraId);
});

for (const [where, part, x, y, story, typed, others] of [
  ['a footnote', null, 30, IN_FOOTNOTE, 'fn:2', 'NoXte', ['body']],
  ['an endnote', null, 30, IN_ENDNOTE, 'en:3', 'EnXdn', ['body']],
  ['another band', { kind: 'footer', rId: 'rId8' }, 30, IN_HEADER, 'hf:rId7', 'HeXad', ['hf:rId8']],
  ['the body from a note', { kind: 'footnote', noteId: 2 }, 30, IN_BODY, 'body', 'SeXed', ['fn:2']],
] as const) {
  test(`typing during a hold on ${where} lands there, and leaves focus alone`, async () => {
    const find = document.createElement('input');
    document.body.append(find);
    try {
      const before = await setup({ part: part as PartEdit | null });
      const { arrive, flush, input, session, text } = before;
      const untouched = others.map((other) => text(other));
      mouse('mousedown', x, y);
      act(() => input.current!.insertText('X'));
      find.focus();
      arrive();
      await settle();
      await flush();
      expect(text(story)).toBe(typed);
      expect(others.map((other) => text(other))).toEqual(untouched);
      expect(session.selection()?.head.story).toBe(story);
      expect(document.activeElement === find).toBe(true);
    } finally {
      find.remove();
    }
  });
}

const FIXTURE = new Uint8Array(
  readFileSync(resolve(import.meta.dir, '__fixtures__/probe-linked-header.docx'))
);

function Seeded(props: {
  generation: number;
  input: RefObject<YrsInputRef | null>;
  onSession: (session: YrsSession) => void;
}) {
  const core = useYrsCoreSession(true, null, null, FIXTURE, props.generation);
  if (!core.session) return null;
  props.onSession(core.session);
  return (
    <Editor
      applyResidentInput={residentInput(false).apply}
      initialPart={null}
      input={props.input}
      queries={null}
      readOnly={false}
      session={core.session}
    />
  );
}

for (const leaving of ['an unmount', 'a new session'] as const) {
  test(`typing during a hold lands at the old caret before ${leaving} destroys the session`, async () => {
    const input = createRef<YrsInputRef>() as RefObject<YrsInputRef | null>;
    const seen: YrsSession[] = [];
    const seeded = (generation: number) => (
      <Seeded
        generation={generation}
        input={input}
        onSession={(session) => {
          if (!seen.includes(session)) seen.push(session);
        }}
      />
    );
    const view = render(seeded(1));
    await waitFor(() => expect(seen.length).toBe(1));
    const session = seen[0];
    const first = session.paragraphs('body')[0];
    session.setSelection({ story: 'body', paraId: first.paraId, offset: 0 });
    await settle();
    let atDestroy: string | null = null;
    const destroy = session.destroy.bind(session);
    session.destroy = () => {
      atDestroy = session.paragraphs('body')[0].text;
      destroy();
    };
    mouse('mousedown', 30, IN_BODY);
    act(() => input.current!.insertText('Q'));
    if (leaving === 'an unmount') view.unmount();
    else view.rerender(seeded(2));
    await waitFor(() => expect(atDestroy).not.toBeNull());
    expect(atDestroy!).toBe(`Q${first.text}`);
  });
}
