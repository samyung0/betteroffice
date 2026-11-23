import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { resolve } from 'node:path';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { createRef, useRef } from 'react';
import type { YrsSession } from '@betteroffice/docx/yrs';
import { createResidentEngineSession } from '@betteroffice/docx/yrs/residentEngineSession';
import { PagedEditor, type PagedEditorRef } from './PagedEditor';
import { useCanvasRenderer } from './hooks/useDisplayList';
import {
  destroySessions,
  docxPackage,
  joinSession,
  loadEditor,
  openSession,
  pagedEditorProps,
} from './pagedEditor.testing';

const { act, cleanup, render, waitFor } = await import('@testing-library/react');

// The resident engine worker's own code, run in this process on a real engine
// session (the wasm this test loaded); it paints nothing, as no canvas is attached.
let startWorker: (scope: unknown, canvas: unknown, harness: unknown) => void;
const originalWorker = globalThis.Worker;

beforeAll(async () => {
  await loadEditor();
  const stubs: Record<string, string> = {
    './residentEngineSession':
      'export const createResidentEngineSession = (clientId) => testHarness.createSession(clientId);',
    '../layout/render/glyphCache': 'export class GlyphCache {}',
    '../layout/render/canvasBackend':
      'export const rasterizeDisplayPageToBackBuffer = () => {}; export const presentOffscreenPageBackBuffer = () => {}; export const presentOffscreenPageBackBufferWithCaret = () => {};',
    '../wasm/fontDecoder': 'export const decodeFontFace = () => { throw new Error("no decoder"); };',
    // No sliced faces here (bundling the fonts package's dist trips Bun on Windows paths).
    './residentProgressiveFonts':
      'export class FontSliceUnavailableError extends Error {} export class ResidentProgressiveFonts {} export const glyphsByFont = () => new Map(); export const sliceFaceOpener = () => null;',
  };
  const result = await Bun.build({
    entrypoints: [resolve(import.meta.dir, '../../../../docx/src/yrs/residentEngineWorker.ts')],
    target: 'bun',
    format: 'iife',
    plugins: [
      {
        name: 'in-process-worker',
        setup(build) {
          build.onResolve({ filter: /.*/ }, ({ path, importer }) =>
            importer.replaceAll('\\', '/').endsWith('/residentEngineWorker.ts') && path in stubs
              ? { path, namespace: 'in-process-worker' }
              : undefined
          );
          build.onLoad({ filter: /.*/, namespace: 'in-process-worker' }, ({ path }) => ({
            contents: stubs[path],
            loader: 'js',
          }));
        },
      },
    ],
  });
  if (!result.success) throw new AggregateError(result.logs, 'Worker bundle failed');
  startWorker = new Function(
    'self',
    'OffscreenCanvas',
    'testHarness',
    await result.outputs[0].text()
  ) as typeof startWorker;
});

afterEach(async () => {
  await act(async () => {
    await new Promise((done) => setTimeout(done, 50));
  });
  cleanup();
  destroySessions();
  globalThis.Worker = originalWorker;
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

/** A Worker whose messages reach the worker's code in this process, asynchronously. */
class InProcessWorker {
  onmessage: ((event: { data: unknown }) => void) | null = null;
  onerror = null;
  onmessageerror = null;
  private readonly scope: { onmessage: ((event: { data: unknown }) => void) | null; postMessage(data: unknown): void };
  constructor() {
    this.scope = {
      onmessage: null,
      postMessage: (data) => queueMicrotask(() => this.onmessage?.({ data })),
    };
    startWorker(this.scope, class {}, { createSession: createResidentEngineSession });
  }
  postMessage(data: unknown): void {
    queueMicrotask(() => this.scope.onmessage?.({ data }));
  }
  terminate(): void {}
}

/** The editor as DocxEditor wires it: the canvas renderer's resident worker lays out every pass. */
function CanvasEditor({ session, editor }: { session: YrsSession; editor: React.Ref<PagedEditorRef> }) {
  const canvas = useCanvasRenderer();
  return (
    <PagedEditor
      ref={editor}
      {...pagedEditorProps(session)}
      residentWorkerLayout
      onLayoutComputed={canvas.onLayoutComputed}
    />
  );
}

/**
 * CanvasEditor with the worker started as DocxEditor's session hook starts it:
 * before the page's first render with the session. `swappedFrom`'s worker is
 * started first, as a room swap inside one frame does; `afterSeed` runs
 * before the page's first layout.
 */
function SeededEditor({
  session,
  update,
  editor,
  swappedFrom,
  afterSeed,
}: {
  session: YrsSession;
  update: Uint8Array;
  editor: React.Ref<PagedEditorRef>;
  swappedFrom?: YrsSession;
  afterSeed?: () => void;
}) {
  const canvas = useCanvasRenderer();
  const started = useRef(false);
  if (!started.current) {
    started.current = true;
    if (swappedFrom) canvas.startWorker(swappedFrom, update);
    canvas.startWorker(session, update);
    afterSeed?.();
  }
  return (
    <PagedEditor
      ref={editor}
      {...pagedEditorProps(session)}
      residentWorkerLayout
      onLayoutComputed={canvas.onLayoutComputed}
    />
  );
}

/** Counts the passes `session` lays out on the page. */
function countPageLayouts(session: YrsSession): () => number {
  let count = 0;
  const layOut = session.layoutDocumentWithRegionsRetainedJson.bind(session);
  session.layoutDocumentWithRegionsRetainedJson = (input) => {
    count += 1;
    return layOut(input);
  };
  return () => count;
}

type Posted = { type: string; snapshot?: { state: Uint8Array; media?: string } };

/** In-process workers that record what each was sent and whether it was terminated. */
function recordWorkers(): Array<{ posted: Posted[]; terminated: boolean }> {
  const workers: Array<{ posted: Posted[]; terminated: boolean }> = [];
  globalThis.Worker = class extends InProcessWorker {
    private readonly record = { posted: [] as Posted[], terminated: false };
    constructor() {
      super();
      workers.push(this.record);
    }
    override postMessage(data: unknown): void {
      this.record.posted.push(data as Posted);
      super.postMessage(data);
    }
    override terminate(): void {
      this.record.terminated = true;
    }
  } as unknown as typeof Worker;
  return workers;
}

/** The layout the page lays out for `session` with no worker, once it has `pages` pages. */
async function pageOnlyLayout(session: YrsSession, pages: (count: number) => boolean): Promise<string> {
  globalThis.Worker = undefined as unknown as typeof Worker;
  const editor = createRef<PagedEditorRef>();
  render(<CanvasEditor session={session} editor={editor} />);
  await waitFor(() => expect(pages(editor.current?.getLayout()?.pages.length ?? 0)).toBe(true), {
    timeout: 10000,
  });
  const layout = JSON.stringify(editor.current!.getLayout());
  cleanup();
  return layout;
}

const BODY = Array.from(
  { length: 90 },
  (_, index) =>
    `<w:p w14:paraId="${(0x10000000 + index).toString(16).toUpperCase()}"><w:r><w:t>Line ${index}</w:t></w:r></w:p>`
).join('');

test('the worker lays the pass out and its reply hands the page its layout; the page lays out nothing', async () => {
  globalThis.Worker = InProcessWorker as unknown as typeof Worker;
  const session = await openSession(docxPackage(BODY), 7401);
  let pageLayouts = 0;
  const layOut = session.layoutDocumentWithRegionsRetainedJson.bind(session);
  session.layoutDocumentWithRegionsRetainedJson = (input) => {
    pageLayouts += 1;
    return layOut(input);
  };
  const editor = createRef<PagedEditorRef>();
  render(<CanvasEditor session={session} editor={editor} />);
  await waitFor(() => expect(editor.current?.getLayout()?.pages.length).toBe(2), { timeout: 10000 });
  expect(pageLayouts).toBe(0);
});

test('a worker started as the session opens loads the state the page loaded: its first frame request ships only what changed since, for the same pages (Q815)', async () => {
  globalThis.Worker = InProcessWorker as unknown as typeof Worker;
  const bytes = docxPackage(BODY);
  const room = (await openSession(bytes, 7500)).encodeState();
  const unseeded = createRef<PagedEditorRef>();
  render(<CanvasEditor session={await joinSession(bytes, room, 7501)} editor={unseeded} />);
  await waitFor(() => expect(unseeded.current?.getLayout()?.pages.length).toBe(2), { timeout: 10000 });
  const expected = JSON.stringify(unseeded.current!.getLayout());
  cleanup();

  const workers = recordWorkers();
  const session = await joinSession(bytes, room, 7502);
  const pageLayouts = countPageLayouts(session);
  const editor = createRef<PagedEditorRef>();
  render(<SeededEditor session={session} update={room} editor={editor} />);
  await waitFor(() => expect(editor.current?.getLayout()?.pages.length).toBe(2), { timeout: 10000 });
  expect(JSON.stringify(editor.current!.getLayout())).toBe(expected);
  // The worker laid it out from the seeded replica.
  expect(pageLayouts()).toBe(0);
  const [seed, bootstrap] = workers[0]!.posted.filter(({ type }) => type === 'seed' || type === 'bootstrap');
  expect([seed?.type, bootstrap?.type]).toEqual(['seed', 'bootstrap']);
  // The page's replica holds what the room sent and nothing more.
  expect(bootstrap!.snapshot!.state.byteLength).toBeLessThan(16);
  expect(bootstrap!.snapshot!.media).toBeUndefined();
});

test("a seeded worker's first frame request ships what changed after the seed: a peer's text and the reader's, laid out as the page lays them out (Q815)", async () => {
  const bytes = docxPackage(BODY);
  const room = (await openSession(bytes, 7520)).encodeState();
  const peer = await joinSession(bytes, room, 7521);
  peer.insertText({ story: 'body', paraId: '10000000', offset: 0 }, 'A peer typed this. '.repeat(160));
  const change = (session: YrsSession) => {
    session.applyUpdate(peer.encodeStateAsUpdate(session.encodeStateVector()));
    session.insertText({ story: 'body', paraId: '10000059', offset: 0 }, 'Typed before the first layout. ');
    return session;
  };
  const expected = await pageOnlyLayout(change(await joinSession(bytes, room, 7522)), (pages) => pages > 2);

  const workers = recordWorkers();
  const session = await joinSession(bytes, room, 7523);
  const pageLayouts = countPageLayouts(session);
  const editor = createRef<PagedEditorRef>();
  render(<SeededEditor session={session} update={room} editor={editor} afterSeed={() => change(session)} />);
  await waitFor(() => expect(JSON.stringify(editor.current?.getLayout())).toBe(expected), { timeout: 10000 });
  expect(pageLayouts()).toBe(0);
  const bootstrap = workers[0]!.posted.find(({ type }) => type === 'bootstrap');
  expect(bootstrap!.snapshot!.state.byteLength).toBeGreaterThan(2);
});

test('a pass over the request and document the font check read last reuses its answer: the pass after the fonts load and a web-font load check nothing again (Q815)', async () => {
  globalThis.Worker = InProcessWorker as unknown as typeof Worker;
  const session = await openSession(docxPackage(BODY), 7403);
  let checks = 0;
  const check = session.layoutFontRequirementsJson.bind(session);
  session.layoutFontRequirementsJson = (input) => {
    checks += 1;
    return check(input);
  };
  let passes = 0;
  const record = session.recordResidentLayout!.bind(session);
  session.recordResidentLayout = (input) => {
    passes += 1;
    return record(input);
  };
  // The editor's web-font load listener (this DOM's font set dispatches nothing).
  const fonts = window.document.fonts as unknown as {
    addEventListener(type: string, listener: () => void): void;
  };
  const listen = fonts.addEventListener;
  let fontsLoaded = () => {};
  fonts.addEventListener = (type, listener) => {
    if (type === 'loadingdone') fontsLoaded = listener;
    listen.call(fonts, type, listener);
  };
  const editor = createRef<PagedEditorRef>();
  try {
    render(<CanvasEditor session={session} editor={editor} />);
  } finally {
    fonts.addEventListener = listen;
  }
  await waitFor(() => expect(editor.current?.getLayout()?.pages.length).toBe(2), { timeout: 10000 });
  await act(async () => fontsLoaded());
  await waitFor(() => expect(passes).toBe(2));
  expect(checks).toBe(1);
});

test('a document change after the font check checks again: a family a peer applies is among the faces the next pass prepares (Q815)', async () => {
  globalThis.Worker = InProcessWorker as unknown as typeof Worker;
  const bytes = docxPackage(BODY);
  const session = await openSession(bytes, 7405);
  const checked: string[] = [];
  const check = session.layoutFontRequirementsJson.bind(session);
  session.layoutFontRequirementsJson = (input) => {
    const requirements = check(input);
    checked.push(requirements);
    return requirements;
  };
  const editor = createRef<PagedEditorRef>();
  render(<CanvasEditor session={session} editor={editor} />);
  await waitFor(() => expect(editor.current?.getLayout()?.pages.length).toBe(2), { timeout: 10000 });
  expect(checked).toHaveLength(1);
  const peer = await joinSession(bytes, session.encodeState(), 7406);
  peer.formatRange(
    { story: 'body', start: { paraId: '10000000', offset: 0 }, end: { paraId: '10000000', offset: 6 } },
    { fontFamily: { ascii: 'Noto Serif JP', hAnsi: 'Noto Serif JP' } }
  );
  act(() => {
    session.applyUpdate(peer.encodeStateAsUpdate(session.encodeStateVector()));
  });
  await waitFor(() => expect(checked).toHaveLength(2));
  expect(checked[1]).toContain('"family":"Noto Serif JP"');
});

test('a worker that fails while it loads the seed hands layout to the page, as a failed worker does (Q815, 2026-10-03)', async () => {
  globalThis.Worker = class extends InProcessWorker {
    override postMessage(data: unknown): void {
      if ((data as { type: string }).type !== 'seed') return super.postMessage(data);
      const failed = this as unknown as { onerror(event: { message: string }): void };
      queueMicrotask(() => failed.onerror({ message: 'the worker crashed' }));
    }
  } as unknown as typeof Worker;
  const bytes = docxPackage(BODY);
  const room = (await openSession(bytes, 7510)).encodeState();
  const session = await joinSession(bytes, room, 7511);
  const pageLayouts = countPageLayouts(session);
  const editor = createRef<PagedEditorRef>();
  render(<SeededEditor session={session} update={room} editor={editor} />);
  await waitFor(() => expect(editor.current?.getLayout()?.pages.length).toBe(2), { timeout: 10000 });
  expect(pageLayouts()).toBeGreaterThan(0);
});

test('a seed the worker refuses, living on, leaves the bootstrap to ship the whole state: the worker lays out the pages the page lays out (Q815)', async () => {
  const bytes = docxPackage(BODY);
  const room = (await openSession(bytes, 7530)).encodeState();
  const expected = await pageOnlyLayout(await joinSession(bytes, room, 7531), (pages) => pages === 2);
  globalThis.Worker = class extends InProcessWorker {
    override postMessage(data: unknown): void {
      const request = data as { type: string; seed?: object };
      // A seed whose state is not an update: the worker answers with an error.
      super.postMessage(
        request.type === 'seed' ? { ...request, seed: { ...request.seed, state: Uint8Array.of(255, 255, 255) } } : data
      );
    }
  } as unknown as typeof Worker;
  const session = await joinSession(bytes, room, 7532);
  const pageLayouts = countPageLayouts(session);
  const errors: unknown[] = [];
  const logError = console.error;
  console.error = (message: unknown) => errors.push(message);
  try {
    const editor = createRef<PagedEditorRef>();
    render(<SeededEditor session={session} update={room} editor={editor} />);
    await waitFor(() => expect(JSON.stringify(editor.current?.getLayout())).toBe(expected), { timeout: 10000 });
  } finally {
    console.error = logError;
  }
  expect(errors).toContain('[CanvasRenderer] Resident engine worker seed failed');
  expect(pageLayouts()).toBe(0);
});

test('an editor closed before its first layout ends the worker it started (Q815, Q791)', async () => {
  const workers = recordWorkers();
  const bytes = docxPackage(BODY);
  const room = (await openSession(bytes, 7540)).encodeState();
  render(<SeededEditor session={await joinSession(bytes, room, 7541)} update={room} editor={createRef()} />);
  cleanup();
  expect(workers.map(({ posted, terminated }) => [posted.map(({ type }) => type), terminated])).toEqual([
    [['seed', 'destroy'], true],
  ]);
});

test("a session swapped in before the first layout ends the first session's worker; the second's lays out (Q815, Q791)", async () => {
  const workers = recordWorkers();
  const bytes = docxPackage(BODY);
  const room = (await openSession(bytes, 7550)).encodeState();
  const first = await joinSession(bytes, room, 7551);
  const second = await joinSession(bytes, room, 7552);
  const pageLayouts = countPageLayouts(second);
  const editor = createRef<PagedEditorRef>();
  render(<SeededEditor session={second} swappedFrom={first} update={room} editor={editor} />);
  await waitFor(() => expect(editor.current?.getLayout()?.pages.length).toBe(2), { timeout: 10000 });
  expect(pageLayouts()).toBe(0);
  expect(workers.map(({ posted, terminated }) => [posted[0]?.type, terminated])).toEqual([
    ['seed', true],
    ['seed', false],
  ]);
  expect(workers[1]!.posted.map(({ type }) => type)).toContain('bootstrap');
});
