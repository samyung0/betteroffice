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

/** CanvasEditor with the worker started as DocxEditor's session hook starts it: before the page's first render with the session. */
function SeededEditor({
  session,
  update,
  editor,
}: {
  session: YrsSession;
  update: Uint8Array;
  editor: React.Ref<PagedEditorRef>;
}) {
  const canvas = useCanvasRenderer();
  const started = useRef(false);
  if (!started.current) {
    started.current = true;
    canvas.startWorker(session, update);
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

  type Posted = { type: string; snapshot?: { state: Uint8Array; media?: string } };
  const posted: Posted[] = [];
  globalThis.Worker = class extends InProcessWorker {
    override postMessage(data: unknown): void {
      posted.push(data as Posted);
      super.postMessage(data);
    }
  } as unknown as typeof Worker;
  const session = await joinSession(bytes, room, 7502);
  const pageLayouts = countPageLayouts(session);
  const editor = createRef<PagedEditorRef>();
  render(<SeededEditor session={session} update={room} editor={editor} />);
  await waitFor(() => expect(editor.current?.getLayout()?.pages.length).toBe(2), { timeout: 10000 });
  expect(JSON.stringify(editor.current!.getLayout())).toBe(expected);
  // The worker laid it out from the seeded replica.
  expect(pageLayouts()).toBe(0);
  const [seed, bootstrap] = posted.filter(({ type }) => type === 'seed' || type === 'bootstrap');
  expect([seed?.type, bootstrap?.type]).toEqual(['seed', 'bootstrap']);
  // The page's replica holds what the room sent and nothing more.
  expect(bootstrap!.snapshot!.state.byteLength).toBeLessThan(16);
  expect(bootstrap!.snapshot!.media).toBeUndefined();
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
