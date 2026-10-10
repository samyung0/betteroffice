import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { createEditSession, preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import type { YrsResidentWorkerProbe, YrsSession } from '@betteroffice/docx/yrs';
import type {
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerResponse,
} from '@betteroffice/docx/yrs/residentEngineWorkerProtocol';
import { useRustDisplayList } from './useDisplayList';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');
const originalWorker = globalThis.Worker;

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(resolve(
  import.meta.dir, '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm'
)))));

afterEach(() => {
  cleanup();
  globalThis.Worker = originalWorker;
});

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

class RecordingWorker {
  static current: RecordingWorker | null = null;
  onmessage: ((event: MessageEvent<ResidentEngineWorkerResponse>) => void) | null = null;
  onerror: ((event: ErrorEvent) => void) | null = null;
  onmessageerror = null;
  posted: ResidentEngineWorkerRequest[] = [];
  constructor() {
    RecordingWorker.current = this;
  }
  postMessage(request: ResidentEngineWorkerRequest): void {
    this.posted.push(request);
  }
  of<T extends ResidentEngineWorkerRequest['type']>(type: T) {
    return this.posted.filter((request) => request.type === type) as Array<
      Extract<ResidentEngineWorkerRequest, { type: T }>
    >;
  }
  reply(request: ResidentEngineWorkerRequest, frame: Uint8Array, layoutRevision: number): void {
    this.onmessage?.({ data: {
      id: request.id, ok: true, frame: frame.slice().buffer, updates: [],
      caret: { frameEpoch: 0, caretRect: null }, selection: null, layoutRevision,
    } } as unknown as MessageEvent<ResidentEngineWorkerResponse>);
  }
  terminate(): void {}
}

/**
 * A main-thread host whose remote updates and layouts the test drives: each
 * `remote()` is one peer update reaching the session, each `layout()` one
 * main-thread layout pass (a new revision and Layout object).
 */
function host() {
  const native = createEditSession(9301);
  native.create_story('body', 'Remote text', 'Normal', 'left');
  const inputs = JSON.parse(native.layout_document_with_regions_json(JSON.stringify({
    bodyStory: 'body',
    regions: { sections: [{ sectionId: 'main', properties: {} }] },
    measurement: { defaults: { fontSize: 11, fontFamily: 'Calibri' } },
    renderEnv: {},
  })));
  // Frames as the worker chains them: a full first frame, then deltas each
  // based on the one before.
  let builtEpoch = 0;
  const nextFrame = () => {
    const frame = native.build_display_list_frame(JSON.stringify(inputs), builtEpoch);
    builtEpoch += 1;
    return frame;
  };
  const listeners = new Set<(update: Uint8Array, origin: 'local' | 'remote') => void>();
  const probe: YrsResidentWorkerProbe = {
    layoutRevision: 1,
    fontsRevision: 0,
    regionLayoutInput: '{"bodyStory":"body"}',
  };
  const engine = {
    residentWorkerProbe: () => ({ ...probe }),
    residentWorkerSnapshot: () => ({
      clientId: 1, state: new Uint8Array(), selection: null, fonts: [], fontsRevision: 0,
      renderInputs: [], measureInputs: [], layoutInput: probe.regionLayoutInput,
      layoutWithRegions: true, layoutRevision: probe.layoutRevision,
    }),
    onUpdate: (listener: (update: Uint8Array, origin: 'local' | 'remote') => void) => {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    selection: () => null,
    applyUpdate: () => null,
  } as unknown as YrsSession;
  let layout = inputs.layout as Layout;
  return {
    engine,
    overrides: { getInputs: () => inputs },
    get layout() {
      return layout;
    },
    nextFrame,
    remote(byte: number) {
      for (const listener of listeners) listener(new Uint8Array([byte]), 'remote');
    },
    layoutPass(): Layout {
      probe.layoutRevision += 1;
      layout = { ...layout };
      return layout;
    },
    free: () => native.free(),
  };
}

async function settle(): Promise<void> {
  for (let i = 0; i < 20; i += 1) await Promise.resolve();
}

test('a burst of remote updates costs one relayout delta, and layouts landing meanwhile one more', async () => {
  globalThis.Worker = RecordingWorker as unknown as typeof Worker;
  const h = host();
  try {
    const { result, rerender, unmount } = renderHook(
      ({ layout }) => useRustDisplayList(layout, h.overrides, undefined, undefined, h.engine),
      { initialProps: { layout: h.layout } }
    );
    await act(settle);
    const worker = RecordingWorker.current!;
    const [bootstrap] = worker.of('bootstrap');
    const first = h.nextFrame();
    await act(async () => worker.reply(bootstrap!, first, 1));
    await waitFor(() => expect(result.current.frame?.frameEpoch).toBe(1));

    // Three peer updates within one frame, then the main thread's one layout.
    for (const byte of [1, 2, 3]) h.remote(byte);
    await act(async () => {
      rerender({ layout: h.layoutPass() });
      await settle();
    });
    expect(worker.of('applyUpdate')).toHaveLength(3);
    expect(worker.of('sync')).toHaveLength(0);
    const relayouts = worker.of('relayout');
    expect(relayouts).toHaveLength(1);
    expect(relayouts[0]).toMatchObject({ expectedFrameEpoch: 1, layoutRevision: 2 });

    // More updates and layouts while that relayout is in flight queue one
    // request, sent only once the first reply is applied.
    for (const byte of [4, 5, 6]) {
      h.remote(byte);
      await act(async () => {
        rerender({ layout: h.layoutPass() });
        await settle();
      });
    }
    expect(worker.of('relayout')).toHaveLength(1);

    // The reply's content is older than the updates since, yet it is applied:
    // the next relayout diffs against it instead of asking for a full frame.
    await act(async () => {
      worker.reply(relayouts[0]!, h.nextFrame(), 2);
      await settle();
    });
    expect(result.current.frame?.frameEpoch).toBe(2);
    const followUp = worker.of('relayout');
    expect(followUp).toHaveLength(2);
    expect(followUp[1]).toMatchObject({ expectedFrameEpoch: 2, layoutRevision: 5 });
    await act(async () => {
      worker.reply(followUp[1]!, h.nextFrame(), 5);
      await settle();
    });
    expect(result.current.frame?.frameEpoch).toBe(3);
    expect(result.current.error).toBeNull();
    expect(result.current.workerSurfacesActive).toBe(true);
    expect(worker.of('relayout')).toHaveLength(2);
    expect(worker.of('sync')).toHaveLength(0);
    unmount();
  } finally {
    h.free();
  }
});

test('a fonts change still syncs the worker in full', async () => {
  globalThis.Worker = RecordingWorker as unknown as typeof Worker;
  const h = host();
  try {
    const { rerender, unmount } = renderHook(
      ({ layout }) => useRustDisplayList(layout, h.overrides, undefined, undefined, h.engine),
      { initialProps: { layout: h.layout } }
    );
    await act(settle);
    const worker = RecordingWorker.current!;
    await act(async () => worker.reply(worker.of('bootstrap')[0]!, h.nextFrame(), 1));
    const probe = h.engine.residentWorkerProbe()!;
    (h.engine as unknown as { residentWorkerProbe: () => YrsResidentWorkerProbe }).residentWorkerProbe =
      () => ({ ...probe, layoutRevision: 2, fontsRevision: 1 });
    await act(async () => {
      rerender({ layout: { ...h.layout } });
      await settle();
    });
    expect(worker.of('relayout')).toHaveLength(0);
    expect(worker.of('sync')).toHaveLength(1);
    unmount();
  } finally {
    h.free();
  }
});

test('a relayout delta shifts later pages in place, so their objects and query-store copies stay', async () => {
  globalThis.Worker = RecordingWorker as unknown as typeof Worker;
  const native = createEditSession(9302);
  native.load_json(JSON.stringify([{
    storyId: 'body',
    paragraphs: Array.from({ length: 150 }, (_, index) => ({ text: `Paragraph ${index} of the body` })),
  }]));
  const request = JSON.stringify({
    bodyStory: 'body',
    regions: { sections: [{ sectionId: 'main', properties: {} }] },
    measurement: { defaults: { fontSize: 11, fontFamily: 'Calibri' } },
    renderEnv: {},
  });
  const inputs = JSON.parse(native.layout_document_with_regions_json(request));
  const first = native.relayout_frame(request, '{}', 0);
  const probe: YrsResidentWorkerProbe = { layoutRevision: 1, fontsRevision: 0, regionLayoutInput: request };
  const listeners = new Set<(update: Uint8Array, origin: 'local' | 'remote') => void>();
  const engine = {
    residentWorkerProbe: () => ({ ...probe }),
    residentWorkerSnapshot: () => ({
      clientId: 1, state: new Uint8Array(), selection: null, fonts: [], fontsRevision: 0,
      renderInputs: [], measureInputs: [], layoutInput: request, layoutWithRegions: true, layoutRevision: 1,
    }),
    onUpdate: (listener: (update: Uint8Array, origin: 'local' | 'remote') => void) => {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    selection: () => null,
    applyUpdate: () => null,
  } as unknown as YrsSession;
  const overrides = { getInputs: () => inputs };
  try {
    const { result, rerender, unmount } = renderHook(
      ({ layout }) => useRustDisplayList(layout, overrides, undefined, undefined, engine),
      { initialProps: { layout: inputs.layout as Layout } }
    );
    await act(settle);
    const worker = RecordingWorker.current!;
    await act(async () => worker.reply(worker.of('bootstrap')[0]!, first, 1));
    const pages = result.current.frame!.pages;
    expect(pages.length).toBeGreaterThan(2);
    const last = pages.at(-1)!.page;

    // A peer types at the document's start: every later page only shifts.
    const [opening] = JSON.parse(native.paragraphs('body')) as Array<{ paraId: string }>;
    native.insert_text('body', opening!.paraId, 0, 'x', undefined, undefined);
    for (const listener of listeners) listener(new Uint8Array([1]), 'remote');
    probe.layoutRevision = 2;
    await act(async () => {
      rerender({ layout: { ...inputs.layout } });
      await settle();
    });
    const [relayout] = worker.of('relayout');
    const delta = native.relayout_frame(request, '{}', relayout!.expectedFrameEpoch);
    await act(async () => {
      worker.reply(relayout!, delta, 2);
      await settle();
    });
    expect(result.current.frame!.frameEpoch).toBe(2);
    expect(result.current.frame!.pages.at(-1)!.page).toBe(last);
    unmount();
  } finally {
    native.free();
  }
});
