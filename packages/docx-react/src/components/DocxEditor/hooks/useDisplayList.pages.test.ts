/**
 * Visible pages first in DOCX Edit (Q815 C): the worker's first frame builds
 * the pages the viewport shows, plus one; the others come a batch at a time,
 * those in view first, then those a caller waits for, until none is left.
 * The worker is the real engine driven through a recording port.
 */
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { DisplayPage } from '@betteroffice/docx/layout/render';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { createEditSession, preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import type { YrsResidentWorkerProbe, YrsSession } from '@betteroffice/docx/yrs';
import type {
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerResponse,
} from '@betteroffice/docx/yrs/residentEngineWorkerProtocol';
import { pageBuildBatch, useRustDisplayList } from './useDisplayList';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');
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
  reply(request: ResidentEngineWorkerRequest, frame: Uint8Array, unbuiltPages: number): void {
    this.onmessage?.({ data: {
      id: request.id, ok: true, frame: frame.slice().buffer, updates: [],
      caret: { frameEpoch: 0, caretRect: null }, selection: null, layoutRevision: 1,
      ...(unbuiltPages > 0 ? { unbuiltPages, media: false } : {}),
    } } as unknown as MessageEvent<ResidentEngineWorkerResponse>);
  }
  terminate(): void {}
}

const settle = async (): Promise<void> => {
  for (let i = 0; i < 20; i += 1) await Promise.resolve();
};

const page = (pageIndex: number, unbuilt = false) =>
  ({ pageIndex, width: 100, height: 100, primitives: [], ...(unbuilt ? { unbuilt } : {}) }) as DisplayPage;

test('a batch takes the pages in view, then those asked for, then the nearest to the view', () => {
  const pages = Array.from({ length: 30 }, (_, index) => page(index, index > 2));
  expect(pageBuildBatch(pages, null, [])).toEqual([3, 4, 5, 6, 7, 8, 9, 10]);
  expect(pageBuildBatch(pages, { first: 20, last: 21 }, [29, 2])).toEqual([20, 21, 29, 22, 19, 23, 18, 24]);
});

test('the first frame builds the pages in view; the rest build until none is left, those in view first', async () => {
  globalThis.Worker = RecordingWorker as unknown as typeof Worker;
  const native = createEditSession(9311);
  native.load_json(JSON.stringify([{
    storyId: 'body',
    paragraphs: Array.from({ length: 900 }, (_, index) => ({ text: `Paragraph ${index} of the body` })),
  }]));
  const request = JSON.stringify({
    bodyStory: 'body',
    regions: { sections: [{ sectionId: 'main', properties: {} }] },
    measurement: { defaults: { fontSize: 11, fontFamily: 'Calibri' } },
    renderEnv: {},
  });
  const inputs = JSON.parse(native.layout_document_with_regions_json(request));
  const probe: YrsResidentWorkerProbe = { layoutRevision: 1, fontsRevision: 0, regionLayoutInput: request };
  const engine = {
    residentWorkerProbe: () => ({ ...probe }),
    onFontSlicesRetry: () => () => {},
    residentWorkerSnapshot: () => ({
      clientId: 1, state: new Uint8Array(), selection: null, fonts: [], fontsRevision: 0,
      renderInputs: [], measureInputs: [], layoutInput: request, layoutWithRegions: true, layoutRevision: 1,
    }),
    onUpdate: () => () => {},
    selection: () => null,
    applyUpdate: () => null,
  } as unknown as YrsSession;
  const viewport = { current: () => ({ height: 600, zoom: 1 }) };
  const overrides = { getInputs: () => inputs };
  try {
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(inputs.layout as Layout, overrides, undefined, undefined, engine, viewport)
    );
    await act(settle);
    const worker = RecordingWorker.current!;
    const [bootstrap] = worker.of('bootstrap');
    expect(bootstrap!.firstPages).toEqual({ height: 600, zoom: 1 });
    const pageCount = native.layout_page_heights().length;
    expect(pageCount).toBeGreaterThan(12);
    const unbuilt = () => result.current.frame!.displayList.pages.flatMap((p, i) => (p.unbuilt ? [i] : []));
    // The worker builds the first pages only.
    await act(async () => {
      worker.reply(bootstrap!, native.build_display_list_frame_first('{}', 0, 2), pageCount - 2);
      await settle();
    });
    expect(unbuilt()).toHaveLength(pageCount - 2);
    expect(result.current.frame!.displayList.pages).toHaveLength(pageCount);

    // A caller waits for the last page; the view moves to pages 9 and 10.
    let lastBuilt = false;
    void result.current.ensurePages([pageCount - 1]).then(() => {
      lastBuilt = true;
    });
    // The first batch was asked for as the frame landed, from the top.
    const builds = () => worker.of('buildPages');
    expect(builds()[0]!.pages).toEqual([2, 3, 4, 5, 6, 7, 8, 9]);
    act(() => result.current.setVisiblePages(9, 10));
    let epoch = 1;
    for (let round = 0; unbuilt().length > 0 && round < 10; round += 1) {
      const build = builds().at(-1)!;
      expect(build.expectedFrameEpoch).toBe(epoch);
      const frame = native.build_unbuilt_display_pages_frame(Uint32Array.from(build.pages), epoch);
      epoch += 1;
      const left = unbuilt().filter((index) => !build.pages.includes(index)).length;
      await act(async () => {
        worker.reply(build, frame, left);
        await settle();
      });
      // The next batch is asked for once the frame is applied.
      await act(settle);
      if (round === 0) {
        // Then the page in view not built yet, the page waited for, and the nearest.
        expect(builds()[1]!.pages.slice(0, 2)).toEqual([10, pageCount - 1]);
      }
    }
    expect(unbuilt()).toEqual([]);
    expect(lastBuilt).toBe(true);
    // Every page built: no request follows.
    const sent = builds().length;
    await act(settle);
    expect(builds()).toHaveLength(sent);
    unmount();
  } finally {
    native.free();
  }
});
