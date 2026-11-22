import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import type { YrsSession } from '@betteroffice/docx/yrs';
import {
  destroySessions,
  docxPackage,
  loadEditor,
  openSession,
  renderPagedEditor,
  residentWorker,
} from './pagedEditor.testing';

const { act, cleanup, waitFor } = await import('@testing-library/react');

beforeAll(loadEditor);
afterEach(async () => {
  await act(async () => {
    await new Promise((done) => setTimeout(done, 50));
  });
  cleanup();
  destroySessions();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

/** Ninety one-line paragraphs: two pages. */
const BODY = Array.from(
  { length: 90 },
  (_, index) =>
    `<w:p w14:paraId="${(0x10000000 + index).toString(16).toUpperCase()}"><w:r><w:t>Line ${index}</w:t></w:r></w:p>`
).join('');

/** Counts the passes this engine lays out on the page. */
function pageLayouts(session: YrsSession): () => number {
  let count = 0;
  const layOut = session.layoutDocumentWithRegionsRetainedJson.bind(session);
  session.layoutDocumentWithRegionsRetainedJson = (input) => {
    count += 1;
    return layOut(input);
  };
  return () => count;
}

test('the page records each pass and takes the layout the worker laid out', async () => {
  const session = await openSession(docxPackage(BODY), 7301);
  const laidOut = pageLayouts(session);
  const worker = residentWorker(session);
  const ref = await renderPagedEditor(session, worker);
  await waitFor(() => expect(ref.current?.getLayout()?.pages.length).toBe(2));
  expect(worker.passes.length).toBeGreaterThan(0);
  expect(laidOut()).toBe(0);
});

test('a worker that fails hands layout back: the page lays out that pass and every later one', async () => {
  const session = await openSession(docxPackage(BODY), 7302);
  const laidOut = pageLayouts(session);
  const worker = residentWorker(session, true);
  const ref = await renderPagedEditor(session, worker);
  await waitFor(() => expect(ref.current?.getLayout()?.pages.length).toBe(2));
  const passes = worker.passes.length;
  expect(laidOut()).toBe(1);
  act(() => ref.current?.relayout());
  expect(laidOut()).toBe(2);
  expect(worker.passes.length).toBe(passes);
});
