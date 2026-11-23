import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { createRef } from 'react';
import type { DisplayListQueries, DisplayListVisualLine } from '@betteroffice/docx/layout/render';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import {
  createYrsInputPositionMap,
  createYrsSession,
  displayPositionToYrsLoc,
  yrsLocToDisplayPosition,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { YrsInput, type YrsInputRef } from './YrsInput';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render } = await import('@testing-library/react');
const sessions: YrsSession[] = [];

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(
        resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
      )
    )
  )
);
afterEach(() => {
  cleanup();
  document.body.innerHTML = '';
  for (const session of sessions.splice(0)) session.destroy();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

// Ten one-line paragraphs on a 100 x 200 page, 20px a line, in a 205px viewport the reader has
// scrolled 400px down, away from the caret on line 2. The editor mounts before its geometry.
async function mount() {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
  );
  const body = Array.from(
    { length: 10 },
    (_, index) => `<w:p><w:r><w:t>Line ${index}</w:t></w:r></w:p>`
  ).join('');
  zip.file(
    'word/document.xml',
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr/></w:body></w:document>`
  );
  const session = await createYrsSession();
  sessions.push(session);
  session.openDocx(await zip.generateAsync({ type: 'uint8array' }), true);
  const map = (story = 'body') => createYrsInputPositionMap(story, session.paragraphSpans(story));
  const lines = (): DisplayListVisualLine[] =>
    map().paragraphs.map((paragraph, index) => ({
      pageIndex: 0,
      x: 0,
      y: index * 20,
      width: 100,
      height: 20,
      baseline: index * 20 + 15,
      from: paragraph.displayStart,
      to: paragraph.displayStart + paragraph.length,
      paraId: paragraph.paraId,
    }));
  const lineAt = (pos: number) =>
    lines().find((line) => line.from <= pos && pos <= line.to) ?? null;
  const queries = {
    isReady: () => true,
    pageSize: () => ({ width: 100, height: 200 }),
    visualLines: lines,
    visualLineAtPosition: lineAt,
    caretRect: (pos: number) => {
      const line = lineAt(pos);
      return (
        line && { pageIndex: line.pageIndex, x: pos - line.from, y: line.y, width: 0, height: 20 }
      );
    },
  } as unknown as DisplayListQueries;

  const scroller = document.createElement('div');
  scroller.style.overflowY = 'auto';
  Object.defineProperty(scroller, 'clientHeight', { value: 205 });
  Object.defineProperty(scroller, 'scrollHeight', { value: 10_000 });
  scroller.getBoundingClientRect = () =>
    ({ top: 0, bottom: 205, left: 0, right: 100, width: 100, height: 205 } as DOMRect);
  const host = document.createElement('div');
  scroller.appendChild(host);
  const page = document.createElement('div');
  page.className = 'canvas-page';
  page.dataset.pageIndex = '0';
  page.getBoundingClientRect = () =>
    ({
      top: -scroller.scrollTop,
      bottom: 200 - scroller.scrollTop,
      left: 0,
      right: 100,
      width: 100,
      height: 200,
    } as DOMRect);
  host.appendChild(page);
  document.body.appendChild(scroller);

  const input = createRef<YrsInputRef>();
  const editor = (displayListQueries: DisplayListQueries | null) => (
    <YrsInput
      ref={input}
      enabled
      readOnly={false}
      session={session}
      inputPositionMap={map}
      displayPositionToLoc={(position, story) => displayPositionToYrsLoc(map(story), position)}
      locToDisplayPosition={(loc) => yrsLocToDisplayPosition(map(loc.story), loc)}
      displayListQueries={displayListQueries}
      layoutUpdateOrigin="remote"
      canvasHostRef={{ current: host }}
      onStateChange={() => {}}
      onDirectInput={() => {}}
    />
  );
  const view = render(editor(null));
  session.setSelection({ story: 'body', paraId: map().paragraphs[2]!.paraId, offset: 3 });
  scroller.scrollTop = 400;
  const press = async (key: string) => {
    fireEvent.keyDown(view.getByTestId('yrs-input'), { key });
    await act(async () => {
      await input.current!.flushPendingInput();
    });
  };
  const geometryArrives = () => view.rerender(editor(queries));
  return { press, geometryArrives, scroller };
}

// Q833 B: once the geometry arrives, the caret the reader edited at comes into view (its line
// at 40px, 24px below the viewport's top).
test('an edit before the first caret pass brings the caret into view once the geometry arrives', async () => {
  const { press, geometryArrives, scroller } = await mount();
  await press('Delete');
  expect(scroller.scrollTop).toBe(400);
  geometryArrives();
  expect(scroller.scrollTop).toBe(16);
});

// A document's first frame with no edit keeps the reader where they are (2026-10-22).
test('a first caret pass with no edit keeps the view', async () => {
  const { geometryArrives, scroller } = await mount();
  geometryArrives();
  expect(scroller.scrollTop).toBe(400);
});
