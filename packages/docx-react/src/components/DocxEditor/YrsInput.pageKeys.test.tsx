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
  for (const session of sessions.splice(0)) session.destroy();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

// 20 one-line paragraphs on two 100 x 200 pages, ten 20px lines a page, the pages 10px apart in a
// 205px viewport: one PageDown is ten lines plus the gap.
async function mount() {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
  );
  const body = Array.from(
    { length: 20 },
    (_, index) => `<w:p><w:r><w:t>Line ${String(index).padStart(2, '0')}</w:t></w:r></w:p>`
  ).join('');
  zip.file(
    'word/document.xml',
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr/></w:body></w:document>`
  );
  const session = await createYrsSession();
  sessions.push(session);
  session.openDocx(await zip.generateAsync({ type: 'uint8array' }), true);
  const map = (story = 'body') => createYrsInputPositionMap(story, session.paragraphSpans(story));
  const lines: DisplayListVisualLine[] = map().paragraphs.map((paragraph, index) => ({
    pageIndex: Math.floor(index / 10),
    x: 0,
    y: (index % 10) * 20,
    width: 100,
    height: 20,
    baseline: (index % 10) * 20 + 15,
    from: paragraph.displayStart,
    to: paragraph.displayStart + paragraph.length,
    paraId: paragraph.paraId,
  }));
  const lineAt = (pos: number) => lines.find((line) => line.from <= pos && pos <= line.to) ?? null;
  const queries = {
    isReady: () => true,
    pageSize: () => ({ width: 100, height: 200 }),
    visualLines: () => lines,
    visualLineAtPosition: lineAt,
    caretRect: (pos: number) => {
      const line = lineAt(pos);
      return (
        line && { pageIndex: line.pageIndex, x: pos - line.from, y: line.y, width: 0, height: 20 }
      );
    },
    hitTestRegions: (pageIndex: number, x: number, y: number) => {
      const line = lines.find((l) => l.pageIndex === pageIndex && l.y <= y && y < l.y + l.height);
      return line ? { region: 'body', pos: Math.min(line.to, line.from + Math.round(x)) } : null;
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
  for (const pageIndex of [0, 1]) {
    const page = document.createElement('div');
    page.className = 'canvas-page';
    page.dataset.pageIndex = String(pageIndex);
    const top = () => pageIndex * 210 - scroller.scrollTop;
    page.getBoundingClientRect = () =>
      ({
        top: top(),
        bottom: top() + 200,
        left: 0,
        right: 100,
        width: 100,
        height: 200,
      } as DOMRect);
    host.appendChild(page);
  }
  document.body.appendChild(scroller);

  const input = createRef<YrsInputRef>();
  const view = render(
    <YrsInput
      ref={input}
      enabled
      readOnly={false}
      session={session}
      inputPositionMap={map}
      displayPositionToLoc={(position, story) => displayPositionToYrsLoc(map(story), position)}
      locToDisplayPosition={(loc) => yrsLocToDisplayPosition(map(loc.story), loc)}
      displayListQueries={queries}
      canvasHostRef={{ current: host }}
      onStateChange={() => {}}
      onDirectInput={() => {}}
    />
  );
  const textarea = view.getByTestId('yrs-input');
  const press = async (key: string, shiftKey = false) => {
    fireEvent.keyDown(textarea, { key, shiftKey });
    await act(async () => {
      await input.current!.flushPendingInput();
    });
  };
  const at = (index: number, offset: number) => ({
    story: 'body',
    paraId: map().paragraphs[index]!.paraId,
    offset,
  });
  return { session, scroller, press, at };
}

// #1538: PageUp/PageDown move the caret about one viewport at its x and the view follows;
// Shift extends the selection.
test('PageDown and PageUp move the caret a viewport, Shift extends, and the view follows', async () => {
  const { session, scroller, press, at } = await mount();
  session.setSelection(at(2, 3));
  const scrollTop = scroller.scrollTop;

  await press('PageDown');
  expect(session.selection()).toEqual({ anchor: at(12, 3), head: at(12, 3) });
  expect(scroller.scrollTop - scrollTop).toBe(210);

  await press('PageDown', true);
  expect(session.selection()).toEqual({ anchor: at(12, 3), head: at(19, 3) });

  await press('PageUp');
  expect(session.selection()?.head).toEqual(at(9, 3));
});
