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

async function docx() {
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
  return zip.generateAsync({ type: 'uint8array' });
}

async function openSession(clientId: number, state?: Uint8Array) {
  const session = await createYrsSession({ clientId });
  sessions.push(session);
  session.openDocx(await docx(), !state);
  if (state) session.loadState(state);
  return session;
}

const mapOf = (session: YrsSession) => (story = 'body') =>
  createYrsInputPositionMap(story, session.paragraphSpans(story));

/** The caret on line 2 of `session`. */
function placeCaret(session: YrsSession) {
  const paraId = mapOf(session)().paragraphs[2]!.paraId;
  session.setSelection({ story: 'body', paraId, offset: 3 });
}

// Ten one-line paragraphs on a 100 x 200 page, 20px a line.
function queriesOf(session: YrsSession): DisplayListQueries {
  const lines = (): DisplayListVisualLine[] =>
    mapOf(session)().paragraphs.map((paragraph, index) => ({
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
  return {
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
}

// The page in a 205px viewport the reader has scrolled 400px down, away from the caret on line
// 2. The editor mounts before its geometry.
async function mount() {
  const session = await openSession(1);
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
  const editor = (on: YrsSession, geometry: boolean) => {
    const map = mapOf(on);
    return (
      <YrsInput
        ref={input}
        enabled
        readOnly={false}
        session={on}
        inputPositionMap={map}
        displayPositionToLoc={(position, story) => displayPositionToYrsLoc(map(story), position)}
        locToDisplayPosition={(loc) => yrsLocToDisplayPosition(map(loc.story), loc)}
        displayListQueries={geometry ? queriesOf(on) : null}
        layoutUpdateOrigin="remote"
        canvasHostRef={{ current: host }}
        onStateChange={() => {}}
        onDirectInput={() => {}}
      />
    );
  };
  const view = render(editor(session, false));
  placeCaret(session);
  scroller.scrollTop = 400;
  const press = async (key: string) => {
    fireEvent.keyDown(view.getByTestId('yrs-input'), { key });
    await act(async () => {
      await input.current!.flushPendingInput();
    });
  };
  /** The geometry arrives, for `on` if the editor has swapped to another session. */
  const geometryArrives = (on = session) => view.rerender(editor(on, true));
  return { session, input, press, geometryArrives, scroller };
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

// The ribbon, Find's replace and Undo edit outside the input and mark it through its ref.
test('an edit made outside the input before the first caret pass brings the caret into view', async () => {
  const { input, geometryArrives, scroller } = await mount();
  input.current!.noteLocalEdit();
  geometryArrives();
  expect(scroller.scrollTop).toBe(16);
});

// A document's first frame with no edit keeps the reader where they are (2026-10-22).
test('a first caret pass with no edit keeps the view', async () => {
  const { geometryArrives, scroller } = await mount();
  geometryArrives();
  expect(scroller.scrollTop).toBe(400);
});

// A peer's edit is no edit of the reader's: the session takes it and the next layout is remote.
test("a peer's edit before the first caret pass keeps the view", async () => {
  const { session, geometryArrives, scroller } = await mount();
  const peer = await openSession(2, session.encodeStateAsUpdate());
  const start = { story: 'body', paraId: mapOf(peer)().paragraphs[0]!.paraId, offset: 0 };
  peer.insertText(start, 'Peer ');
  session.applyUpdate(peer.encodeStateAsUpdate(session.encodeStateVector()));
  geometryArrives();
  expect(scroller.scrollTop).toBe(400);
});

// A swapped-in or re-seeded session keeps the reader's place (2026-10-22): an edit on the
// session it replaced is not its own.
test('a session that replaced the edited one keeps the view at its first caret pass', async () => {
  const { press, geometryArrives, scroller } = await mount();
  await press('Delete');
  const next = await openSession(3);
  placeCaret(next);
  geometryArrives(next);
  expect(scroller.scrollTop).toBe(400);
});
