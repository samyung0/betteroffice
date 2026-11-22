import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import {
  destroySessions,
  docxPackage,
  loadEditor,
  openSession,
} from '../pagedEditor.testing';
import {
  CanvasCitationOverlay,
  citedRange,
  findCitedParagraph,
  type DocxCitation,
} from './CanvasCitationOverlay';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, render } = await import('@testing-library/react');

beforeAll(loadEditor);
afterEach(() => {
  cleanup();
  destroySessions();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const QUOTE = 'cited passage sits';
const CITED = `The ${QUOTE} in its own paragraph.`;
const offered: string[][] = [];
/** A host rule: the one paragraph holding the quote once; records what it was offered. */
const citation: DocxCitation = {
  color: 'yellow',
  match: <T extends { text: string }>(items: T[]) => {
    offered.push(items.map((item) => item.text));
    const holders = items.filter((item) => item.text.split(QUOTE).length === 2);
    return holders.length === 1 ? holders[0] : null;
  },
  probe: 'passage',
};
const toDisplay = (loc: { offset: number }) => 10 + loc.offset;

// A citation is display only: finding it reads the document and changes
// nothing a save, the undo history or a collaborator would see.
test('the host rule picks among the paragraphs holding the probe and leaves the document untouched', async () => {
  const session = await openSession(
    docxPackage(
      `<w:p w14:paraId="10000001"><w:r><w:t>An introduction.</w:t></w:r></w:p>` +
        `<w:p w14:paraId="10000002"><w:r><w:t>${CITED}</w:t></w:r></w:p>`
    ),
    7601
  );
  const state = session.encodeStateAsUpdate();
  const selection = session.selection();
  const canUndo = session.canUndo();
  let updates = 0;
  const stop = session.onUpdate(() => {
    updates += 1;
  });
  offered.length = 0;
  const cited = findCitedParagraph(session, citation);
  const range = cited && citedRange(session, cited, citation, toDisplay);
  stop();

  expect(offered[0]).toEqual([CITED]);
  expect(cited).toEqual({ paraId: '10000002', story: 'body' });
  expect(range).toEqual({ from: 10, to: 10 + CITED.length });
  expect(updates).toBe(0);
  expect(session.encodeStateAsUpdate()).toEqual(state);
  expect(session.selection()).toEqual(selection);
  expect(session.canUndo()).toBe(canUndo);
});

// Q809 A: checked on every change with the same rule, never moved.
test('the mark goes once its paragraph no longer holds the quote, and with the paragraph', async () => {
  const session = await openSession(
    docxPackage(
      `<w:p w14:paraId="10000001"><w:r><w:t>${CITED}</w:t></w:r></w:p>` +
        `<w:p w14:paraId="10000002"><w:r><w:t>Then the ${QUOTE} again.</w:t></w:r></w:p>`
    ),
    7603
  );
  const cited = { paraId: '10000001', story: 'body' };
  expect(citedRange(session, cited, citation, toDisplay)).not.toBeNull();
  // Typing elsewhere in it keeps the mark at its paragraph's range.
  session.insertText({ ...cited, offset: 0 }, 'Now: ');
  expect(citedRange(session, cited, citation, toDisplay)?.to).toBe(10 + 5 + CITED.length);
  // Joined with the next paragraph, it holds the quote twice: no longer the match.
  session.mergeParagraphs('body', cited.paraId, 'forward');
  expect(citedRange(session, cited, citation, toDisplay)).toBeNull();
  // Replaced text: gone too.
  const text = session.paragraphUnitText('body', cited.paraId);
  session.replaceRange(
    { end: { ...cited, offset: text.length }, start: { ...cited, offset: 0 }, story: 'body' },
    'Other words.'
  );
  expect(citedRange(session, cited, citation, toDisplay)).toBeNull();
  expect(citedRange(session, { paraId: 'FFFFFFFF', story: 'body' }, citation, toDisplay)).toBeNull();
});

test('a paragraph outside the body is not marked', async () => {
  const session = await openSession(docxPackage(`<w:p><w:r><w:t>Body text.</w:t></w:r></w:p>`), 7602);
  session.createStory('hf:rId1', CITED);
  expect(findCitedParagraph(session, citation)).toBeNull();
});

const target = document.createElement('div');
const host = document.createElement('div');
const hostRef = { current: host };
function overlay(
  range: { from: number; to: number },
  scrollKey: unknown,
  { presented = true, zoom = 1 }: { presented?: boolean; zoom?: number } = {}
) {
  const queries = {
    pageSize: () => ({ width: 800, height: 1000 }),
    rangeRects: (from: number, to: number) => [
      { pageIndex: 0, x: from, y: 100, width: to - from, height: 20 },
      { pageIndex: 0, x: from, y: 120, width: 100, height: 20 },
    ],
  } as unknown as DisplayListQueries;
  return (
    <CanvasCitationOverlay
      range={range}
      color="rgba(255, 196, 0, 0.3)"
      scrollKey={scrollKey}
      presented={presented}
      overlayTarget={target}
      canvasHostRef={hostRef}
      displayListQueries={queries}
      sidebarOpen={false}
      zoom={zoom}
    />
  );
}

test('the mark is drawn in its colour, centred once the pages are presented and again as the zoom settles', () => {
  document.body.append(target, host);
  const canvas = document.createElement('canvas');
  canvas.dataset.pageIndex = '0';
  // The page shown at half size, 10px in from the overlay's corner.
  canvas.getBoundingClientRect = () => ({ left: 10, top: 10, width: 400, height: 500 }) as DOMRect;
  host.append(canvas);
  const scrolled: unknown[] = [];
  const scrollIntoView = HTMLElement.prototype.scrollIntoView;
  HTMLElement.prototype.scrollIntoView = function (options) {
    scrolled.push([this.dataset.citationHighlight, options]);
  };
  const Observer = globalThis.ResizeObserver;
  let observers = 0;
  globalThis.ResizeObserver = class extends Observer {
    constructor(callback: ResizeObserverCallback) {
      super(callback);
      observers += 1;
    }
  };
  try {
    const first = {};
    const view = render(overlay({ from: 72, to: 372 }, first, { presented: false }));
    const marks = [...target.querySelectorAll<HTMLElement>('[data-citation-highlight]')];
    expect(marks.map((mark) => [mark.style.left, mark.style.top, mark.style.width])).toEqual([
      ['46px', '60px', '150px'],
      ['46px', '70px', '50px'],
    ]);
    expect(marks[0].style.background).toBe('rgba(255, 196, 0, 0.3)');
    expect(scrolled).toEqual([]);
    view.rerender(overlay({ from: 72, to: 372 }, first));
    expect(scrolled).toEqual([['true', { block: 'center' }]]);

    // Typing moves the paragraph: the mark follows it on the same observer,
    // without scrolling again.
    view.rerender(overlay({ from: 80, to: 400 }, first));
    expect(target.querySelector<HTMLElement>('[data-citation-highlight]')?.style.left).toBe('50px');
    expect(scrolled).toHaveLength(1);
    expect(observers).toBe(1);

    // A fitted zoom settling centres it again, until the reader moves.
    view.rerender(overlay({ from: 80, to: 400 }, first, { zoom: 0.5 }));
    expect(scrolled).toHaveLength(2);
    document.dispatchEvent(new Event('wheel'));
    view.rerender(overlay({ from: 80, to: 400 }, first, { zoom: 0.75 }));
    expect(scrolled).toHaveLength(2);

    // The next citation centres its passage again.
    view.rerender(overlay({ from: 80, to: 400 }, {}, { zoom: 0.75 }));
    expect(scrolled).toHaveLength(3);
  } finally {
    HTMLElement.prototype.scrollIntoView = scrollIntoView;
    globalThis.ResizeObserver = Observer;
    target.remove();
    host.remove();
  }
});
