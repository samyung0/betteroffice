import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import {
  destroySessions,
  docxPackage,
  loadEditor,
  openSession,
} from '../pagedEditor.testing';
import { CanvasCitationOverlay, citedRange, findCitedParagraph } from './CanvasCitationOverlay';

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

const CITED = 'The cited passage sits in its own paragraph.';
const pick = <T extends { text: string }>(items: T[]) =>
  items.find((item) => item.text.includes('cited passage')) ?? null;

// A citation is display only: finding it reads the document and changes
// nothing a save, the undo history or a collaborator would see.
test('the host rule picks a body paragraph by its text and leaves the document untouched', async () => {
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
  const cited = findCitedParagraph(session, { color: 'yellow', match: pick });
  const range = cited && citedRange(session, cited, (loc) => 10 + loc.offset);
  stop();

  expect(cited).toEqual({ paraId: '10000002', story: 'body' });
  expect(range).toEqual({ from: 10, to: 10 + CITED.length });
  expect(updates).toBe(0);
  expect(session.encodeStateAsUpdate()).toEqual(state);
  expect(session.selection()).toEqual(selection);
  expect(session.canUndo()).toBe(canUndo);
  expect(findCitedParagraph(session, { color: 'yellow', match: () => null })).toBeNull();
});

test('a paragraph outside the body is not marked, and a removed one has no range', async () => {
  const session = await openSession(docxPackage(`<w:p><w:r><w:t>Body text.</w:t></w:r></w:p>`), 7602);
  session.createStory('header:rId1', CITED);
  expect(findCitedParagraph(session, { color: 'yellow', match: pick })).toBeNull();
  expect(citedRange(session, { paraId: 'FFFFFFFF', story: 'body' }, () => 0)).toBeNull();
});

function overlay(range: { from: number; to: number }, scrollKey: unknown) {
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
      overlayTarget={target}
      canvasHostRef={{ current: host }}
      displayListQueries={queries}
      sidebarOpen={false}
      zoom={1}
    />
  );
}
const target = document.createElement('div');
const host = document.createElement('div');

test('the mark is drawn in its colour over the page and centred once per citation', () => {
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
  try {
    const first = {};
    const view = render(overlay({ from: 72, to: 372 }, first));
    const marks = [...target.querySelectorAll<HTMLElement>('[data-citation-highlight]')];
    expect(marks.map((mark) => [mark.style.left, mark.style.top, mark.style.width])).toEqual([
      ['46px', '60px', '150px'],
      ['46px', '70px', '50px'],
    ]);
    expect(marks[0].style.background).toBe('rgba(255, 196, 0, 0.3)');
    expect(scrolled).toEqual([['true', { block: 'center' }]]);

    // Typing moves the paragraph: the mark follows it without scrolling again.
    view.rerender(overlay({ from: 80, to: 400 }, first));
    expect(target.querySelector<HTMLElement>('[data-citation-highlight]')?.style.left).toBe('50px');
    expect(scrolled).toHaveLength(1);

    // The next citation centres its passage again.
    view.rerender(overlay({ from: 80, to: 400 }, {}));
    expect(scrolled).toHaveLength(2);
  } finally {
    HTMLElement.prototype.scrollIntoView = scrollIntoView;
    target.remove();
    host.remove();
  }
});
