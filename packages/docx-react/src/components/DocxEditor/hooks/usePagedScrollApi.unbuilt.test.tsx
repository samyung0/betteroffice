/**
 * A jump to a position on a page not built yet (a Find hit, an outline entry;
 * Q815 C): the query asks for the page, and the scroll follows once it is
 * built.
 */
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import { usePagedScrollApi } from './usePagedScrollApi';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');

afterEach(() => cleanup());
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const PAGE = { width: 800, height: 1000 };

function queries(built: boolean): DisplayListQueries {
  return {
    displayList: { pages: [{ pageIndex: 0, ...PAGE, primitives: [], ...(built ? {} : { unbuilt: true }) }] },
    pageSize: () => PAGE,
    anchorRect: () => (built ? { pageIndex: 0, x: 10, y: 700, width: 5, height: 20 } : null),
  } as unknown as DisplayListQueries;
}

test('a jump to a page not built yet scrolls once the page is built', () => {
  const host = document.createElement('div');
  host.className = 'canvas-pages';
  const canvas = document.createElement('canvas');
  canvas.dataset.pageIndex = '0';
  canvas.getBoundingClientRect = () =>
    ({ left: 0, top: 0, right: PAGE.width, bottom: PAGE.height, ...PAGE }) as DOMRect;
  host.append(canvas);
  const scrolls: number[] = [];
  const scroller = {
    scrollTop: 0,
    clientHeight: 400,
    getBoundingClientRect: () => ({ top: 0 }) as DOMRect,
    scrollTo: ({ top }: { top: number }) => scrolls.push(top),
  } as unknown as HTMLDivElement;
  const { result, rerender } = renderHook(
    ({ displayListQueries }) =>
      usePagedScrollApi({
        pagesContainerRef: { current: null },
        yrsInputRef: { current: null },
        yrsSession: null,
        yrsLocToDisplayPosition: () => null,
        getScrollContainer: () => scroller,
        displayListQueries,
        canvasHostRef: { current: host },
      }),
    { initialProps: { displayListQueries: queries(false) } }
  );
  act(() => result.current.scrollToPositionImpl(150));
  expect(scrolls).toEqual([]);
  rerender({ displayListQueries: queries(true) });
  // The rect's middle (710) to the pane's middle (200).
  expect(scrolls).toEqual([510]);
  rerender({ displayListQueries: queries(true) });
  expect(scrolls).toEqual([510]);
});
