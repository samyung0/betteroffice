import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import type { DisplayPage } from '@betteroffice/docx/layout/render';
import { CanvasPageMirror } from './CanvasPageMirror';

const { act, cleanup, render } = await import('@testing-library/react');

afterEach(() => {
  cleanup();
});

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const page = {
  pageIndex: 0,
  width: 500,
  height: 500,
  primitives: [
    {
      kind: 'text',
      text: 'Hello',
      x: 100,
      baselineY: 100,
      width: 40,
      font: '400 16px Calibri',
      color: '#000000',
      docStart: 1,
      docEnd: 6,
      blockId: 0,
      lineIndex: 0,
    },
  ],
} as unknown as DisplayPage;

const mirrorPage = (container: HTMLElement) => container.querySelector('.layout-page-mirror');

test('swaps between the positioned and the plain-text mirror at idle time', async () => {
  const { container, rerender } = render(<CanvasPageMirror page={page} revision={0} full />);
  expect(mirrorPage(container)?.classList.contains('layout-page-mirror-text')).toBe(false);

  // leaving the viewport window: the positioned mirror stays until idle time
  rerender(<CanvasPageMirror page={page} revision={0} full={false} />);
  expect(mirrorPage(container)?.classList.contains('layout-page-mirror-text')).toBe(false);
  await act(() => new Promise((resolve) => setTimeout(resolve, 200)));
  expect(mirrorPage(container)?.classList.contains('layout-page-mirror-text')).toBe(true);
  expect(mirrorPage(container)?.textContent).toBe('Hello');

  rerender(<CanvasPageMirror page={page} revision={0} full />);
  await act(() => new Promise((resolve) => setTimeout(resolve, 200)));
  expect(mirrorPage(container)?.classList.contains('layout-page-mirror-text')).toBe(false);
  expect(mirrorPage(container)?.querySelector('.layout-run-text')?.textContent).toBe('Hello');
});

test('a selectable page holding a selection end keeps its positioned mirror', async () => {
  const { container, rerender } = render(<CanvasPageMirror page={page} revision={0} full selectable />);
  expect(container.querySelector('.canvas-page-mirror--selectable')).not.toBeNull();
  const text = container.querySelector('.layout-run-text')?.firstChild as Text;
  document.getSelection()?.setBaseAndExtent(text, 1, text, 4);

  rerender(<CanvasPageMirror page={page} revision={0} full={false} selectable />);
  await act(() => new Promise((resolve) => setTimeout(resolve, 200)));
  expect(mirrorPage(container)?.classList.contains('layout-page-mirror-text')).toBe(false);
  expect(document.getSelection()?.toString()).toBe('ell');

  // once the selection lets go, the page goes plain
  document.getSelection()?.removeAllRanges();
  await act(() => new Promise((resolve) => setTimeout(resolve, 50)));
  expect(mirrorPage(container)?.classList.contains('layout-page-mirror-text')).toBe(true);
});

test('a new version of the page keeps the mounted mirror and its unchanged runs', () => {
  const twoRuns = (last: string) =>
    ({
      ...page,
      primitives: [
        page.primitives[0],
        { ...page.primitives[0], text: last, x: 150, docStart: 6, docEnd: 6 + last.length },
      ],
    }) as unknown as DisplayPage;
  const { container, rerender } = render(<CanvasPageMirror page={twoRuns('world')} revision={0} full />);
  const root = mirrorPage(container);
  const first = container.querySelector('.layout-run-text');

  rerender(<CanvasPageMirror page={twoRuns('worlds')} revision={0} full />);
  expect(mirrorPage(container)).toBe(root);
  expect(container.querySelector('.layout-run-text')).toBe(first);
  expect(root?.textContent).toBe('Helloworlds');
});

test('a selection end in a plain-text page does not keep it from turning positioned', async () => {
  const { container, rerender } = render(<CanvasPageMirror page={page} revision={0} full={false} selectable />);
  const plain = container.querySelector('.layout-page-mirror-text')!;
  const text = document.createTreeWalker(plain, NodeFilter.SHOW_TEXT).nextNode() as Text;
  document.getSelection()?.setBaseAndExtent(text, 1, text, 4);

  rerender(<CanvasPageMirror page={page} revision={0} full selectable />);
  await act(() => new Promise((resolve) => setTimeout(resolve, 200)));
  expect(mirrorPage(container)?.classList.contains('layout-page-mirror-text')).toBe(false);
  document.getSelection()?.removeAllRanges();
});

test("the editor's mirror reads a shaped line as one run, a viewer's text layer keeps one per glyph cluster", () => {
  const text = 'Marker01 keeps';
  const shaped = {
    ...page,
    primitives: [
      {
        kind: 'glyphRun',
        fontId: 0,
        size: 16,
        color: '#000000',
        text,
        glyphs: Array.from(text, (_, index) => ({ id: 7, x: 100 + index * 8, y: 100, cluster: index, advance: 8 })),
        docStart: 1,
        docEnd: 1 + text.length,
        blockId: 0,
        lineIndex: 0,
        clusterRuns: true,
      },
    ],
  } as unknown as DisplayPage;
  const runs = (container: HTMLElement) =>
    [...container.querySelectorAll('.layout-run-text')].map((run) => run.textContent);
  const editor = render(<CanvasPageMirror page={shaped} revision={0} full />);
  expect(runs(editor.container)).toEqual([text]);
  const viewer = render(<CanvasPageMirror page={shaped} revision={0} full selectable />);
  expect(runs(viewer.container)).toEqual(Array.from(text));
});
