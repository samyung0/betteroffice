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
