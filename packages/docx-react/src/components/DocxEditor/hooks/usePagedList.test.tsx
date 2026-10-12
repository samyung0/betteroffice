import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import type { DisplayList } from '@betteroffice/docx/layout/render';
import { usePagedList } from './usePagedList';

const { act, cleanup, renderHook } = await import('@testing-library/react');

afterEach(() => cleanup());
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const listOf = (pages: number) =>
  ({ pages: Array.from({ length: pages }, (_, pageIndex) => ({ pageIndex })) }) as unknown as DisplayList;

test('a list with pages is drawn at once', async () => {
  const ready = listOf(2);
  const { result } = renderHook(() => usePagedList(ready));
  expect(await result.current()).toBe(ready);
});

test('a draw asked for before the layout lands waits for the list with pages, once', async () => {
  const { result, rerender } = renderHook(({ list }) => usePagedList(list), {
    initialProps: { list: null as DisplayList | null },
  });
  const drawn: DisplayList[] = [];
  const pending = result.current().then((list) => drawn.push(list));
  await act(async () => rerender({ list: listOf(0) }));
  expect(drawn).toEqual([]);
  const landed = listOf(3);
  await act(async () => rerender({ list: landed }));
  await pending;
  expect(drawn).toEqual([landed]);
  // A later list resolves nothing more.
  await act(async () => rerender({ list: listOf(4) }));
  expect(drawn).toEqual([landed]);
});
