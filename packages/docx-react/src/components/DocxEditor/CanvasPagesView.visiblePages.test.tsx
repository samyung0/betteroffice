import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import type { DisplayList } from '@betteroffice/docx/layout/render';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { CanvasPagesView } from './CanvasPagesView';

const { cleanup, render } = await import('@testing-library/react');

afterEach(() => cleanup());
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('a long document reports the pages in view as it measures them, so those not built yet build first (Q815 C)', () => {
  const displayList = {
    pages: Array.from({ length: 30 }, (_, pageIndex) => ({
      pageIndex,
      width: 100,
      height: 100,
      primitives: [],
      unbuilt: true,
    })),
  } as DisplayList;
  const seen: Array<[number, number]> = [];
  render(
    <CanvasPagesView
      displayList={displayList}
      glyphOutlineProvider={false}
      interactive
      onVisiblePages={(first, last) => seen.push([first, last])}
    />
  );
  expect(seen.length).toBeGreaterThan(0);
  const [first, last] = seen[0]!;
  expect(first).toBe(0);
  expect(last).toBeGreaterThanOrEqual(first);
  expect(last).toBeLessThan(30);
});
