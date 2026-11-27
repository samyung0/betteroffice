/**
 * A caret that moves while it shows sets no state: a no-op setVisible(true)
 * from its position effect was queued at the Default lane without a render,
 * and the next urgent render of the caret (the next key) skipped it and left
 * Default work pending, so React's nested-update count climbed while typing
 * until it threw #185.
 */
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import { Profiler } from 'react';
import type { CaretPosition } from '@betteroffice/docx/layout';
import { SelectionOverlay } from './SelectionOverlay';

// React's nested-update count as React keeps it (as useDisplayList.pages.test.ts
// counts it): a commit that leaves Sync, InputContinuous or Default work
// (lanes 42) pending adds one, any other resets it; past 50 React throws #185.
// A commit whose passive effects flush at once counts the lanes after that
// flush. Installed before react-dom loads.
const nested = { count: 0, max: 0 };
let open: { lanes: number } | null = null;
const settleCommit = (lanes: number): void => {
  nested.count = lanes & 42 ? nested.count + 1 : 0;
  nested.max = Math.max(nested.max, nested.count);
};
(globalThis as { __REACT_DEVTOOLS_GLOBAL_HOOK__?: unknown }).__REACT_DEVTOOLS_GLOBAL_HOOK__ = {
  supportsFiber: true,
  isDisabled: false,
  renderers: new Map(),
  inject(this: { renderers: Map<number, unknown> }, internals: unknown) {
    this.renderers.set(this.renderers.size + 1, internals);
    return this.renderers.size;
  },
  onCommitFiberRoot(_id: number, root: { pendingLanes: number }) {
    if (open) settleCommit(open.lanes);
    const commit = { lanes: root.pendingLanes };
    open = commit;
    queueMicrotask(() => {
      if (open === commit) {
        settleCommit(commit.lanes);
        open = null;
      }
    });
  },
  onPostCommitFiberRoot(_id: number, root: { pendingLanes: number }) {
    if (!open) return;
    settleCommit(root.pendingLanes);
    open = null;
  },
  onCommitFiberUnmount() {},
  setStrictMode() {},
  checkDCE() {},
};

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render } = await import('@testing-library/react');
const { flushSync } = await import('react-dom');
afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('a visible caret that moves leaves no update for the next render to skip', async () => {
  let commits = 0;
  const view = (caret: CaretPosition) => (
    <Profiler id="overlay" onRender={() => (commits += 1)}>
      <SelectionOverlay selectionRects={[]} caretPosition={caret} isFocused />
    </Profiler>
  );
  const shown = render(view({ pageIndex: 0, x: 100, y: 100, height: 20 }));
  await act(async () => {});
  // A key moves the caret; its position effect runs while the caret shows.
  shown.rerender(view({ pageIndex: 0, x: 108, y: 100, height: 20 }));
  await act(async () => {});
  commits = 0;
  nested.count = 0;
  nested.max = 0;
  // The next urgent render with the same props (the next key's first commit).
  await act(async () => {
    flushSync(() => shown.rerender(view({ pageIndex: 0, x: 108, y: 100, height: 20 })));
  });
  expect(commits).toBe(1);
  expect(nested.max).toBe(0);
});
