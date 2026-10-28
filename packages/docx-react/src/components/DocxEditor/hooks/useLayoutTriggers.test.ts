import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import type { HeaderFooter } from '@betteroffice/docx/types/document';
import type { YrsRenderEnv } from '@betteroffice/docx/yrs';
import { useLayoutTriggers } from './useLayoutTriggers';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, renderHook } = await import('@testing-library/react');
// happy-dom has no FontFaceSet; the font trigger only listens on it.
if (!document.fonts)
  Object.defineProperty(document, 'fonts', {
    configurable: true,
    value: { addEventListener: () => {}, removeEventListener: () => {} },
  });

afterEach(cleanup);
afterAll(() => {
  if (ownsDom) GlobalRegistrator.unregister();
});

// Only the reader's hidden-text toggle is a local pass here: a peer's change
// (a comment anywhere renumbers the env, a peer's header edit) keeps the
// reader's view, and the reader's header or footer save schedules its own
// local pass (Epo 2026-10-23).
test('lays out only the hidden-text toggle as a local pass', () => {
  const calls: string[] = [];
  type Props = { header: HeaderFooter | null; renderEnv: YrsRenderEnv };
  const { rerender } = renderHook(
    ({ header, renderEnv }: Props) =>
      useLayoutTriggers({
        headerContent: header,
        renderEnv,
        runLayoutPipeline: () => calls.push('run'),
        scheduleLayout: (origin?: string) => calls.push(`schedule ${origin}`),
        updateSelectionOverlay: () => {},
      }),
    { initialProps: { header: null, renderEnv: { showHiddenText: false } } as Props }
  );
  expect(calls).toEqual([]);
  // A peer's comment: a new env with the same toggle.
  rerender({ header: null, renderEnv: { numericIds: { c1: 1 }, showHiddenText: false } });
  expect(calls).toEqual(['run']);
  // A peer's header edit.
  rerender({
    header: { type: 'header' } as unknown as HeaderFooter,
    renderEnv: { numericIds: { c1: 1 }, showHiddenText: false },
  });
  expect(calls).toEqual(['run', 'run']);
  calls.length = 0;
  rerender({
    header: { type: 'header' } as unknown as HeaderFooter,
    renderEnv: { numericIds: { c1: 1 }, showHiddenText: true },
  });
  expect(calls).toEqual(['schedule local', 'run']);
});
