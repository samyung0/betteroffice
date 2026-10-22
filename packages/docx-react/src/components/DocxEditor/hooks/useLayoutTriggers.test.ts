import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import type { HeaderFooter } from '@betteroffice/docx/types/document';
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

// A header or footer save and the hidden-text toggle are the reader's own
// actions: their pass is local, so the caret stays in view (Epo 2026-10-23).
test('lays out a header/footer or render change as a local pass', () => {
  const calls: string[] = [];
  const options = (headerContent: HeaderFooter | null) => ({
    headerContent,
    runLayoutPipeline: () => calls.push('run'),
    scheduleLayout: (origin?: string) => calls.push(`schedule ${origin}`),
    updateSelectionOverlay: () => {},
  });
  const { rerender } = renderHook((header: HeaderFooter | null) => useLayoutTriggers(options(header)), {
    initialProps: null as HeaderFooter | null,
  });
  expect(calls).toEqual([]);
  rerender({ type: 'header' } as unknown as HeaderFooter);
  expect(calls).toEqual(['schedule local', 'run']);
});
