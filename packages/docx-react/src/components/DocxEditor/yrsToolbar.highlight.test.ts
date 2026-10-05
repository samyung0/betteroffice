import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsInputPositionMap, createYrsSession } from '@betteroffice/docx/yrs';
import { useSelectionTracker, type SelectionStateDelta } from './hooks/useSelectionTracker';
import {
  applyYrsToolbarFormatting,
  currentYrsToolbarSelection,
  withStoredYrsFormatting,
} from './yrsToolbar';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { renderHook } = await import('@testing-library/react');

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const WASM = resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm');

test('the selection reports its highlight as it reports its colour', async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  const session = await createYrsSession({ clientId: 8 });
  try {
    session.createStory('body', 'one two six');
    const [paragraph] = session.paragraphs('body');
    const at = (offset: number) => ({ story: 'body', paraId: paragraph.paraId, offset });
    const map = createYrsInputPositionMap('body', [{ paraId: paragraph.paraId, length: 11 }]);
    const highlight = (from: number, to: number) => {
      session.setSelection(at(from), at(to));
      return currentYrsToolbarSelection(session, map)!.context.highlight;
    };
    // What the toolbar's highlight picker is given for a selection.
    const deltas: SelectionStateDelta[] = [];
    const { result } = renderHook(() =>
      useSelectionTracker({
        borderSpecRef: { current: { style: 'single', size: 4, color: { rgb: '000000' } } },
        theme: null,
        historyStateRef: { current: null },
        getCachedStyleResolver: () => {
          throw new Error('unused');
        },
        setFloatingCommentBtn: () => {},
        applySelectionDelta: (delta) => deltas.push(delta),
        recomputeFloatingCommentBtn: () => {},
        onSelectionChange: undefined,
        selectionChangeSubscribersRef: { current: new Set() },
      })
    );
    const picked = (from: number, to: number) => {
      session.setSelection(at(from), at(to));
      result.current.handleYrsSelectionChange(currentYrsToolbarSelection(session, map)!);
      return deltas[deltas.length - 1].selectionFormatting.highlight;
    };

    session.setSelection(at(0), at(3));
    applyYrsToolbarFormatting(session, map, { type: 'highlightColor', value: 'FFFF00' }, () => ({}));
    session.setSelection(at(4), at(7));
    applyYrsToolbarFormatting(session, map, { type: 'highlightColor', value: 'F4CCCC' }, () => ({}));

    // Word's palette colours come back as their names, any other colour as its hex.
    expect(highlight(0, 3)).toBe('yellow');
    expect(highlight(4, 7)).toBe('F4CCCC');
    expect(highlight(2, 2)).toBe('yellow');
    expect(highlight(0, 7)).toBeNull();
    expect(highlight(8, 11)).toBeNull();
    expect([picked(0, 3), picked(4, 7), picked(0, 7)]).toEqual(['yellow', 'F4CCCC', undefined]);

    // A highlight picked at a collapsed caret shows before anything is typed.
    session.setSelection(at(9), at(9));
    const caret = currentYrsToolbarSelection(session, map)!;
    expect(withStoredYrsFormatting(caret, { clear: false, delta: { highlight: 'cyan' } }).context.highlight).toBe('cyan');
    expect(withStoredYrsFormatting(caret, { clear: true, delta: {} }).context.highlight).toBeNull();
  } finally {
    session.destroy();
  }
});
