import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsInputPositionMap, createYrsSession } from '@betteroffice/docx/yrs';
import { applyYrsToolbarFormatting } from './yrsToolbar';

const WASM = resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm');

test('the toolbar toggles superscript and subscript on the selection, each replacing the other', async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  const session = await createYrsSession({ clientId: 7 });
  try {
    session.createStory('body', 'AXZ');
    const [paragraph] = session.paragraphs('body');
    const at = (offset: number) => ({ story: 'body', paraId: paragraph.paraId, offset });
    session.setSelection(at(1), at(2));
    const map = createYrsInputPositionMap('body', [{ paraId: paragraph.paraId, length: 3 }]);
    const x = () => {
      const [block] = session.yrsBlocksForStory('body') as {
        runs: { text: string; superscript?: boolean; subscript?: boolean }[];
      }[];
      const run = block.runs.find((candidate) => candidate.text === 'X');
      return { superscript: !!run?.superscript, subscript: !!run?.subscript };
    };

    expect(applyYrsToolbarFormatting(session, map, 'superscript')).toBe(true);
    expect(x()).toEqual({ superscript: true, subscript: false });
    applyYrsToolbarFormatting(session, map, 'subscript');
    expect(x()).toEqual({ superscript: false, subscript: true });
    applyYrsToolbarFormatting(session, map, 'subscript');
    expect(x()).toEqual({ superscript: false, subscript: false });
  } finally {
    session.destroy();
  }
});
