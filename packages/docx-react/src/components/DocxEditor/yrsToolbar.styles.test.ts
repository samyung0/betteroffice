import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createStyleResolver } from '@betteroffice/docx/styles';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsInputPositionMap, createYrsSession } from '@betteroffice/docx/yrs';
import { applyYrsToolbarFormatting } from './yrsToolbar';

const WASM = resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm');

const styles = createStyleResolver({
  docDefaults: { pPr: { spaceAfter: 160 } },
  styles: [
    { styleId: 'Normal', type: 'paragraph', default: true },
    { styleId: 'ListParagraph', type: 'paragraph', pPr: { indentLeft: 720, contextualSpacing: true } },
  ],
});

test('outdent stores 0 where the style indents, and applying a style writes its paragraph values', async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  const session = await createYrsSession({ clientId: 8 });
  try {
    session.createStory('body', 'StyledPlain');
    const [first] = session.paragraphs('body');
    session.splitParagraph({ story: 'body', paraId: first.paraId, offset: 6 });
    const [styled, plain] = session.paragraphs('body');
    session.setParagraphAttr(styled.paraId, 'pStyle', 'ListParagraph');
    session.setParagraphAttr(styled.paraId, 'indentLeft', 720);
    session.setParagraphAttr(plain.paraId, 'indentLeft', 720);
    const map = createYrsInputPositionMap('body', [
      { paraId: styled.paraId, length: 6 },
      { paraId: plain.paraId, length: 5 },
    ]);
    session.setSelection({ story: 'body', paraId: styled.paraId, offset: 0 }, { story: 'body', paraId: plain.paraId, offset: 1 });
    expect(applyYrsToolbarFormatting(session, map, 'outdent', styles)).toBe(true);
    const [afterStyled, afterPlain] = session.paragraphs('body');
    expect(afterStyled.properties.indentLeft).toBe(0);
    expect(afterPlain.properties.indentLeft).toBeUndefined();

    session.setSelection({ story: 'body', paraId: plain.paraId, offset: 0 });
    expect(applyYrsToolbarFormatting(session, map, { type: 'applyStyle', value: 'ListParagraph' }, styles)).toBe(true);
    expect(session.paragraphs('body')[1]!.properties).toMatchObject({
      pStyle: 'ListParagraph',
      indentLeft: 720,
      contextualSpacing: true,
      spaceAfter: 160,
    });
  } finally {
    session.destroy();
  }
});
