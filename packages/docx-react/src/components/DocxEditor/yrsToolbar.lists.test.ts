import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { toBytes } from '@betteroffice/docx/docx/rezip/parts';
import { createYrsInputPositionMap, type YrsSession } from '@betteroffice/docx/yrs';
import { destroySessions, docxPackage, loadEditor, openSession, savedBytes } from './pagedEditor.testing';
import { applyYrsToolbarFormatting } from './yrsToolbar';

beforeAll(loadEditor);
afterEach(destroySessions);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const level = (format: string, text: string) =>
  `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${text}"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>`;
// The bullet and number buttons write numId 1 and 2 (setList); this file
// defines those as bullets and numbers.
const NUMBERING = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0">${level('bullet', '•')}</w:abstractNum><w:abstractNum w:abstractNumId="1">${level('decimal', '%1.')}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;

function listDocx(): Uint8Array {
  const body =
    '<w:p w14:paraId="10000001"><w:r><w:t>Apples</w:t></w:r></w:p>' +
    '<w:p w14:paraId="10000002"><w:r><w:t>First step</w:t></w:r></w:p>';
  return docxPackage(body, {
    parts: { 'word/numbering.xml': toBytes(NUMBERING) },
    rels: '<Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>',
    contentTypes:
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>',
  });
}

/** Bullets on the first paragraph and numbers on the second, from the toolbar. */
function toggleLists(session: YrsSession): void {
  const map = () => createYrsInputPositionMap('body', session.paragraphSpans('body'));
  session.setSelection({ story: 'body', paraId: '10000001', offset: 2 });
  expect(applyYrsToolbarFormatting(session, map(), 'bulletList', () => ({}))).toBe(true);
  session.setSelection({ story: 'body', paraId: '10000002', offset: 2 });
  expect(applyYrsToolbarFormatting(session, map(), 'numberedList', () => ({}))).toBe(true);
}

/** What a paragraph shows as its list: bullet, a number format, or nothing. */
function lists(session: YrsSession): string[] {
  return session.paragraphs('body').map(({ properties }) =>
    !properties.numPr
      ? 'none'
      : properties.listIsBullet
        ? 'bullet'
        : String(properties.listNumFmt ?? 'unknown')
  );
}

test('toolbar bullets and numbers save and reopen as the editor shows them', async () => {
  const session = await openSession(listDocx(), 7601);
  toggleLists(session);
  expect(lists(session)).toEqual(['bullet', 'decimal']);
  const reopened = await openSession(await savedBytes(session), 7602);
  expect(lists(reopened)).toEqual(['bullet', 'decimal']);
});
