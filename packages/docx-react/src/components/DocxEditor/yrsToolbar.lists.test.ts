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
/** numbering.xml whose numId 1 and 2 have the given first levels. */
const numbering = (one: string, two: string) =>
  `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0">${one}</w:abstractNum><w:abstractNum w:abstractNumId="1">${two}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;
const BULLET = level('bullet', '•');
const DECIMAL = level('decimal', '%1.');

function listDocx(numberingXml?: string): Uint8Array {
  const body =
    '<w:p w14:paraId="10000001"><w:r><w:t>Apples</w:t></w:r></w:p>' +
    '<w:p w14:paraId="10000002"><w:r><w:t>First step</w:t></w:r></w:p>';
  if (!numberingXml) return docxPackage(body);
  return docxPackage(body, {
    parts: { 'word/numbering.xml': toBytes(numberingXml) },
    rels: '<Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>',
    contentTypes:
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>',
  });
}

/** Presses the bullet or number button with the caret in paragraph `paraId`, as PagedEditor does. */
function press(session: YrsSession, paraId: string, action: 'bulletList' | 'numberedList') {
  session.setSelection({ story: 'body', paraId, offset: 2 });
  const map = createYrsInputPositionMap('body', session.paragraphSpans('body'));
  const source = session.materializeDocx()?.package.numbering;
  expect(applyYrsToolbarFormatting(session, map, action, () => ({}), undefined, source)).toBe(true);
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

for (const [name, numberingXml] of [
  ['no numbering part', undefined],
  ['numId 1 numbers and 2 bullets', numbering(DECIMAL, BULLET)],
  ['numId 1 bullets and 2 numbers', numbering(BULLET, DECIMAL)],
] as const) {
  test(`toolbar bullets and numbers save and reopen as shown (${name})`, async () => {
    const session = await openSession(listDocx(numberingXml), 7601);
    press(session, '10000001', 'bulletList');
    press(session, '10000002', 'numberedList');
    expect(lists(session)).toEqual(['bullet', 'decimal']);
    const reopened = await openSession(await savedBytes(session), 7602);
    expect(lists(reopened)).toEqual(['bullet', 'decimal']);
    // Each button turns its own kind of list off again.
    press(reopened, '10000001', 'bulletList');
    press(reopened, '10000002', 'numberedList');
    expect(lists(reopened)).toEqual(['none', 'none']);
  });
}
