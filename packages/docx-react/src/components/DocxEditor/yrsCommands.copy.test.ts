import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '@betteroffice/docx/docx/rezip/parts';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession, type YrsSession } from '@betteroffice/docx/yrs';
import { yrsSelectionText } from './yrsCommands';

const WASM = resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm');
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

function docx(body: string): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set(
    '[Content_Types].xml',
    toBytes(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        `<Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/></Types>`
    )
  );
  parts.set(
    '_rels/.rels',
    toBytes(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    )
  );
  parts.set('word/document.xml', toBytes(`<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr/></w:body></w:document>`));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

const p = (inner: string) => `<w:p>${inner}</w:p>`;
const t = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
const br = (kind: string) => `<w:r><w:br w:type="${kind}"/></w:r>`;

async function open(body: string): Promise<YrsSession> {
  const session = await createYrsSession({ clientId: 31 });
  session.openDocx(docx(body), true);
  return session;
}

/** The copy of the selection from `[paragraph, offset]` to `[paragraph, offset]`, in story units. */
function copy(session: YrsSession, [from, start]: [number, number], [to, end]: [number, number]) {
  const paragraphs = session.paragraphs('body');
  session.setSelection(
    { story: 'body', paraId: paragraphs[from]!.paraId, offset: start },
    { story: 'body', paraId: paragraphs[to]!.paraId, offset: end }
  );
  return yrsSelectionText(session);
}

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

// As view mode and LibreOffice's text export: a newline where the text continues after the break.
describe('edit-mode copy over a page or column break inside a paragraph', () => {
  test.each(['page', 'column'])('writes a newline at a %s break and stays not plain', async (kind) => {
    const session = await open(p(t('Lead')) + p(t('Aa') + br(kind) + t('Bb')) + p(t('Cc')));
    try {
      // Units: A a [break] B b.
      expect(copy(session, [0, 0], [2, 2])).toEqual({ text: 'Lead\nAa\nBb\nCc', plain: false });
      expect(copy(session, [1, 1], [1, 4])).toEqual({ text: 'a\nB', plain: false });
      expect(copy(session, [1, 3], [1, 5])).toEqual({ text: 'Bb', plain: true });
    } finally {
      session.destroy();
    }
  });

  test('writes one newline at a break opening its paragraph', async () => {
    const session = await open(p(t('Lead')) + p(br('page') + t('Bb')));
    try {
      const paragraphs = session.paragraphs('body');
      const end = session.paragraphSpans('body').find((span) => span.paraId === paragraphs[1]!.paraId)!.length;
      expect(copy(session, [0, 0], [1, end])).toEqual({ text: 'Lead\nBb', plain: false });
    } finally {
      session.destroy();
    }
  });
});
