import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import type { Document, Paragraph } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');

function fixture(): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set(
    '[Content_Types].xml',
    toBytes(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    )
  );
  parts.set(
    '_rels/.rels',
    toBytes(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    )
  );
  // Word stamps the first paragraph; the other two carry no w14:paraId.
  parts.set(
    'word/document.xml',
    toBytes(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="1A2B3C4D"><w:r><w:t>Alpha</w:t></w:r></w:p><w:p><w:r><w:t>Beta gamma</w:t></w:r></w:p><w:p><w:r><w:t>Delta</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>'
    )
  );
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

const paraIds = (document: Document) =>
  document.package.document.content
    .filter((block): block is Paragraph => block.type === 'paragraph')
    .map((paragraph) => paragraph.paraId);

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

test('editor paragraph ids save as unique hex w14:paraIds and source ids stay', async () => {
  const bytes = fixture();
  const parsed = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });
  const session = await createYrsSession({ clientId: 2001 });
  let saved: Document;
  let again: Document;
  try {
    session.seedFromDocx(bytes);
    const [, beta] = session.paragraphs('body');
    // The second half takes an editor id, and "Delta" keeps body:p2 one slot later.
    session.splitParagraph({ story: 'body', paraId: beta.paraId, offset: 4 });
    expect(session.paragraphs('body').map((p) => p.paraId)).toEqual([
      '1A2B3C4D',
      'body:p1',
      '2001:0',
      'body:p2',
    ]);
    saved = yrsToDocument(session, parsed);
    again = yrsToDocument(session, parsed, { onStory: () => {} });
  } finally {
    session.destroy();
  }

  const ids = paraIds(saved);
  // A seeded id still in its own slot saves without one, as the source had none.
  expect(ids.slice(0, 2)).toEqual(['1A2B3C4D', undefined]);
  const minted = ids.slice(2) as string[];
  for (const id of minted) {
    expect(id).toMatch(/^[0-9A-F]{8}$/);
    expect(parseInt(id, 16)).toBeGreaterThan(0);
    expect(parseInt(id, 16)).toBeLessThan(0x7fffffff);
  }
  expect(new Set(['1A2B3C4D', ...minted]).size).toBe(3);
  expect(paraIds(again)).toEqual(ids);

  // The parser keeps every saved id instead of re-minting it.
  const reopened = await parseDocx(await repackDocx(saved), { preloadFonts: false });
  expect(paraIds(reopened)).toEqual(ids);
});
