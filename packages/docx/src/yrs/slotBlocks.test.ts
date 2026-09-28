import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { buildResidentRegionLayoutRequest } from '../editor/computeLayout';
import type { ResidentFontRequirement } from '../layout/measure';
import { decodeFrameDelta } from '../layout/render/frameDelta';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm } from '../wasm/opc';
import { preloadParseWasm } from '../wasm/parse';
import { createYrsSession, type YrsSession } from './index';

// Text never goes ahead of the tables or page breaks that open a paragraph
// slot: the render bridge refuses that state, and Word never produces it.

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FONT = resolve(
  import.meta.dir,
  '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
);
const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument';
const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.wordprocessingml.document.main+xml"/></Types>`;
const PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPkg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;

function fixture(body: string): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes(CONTENT_TYPES));
  parts.set('_rels/.rels', toBytes(PACKAGE_RELS));
  parts.set(
    'word/document.xml',
    toBytes(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`
    )
  );
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}
const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const PAGE_BREAK = `${paragraph('Before')}<w:p><w:r><w:br w:type="page"/></w:r></w:p>${paragraph(
  'Chapter'
)}`;
const TBL = `<w:tbl><w:tr><w:tc>${paragraph('cell')}</w:tc></w:tr></w:tbl>`;
const TABLE = `${paragraph('Before')}${TBL}${paragraph('After')}`;
const SUGGESTING = { name: 'Bob', date: '2026-09-29T00:00:00Z' };

let clientId = 69100;
async function open(body: string): Promise<YrsSession> {
  const session = await createYrsSession({ clientId: (clientId += 1) });
  session.openDocx(fixture(body), true);
  return session;
}
const units = (session: YrsSession): string =>
  session
    .storySegments('body')
    .map((segment) =>
      segment.kind === 'text'
        ? segment.text
        : segment.kind === 'pilcrow'
        ? '¶'
        : `[${segment.embedKind}]`
    )
    .join('');
const paraId = (session: YrsSession, index: number) => session.paragraphs('body')[index]!.paraId;
const caretAt = (paraId: string, offset: number) => ({ story: 'body', paraId, offset });

describe('paragraph slots that open with a table or page break', () => {
  beforeAll(() =>
    Promise.all([
      preloadEditWasm(new Uint8Array(readFileSync(WASM))),
      preloadOpcWasm(),
      preloadParseWasm(),
    ])
  );

  it.each([
    ['Delete at the end of the paragraph before', 1, 'forward'],
    ['Backspace at the start of the paragraph opened by', 2, 'backward'],
  ] as const)('%s a page break removes the break', async (_, index, direction) => {
    const session = await open(PAGE_BREAK);
    try {
      expect(units(session)).toBe('Before¶¶[pageBreak]Chapter¶');
      const target = paraId(session, index);
      const { caret } = session.mergeParagraphs('body', target, direction);
      expect(units(session)).toBe('Before¶¶Chapter¶');
      expect(caret).toEqual(caretAt(target, 0));
      expect(session.undo()).toBe(true);
      expect(units(session)).toBe('Before¶¶[pageBreak]Chapter¶');
    } finally {
      session.destroy();
    }
  });

  it('Delete or Backspace next to a table changes nothing when the paragraph before it has content', async () => {
    const session = await open(TABLE);
    try {
      const [first, second] = [paraId(session, 0), paraId(session, 1)];
      const before = session.encodeState();
      expect(session.mergeParagraphs('body', first, 'forward').caret).toEqual(caretAt(first, 6));
      expect(session.mergeParagraphs('body', second, 'backward', SUGGESTING).caret).toEqual(
        caretAt(second, 0)
      );
      expect(units(session)).toBe('Before¶[table]After¶');
      expect(session.encodeState()).toEqual(before);
    } finally {
      session.destroy();
    }
  });

  it('Delete in an empty paragraph before a table removes it, and suggesting marks it', async () => {
    const session = await open(`${paragraph('Before')}<w:p/>${TBL}${paragraph('After')}`);
    try {
      expect(units(session)).toBe('Before¶¶[table]After¶');
      const [empty, slot] = [paraId(session, 1), paraId(session, 2)];
      const plain = session.mergeParagraphs('body', empty, 'forward');
      expect(units(session)).toBe('Before¶[table]After¶');
      expect(plain.caret).toEqual(caretAt(slot, 0));
      expect(session.undo()).toBe(true);
      // Backspace in suggesting mode leaves the caret at the marked paragraph.
      const suggested = session.mergeParagraphs('body', slot, 'backward', SUGGESTING);
      expect(suggested.revisionId).not.toBeNull();
      expect(suggested.caret).toEqual(caretAt(empty, 0));
      expect(units(session)).toBe('Before¶¶[table]After¶');
    } finally {
      session.destroy();
    }
  });

  it('Delete or Backspace between two tables never joins them', async () => {
    const session = await open(`${TBL}<w:p/>${TBL}${paragraph('After')}`);
    try {
      expect(units(session)).toBe('[table]¶[table]After¶');
      const before = session.encodeState();
      session.mergeParagraphs('body', paraId(session, 0), 'forward');
      session.mergeParagraphs('body', paraId(session, 1), 'backward');
      expect(session.encodeState()).toEqual(before);
    } finally {
      session.destroy();
    }
  });

  it.each([
    ['page break', PAGE_BREAK, 2, 'Before¶¶[pageBreak]Typed Chapter¶'],
    ['table', TABLE, 1, 'Before¶[table]Typed After¶'],
  ])(
    'an insert at offset 0 of a slot that opens with a %s lands after it',
    async (_, body, index, typed) => {
      const session = await open(body);
      try {
        const target = paraId(session, index);
        const receipt = session.insertText({ story: 'body', paraId: target, offset: 0 }, 'Typed ');
        expect(units(session)).toBe(typed);
        expect(receipt.range).toEqual({
          story: 'body',
          start: { paraId: target, offset: 1 },
          end: { paraId: target, offset: 7 },
        });
      } finally {
        session.destroy();
      }
    }
  );

  it('a selection delete or cut ending at a table slot keeps the mark before it, and Undo restores the text', async () => {
    const session = await open(TABLE);
    try {
      // The editor's selection delete and cut both delete the selected range.
      session.deleteRange({
        story: 'body',
        start: { paraId: paraId(session, 0), offset: 3 },
        end: { paraId: paraId(session, 1), offset: 0 },
      });
      expect(units(session)).toBe('Bef¶[table]After¶');
      expect(session.undo()).toBe(true);
      expect(units(session)).toBe('Before¶[table]After¶');
    } finally {
      session.destroy();
    }
  });

  it('typing at a resident caret ahead of a page break lands after it and keeps its order', async () => {
    const session = await open(PAGE_BREAK);
    try {
      const document = session.materializeDocx()!;
      const fontId = session.registerFont(new Uint8Array(readFileSync(FONT)));
      const request = buildResidentRegionLayoutRequest(document, 24, {});
      const requirements = JSON.parse(
        session.layoutFontRequirementsJson(JSON.stringify(request))
      ) as ResidentFontRequirement[];
      request.measurement = {
        fontChains: Object.fromEntries(
          requirements.map((requirement) => [requirement.key, [fontId]])
        ),
        defaults: { fontSize: 11, fontFamily: 'Calibri' },
        compat: { noLeading: false, doNotExpandShiftReturn: false },
        authoritativeShaping: true,
      };
      session.layoutDocumentWithRegionsRetainedJson(JSON.stringify(request));
      session.setSelection({ story: 'body', paraId: paraId(session, 2), offset: 0 });
      let epoch = decodeFrameDelta(session.buildDisplayListFrame('{}', 0)).frameEpoch;
      for (const key of ['a', 'b'])
        epoch = decodeFrameDelta(session.applyInput(key, epoch)).frameEpoch;
      expect(units(session)).toBe('Before¶¶[pageBreak]abChapter¶');
    } finally {
      session.destroy();
    }
  });
});
