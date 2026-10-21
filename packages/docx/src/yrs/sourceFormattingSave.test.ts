import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import type { Document, Paragraph, Run, Table } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { documentToYrs } from './documentToYrs';
import { createYrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';

// A save writes back only what an edit changed (Epo 2026-10-21): runs the
// session left keep their source formatting and cells their own borders and
// margins, so styles keep applying in Word. office-service's
// tests/export_source_formatting.rs checks the native save the same way.

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))
  )
);

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const LINE = '<w:VAL w:val="single" w:sz="4" w:space="0" w:color="000000"/>';
const borders = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
  .map((side) => LINE.replace('VAL', side))
  .join('');
const cell = (text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;

/** A red bold heading, a plain paragraph with an italic run, and a table styled with borders and margins. */
async function styled(): Promise<Document> {
  const parts: PartsMap = new Map();
  parts.set(
    '[Content_Types].xml',
    toBytes(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${OFFICE}.styles+xml"/></Types>`
    )
  );
  parts.set(
    '_rels/.rels',
    toBytes(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`
    )
  );
  parts.set(
    'word/styles.xml',
    toBytes(
      `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:color w:val="FF0000"/><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="table" w:styleId="Grid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders>${borders}</w:tblBorders><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>`
    )
  );
  parts.set(
    'word/document.xml',
    toBytes(
      `<w:document xmlns:w="${W}"><w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">Plain words </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>slanted</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblStyle w:val="Grid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr>${cell('A')}${cell('B')}</w:tr><w:tr>${cell('C')}${cell('D')}</w:tr></w:tbl><w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`
    )
  );
  return parseDocx(rezipPartsToArrayBuffer(parts), { preloadFonts: false });
}

const paragraphs = (document: Document) =>
  document.package.document.content.filter((block): block is Paragraph => block.type === 'paragraph');
const runs = (paragraph: Paragraph) =>
  paragraph.content.filter((child): child is Run => child.type === 'run');
const cells = (document: Document) =>
  (document.package.document.content.find((block) => block.type === 'table') as Table).rows.flatMap(
    (row) => row.cells
  );

test('runs and cells the session left save what the source wrote', async () => {
  const source = await styled();
  const session = await createYrsSession({ clientId: 61021 });
  try {
    documentToYrs(session, source);
    const saved = yrsToDocument(session, source);
    paragraphs(saved).forEach((paragraph, index) => {
      expect(runs(paragraph).map((run) => run.formatting)).toEqual(
        runs(paragraphs(source)[index]!).map((run) => run.formatting)
      );
    });
    expect(cells(saved).map((cell) => cell.formatting)).toEqual(
      cells(source).map((cell) => cell.formatting)
    );
  } finally {
    session.destroy();
  }
});

test('an edited run and cell save their edits', async () => {
  const source = await styled();
  const session = await createYrsSession({ clientId: 61022 });
  try {
    documentToYrs(session, source);
    const { paraId } = session.paragraphs('body')[1]!;
    session.formatRange(
      { story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 5 } },
      { bold: true }
    );
    const thick = { style: 'double', size: 12, color: { rgb: '00FF00' } };
    const corner = { story: 'body', tableIndex: 0, row: 0, column: 0 };
    session.setCellBorders({ anchor: corner, head: corner }, { top: thick });
    const saved = yrsToDocument(session, source);
    const [plain, words, slanted] = runs(paragraphs(saved)[1]!);
    expect(plain!.formatting?.bold).toBe(true);
    expect(words!.formatting).toBeUndefined();
    expect(slanted!.formatting).toEqual(runs(paragraphs(source)[1]!)[1]!.formatting);
    const [first, ...others] = cells(saved);
    expect(Object.keys(first!.formatting?.borders ?? {})).toEqual(['top']);
    expect(first!.formatting?.borders?.top).toMatchObject({ style: 'double' });
    for (const other of others) expect(other.formatting?.borders).toBeUndefined();
  } finally {
    session.destroy();
  }
});
