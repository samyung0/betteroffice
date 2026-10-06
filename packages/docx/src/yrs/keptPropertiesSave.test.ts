import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { unzipContainer } from '../wasm/opc';
import type { Document } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { createStyleResolver } from '../styles';
import {
  applyStyleValues,
  cellParagraphFormatting,
  createYrsSession,
  applyNextStyle,
  endEmptyListItem,
  inOneUndoStep,
  styleNewCells,
  styleParagraphValues,
  type ParagraphStyleValues,
  type YrsSession,
} from './index';
import { yrsToDocument } from './yrsToDocument';

// What Word reads that the editor's model does not hold has to come back on
// save as the source wrote it. The oracle is the source XML.

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

const STYLES =
  `<w:styles xmlns:w="${W}"><w:docDefaults><w:pPrDefault><w:pPr>` +
  '<w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>' +
  '<w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr><w:contextualSpacing/></w:pPr></w:style>' +
  '<w:style w:type="table" w:styleId="Grid"><w:name w:val="Table Grid"/>' +
  '<w:tblStylePr w:type="firstRow"><w:pPr><w:jc w:val="center"/></w:pPr></w:tblStylePr></w:style>' +
  '<w:style w:type="table" w:styleId="Edged"><w:name w:val="Edged"/>' +
  '<w:tblStylePr w:type="firstRow"><w:pPr><w:jc w:val="center"/></w:pPr></w:tblStylePr>' +
  '<w:tblStylePr w:type="lastRow"><w:pPr><w:spacing w:before="120"/></w:pPr></w:tblStylePr>' +
  '<w:tblStylePr w:type="lastCol"><w:pPr><w:jc w:val="right"/></w:pPr></w:tblStylePr></w:style></w:styles>';

const NUMBERING =
  `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/>` +
  '<w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>' +
  '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>' +
  '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>';

function fixture(body: string): Uint8Array<ArrayBuffer> {
  const parts = new Map<string, Uint8Array>();
  parts.set(
    '[Content_Types].xml',
    toBytes(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        `<Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/>` +
        `<Override PartName="/word/styles.xml" ContentType="${OFFICE}.styles+xml"/>` +
        `<Override PartName="/word/numbering.xml" ContentType="${OFFICE}.numbering+xml"/></Types>`
    )
  );
  parts.set(
    '_rels/.rels',
    toBytes(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    )
  );
  parts.set(
    'word/_rels/document.xml.rels',
    toBytes(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/>` +
        `<Relationship Id="numbering" Type="${R}/numbering" Target="numbering.xml"/></Relationships>`
    )
  );
  parts.set(
    'word/document.xml',
    toBytes(`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr/></w:body></w:document>`)
  );
  parts.set('word/styles.xml', toBytes(STYLES));
  parts.set('word/numbering.xml', toBytes(NUMBERING));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

/** As Capy opens a file: the native seed, with the base materialized from the source. */
async function open(bytes: Uint8Array, clientId = 7) {
  const session = await createYrsSession({ clientId });
  session.openDocx(bytes, true);
  const base = session.materializeDocx()!;
  return { session, base };
}

async function savedDocumentXml(session: YrsSession, base: Document): Promise<string> {
  const bytes = new Uint8Array(await repackDocx(yrsToDocument(session, base)));
  const files = unzipContainer(bytes) as Record<string, Uint8Array>;
  return new TextDecoder().decode(files['word/document.xml']);
}

/** Each saved `w:p`'s pPr, in document order ('' when it has none). */
function pPrs(xml: string): string[] {
  return [...xml.matchAll(/<w:p(?: [^>]*)?>(.*?)<\/w:p>|<w:p(?: [^>]*)?\/>/g)].map(
    (match) => match[1]?.match(/^<w:pPr>.*<\/w:pPr>/)?.[0] ?? ''
  );
}

/** Style values as the editor reads them (PagedEditor's `paragraphStyleValues`, the materialized source as its cells' source). */
function styleValuesFor(session: YrsSession, base: Document): ParagraphStyleValues {
  const resolver = createStyleResolver(base.package.styles!);
  return (styleId, story, list = true) =>
    styleParagraphValues(resolver, styleId, {
      cell: cellParagraphFormatting(session, resolver, story, base),
      numbering: list ? base.package.numbering : undefined,
      numPr: typeof list === 'object' ? list : undefined,
    });
}

function at(session: YrsSession, story: string, index: number, offset = 0) {
  return { story, paraId: session.paragraphs(story)[index]!.paraId, offset };
}

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))
  )
);

describe('pPr children the model has no field for', () => {
  // Every CT_PPrBase child the model lacks, in schema order between modeled ones.
  const KEPT =
    '<w:pPr><w:suppressAutoHyphens/><w:kinsoku w:val="0"/><w:wordWrap w:val="0"/>' +
    '<w:overflowPunct w:val="0"/><w:topLinePunct/><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/>' +
    '<w:adjustRightInd w:val="0"/><w:snapToGrid w:val="0"/><w:spacing w:after="0"/><w:ind w:left="100"/>' +
    '<w:mirrorIndents/><w:suppressOverlap/><w:jc w:val="both"/><w:textDirection w:val="tbRl"/>' +
    '<w:textAlignment w:val="center"/><w:textboxTightWrap w:val="allLines"/><w:outlineLvl w:val="8"/>' +
    '<w:divId w:val="123"/><w:cnfStyle w:val="100000000000"/></w:pPr>';
  const FRAME =
    '<w:pPr><w:framePr w:hAnchor="text" w:vAnchor="text" w:wrap="around" w:dropCap="drop" w:lines="3" ' +
    'w:hSpace="72"/></w:pPr>';
  const CELL = '<w:pPr><w:kinsoku w:val="0"/><w:cnfStyle w:val="101000000000"/></w:pPr>';
  const bytes = fixture(
    `<w:p>${KEPT}<w:r><w:t>Hello world</w:t></w:r></w:p>` +
      `<w:p>${FRAME}<w:r><w:t>D</w:t></w:r></w:p>` +
      '<w:p><w:r><w:t>plain</w:t></w:r></w:p>' +
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>' +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p>${CELL}<w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
      '<w:p/>'
  );

  it('come back in schema order on an untouched paragraph, a framed one and a cell', async () => {
    const { session, base } = await open(bytes);
    session.insertText(at(session, 'body', 2), 'X');
    const cell = session.storyIds().find((id) => id.startsWith('body:t'))!;
    session.insertText(at(session, cell, 0), 'z');
    expect(pPrs(await savedDocumentXml(session, base))).toEqual([KEPT, FRAME, '', CELL, '']);
    session.destroy();
  });

  it('stay on a paragraph the editor changed and on both halves of a split', async () => {
    const { session, base } = await open(bytes);
    const { paraId } = session.paragraphs('body')[0]!;
    session.setParagraphAttrs(
      { story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 0 } },
      { alignment: 'center' }
    );
    session.insertText(at(session, 'body', 0, 5), 'X');
    session.splitParagraph(at(session, 'body', 0, 3));
    const centred = KEPT.replace('<w:jc w:val="both"/>', '<w:jc w:val="center"/>');
    expect(pPrs(await savedDocumentXml(session, base)).slice(0, 3)).toEqual([centred, centred, FRAME]);
    session.destroy();
  });
});

describe('a vertically merged continuation cell', () => {
  const CONTINUED = '<w:pPr><w:spacing w:after="0"/><w:rPr><w:sz w:val="16"/></w:rPr></w:pPr>';
  const cell = (merge: string, content: string) =>
    `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>${merge}</w:tcPr>${content}</w:tc>`;
  const bytes = fixture(
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
      `<w:tr>${cell('<w:vMerge w:val="restart"/>', '<w:p><w:r><w:t>Merged</w:t></w:r></w:p>')}${cell('', '<w:p><w:r><w:t>A</w:t></w:r></w:p>')}</w:tr>` +
      `<w:tr>${cell('<w:vMerge/>', `<w:p>${CONTINUED}</w:p>`)}${cell('', '<w:p><w:r><w:t>B</w:t></w:r></w:p>')}</w:tr>` +
      '</w:tbl><w:p/>'
  );

  it('keeps its paragraph properties, also after rows move', async () => {
    const { session, base } = await open(bytes);
    // The cell after the continuation is the row's first story, as the seed numbers it.
    expect(session.paragraphs('body:t0:r1c0')[0]!.text).toBe('B');
    session.insertText(at(session, 'body:t0:r1c0', 0), 'X');
    expect(pPrs(await savedDocumentXml(session, base))).toEqual(['', '', CONTINUED, '', '']);
    session.insertRow({ story: 'body', tableIndex: 0, row: 0, column: 1 }, 'above');
    expect(pPrs(await savedDocumentXml(session, base)).slice(2)).toEqual(['', '', CONTINUED, '', '']);
    session.destroy();
  });
});

/** Sends each peer what the other has. */
function sync(left: YrsSession, right: YrsSession): void {
  left.applyUpdate(right.encodeStateAsUpdate(left.encodeStateVector()));
  right.applyUpdate(left.encodeStateAsUpdate(right.encodeStateVector()));
}

const strip = (pPr: string) => pPr.replace(/ w:(author|date)="[^"]*"/g, '');
const revisionIds = (pPr: string) => [...pPr.matchAll(/<w:(?:pPrChange|ins|del) w:id="(\d+)"/g)].map((m) => m[1]);

describe('Enter in a paragraph with a tracked revision', () => {
  const BY = 'w:author="Rev" w:date="2026-01-01T00:00:00Z"';
  const CHANGE = `<w:pPrChange w:id="5" ${BY}><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange>`;
  const bytes = fixture(
    `<w:p><w:pPr><w:jc w:val="center"/>${CHANGE}</w:pPr><w:r><w:t>Changed para</w:t></w:r></w:p>` +
      `<w:p><w:pPr><w:rPr><w:ins w:id="7" ${BY}/></w:rPr></w:pPr><w:r><w:t>Inserted mark</w:t></w:r></w:p>` +
      `<w:p><w:pPr><w:rPr><w:del w:id="8" ${BY}/></w:rPr></w:pPr><w:r><w:t>Deleted mark</w:t></w:r></w:p>` +
      '<w:p/>'
  );

  it('leaves the mark revision on the source mark and gives the copied change its own id', async () => {
    const { session, base } = await open(bytes);
    for (const index of [2, 1, 0]) session.splitParagraph(at(session, 'body', index, 4));
    const saved = pPrs(await savedDocumentXml(session, base)).map(strip);
    expect(saved.slice(2)).toEqual([
      '',
      '<w:pPr><w:rPr><w:ins w:id="7"/></w:rPr></w:pPr>',
      '',
      '<w:pPr><w:rPr><w:del w:id="8"/></w:rPr></w:pPr>',
      '',
    ]);
    const [copy, source] = saved.slice(0, 2);
    expect(source).toBe(`<w:pPr><w:jc w:val="center"/>${strip(CHANGE)}</w:pPr>`);
    // The copy saves in the upper half of int32, above the small ids Word writes.
    const [id] = revisionIds(copy!).map(Number);
    expect(id).toBeGreaterThanOrEqual(2 ** 30);
    expect(id).toBeLessThan(2 ** 31);
    expect(copy).toBe(source!.replace('w:id="5"', `w:id="${id}"`));
    session.destroy();
  });

  it('gives each peer’s split of one paragraph its own revision id', async () => {
    const left = (await open(bytes, 81)).session;
    const { session: right, base } = await open(bytes, 82);
    left.splitParagraph(at(left, 'body', 0, 3));
    right.splitParagraph(at(right, 'body', 0, 8));
    sync(left, right);
    const saved = await savedDocumentXml(left, base);
    expect(await savedDocumentXml(right, base)).toBe(saved);
    const ids = pPrs(saved).slice(0, 3).flatMap(revisionIds).map(Number);
    expect(new Set(ids).size).toBe(3);
    expect(ids.filter((id) => id >= 2 ** 30)).toHaveLength(2);
    left.destroy();
    right.destroy();
  });
});

describe('table rows and new cells', () => {
  const tc = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const SKIPPED = '<w:trPr><w:gridBefore w:val="1"/><w:wBefore w:w="1500" w:type="dxa"/></w:trPr>';
  const bytes = fixture(
    '<w:p><w:r><w:t>x</w:t></w:r></w:p>' +
      '<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0" w:firstRow="1"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>' +
      `<w:tr>${tc('A0')}${tc('A1')}${tc('A2')}</w:tr><w:tr>${SKIPPED}${tc('D1')}${tc('D2')}</w:tr></w:tbl><w:p/>`
  );
  const rows = (xml: string) => [...xml.matchAll(/<w:tr>.*?<\/w:tr>/g)].map((m) => m[0]);

  it('keeps the grid columns a row skips while they fit the grid', async () => {
    const { session, base } = await open(bytes);
    session.insertText(at(session, 'body', 0), 'y');
    expect(rows(await savedDocumentXml(session, base))[1]).toStartWith(`<w:tr>${SKIPPED}<w:tc>`);
    // A row added next to it has every column, so it skips none.
    session.insertRow({ story: 'body', tableIndex: 0, row: 1, column: 0 }, 'below');
    const added = rows(await savedDocumentXml(session, base));
    expect(added[1]).toStartWith(`<w:tr>${SKIPPED}<w:tc>`);
    expect(added[2]).not.toContain('<w:trPr>');
    expect(added[2]!.match(/<w:tc>/g)).toHaveLength(3);
    session.destroy();
  });

  it('drops the skipped grid columns of a row a column deletion leaves too wide', async () => {
    const { session, base } = await open(bytes);
    const last = { story: 'body', tableIndex: 0, row: 0, column: 2 };
    session.deleteColumn({ anchor: last, head: last });
    const saved = rows(await savedDocumentXml(session, base));
    expect(saved[1]).not.toContain('w:gridBefore');
    expect(saved[1]!.match(/<w:tc>/g)).toHaveLength(2);
    session.destroy();
  });

  it('gives a new row’s cells no alignment of their own, for each peer', async () => {
    const left = (await open(bytes, 81)).session;
    const { session: right, base } = await open(bytes, 82);
    left.insertRow({ story: 'body', tableIndex: 0, row: 0, column: 0 }, 'above');
    sync(left, right);
    const saved = await savedDocumentXml(right, base);
    expect(await savedDocumentXml(left, base)).toBe(saved);
    expect(rows(saved)).toHaveLength(3);
    expect(rows(saved)[0]).not.toContain('<w:jc ');
    left.destroy();
    right.destroy();
  });
});

describe('continuation cells after rows and tables change', () => {
  const cell = (merge: string, content: string) =>
    `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>${merge}</w:tcPr>${content}</w:tc>`;
  const merged = (rows: ReadonlyArray<readonly [string, string, string]>) =>
    fixture(
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
        rows.map(([merge, content, text]) => `<w:tr>${cell(merge, content)}${cell('', `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`)}</w:tr>`).join('') +
        '</w:tbl><w:p/>'
    );
  const after = (n: number) => `<w:p><w:pPr><w:spacing w:after="${n}"/></w:pPr></w:p>`;
  const RESTART = '<w:vMerge w:val="restart"/>';
  const CONTINUE = '<w:vMerge/>';
  /** Each saved row as `text:continuation spacing after`. */
  const rows = (xml: string) =>
    [...xml.matchAll(/<w:tr>.*?<\/w:tr>/g)].map(
      ([row]) =>
        `${[...row.matchAll(/<w:t>([^<]*)<\/w:t>/g)].map((text) => text[1]).join('')}:${row.match(/w:after="(\d+)"/)?.[1] ?? '-'}`
    );

  it('follow their own rows when a row inside the merge goes', async () => {
    const bytes = merged([
      [RESTART, '<w:p><w:r><w:t>M</w:t></w:r></w:p>', 'A'],
      [CONTINUE, after(101), 'B'],
      [CONTINUE, after(102), 'C'],
      [CONTINUE, after(103), 'D'],
    ]);
    const { session, base } = await open(bytes);
    const c = { story: 'body', tableIndex: 0, row: 2, column: 1 };
    session.deleteRow({ anchor: c, head: c });
    expect(rows(await savedDocumentXml(session, base))).toEqual(['MA:-', 'B:101', 'D:103']);
    session.destroy();
  });

  it('keep to their own merge when rows go into another one in the column', async () => {
    const bytes = merged([
      [RESTART, '<w:p><w:r><w:t>M</w:t></w:r></w:p>', 'A'],
      [CONTINUE, after(201), 'B'],
      [RESTART, '<w:p><w:r><w:t>N</w:t></w:r></w:p>', 'C'],
      [CONTINUE, after(301), 'D'],
    ]);
    const { session, base } = await open(bytes);
    session.insertRow({ story: 'body', tableIndex: 0, row: 0, column: 1 }, 'below');
    expect(rows(await savedDocumentXml(session, base))).toEqual(['MA:-', ':-', 'B:201', 'NC:-', 'D:301']);
    session.destroy();
  });

  it('take nothing from a deleted table whose story ids a new table reuses', async () => {
    const bytes = merged([
      [RESTART, '<w:p><w:r><w:t>M</w:t></w:r></w:p>', 'A'],
      [CONTINUE, after(777), 'B'],
    ]);
    const { session, base } = await open(bytes);
    session.deleteTable({ story: 'body', tableIndex: 0 });
    session.insertTable(at(session, 'body', 0), 2, 2);
    const a = { story: 'body', tableIndex: 0, row: 0, column: 0 };
    const b = { story: 'body', tableIndex: 0, row: 1, column: 0 };
    session.mergeCells({ anchor: a, head: b });
    expect(await savedDocumentXml(session, base)).not.toContain('w:after="777"');
    session.destroy();
  });
});

describe('a new row in a table with a header row style', () => {
  const tc = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const bytes = fixture(
    '<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0" w:firstRow="1"/></w:tblPr>' +
      `<w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid><w:tr>${tc('H0')}${tc('H1')}</w:tr><w:tr>${tc('B0')}${tc('B1')}</w:tr></w:tbl><w:p/>`
  );

  it('shows its header cells centred, as the file and Word do', async () => {
    const { session, base } = await open(bytes);
    const values = styleValuesFor(session, base);
    const receipt = session.insertRow({ story: 'body', tableIndex: 0, row: 0, column: 0 }, 'above');
    styleNewCells(session, receipt.createdStoryIds, values);
    const added = receipt.createdStoryIds.map((story) => session.paragraphs(story)[0]!.properties.alignment);
    expect(added).toEqual(['center', 'center']);
    const xml = await savedDocumentXml(session, base);
    expect(xml.match(/<w:tr>.*?<\/w:tr>/)![0]).not.toContain('<w:jc ');
    session.destroy();
  });
});

describe('cells the table had when it opened, after rows and columns move', () => {
  // Header row centred, last row 120 before, last column right: what the seed
  // gave each cell where it was, which the editor keeps showing after a move.
  const tc = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const bytes = fixture(
    '<w:tbl><w:tblPr><w:tblStyle w:val="Edged"/><w:tblW w:w="0" w:type="auto"/>' +
      '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="1" w:firstColumn="0" w:lastColumn="1" w:noHBand="1" w:noVBand="1"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>' +
      `<w:tr>${tc('H0')}${tc('H1')}</w:tr><w:tr>${tc('B0')}${tc('B1')}</w:tr><w:tr>${tc('L0')}${tc('L1')}</w:tr></w:tbl><w:p/>`
  );
  const SOURCE = ['H0', 'H1', 'B0', 'B1', 'L0', 'L1'];
  const NEW_CELL = '<w:pPr><w:pStyle w:val="Normal"/></w:pPr>';
  /** Each saved paragraph's text and pPr ('' for none); a new cell's text is ''. */
  const savedCells = (xml: string) =>
    [...xml.matchAll(/<w:p(?: [^>]*)?>(.*?)<\/w:p>|<w:p(?: [^>]*)?\/>/g)].map(
      (match) =>
        [
          match[1]?.match(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/)?.[1] ?? '',
          match[1]?.match(/^<w:pPr>.*<\/w:pPr>/)?.[0] ?? '',
        ] as const
    );
  /** Saved paragraphs with a pPr besides a new cell's style. */
  const formatted = (xml: string) =>
    savedCells(xml).filter(([text, pPr]) => pPr !== '' && !(text === '' && pPr === NEW_CELL));
  const table = { story: 'body', tableIndex: 0 };
  const moves: Array<[string, (session: YrsSession) => readonly string[]]> = [
    ['a row inserted above the header row', (session) => session.insertRow({ ...table, row: 0, column: 0 }, 'above').createdStoryIds],
    ['a row added below the last row', (session) => session.insertRow({ ...table, row: 2, column: 0 }, 'below').createdStoryIds],
    ['a column added after the last', (session) => session.insertColumn({ ...table, row: 0, column: 1 }, 'right').createdStoryIds],
    ['the header row deleted', (session) => {
      session.deleteRow({ anchor: { ...table, row: 0, column: 0 }, head: { ...table, row: 0, column: 1 } });
      return [];
    }],
  ];

  for (const [move, edit] of moves)
    it(`save untouched as they were, so Word gives them their new place's look, after ${move}`, async () => {
      const { session, base } = await open(bytes);
      const before = new Map(session.storyIds().map((story) => [story, session.paragraphs(story)[0]?.properties]));
      const created = edit(session);
      styleNewCells(session, created, styleValuesFor(session, base));
      // The editor still shows the look each cell seeded with.
      for (const [story, properties] of before)
        if (session.storyIds().includes(story)) expect(session.paragraphs(story)[0]?.properties).toEqual(properties);
      // The source's cells save no pPr, as the source has none, and new ones
      // (styled for their new place) only their style.
      const xml = await savedDocumentXml(session, base);
      expect(formatted(xml)).toEqual([]);
      expect(savedCells(xml).map(([text]) => text).filter((text) => text !== '')).toEqual(
        SOURCE.filter((text) => !text.startsWith('H') || !move.includes('deleted'))
      );
      session.destroy();
    });

  it('saves what the user changed in such a cell against the look it seeded with', async () => {
    const { session, base } = await open(bytes);
    session.insertRow({ ...table, row: 0, column: 0 }, 'above');
    const story = session.storyIds().find((id) => session.paragraphs(id)[0]?.text === 'H0')!;
    const { paraId } = session.paragraphs(story)[0]!;
    session.setParagraphAttrs({ story, start: { paraId, offset: 0 }, end: { paraId, offset: 0 } }, { alignment: 'left' });
    const saved = Object.fromEntries(savedCells(await savedDocumentXml(session, base)));
    expect(saved.H0).toBe('<w:pPr><w:jc w:val="left"/></w:pPr>');
    expect(saved.H1).toBe('');
    session.destroy();
  });

  it('take a style picked in them by their seeded look, so the save writes only the style', async () => {
    const { session, base } = await open(bytes);
    // B0 becomes the first row, whose look centres; it seeded without.
    session.deleteRow({ anchor: { ...table, row: 0, column: 0 }, head: { ...table, row: 0, column: 1 } });
    const story = session.storyIds().find((id) => session.paragraphs(id)[0]?.text === 'B0')!;
    const { paraId } = session.paragraphs(story)[0]!;
    applyStyleValues(session, { story, start: { paraId, offset: 0 }, end: { paraId, offset: 0 } }, 'Heading1', styleValuesFor(session, base));
    expect(session.paragraphs(story)[0]!.properties.alignment ?? null).toBeNull();
    const saved = Object.fromEntries(savedCells(await savedDocumentXml(session, base)));
    expect(saved.B0).toBe('<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>');
    session.destroy();
  });

  it('save against the source on a later export whose base is the previous one', async () => {
    const { session, base } = await open(bytes);
    session.insertRow({ ...table, row: 0, column: 0 }, 'above');
    // As the editor's exporter does: each export's base is the previous projection.
    const first = yrsToDocument(session, base, { source: base });
    const story = session.storyIds().find((id) => session.paragraphs(id)[0]?.text === 'H0')!;
    session.insertText(at(session, story, 0, 2), '!');
    const second = yrsToDocument(session, first, { source: base });
    const files = unzipContainer(new Uint8Array(await repackDocx(second))) as Record<string, Uint8Array>;
    const saved = Object.fromEntries(savedCells(new TextDecoder().decode(files['word/document.xml'])));
    expect(saved['H0!']).toBe('');
    session.destroy();
  });

  it('saves the same for each peer when the other moved the rows', async () => {
    const left = (await open(bytes, 83)).session;
    const { session: right, base } = await open(bytes, 84);
    left.insertRow({ ...table, row: 2, column: 0 }, 'below');
    sync(left, right);
    const saved = await savedDocumentXml(right, base);
    expect(await savedDocumentXml(left, base)).toBe(saved);
    expect(formatted(saved)).toEqual([]);
    left.destroy();
    right.destroy();
  });
});

describe('Enter at the end of a paragraph whose style gives its list', () => {
  const bytes = fixture('<w:p><w:pPr><w:pStyle w:val="ListBullet"/></w:pPr><w:r><w:t>Bullet</w:t></w:r></w:p><w:p/>');

  it('keeps the list in the editor and saves only the style, for each peer', async () => {
    const left = (await open(bytes, 81)).session;
    const { session: right, base } = await open(bytes, 82);
    left.splitParagraph(at(left, 'body', 0, 'Bullet'.length));
    sync(left, right);
    for (const session of [left, right]) {
      const added = session.paragraphs('body')[1]!.properties;
      expect(added).toMatchObject({ pStyle: 'ListBullet', numPr: { numId: 1 }, listMarker: '•', indentLeft: 720 });
      expect(pPrs(await savedDocumentXml(session, base)).slice(0, 2)).toEqual([
        '<w:pPr><w:pStyle w:val="ListBullet"/></w:pPr>',
        '<w:pPr><w:pStyle w:val="ListBullet"/></w:pPr>',
      ]);
    }
    left.destroy();
    right.destroy();
  });
});

describe('a run’s fonts', () => {
  const run = (fonts: string, text: string) =>
    `<w:p><w:r><w:rPr><w:rFonts ${fonts}/></w:rPr><w:t>${text}</w:t></w:r><w:r><w:t> tail</w:t></w:r></w:p>`;
  const LATIN = 'w:ascii="Calibri" w:hAnsi="Calibri"';
  const EAST = 'w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="PMingLiU"';
  const THEMED =
    'w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"';
  const bytes = fixture(run(LATIN, 'latin') + run(EAST, 'east中') + run(THEMED, 'themed') + '<w:p/>');
  const fonts = (xml: string) => [...xml.matchAll(/<w:rFonts ([^>]*)\/>/g)].map((match) => match[1]);

  it('save as the source wrote them when only the text changed', async () => {
    const { session, base } = await open(bytes);
    for (const index of [0, 1, 2]) session.insertText(at(session, 'body', index, 2), 'X');
    expect(fonts(await savedDocumentXml(session, base))).toEqual([LATIN, EAST, THEMED]);
    session.destroy();
  });

  it('take a picked font in the Latin and complex-script slots and keep the East Asian one', async () => {
    const { session, base } = await open(bytes);
    for (const [index, length] of [[0, 5], [1, 5], [2, 6]] as const) {
      const { paraId } = session.paragraphs('body')[index]!;
      session.formatRange(
        { story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: length } },
        { fontFamily: { ascii: 'Arial', hAnsi: 'Arial' } }
      );
    }
    const ARIAL = 'w:ascii="Arial" w:hAnsi="Arial"';
    expect(fonts(await savedDocumentXml(session, base))).toEqual([
      `${ARIAL} w:cs="Arial"`,
      `${ARIAL} w:eastAsia="PMingLiU" w:cs="Arial"`,
      `${ARIAL} w:cs="Arial" w:eastAsiaTheme="minorEastAsia"`,
    ]);
    session.destroy();
  });
});

describe('a font picked across runs with different fonts', () => {
  const bytes = fixture(
    '<w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="PMingLiU"/></w:rPr><w:t>中一</w:t></w:r>' +
      '<w:r><w:t>plain</w:t></w:r></w:p>' +
      '<w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="PMingLiU"/></w:rPr><w:t>中二</w:t></w:r></w:p>' +
      '<w:p><w:r><w:rPr><w:rFonts w:eastAsia="SimSun"/></w:rPr><w:t>中三</w:t></w:r></w:p><w:p/>'
  );

  it('keeps each run’s East Asian font in one pick', async () => {
    const { session, base } = await open(bytes);
    const [first, , third] = session.paragraphs('body');
    session.formatRange(
      { story: 'body', start: { paraId: first!.paraId, offset: 0 }, end: { paraId: third!.paraId, offset: 2 } },
      { fontFamily: { ascii: 'Arial', hAnsi: 'Arial' } }
    );
    const xml = await savedDocumentXml(session, base);
    const runs = [...xml.matchAll(/<w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:t>([^<]*)<\/w:t><\/w:r>/g)].map((match) => [
      match[1],
      match[0].match(/<w:rFonts ([^>]*)\/>/)?.[1],
    ]);
    const ARIAL = 'w:ascii="Arial" w:hAnsi="Arial"';
    expect(runs).toEqual([
      ['中一', `${ARIAL} w:eastAsia="PMingLiU" w:cs="Arial"`],
      ['plain', `${ARIAL} w:cs="Arial"`],
      ['中二', `${ARIAL} w:eastAsia="PMingLiU" w:cs="Arial"`],
      ['中三', `${ARIAL} w:eastAsia="SimSun" w:cs="Arial"`],
    ]);
    session.destroy();
  });
});

describe('an East Asian font picked', () => {
  const bytes = fixture(
    '<w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="PMingLiU" w:eastAsiaTheme="minorEastAsia"/></w:rPr><w:t>中文</w:t></w:r></w:p><w:p/>'
  );

  it('sets the run’s East Asian font too, as Word’s font box does', async () => {
    const { session, base } = await open(bytes);
    const { paraId } = session.paragraphs('body')[0]!;
    session.formatRange(
      { story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 2 } },
      { fontFamily: { ascii: 'SimSun', hAnsi: 'SimSun' } }
    );
    expect((await savedDocumentXml(session, base)).match(/<w:rFonts [^>]*\/>/)?.[0]).toBe(
      '<w:rFonts w:ascii="SimSun" w:hAnsi="SimSun" w:eastAsia="SimSun" w:cs="SimSun"/>'
    );
    session.destroy();
  });
});

describe('a tracked deletion of a note reference or a simple field', () => {
  const BY = 'w:author="Rev" w:date="2026-01-01T00:00:00Z"';
  const NOTE = '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>';
  const FIELD = '<w:fldSimple w:instr=" PAGE "><w:r><w:delText>9</w:delText></w:r></w:fldSimple>';
  const bytes = fixture(
    `<w:p><w:r><w:t>a </w:t></w:r><w:del w:id="91" ${BY}>${NOTE}</w:del><w:r><w:t>b</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>c </w:t></w:r><w:del w:id="92" ${BY}>${FIELD}</w:del><w:r><w:t>d</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>e</w:t></w:r>${NOTE}</w:p>`
  );

  it('stays deleted, with its result and the reference style, after an edit to its paragraph', async () => {
    const { session, base } = await open(bytes);
    for (const index of [0, 1, 2]) session.insertText(at(session, 'body', index), 'Q');
    const xml = strip(await savedDocumentXml(session, base));
    expect(xml).toContain(`<w:del w:id="91">${NOTE}</w:del>`);
    expect(xml).toContain(`<w:del w:id="92">${FIELD}</w:del>`);
    expect(xml).toContain(`<w:t>Qe</w:t></w:r>${NOTE}`);
    session.destroy();
  });
});

describe('Enter in a list', () => {
  const NUMBERED = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>';
  const bytes = fixture(
    `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/>${NUMBERED}</w:pPr><w:r><w:t>One</w:t></w:r></w:p>` +
      '<w:p><w:pPr><w:pStyle w:val="ListBullet"/></w:pPr><w:r><w:t>Bul</w:t></w:r></w:p><w:p/>'
  );
  const DIRECT = `<w:pPr><w:pStyle w:val="ListParagraph"/>${NUMBERED}</w:pPr>`;
  const STYLED = '<w:pPr><w:pStyle w:val="ListBullet"/></w:pPr>';

  it('goes on at the end of an item and ends on an empty one, for direct and style lists', async () => {
    const { session, base } = await open(bytes);
    const values = styleValuesFor(session, base);
    session.splitParagraph(at(session, 'body', 1, 3));
    session.splitParagraph(at(session, 'body', 0, 3));
    expect(session.paragraphs('body').map(({ text, properties }) => [text, properties.listMarker ?? null])).toEqual([
      ['One', '•'],
      ['', '•'],
      ['Bul', '•'],
      ['', '•'],
      ['', null],
    ]);
    expect(pPrs(await savedDocumentXml(session, base)).slice(0, 4)).toEqual([DIRECT, DIRECT, STYLED, STYLED]);

    expect(endEmptyListItem(session, 'body', session.paragraphs('body')[0]!.paraId, values)).toBe(false);
    for (const index of [1, 3]) {
      expect(endEmptyListItem(session, 'body', session.paragraphs('body')[index]!.paraId, values)).toBe(true);
    }
    const [, direct, , styled] = session.paragraphs('body');
    expect([direct!.properties.listMarker ?? null, direct!.properties.indentLeft]).toEqual([null, 720]);
    expect([styled!.properties.listMarker ?? null, styled!.properties.indentLeft ?? null]).toEqual([null, null]);
    expect(pPrs(await savedDocumentXml(session, base)).slice(0, 4)).toEqual([
      DIRECT,
      '<w:pPr><w:pStyle w:val="ListParagraph"/></w:pPr>',
      STYLED,
      '<w:pPr><w:pStyle w:val="ListBullet"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr></w:pPr>',
    ]);
    session.destroy();
  });

  it('ends for both peers when one types into the empty item meanwhile', async () => {
    const left = (await open(bytes, 81)).session;
    const { session: right, base } = await open(bytes, 82);
    left.splitParagraph(at(left, 'body', 0, 3));
    sync(left, right);
    const { paraId } = left.paragraphs('body')[1]!;
    expect(endEmptyListItem(left, 'body', paraId, styleValuesFor(left, base))).toBe(true);
    right.insertText({ story: 'body', paraId, offset: 0 }, 'x');
    sync(left, right);
    const saved = await savedDocumentXml(left, base);
    expect(await savedDocumentXml(right, base)).toBe(saved);
    expect(pPrs(saved)[1]).toBe('<w:pPr><w:pStyle w:val="ListParagraph"/></w:pPr>');
    expect(left.paragraphs('body')[1]!.text).toBe('x');
    left.destroy();
    right.destroy();
  });
});

describe('revision ids the editor makes', () => {
  const BY = 'w:author="Rev" w:date="2026-01-01T00:00:00Z"';
  const changed = (id: number) =>
    `<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="${id}" ${BY}><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr><w:r><w:t>Changed para</w:t></w:r></w:p>`;

  it('save in the upper half of int32, the same in every save, without reading the session’s revisions', async () => {
    const bytes = fixture(changed(5) + changed(6) + '<w:p/>');
    const { session, base } = await open(bytes);
    session.listRevisions = () => {
      throw new Error('the save read every revision');
    };
    session.splitParagraph(at(session, 'body', 0, 3));
    const first = revisionIds(pPrs(await savedDocumentXml(session, base))[0]!).map(Number);
    session.splitParagraph(at(session, 'body', 2, 3));
    const second = pPrs(await savedDocumentXml(session, base)).slice(0, 4).flatMap(revisionIds).map(Number);
    for (const id of [first[0]!, second[2]!]) {
      expect(id).toBeGreaterThanOrEqual(2 ** 30);
      expect(id).toBeLessThan(2 ** 31);
    }
    expect(second[0]).toBe(first[0]!);
    expect(second[2]).not.toBe(first[0]);
    session.destroy();
  });

  it('cost no more as the document holds more revisions', async () => {
    let body = '';
    for (let index = 0; index < 300; index += 1) {
      body +=
        `<w:p><w:ins w:id="${3 * index + 1}" ${BY}><w:r><w:t>one </w:t></w:r></w:ins>` +
        `<w:del w:id="${3 * index + 2}" ${BY}><w:r><w:delText>two </w:delText></w:r></w:del>` +
        `<w:ins w:id="${3 * index + 3}" ${BY}><w:r><w:t>three</w:t></w:r></w:ins></w:p>`;
    }
    const { session, base } = await open(fixture(body + '<w:p/>'));
    yrsToDocument(session, base);
    session.insertText(at(session, 'body', 150, 0), 'x');
    const started = performance.now();
    yrsToDocument(session, base);
    // Reading the session's 900 revisions took about 2 s; the save alone takes about 15 ms.
    expect(performance.now() - started).toBeLessThan(500);
    session.destroy();
  });
});

describe('cells a table op makes', () => {
  const tc = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const bytes = fixture(
    '<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0" w:firstRow="1"/></w:tblPr>' +
      `<w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid><w:tr>${tc('H0')}${tc('H1')}</w:tr></w:tbl><w:p/>`
  );

  it('go with their styling in one Undo', async () => {
    const { session, base } = await open(bytes);
    session.beginUndoCapture();
    const cells = () => session.storyIds().filter((story) => story.startsWith('body:t0:')).length;
    inOneUndoStep(session, () => {
      const receipt = session.insertRow({ story: 'body', tableIndex: 0, row: 0, column: 0 }, 'above');
      styleNewCells(session, receipt.createdStoryIds, styleValuesFor(session, base));
    });
    expect(cells()).toBe(4);
    expect(session.undo()).toBe(true);
    expect(cells()).toBe(2);
    expect(session.undoCaptureMode()).toBe('auto');
    session.destroy();
  });
});

describe('Enter around list items', () => {
  const N = (level: number, numId = 1) => `<w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${numId}"/></w:numPr>`;
  const bytes = fixture(
    `<w:p><w:pPr><w:pStyle w:val="Heading1"/>${N(0)}</w:pPr><w:r><w:t>Head</w:t></w:r></w:p>` +
      `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/>${N(0)}</w:pPr><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>` +
      `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/>${N(0)}</w:pPr><w:r><w:br/></w:r></w:p>` +
      `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/>${N(1)}</w:pPr></w:p>` +
      `<w:p><w:pPr><w:pStyle w:val="ListBullet"/>${N(0, 2)}</w:pPr></w:p><w:p/>`
  );

  it('starts the next style clean after a heading numbered on itself', async () => {
    const { session, base } = await open(bytes);
    const values = styleValuesFor(session, base);
    const { secondParaId: paraId } = session.splitParagraph(at(session, 'body', 0, 4));
    applyNextStyle(
      session,
      { story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 0 } },
      'Normal',
      'Heading1',
      values
    );
    const properties = session.paragraphs('body')[1]!.properties;
    expect([properties.pStyle, properties.numPr ?? null, properties.listMarker ?? null]).toEqual(['Normal', null, null]);
    expect(pPrs(await savedDocumentXml(session, base))[1]).toBe('<w:pPr><w:pStyle w:val="Normal"/></w:pPr>');
    session.destroy();
  });

  it('keeps an item holding only a field or a break, moves a nested item up, and turns a style’s list off', async () => {
    const { session, base } = await open(bytes);
    const values = styleValuesFor(session, base);
    const end = (index: number) => endEmptyListItem(session, 'body', session.paragraphs('body')[index]!.paraId, values);
    expect([end(1), end(2)]).toEqual([false, false]);
    expect(end(3)).toBe(true);
    const nested = session.paragraphs('body')[3]!.properties;
    expect([nested.numPr, nested.listMarker, nested.indentLeft]).toEqual([{ numId: 1, ilvl: 0 }, '•', 720]);
    expect(end(3)).toBe(true);
    expect(session.paragraphs('body')[3]!.properties.numPr ?? null).toBe(null);
    // ListBullet gives a list, so ending the item's own list turns the style's off too.
    expect(end(4)).toBe(true);
    const saved = pPrs(await savedDocumentXml(session, base));
    expect(saved.slice(3, 5)).toEqual([
      '<w:pPr><w:pStyle w:val="ListParagraph"/></w:pPr>',
      '<w:pPr><w:pStyle w:val="ListBullet"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr></w:pPr>',
    ]);
    session.destroy();
  });
});

describe('Enter at the end of a paragraph', () => {
  // Word copies the paragraph mark, so the new paragraph has every direct pPr
  // the paragraph has, the mark's run properties too. The oracle is the pPr the
  // save writes for the untouched paragraph.
  const DIRECT =
    '<w:pPr><w:keepNext/><w:keepLines/><w:pageBreakBefore/><w:widowControl w:val="0"/><w:suppressLineNumbers/>' +
    '<w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="auto"/><w:bottom w:val="double" w:sz="6" w:space="1" w:color="FF0000"/></w:pBdr>' +
    '<w:shd w:val="clear" w:color="auto" w:fill="DDEEFF"/><w:tabs><w:tab w:val="center" w:pos="4320"/></w:tabs>' +
    '<w:kinsoku w:val="0"/><w:snapToGrid w:val="0"/><w:spacing w:before="120" w:after="60"/>' +
    '<w:ind w:left="360" w:right="240" w:firstLine="180"/><w:contextualSpacing/><w:jc w:val="center"/>' +
    '<w:textAlignment w:val="center"/><w:outlineLvl w:val="2"/><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr></w:pPr>';
  const COPIED = [
    'pStyle',
    'alignment',
    'spaceBefore',
    'spaceAfter',
    'indentLeft',
    'indentRight',
    'indentFirstLine',
    'hangingIndent',
    'borders',
    'shading',
    'tabs',
    'keepNext',
    'keepLines',
    'widowControl',
    'pageBreakBefore',
    'contextualSpacing',
    'outlineLevel',
    'snapToGrid',
    'defaultTextFormatting',
    '_originalFormatting',
  ];
  const copied = (properties: Record<string, unknown>) =>
    Object.fromEntries(COPIED.map((key) => [key, properties[key] ?? null]));
  const bytes = fixture(`<w:p>${DIRECT}<w:r><w:t>Source</w:t></w:r></w:p><w:p/>`);
  const end = (session: YrsSession, index: number) =>
    at(session, 'body', index, session.paragraphs('body')[index]!.text.length);

  async function savedBytes(session: YrsSession, base: Document): Promise<Uint8Array<ArrayBuffer>> {
    return new Uint8Array(await repackDocx(yrsToDocument(session, base)));
  }

  it('copies every direct property, the mark’s run properties and the kept pPr children', async () => {
    const { session, base } = await open(bytes);
    const [source] = pPrs(await savedDocumentXml(session, base));
    expect(source).toContain('<w:pBdr>');
    session.splitParagraph(end(session, 0));
    const [first, added] = session.paragraphs('body');
    expect(added!.text).toBe('');
    expect(copied(added!.properties)).toEqual(copied(first!.properties));
    expect(added!.properties.defaultTextFormatting).toMatchObject({ bold: true, color: { rgb: 'FF0000' } });

    const saved = await savedBytes(session, base);
    const xml = new TextDecoder().decode((unzipContainer(saved) as Record<string, Uint8Array>)['word/document.xml']);
    expect(pPrs(xml).slice(0, 2)).toEqual([source, source]);
    // Reopening the file gives the new paragraph what the editor holds.
    const reopened = (await open(saved, 9)).session;
    expect(copied(reopened.paragraphs('body')[1]!.properties)).toEqual(copied(added!.properties));
    reopened.destroy();
    session.destroy();
  });

  it('undoes in one step', async () => {
    const { session, base } = await open(bytes);
    const untouched = await savedDocumentXml(session, base);
    const before = session.paragraphs('body')[0]!.properties;
    session.beginUndoCapture();
    session.splitParagraph(end(session, 0));
    expect(session.undo()).toBe(true);
    expect(session.paragraphs('body').map(({ text }) => text)).toEqual(['Source', '']);
    expect(session.paragraphs('body')[0]!.properties).toEqual(before);
    expect(await savedDocumentXml(session, base)).toBe(untouched);
    session.destroy();
  });

  it('leaves the section with the mark that ends it', async () => {
    const sectioned = fixture(
      '<w:p><w:pPr><w:jc w:val="right"/><w:sectPr><w:type w:val="continuous"/></w:sectPr></w:pPr><w:r><w:t>Section end</w:t></w:r></w:p>' +
        '<w:p><w:r><w:t>Next</w:t></w:r></w:p>'
    );
    const { session, base } = await open(sectioned);
    const [source] = pPrs(await savedDocumentXml(session, base));
    expect(source).toContain('<w:sectPr');
    session.splitParagraph(end(session, 0));
    expect(pPrs(await savedDocumentXml(session, base)).slice(0, 2)).toEqual([
      '<w:pPr><w:jc w:val="right"/></w:pPr>',
      source,
    ]);
    session.destroy();
  });

  it('keeps a tracked mark revision with the text and gives the copied property change its own id', async () => {
    const BY = 'w:author="Rev" w:date="2026-01-01T00:00:00Z"';
    const CHANGE = `<w:pPrChange w:id="5" ${BY}><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange>`;
    const tracked = fixture(
      `<w:p><w:pPr><w:jc w:val="center"/>${CHANGE}</w:pPr><w:r><w:t>Changed para</w:t></w:r></w:p>` +
        `<w:p><w:pPr><w:rPr><w:ins w:id="7" ${BY}/></w:rPr></w:pPr><w:r><w:t>Inserted mark</w:t></w:r></w:p>` +
        `<w:p><w:pPr><w:rPr><w:del w:id="8" ${BY}/></w:rPr></w:pPr><w:r><w:t>Deleted mark</w:t></w:r></w:p>` +
        '<w:p/>'
    );
    const { session, base } = await open(tracked);
    for (const index of [2, 1, 0]) session.splitParagraph(end(session, index));
    const saved = pPrs(await savedDocumentXml(session, base)).map(strip);
    expect(saved.slice(2)).toEqual([
      '<w:pPr><w:rPr><w:ins w:id="7"/></w:rPr></w:pPr>',
      '',
      '<w:pPr><w:rPr><w:del w:id="8"/></w:rPr></w:pPr>',
      '',
      '',
    ]);
    const [source, copy] = saved.slice(0, 2);
    expect(source).toBe(`<w:pPr><w:jc w:val="center"/>${strip(CHANGE)}</w:pPr>`);
    const [id] = revisionIds(copy!).map(Number);
    expect(id).toBeGreaterThanOrEqual(2 ** 30);
    expect(copy).toBe(source!.replace('w:id="5"', `w:id="${id}"`));
    session.destroy();
  });

  it('starts a different next style clean, in one Undo step', async () => {
    const heading = fixture(
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr>' +
        '<w:kinsoku w:val="0"/><w:spacing w:after="0"/><w:jc w:val="center"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>Head</w:t></w:r></w:p><w:p/>'
    );
    const { session, base } = await open(heading);
    const untouched = await savedDocumentXml(session, base);
    session.beginUndoCapture();
    const { secondParaId: paraId } = session.splitParagraph(end(session, 0));
    const range = { story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 0 } };
    applyNextStyle(session, range, 'Normal', 'Heading1', styleValuesFor(session, base));
    const properties = session.paragraphs('body')[1]!.properties;
    expect([properties.pStyle, properties.alignment ?? null, properties.borders ?? null]).toEqual(['Normal', null, null]);
    expect(properties._originalFormatting ?? null).toBe(null);
    expect((properties.defaultTextFormatting as Record<string, unknown> | undefined)?.bold ?? null).toBe(null);
    expect(pPrs(await savedDocumentXml(session, base))[1]).toBe('<w:pPr><w:pStyle w:val="Normal"/></w:pPr>');
    expect(session.undo()).toBe(true);
    expect(await savedDocumentXml(session, base)).toBe(untouched);
    session.destroy();
  });

  it('converges with a peer’s alignment change, which lands on the new paragraph', async () => {
    const left = (await open(bytes, 81)).session;
    const { session: right, base } = await open(bytes, 82);
    left.splitParagraph(end(left, 0));
    const { paraId } = right.paragraphs('body')[0]!;
    right.setParagraphAttrs({ story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 0 } }, { alignment: 'right' });
    sync(left, right);
    const saved = await savedDocumentXml(left, base);
    expect(await savedDocumentXml(right, base)).toBe(saved);
    // The new paragraph ends with the source's own mark, which the peer changed.
    expect(left.paragraphs('body').map(({ text, properties }) => [text, properties.alignment ?? null])).toEqual([
      ['Source', 'center'],
      ['', 'right'],
      ['', null],
    ]);
    left.destroy();
    right.destroy();
  });
});
