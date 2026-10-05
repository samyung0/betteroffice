import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { unzipContainer } from '../wasm/opc';
import type { Document } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsSession } from './index';
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
  '<w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr><w:contextualSpacing/></w:pPr></w:style>' +
  '<w:style w:type="table" w:styleId="Grid"><w:name w:val="Table Grid"/>' +
  '<w:tblStylePr w:type="firstRow"><w:pPr><w:jc w:val="center"/></w:pPr></w:tblStylePr></w:style></w:styles>';

const NUMBERING =
  `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/>` +
  '<w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>' +
  '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>';

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
    expect(copy!.replace(/w:id="\d+"/, 'w:id="5"')).toBe(source);
    expect(revisionIds(copy!)).not.toEqual(['5']);
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
    const ids = pPrs(saved).slice(0, 3).flatMap(revisionIds);
    expect(ids).toHaveLength(3);
    expect(new Set(ids).size).toBe(3);
    expect(ids).toContain('5');
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

  it('keeps the grid columns a row skips', async () => {
    const { session, base } = await open(bytes);
    session.insertText(at(session, 'body', 0), 'y');
    expect(rows(await savedDocumentXml(session, base))[1]).toStartWith(`<w:tr>${SKIPPED}<w:tc>`);
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
