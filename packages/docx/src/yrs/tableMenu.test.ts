import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { readDocxContainer } from '../docx/zipContainer';
import type { Document, Table } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import {
  compatibilityModeFromDocument,
  createYrsSession,
  type YrsSession,
  type YrsTableRange,
} from './index';
import { yrsToDocument } from './yrsToDocument';

// The table menu's six items: what Word writes for each, a reopen showing the
// same, an untouched table saving the same bytes, one Undo step, and peers.

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

const cell = (width: number, text: string, extra = '') =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${extra}</w:tcPr>` +
  `<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;

// A Word table with a fixed layout, then one the edits never touch.
const EDITED =
  '<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/>' +
  '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  `<w:tr>${cell(1000, 'H1')}${cell(2000, 'H2')}${cell(3000, 'H3')}</w:tr>` +
  `<w:tr>${cell(1000, 'A')}${cell(2000, 'B')}${cell(3000, 'C')}</w:tr></w:tbl>`;
const UNTOUCHED =
  '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:jc w:val="center"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
  `<w:tr><w:trPr><w:tblHeader/></w:trPr>${cell(2000, 'X', '<w:noWrap/>')}${cell(2000, 'Y', '<w:vAlign w:val="bottom"/>')}</w:tr></w:tbl>`;
const BODY = `<w:p><w:r><w:t>Before</w:t></w:r></w:p>${EDITED}<w:p/>${UNTOUCHED}<w:p/>`;

/** The fixture, or `body` in a document whose settings give `compatibilityMode`. */
function fixture(body = BODY, compatibilityMode?: number): Uint8Array<ArrayBuffer> {
  const settings = compatibilityMode !== undefined;
  const parts = new Map<string, Uint8Array>();
  parts.set(
    '[Content_Types].xml',
    toBytes(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        `<Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/>` +
        (settings
          ? `<Override PartName="/word/settings.xml" ContentType="${OFFICE}.settings+xml"/>`
          : '') +
        '</Types>'
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
    'word/document.xml',
    toBytes(`<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr/></w:body></w:document>`)
  );
  if (settings) {
    parts.set(
      'word/_rels/document.xml.rels',
      toBytes(
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          `<Relationship Id="settings" Type="${R}/settings" Target="settings.xml"/></Relationships>`
      )
    );
    parts.set(
      'word/settings.xml',
      toBytes(
        `<w:settings xmlns:w="${W}"><w:compat><w:compatSetting w:name="compatibilityMode" ` +
          `w:uri="http://schemas.microsoft.com/office/word" w:val="${compatibilityMode}"/></w:compat></w:settings>`
      )
    );
  }
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

const TABLE = { story: 'body', tableIndex: 0 };
const at = (row: number, column: number) => ({ ...TABLE, row, column });
const range = (row: number, column: number, toRow = row, toColumn = column): YrsTableRange => ({
  anchor: at(row, column),
  head: at(toRow, toColumn),
});

async function open(clientId: number, bytes = fixture()) {
  const session = await createYrsSession({ clientId });
  session.seedFromDocx(bytes);
  session.beginUndoCapture();
  return session;
}

async function saved(session: YrsSession, source: Document): Promise<ArrayBuffer> {
  return repackDocx(yrsToDocument(session, source));
}

function tablesXml(bytes: ArrayBuffer): string[] {
  const xml = readDocxContainer(bytes).text('word/document.xml') ?? '';
  return xml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/g) ?? [];
}

function tables(document: Document): Table[] {
  return document.package.document.content.filter((block): block is Table => block.type === 'table');
}

function payload(session: YrsSession) {
  return session.storyTables('body')[0] as {
    tblPr: Record<string, unknown>;
    grid: number[];
    rows: Array<{ trPr: Record<string, unknown>; cells: Array<{ tcPr: Record<string, unknown> }> }>;
  };
}

/** Two replicas of the fixture, the second seeded from the first. */
async function peers(leftId: number, rightId: number): Promise<[YrsSession, YrsSession]> {
  const left = await open(leftId);
  const right = await createYrsSession({ clientId: rightId });
  right.applyUpdate(left.encodeStateAsUpdate());
  return [left, right];
}

function sync(left: YrsSession, right: YrsSession): void {
  left.applyUpdate(right.encodeStateAsUpdate(left.encodeStateVector()));
  right.applyUpdate(left.encodeStateAsUpdate(right.encodeStateVector()));
}

/** Every row's cells cover exactly the grid (the fixture has no merges). */
function consistent(table: ReturnType<typeof payload>): boolean {
  return table.rows.every(
    (row) =>
      row.cells.reduce((sum, cell) => sum + Number(cell.tcPr.colspan ?? 1), 0) === table.grid.length
  );
}

/** A peer's structural edits that race the table menu's. */
const STRUCTURAL: Array<[string, (session: YrsSession) => void]> = [
  ['adds a row', (session) => session.insertRow(at(1, 0), 'below')],
  ['adds a column', (session) => session.insertColumn(at(0, 1), 'right')],
  ['deletes a column', (session) => session.deleteColumn(range(0, 1))],
];

interface Item {
  name: string;
  run: (session: YrsSession) => void;
  /** What Word writes into the edited table. */
  xml: RegExp[];
  absent?: RegExp[];
  /** The reopened table shows it. */
  reopened: (table: Table) => void;
}

const ITEMS: Item[] = [
  {
    name: 'vertical alignment',
    run: (session) => session.setCellTextFormat(range(1, 0, 1, 1), { verticalAlign: 'center' }),
    xml: [/<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"\/><w:vAlign w:val="center"\/><\/w:tcPr><w:p><w:r><w:t>A</],
    absent: [/H1<\/w:t>[\s\S]*vAlign[\s\S]*<\/w:tr><w:tr>/],
    reopened: (table) => {
      expect(table.rows[1].cells.map((cell) => cell.formatting?.verticalAlign)).toEqual([
        'center',
        'center',
        undefined,
      ]);
    },
  },
  {
    name: 'table alignment',
    run: (session) => session.setTableAlignment(TABLE, 'center'),
    xml: [/<w:tblW w:w="6000" w:type="dxa"\/><w:jc w:val="center"\/>/],
    reopened: (table) => expect(table.formatting?.justification).toBe('center'),
  },
  {
    name: 'header row',
    run: (session) => session.setHeaderRows(TABLE, 1),
    xml: [/<w:tr><w:trPr><w:tblHeader\/><\/w:trPr><w:tc><w:tcPr><w:tcW w:w="1000"/],
    absent: [/<w:t>H3<\/w:t>[\s\S]*tblHeader/],
    reopened: (table) =>
      expect(table.rows.map((row) => row.formatting?.header ?? false)).toEqual([true, false]),
  },
  {
    name: 'distribute columns',
    run: (session) => session.distributeColumns(range(0, 0)),
    xml: [
      /<w:tblGrid><w:gridCol w:w="2000"\/><w:gridCol w:w="2000"\/><w:gridCol w:w="2000"\/><\/w:tblGrid>/,
      /<w:tcW w:w="2000" w:type="dxa"\/><\/w:tcPr><w:p><w:r><w:t>C</,
    ],
    absent: [/w:w="1000"/, /w:w="3000"/],
    reopened: (table) => {
      expect(table.columnWidths).toEqual([2000, 2000, 2000]);
      expect(table.rows[0].cells.map((cell) => cell.formatting?.width?.value)).toEqual([2000, 2000, 2000]);
    },
  },
  {
    name: 'auto-fit',
    run: (session) => session.autofitTable(TABLE),
    xml: [/<w:tblPr><w:tblW w:w="0" w:type="auto"\/>/, /<w:tcW w:w="0" w:type="auto"\/><\/w:tcPr><w:p><w:r><w:t>H1</],
    absent: [/tblLayout/, /w:type="dxa"/],
    reopened: (table) => {
      expect(table.formatting?.layout).toBeUndefined();
      expect(table.formatting?.width).toEqual({ value: 0, type: 'auto' });
      for (const row of table.rows)
        for (const cell of row.cells) expect(cell.formatting?.width).toEqual({ value: 0, type: 'auto' });
    },
  },
  {
    name: 'wrap text off',
    run: (session) => session.setCellTextFormat(range(0, 2), { noWrap: true }),
    xml: [/<w:tcW w:w="3000" w:type="dxa"\/><w:noWrap\/><\/w:tcPr><w:p><w:r><w:t>H3</],
    reopened: (table) =>
      expect(table.rows[0].cells.map((cell) => cell.formatting?.noWrap ?? false)).toEqual([
        false,
        false,
        true,
      ]),
  },
];

describe('the DOCX table menu', () => {
  let source: Document;
  let baseline: string[];

  beforeAll(async () => {
    preloadEditWasm(new Uint8Array(readFileSync(WASM)));
    source = await parseDocx(fixture().buffer, { preloadFonts: false });
    const session = await open(61000);
    try {
      baseline = tablesXml(await saved(session, source));
      expect(baseline[1]).toBe(UNTOUCHED);
    } finally {
      session.destroy();
    }
  });

  for (const [index, item] of ITEMS.entries()) {
    it(`${item.name}: saves what Word writes, reopens the same and keeps the other table`, async () => {
      const session = await open(61001 + index);
      let bytes: ArrayBuffer;
      try {
        item.run(session);
        bytes = await saved(session, source);
      } finally {
        session.destroy();
      }
      const [edited, untouched] = tablesXml(bytes);
      for (const pattern of item.xml) expect(edited).toMatch(pattern);
      for (const pattern of item.absent ?? []) expect(edited).not.toMatch(pattern);
      expect(untouched).toBe(baseline[1]);
      const reopened = await parseDocx(bytes, { preloadFonts: false });
      item.reopened(tables(reopened)[0]);
      // A reopened editor seeds the same values.
      const again = await open(61101 + index, new Uint8Array(bytes));
      try {
        const resaved = tablesXml(await saved(again, reopened));
        expect(resaved[0]).toBe(edited);
      } finally {
        again.destroy();
      }
    });

    it(`${item.name}: undoes in one step`, async () => {
      const session = await open(61201 + index);
      try {
        const before = payload(session);
        item.run(session);
        expect(payload(session)).not.toEqual(before);
        expect(session.undo()).toBe(true);
        expect(payload(session)).toEqual(before);
        expect(tablesXml(await saved(session, source))).toEqual(baseline);
      } finally {
        session.destroy();
      }
    });

    it(`${item.name}: keeps a peer's typing made at the same time`, async () => {
      const [left, right] = await peers(61301 + index, 61401 + index);
      try {
        const story = 'body:t0:r1c1';
        const paraId = right.paragraphs(story)[0].paraId;
        right.insertText({ story, paraId, offset: 1 }, ' typed');
        item.run(left);
        sync(left, right);
        expect(payload(right)).toEqual(payload(left));
        const [both] = tablesXml(await saved(right, source));
        for (const pattern of item.xml) expect(both).toMatch(pattern);
        expect(both).toContain('B typed');
      } finally {
        left.destroy();
        right.destroy();
      }
    });

    for (const [change, structural] of STRUCTURAL) {
      it(`${item.name}: racing a peer that ${change} converges on a consistent table`, async () => {
        // Both client-id orders, so either edit can win the rows.
        for (const [a, b] of [
          [61501, 61502],
          [61502, 61501],
        ]) {
          const [left, right] = await peers(a + index * 10, b + index * 10);
          try {
            item.run(left);
            structural(right);
            sync(left, right);
            expect(payload(right)).toEqual(payload(left));
            expect(consistent(payload(left))).toBe(true);
            const [edited] = tablesXml(await saved(left, source));
            const columns = edited.match(/<w:gridCol /g)?.length;
            for (const row of edited.match(/<w:tr>[\s\S]*?<\/w:tr>/g) ?? []) {
              expect(row.match(/<w:tc>/g)?.length).toBe(columns);
            }
          } finally {
            left.destroy();
            right.destroy();
          }
        }
      });
    }
  }

  it('keeps a peer’s added row when the table is aligned at the same time', async () => {
    const left = await open(61500);
    const right = await createYrsSession({ clientId: 61501 });
    try {
      right.applyUpdate(left.encodeStateAsUpdate());
      left.setTableAlignment(TABLE, 'right');
      right.insertRow(at(1, 0), 'below');
      left.applyUpdate(right.encodeStateAsUpdate(left.encodeStateVector()));
      right.applyUpdate(left.encodeStateAsUpdate(right.encodeStateVector()));
      for (const session of [left, right]) {
        expect(payload(session).tblPr.justification).toBe('right');
        expect(payload(session).rows).toHaveLength(3);
      }
    } finally {
      left.destroy();
      right.destroy();
    }
  });

  it('gives a new table the document’s compatibility mode, so a reopen draws it alike', async () => {
    for (const mode of [15, 14]) {
      const bytes = fixture('<w:p><w:r><w:t>Body</w:t></w:r></w:p>', mode);
      const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
      expect(compatibilityModeFromDocument(parsed)).toBe(mode);
      const session = await open(61700 + mode, bytes);
      const lowered = (replica: YrsSession) => {
        const table = (replica.yrsBlocksForStory('body') as Array<Record<string, unknown>>).find(
          (block) => block.kind === 'table'
        );
        return {
          mode: table?.compatibilityMode,
          margin: Math.round(Number(table?.cellMarginLeft) * 10) / 10,
        };
      };
      let reopened: YrsSession | undefined;
      try {
        const anchor = session.paragraphs('body')[0];
        session.insertTable(
          { story: 'body', paraId: anchor.paraId, offset: 0 },
          1,
          1,
          undefined,
          compatibilityModeFromDocument(parsed)
        );
        expect(lowered(session)).toEqual({ mode, margin: 7.2 });
        reopened = await open(61800 + mode, new Uint8Array(await saved(session, parsed)));
        expect(lowered(reopened)).toEqual(lowered(session));
      } finally {
        session.destroy();
        reopened?.destroy();
      }
    }
  });

  it('pins header rows down to the selection and unpins them all', async () => {
    const session = await open(61600);
    try {
      session.setHeaderRows(TABLE, 2);
      expect(payload(session).rows.map((row) => row.trPr.isHeader)).toEqual([true, true]);
      session.setHeaderRows(TABLE, 0);
      expect(payload(session).rows.map((row) => row.trPr.isHeader)).toEqual([false, false]);
      expect(() => session.setHeaderRows(TABLE, 3)).toThrow();
    } finally {
      session.destroy();
    }
  });
});
