import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import type { BlockContent, Document, Paragraph, ParagraphFormatting } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { createStyleResolver } from '../styles';
import { styleParagraphValues } from './documentToYrs';
import { createYrsSession, type YrsParagraphAttrs, type YrsParagraphTabStop, type YrsSession } from './index';
import { explicitParagraphAttrs } from './paragraphSeed';
import { yrsToDocument } from './yrsToDocument';

// Word reads the saved file; the oracle is that reopening it (a publication
// reseeds from the saved bytes) shows what the editor showed before the save.

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

const BODY = [
  '<w:p><w:r><w:t>Plain</w:t></w:r></w:p>',
  '<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>',
  '<w:p><w:pPr><w:pStyle w:val="Title"/><w:keepNext/><w:keepLines/>' +
    '<w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="FF0000"/></w:pBdr>' +
    '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/><w:tabs><w:tab w:val="left" w:pos="2000"/></w:tabs>' +
    '<w:spacing w:before="100" w:after="200" w:line="300" w:lineRule="auto"/>' +
    '<w:ind w:left="400" w:right="300" w:firstLine="200"/><w:contextualSpacing/><w:jc w:val="center"/>' +
    '<w:outlineLvl w:val="1"/></w:pPr><w:r><w:t>Full</w:t></w:r></w:p>',
  '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
    '<w:r><w:t>Item one</w:t></w:r></w:p>',
  '<w:p><w:pPr><w:ind w:left="420" w:leftChars="200"/></w:pPr><w:r><w:t>Chars</w:t></w:r></w:p>',
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' +
    '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
  '<w:p><w:r><w:t>Tail</w:t></w:r></w:p>',
  '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/></w:pPr><w:r><w:t>Styled indent</w:t></w:r></w:p>',
  '<w:p><w:pPr><w:pStyle w:val="Tabbed"/></w:pPr><w:r><w:t>Styled tabs</w:t></w:r></w:p>',
].join('');

const STYLES =
  `<w:styles xmlns:w="${W}"><w:docDefaults><w:pPrDefault><w:pPr>` +
  '<w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Tabbed"><w:name w:val="Tabbed"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:keepNext/><w:tabs><w:tab w:val="left" w:pos="1000"/><w:tab w:val="right" w:pos="5000"/></w:tabs>' +
  '<w:ind w:firstLine="360"/></w:pPr></w:style>' +
  '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/>' +
  '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style></w:styles>';

const NUMBERING =
  `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="1">` +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>' +
  '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2."/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>' +
  '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>';

function fixture(): Uint8Array<ArrayBuffer> {
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
    toBytes(`<w:document xmlns:w="${W}"><w:body>${BODY}<w:sectPr/></w:body></w:document>`)
  );
  parts.set('word/styles.xml', toBytes(STYLES));
  parts.set('word/numbering.xml', toBytes(NUMBERING));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

/** Every paragraph in document order, table cells included. */
function paragraphs(document: Document): Paragraph[] {
  const found: Paragraph[] = [];
  const visit = (blocks: readonly BlockContent[]) => {
    for (const block of blocks) {
      if (block.type === 'paragraph') found.push(block);
      else if (block.type === 'table') {
        for (const row of block.rows) for (const cell of row.cells) visit(cell.content);
      }
    }
  };
  visit(document.package.document.content);
  return found;
}

/** The paragraph properties an editor shows: what the seed resolves and the layout reads. */
const SHOWN = [
  'pStyle',
  'alignment',
  'spaceBefore',
  'spaceAfter',
  'spaceBeforeLines',
  'spaceAfterLines',
  'beforeAutospacing',
  'afterAutospacing',
  'lineSpacing',
  'lineSpacingRule',
  'indentLeft',
  'indentRight',
  'indentFirstLine',
  'hangingIndent',
  'borders',
  'shading',
  'tabs',
  'pageBreakBefore',
  'keepNext',
  'keepLines',
  'widowControl',
  'contextualSpacing',
  'snapToGrid',
  'autoSpaceDE',
  'autoSpaceDN',
  'outlineLevel',
  'bidi',
  'numPr',
] as const;

/** Tab stops compare by position, alignment and leader whatever shape the editor stored them in. */
function shown(session: YrsSession, story: string): Array<Record<string, unknown>> {
  return session.paragraphs(story).map(({ properties }) =>
    Object.fromEntries(
      SHOWN.flatMap((key) => {
        let value = properties[key];
        if (value == null || (key === 'hangingIndent' && value === false)) return [];
        if (key === 'tabs' && Array.isArray(value)) {
          value = value.map((tab: Record<string, unknown>) => ({
            position: tab.position ?? tab.pos,
            alignment: tab.alignment ?? tab.val,
            ...(tab.leader ? { leader: tab.leader } : {}),
          }));
        }
        return [[key, value]];
      })
    )
  );
}

/** Body story paragraphs then the table cell's, as `paragraphs()` orders the saved file. */
function shownAll(session: YrsSession): Array<Record<string, unknown>> {
  const cell = session.storyIds().find((id) => id.startsWith('body:t'));
  const body = shown(session, 'body');
  return [...body.slice(0, 5), ...(cell ? shown(session, cell) : []), ...body.slice(5)];
}

async function save(session: YrsSession, base: Document): Promise<Uint8Array<ArrayBuffer>> {
  return new Uint8Array(await repackDocx(yrsToDocument(session, base)));
}

async function reopen(bytes: Uint8Array<ArrayBuffer>, clientId: number) {
  const document = await parseDocx(bytes.buffer, { preloadFonts: false });
  const session = await createYrsSession({ clientId });
  session.seedFromDocx(bytes);
  return { document, session };
}

function range(session: YrsSession, index: number) {
  const { paraId } = session.paragraphs('body')[index]!;
  return { story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 0 } };
}

const EDITED = [0, 1, 2] as const;

/** One editor change per paragraph property; booleans flip what the paragraph shows. */
const CASES: Array<[string, (current: Record<string, unknown>) => YrsParagraphAttrs, Partial<ParagraphFormatting>]> = [
  ['alignment', () => ({ alignment: 'right' }), { alignment: 'right' }],
  ['left indent', () => ({ indentLeft: 1234 }), { indentLeft: 1234 }],
  ['right indent', () => ({ indentRight: 567 }), { indentRight: 567 }],
  ['first line indent', () => ({ indentFirstLine: 345 }), { indentFirstLine: 345 }],
  ['line spacing', () => ({ lineSpacing: 480, lineSpacingRule: 'auto' }), { lineSpacing: 480 }],
  ['exact line spacing', () => ({ lineSpacing: 300, lineSpacingRule: 'exact' }), { lineSpacing: 300, lineSpacingRule: 'exact' }],
  ['space before', () => ({ spaceBefore: 240 }), { spaceBefore: 240 }],
  ['space after', () => ({ spaceAfter: 360 }), { spaceAfter: 360 }],
  ['tabs', () => ({ tabs: [{ position: 3000, alignment: 'center' }] }), { tabs: [{ position: 3000, alignment: 'center' }] }],
  ['right to left', () => ({ bidi: true }), { bidi: true }],
  ['outline level', () => ({ other: { outlineLevel: 3 } }), { outlineLevel: 3 }],
  [
    'borders',
    () => ({ other: { borders: { bottom: { style: 'double', size: 6, space: 1, color: { rgb: '0000FF' } } } } }),
    { borders: { bottom: { style: 'double', size: 6, space: 1, color: { rgb: '0000FF' } } } },
  ],
  [
    'shading',
    () => ({ other: { shading: { pattern: 'clear', fill: { rgb: '00FF00' } } } }),
    { shading: { pattern: 'clear', fill: { rgb: '00FF00' } } },
  ],
  ...(['keepNext', 'keepLines', 'pageBreakBefore', 'widowControl', 'contextualSpacing', 'snapToGrid', 'autoSpaceDE'] as const).map(
    (key): (typeof CASES)[number] => [key, (current) => ({ other: { [key]: current[key] !== true } }), {}]
  ),
];

describe('paragraph properties on save', () => {
  beforeAll(() =>
    preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))))
  );

  it('keeps every untouched paragraph exactly as the source wrote it', async () => {
    const bytes = fixture();
    const source = await parseDocx(bytes.buffer, { preloadFonts: false });
    const session = await createYrsSession({ clientId: 81001 });
    try {
      session.seedFromDocx(bytes);
      const first = yrsToDocument(session, source);
      // The editor's later saves project over the previous projection; typing reprojects the item.
      const item = session.paragraphs('body')[3]!;
      session.insertText({ story: 'body', paraId: item.paraId, offset: 0 }, 'An ');
      const saved = await parseDocx(await repackDocx(yrsToDocument(session, first)), { preloadFonts: false });
      expect(paragraphs(saved).map((paragraph) => paragraph.formatting)).toEqual(
        paragraphs(source).map((paragraph) => paragraph.formatting)
      );
    } finally {
      session.destroy();
    }
  });

  for (const [name, change, expected] of CASES) {
    it(`saves an edited ${name} on a plain, a styled and a fully formatted paragraph`, async () => {
      const bytes = fixture();
      const source = await parseDocx(bytes.buffer, { preloadFonts: false });
      const session = await createYrsSession({ clientId: 81002 });
      let reopened: Awaited<ReturnType<typeof reopen>> | undefined;
      try {
        session.seedFromDocx(bytes);
        for (const index of EDITED) {
          session.setParagraphAttrs(range(session, index), change(shown(session, 'body')[index]!));
        }
        const editor = shownAll(session);
        reopened = await reopen(await save(session, source), 81003);
        expect(shownAll(reopened.session)).toEqual(editor);
        const before = paragraphs(source);
        const after = paragraphs(reopened.document);
        for (const index of EDITED) {
          expect(after[index]!.formatting).toMatchObject(expected);
          expect(after[index]!.formatting?.styleId).toBe(before[index]!.formatting?.styleId);
        }
        // What the change left alone stays as the source wrote it.
        expect(after[2]!.formatting?.borders ?? null).toEqual(
          name === 'borders' ? expected.borders! : before[2]!.formatting!.borders!
        );
        expect(after.slice(3).map((paragraph) => paragraph.formatting)).toEqual(
          before.slice(3).map((paragraph) => paragraph.formatting)
        );
      } finally {
        session.destroy();
        reopened?.session.destroy();
      }
    });
  }

  it('removes a cleared source property', async () => {
    const bytes = fixture();
    const source = await parseDocx(bytes.buffer, { preloadFonts: false });
    const session = await createYrsSession({ clientId: 81004 });
    try {
      session.seedFromDocx(bytes);
      session.setParagraphAttrs(range(session, 2), { tabs: null, other: { keepNext: null, borders: null } });
      const saved = paragraphs(await parseDocx((await save(session, source)).buffer, { preloadFonts: false }))[2]!;
      expect(saved.formatting?.tabs).toBeUndefined();
      expect(saved.formatting?.keepNext).toBeUndefined();
      expect(saved.formatting?.borders).toBeUndefined();
      expect(saved.formatting).toMatchObject({ keepLines: true, indentLeft: 400, alignment: 'center' });
    } finally {
      session.destroy();
    }
  });

  it('drops the character-unit twin of an edited indent', async () => {
    const bytes = fixture();
    const source = await parseDocx(bytes.buffer, { preloadFonts: false });
    const session = await createYrsSession({ clientId: 81005 });
    try {
      session.seedFromDocx(bytes);
      session.setParagraphAttrs(range(session, 4), { indentLeft: 1000 });
      const saved = paragraphs(await parseDocx((await save(session, source)).buffer, { preloadFonts: false }))[4]!;
      expect(saved.formatting?.indentLeft).toBe(1000);
      expect(saved.formatting?.indentLeftChars).toBeUndefined();
    } finally {
      session.destroy();
    }
  });

  it('writes no list or table indent and spacing for split and edited paragraphs', async () => {
    const bytes = fixture();
    const source = await parseDocx(bytes.buffer, { preloadFonts: false });
    const session = await createYrsSession({ clientId: 81006 });
    let reopened: Awaited<ReturnType<typeof reopen>> | undefined;
    try {
      session.seedFromDocx(bytes);
      session.splitParagraph({ story: 'body', paraId: session.paragraphs('body')[3]!.paraId, offset: 4 });
      const cell = session.storyIds().find((id) => id.startsWith('body:t'))!;
      const { paraId } = session.paragraphs(cell)[0]!;
      session.setParagraphAttrs(
        { story: cell, start: { paraId, offset: 0 }, end: { paraId, offset: 0 } },
        { indentLeft: 200 }
      );
      const before = paragraphs(source);
      const editor = [...shown(session, 'body'), ...shown(session, cell)];
      reopened = await reopen(await save(session, source), 81007);
      expect([...shown(reopened.session, 'body'), ...shown(reopened.session, cell)]).toEqual(editor);
      const after = paragraphs(reopened.document);
      expect(after[3]!.formatting).toEqual(before[3]!.formatting);
      expect(after[4]!.formatting).toEqual(before[3]!.formatting);
      expect(after[6]!.formatting).toEqual({ ...before[5]!.formatting, indentLeft: 200 });
    } finally {
      session.destroy();
      reopened?.session.destroy();
    }
  });

  it('applies a style’s paragraph values and reopens as the editor showed them', async () => {
    const bytes = fixture();
    const source = await parseDocx(bytes.buffer, { preloadFonts: false });
    const styles = createStyleResolver(source.package.styles);
    const session = await createYrsSession({ clientId: 81008 });
    let reopened: Awaited<ReturnType<typeof reopen>> | undefined;
    try {
      session.seedFromDocx(bytes);
      for (const [index, styleId] of [[0, 'Title'], [1, 'Normal'], [2, 'Tabbed']] as const) {
        session.applyParagraphStyle(range(session, index), styleId, styleParagraphValues(styles, styleId));
      }
      const editor = shownAll(session);
      // A paragraph styled anew shows exactly what a fresh one with that style seeds as.
      expect(editor[0]).toMatchObject({ pStyle: 'Title', contextualSpacing: true, spaceAfter: 0, lineSpacing: 240 });
      expect(editor[1]).toMatchObject({ pStyle: 'Normal', spaceAfter: 160, lineSpacing: 259 });
      expect(editor[1]!.contextualSpacing).toBeUndefined();
      expect(editor[2]).toMatchObject({ pStyle: 'Tabbed', keepNext: true, indentFirstLine: 360 });
      expect(editor[2]!.keepLines).toBeUndefined();
      reopened = await reopen(await save(session, source), 81009);
      expect(shownAll(reopened.session)).toEqual(editor);
      // Nothing the style gives is copied into a paragraph that had no direct formatting.
      expect(paragraphs(reopened.document)[0]!.formatting).toEqual({ styleId: 'Title' });
    } finally {
      session.destroy();
      reopened?.session.destroy();
    }
  });

  it('stores a cleared style value as 0, false or a clear tab stop, and reopens it', async () => {
    const bytes = fixture();
    const source = await parseDocx(bytes.buffer, { preloadFonts: false });
    const styles = createStyleResolver(source.package.styles);
    const session = await createYrsSession({ clientId: 81012 });
    let reopened: Awaited<ReturnType<typeof reopen>> | undefined;
    // What the ruler, the outdent button and tab removal send, made explicit as the editor does.
    const set = (index: number, attrs: (properties: Record<string, unknown>) => YrsParagraphAttrs) => {
      const { properties } = session.paragraphs('body')[index]!;
      const style = styleParagraphValues(styles, properties.pStyle as string);
      session.setParagraphAttrs(range(session, index), explicitParagraphAttrs(attrs(properties), style));
    };
    const removeTab = (position: number) => (properties: Record<string, unknown>) => {
      const kept = (properties.tabs as YrsParagraphTabStop[]).filter((tab) => tab.position !== position);
      return { tabs: kept.length > 0 ? kept : null };
    };
    try {
      session.seedFromDocx(bytes);
      set(6, () => ({ indentLeft: null }));
      set(7, () => ({ indentFirstLine: null, hangingIndent: false }));
      set(7, removeTab(1000));
      set(7, removeTab(5000));
      // A paragraph whose style sets nothing there clears to unset.
      set(2, () => ({ indentLeft: null }));
      const editor = shown(session, 'body');
      expect(editor[6]).toMatchObject({ indentLeft: 0 });
      expect(editor[7]).toMatchObject({
        indentFirstLine: 0,
        tabs: [{ position: 1000, alignment: 'clear' }, { position: 5000, alignment: 'clear' }],
      });
      expect(editor[2]!.indentLeft).toBeUndefined();
      const saved = await save(session, source);
      reopened = await reopen(saved, 81013);
      expect(shown(reopened.session, 'body')).toEqual(editor);
      const after = paragraphs(reopened.document);
      expect(after[7]!.formatting).toMatchObject({ indentLeft: 0 });
      expect(after[8]!.formatting).toMatchObject({
        indentFirstLine: 0,
        tabs: [{ position: 1000, alignment: 'clear' }, { position: 5000, alignment: 'clear' }],
      });
    } finally {
      session.destroy();
      reopened?.session.destroy();
    }
  });

  it('saves a hanging first line set from the ruler', async () => {
    const bytes = fixture();
    const source = await parseDocx(bytes.buffer, { preloadFonts: false });
    const session = await createYrsSession({ clientId: 81014 });
    let reopened: Awaited<ReturnType<typeof reopen>> | undefined;
    try {
      session.seedFromDocx(bytes);
      session.setParagraphAttrs(range(session, 0), { indentFirstLine: 360, hangingIndent: true });
      const editor = shown(session, 'body');
      expect(editor[0]).toMatchObject({ indentFirstLine: 360, hangingIndent: true });
      reopened = await reopen(await save(session, source), 81015);
      // The parser signs a hanging first line negative; the editor draws its magnitude either way.
      const hanging = (value: Record<string, unknown> | undefined) => [Math.abs(Number(value?.indentFirstLine)), value?.hangingIndent];
      expect(hanging(paragraphs(reopened.document)[0]!.formatting as Record<string, unknown>)).toEqual([360, true]);
      expect(hanging(shown(reopened.session, 'body')[0])).toEqual([360, true]);
    } finally {
      session.destroy();
      reopened?.session.destroy();
    }
  });

  it('keeps both peers’ changes to one paragraph', async () => {
    const bytes = fixture();
    const source = await parseDocx(bytes.buffer, { preloadFonts: false });
    const left = await createYrsSession({ clientId: 81010 });
    const right = await createYrsSession({ clientId: 81011 });
    try {
      left.seedFromDocx(bytes);
      right.applyUpdate(left.encodeStateAsUpdate());
      left.setParagraphAttrs(range(left, 1), { indentLeft: 900 });
      right.setParagraphAttrs(range(right, 1), { lineSpacing: 360, lineSpacingRule: 'auto' });
      left.applyUpdate(right.encodeStateAsUpdate(left.encodeStateVector()));
      right.applyUpdate(left.encodeStateAsUpdate(right.encodeStateVector()));
      for (const session of [left, right]) {
        const saved = paragraphs(await parseDocx((await save(session, source)).buffer, { preloadFonts: false }))[1]!;
        expect(saved.formatting).toMatchObject({ styleId: 'Title', indentLeft: 900, lineSpacing: 360 });
      }
    } finally {
      left.destroy();
      right.destroy();
    }
  });
});
