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
import {
  applyStyleValues,
  createYrsSession,
  type YrsParagraphAttrs,
  type YrsParagraphTabStop,
  type YrsSession,
} from './index';
import { explicitParagraphAttrs } from './paragraphSeed';
import { cellParagraphFormatting } from './tableParagraphFormatting';
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

function fixture(body = BODY, styles = STYLES): Uint8Array<ArrayBuffer> {
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
    toBytes(`<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr/></w:body></w:document>`)
  );
  parts.set('word/styles.xml', toBytes(styles));
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
      const values = (styleId: string | null) => styleParagraphValues(styles, styleId);
      for (const [index, styleId] of [[0, 'Title'], [1, 'Normal'], [2, 'Tabbed']] as const) {
        applyStyleValues(session, range(session, index), styleId, values);
      }
      const editor = shownAll(session);
      // A paragraph styled anew shows exactly what a fresh one with that style seeds as.
      expect(editor[0]).toMatchObject({ pStyle: 'Title', contextualSpacing: true, spaceAfter: 0, lineSpacing: 240 });
      expect(editor[1]).toMatchObject({ pStyle: 'Normal', spaceAfter: 160, lineSpacing: 259 });
      expect(editor[1]!.contextualSpacing).toBeUndefined();
      // Direct formatting stays, as Word keeps it.
      expect(editor[2]).toMatchObject({ pStyle: 'Tabbed', keepLines: true, indentFirstLine: 200, alignment: 'center' });
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

// Word's built-in looks: Title and Intense Quote have borders, Header tabs at
// 4680/9360, TOC 1 a dot-leader tab; Grid is Table Grid's spacing with a
// centred header row.
const WORD_STYLES =
  `<w:styles xmlns:w="${W}"><w:docDefaults><w:pPrDefault><w:pPr>` +
  '<w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="0"/><w:outlineLvl w:val="0"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:pBdr><w:bottom w:val="single" w:sz="8" w:space="4" w:color="4F81BD"/></w:pBdr>' +
  '<w:spacing w:after="300" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="IntenseQuote"><w:name w:val="Intense Quote"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:pBdr><w:bottom w:val="single" w:sz="4" w:space="4" w:color="4F81BD"/></w:pBdr>' +
  '<w:shd w:val="clear" w:color="auto" w:fill="EEF3FA"/><w:spacing w:before="200" w:after="280"/><w:ind w:left="936" w:right="936"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:tabs><w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/></w:tabs>' +
  '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs><w:spacing w:after="100"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="RtlPara"><w:name w:val="RTL Paragraph"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:bidi/><w:snapToGrid w:val="0"/><w:autoSpaceDE w:val="0"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="ListNumber"><w:name w:val="List Number"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr><w:contextualSpacing/></w:pPr></w:style>' +
  '<w:style w:type="table" w:styleId="Grid"><w:name w:val="Table Grid"/>' +
  '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' +
  '<w:tblStylePr w:type="firstRow"><w:pPr><w:jc w:val="center"/></w:pPr></w:tblStylePr>' +
  '<w:tblStylePr w:type="lastRow"><w:pPr><w:jc w:val="right"/></w:pPr></w:tblStylePr></w:style></w:styles>';

const WORD_LOOKS = ['Title', 'IntenseQuote', 'Header', 'TOC1', 'RtlPara'] as const;
const styled = (styleId: string, text: string) =>
  `<w:p><w:pPr><w:pStyle w:val="${styleId}"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
const plain = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const gridCell = (text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${plain(text)}</w:tc>`;
const WORD_BODY = [
  ...WORD_LOOKS.map((styleId) => styled(styleId, styleId)),
  ...WORD_LOOKS.map((styleId) => plain(`to ${styleId}`)),
  plain('Contested'),
  '<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/><w:tblW w:w="0" w:type="auto"/>' +
    '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="1" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/></w:tblPr>' +
    `<w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr>${gridCell('Head')}</w:tr><w:tr>${gridCell('Body')}</w:tr>` +
    `<w:tr>${gridCell('Last')}</w:tr></w:tbl>`,
  plain('Tail'),
  styled('ListNumber', 'Numbered by style'),
  plain('To number'),
  '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Numbered directly</w:t></w:r></w:p>',
  // Direct formatting a style change keeps (the review's direct-keep shapes), then a mark with its own size.
  '<w:p><w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr><w:r><w:t>RTL right</w:t></w:r></w:p>',
  '<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="000000"/>' +
    '<w:bottom w:val="single" w:sz="4" w:space="1" w:color="000000"/></w:pBdr>' +
    '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/></w:pPr><w:r><w:t>Boxed</w:t></w:r></w:p>',
  '<w:p><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8000"/></w:tabs></w:pPr><w:r><w:t>Tabbed</w:t></w:r></w:p>',
  '<w:p><w:pPr><w:ind w:left="1440"/><w:jc w:val="center"/></w:pPr><w:r><w:t>Centred</w:t></w:r></w:p>',
  '<w:p><w:pPr><w:spacing w:before="80" w:beforeLines="100" w:beforeAutospacing="1"/></w:pPr><w:r><w:t>Spaced</w:t></w:r></w:p>',
  '<w:p><w:pPr><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="40"/></w:rPr></w:pPr><w:r><w:t>Big mark</w:t></w:r></w:p>',
  '<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:spacing w:before="240"/></w:pPr><w:r><w:t>Spaced heading</w:t></w:r></w:p>',
].join('');
const NUMBERED = WORD_LOOKS.length * 2 + 2;
const DIRECT = NUMBERED + 3;

describe('applying a style', () => {
  beforeAll(() =>
    preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))))
  );

  /** What the editor's style picker does: the style's values with the cell's table style and the package's numbering. */
  const applyStyle = (session: YrsSession, source: Document, story: string, index: number, styleId: string) => {
    const { paraId } = session.paragraphs(story)[index]!;
    const styles = createStyleResolver(source.package.styles);
    const values = (id: string | null, at: string) =>
      styleParagraphValues(styles, id, {
        cell: cellParagraphFormatting(session, styles, at),
        numbering: source.package.numbering,
      });
    applyStyleValues(session, { story, start: { paraId, offset: 0 }, end: { paraId, offset: 0 } }, styleId, values);
  };
  /** The body, then its table's cell stories in row order (ids say where a cell was made, not where it is). */
  const stories = (session: YrsSession) => [
    'body',
    ...session.storySegments('body').flatMap((segment) =>
      segment.kind === 'embed' && segment.embedKind === 'table'
        ? ((segment.payload.rows ?? []) as Array<{ cells?: Array<{ story: string }> }>).flatMap((row) =>
            (row.cells ?? []).map((cell) => cell.story)
          )
        : []
    ),
  ];
  const shownStories = (session: YrsSession) => stories(session).map((story) => shown(session, story));

  it('takes and leaves Word’s borders, tabs and right-to-left with the style alone', async () => {
    const bytes = fixture(WORD_BODY, WORD_STYLES);
    const source = await parseDocx(bytes.buffer, { preloadFonts: false });
    const session = await createYrsSession({ clientId: 82001 });
    const fresh = await createYrsSession({ clientId: 82002 });
    let reopened: Awaited<ReturnType<typeof reopen>> | undefined;
    try {
      session.seedFromDocx(bytes);
      fresh.seedFromDocx(bytes);
      WORD_LOOKS.forEach((styleId, index) => {
        applyStyle(session, source, 'body', index, 'Normal');
        applyStyle(session, source, 'body', WORD_LOOKS.length + index, styleId);
      });
      const editor = shown(session, 'body');
      const seeded = shown(fresh, 'body');
      // Restyled paragraphs show what a paragraph seeded with the new style shows.
      WORD_LOOKS.forEach((_, index) => {
        expect(editor[index]).toEqual({ ...seeded[WORD_LOOKS.length + index], pStyle: 'Normal' });
        expect(editor[WORD_LOOKS.length + index]).toEqual(seeded[index]);
      });
      expect(editor[0]!.borders).toBeUndefined();
      expect(editor[WORD_LOOKS.length + 3]).toMatchObject({ tabs: [{ position: 9350, alignment: 'right', leader: 'dot' }] });
      expect(editor[WORD_LOOKS.length + 4]).toMatchObject({ bidi: true, snapToGrid: false, autoSpaceDE: false });
      reopened = await reopen(await save(session, source), 82003);
      expect(shown(reopened.session, 'body')).toEqual(editor);
      const after = paragraphs(reopened.document);
      WORD_LOOKS.forEach((styleId, index) => {
        expect(after[index]!.formatting).toEqual({ styleId: 'Normal' });
        expect(after[WORD_LOOKS.length + index]!.formatting).toEqual({ styleId });
      });
    } finally {
      session.destroy();
      fresh.destroy();
      reopened?.session.destroy();
    }
  });

  it('keeps a table style’s spacing and header-row alignment in a cell', async () => {
    const bytes = fixture(WORD_BODY, WORD_STYLES);
    const source = await parseDocx(bytes.buffer, { preloadFonts: false });
    const session = await createYrsSession({ clientId: 82004 });
    let reopened: Awaited<ReturnType<typeof reopen>> | undefined;
    try {
      session.seedFromDocx(bytes);
      const before = shownStories(session);
      const cells = stories(session).filter((id) => id !== 'body');
      for (const cell of cells) applyStyle(session, source, cell, 0, 'Normal');
      const editor = shownStories(session);
      expect(editor).toEqual(before.map((story, index) => (index === 0 ? story : [{ ...story[0], pStyle: 'Normal' }])));
      expect(editor[1]![0]).toMatchObject({ alignment: 'center', spaceAfter: 0, lineSpacing: 240 });
      expect(editor[2]![0]).toMatchObject({ spaceAfter: 0, lineSpacing: 240 });
      expect(editor[3]![0]).toMatchObject({ alignment: 'right', spaceAfter: 0 });
      reopened = await reopen(await save(session, source), 82005);
      expect(shownStories(reopened.session)).toEqual(editor);
      const saved = paragraphs(reopened.document).slice(WORD_LOOKS.length * 2 + 1, WORD_LOOKS.length * 2 + 4);
      expect(saved.map((paragraph) => paragraph.formatting)).toEqual(cells.map(() => ({ styleId: 'Normal' })));
    } finally {
      session.destroy();
      reopened?.session.destroy();
    }
  });

  it('converges on one style’s values when two peers apply different styles', async () => {
    const bytes = fixture(WORD_BODY, WORD_STYLES);
    const source = await parseDocx(bytes.buffer, { preloadFonts: false });
    const left = await createYrsSession({ clientId: 82006 });
    const right = await createYrsSession({ clientId: 82007 });
    const alone = await createYrsSession({ clientId: 82008 });
    const contested = WORD_LOOKS.length * 2;
    try {
      left.seedFromDocx(bytes);
      right.applyUpdate(left.encodeStateAsUpdate());
      alone.applyUpdate(left.encodeStateAsUpdate());
      applyStyle(left, source, 'body', contested, 'Heading1');
      applyStyle(right, source, 'body', contested, 'IntenseQuote');
      left.applyUpdate(right.encodeStateAsUpdate(left.encodeStateVector()));
      right.applyUpdate(left.encodeStateAsUpdate(right.encodeStateVector()));
      const merged = shown(left, 'body')[contested]!;
      expect(shown(right, 'body')[contested]).toEqual(merged);
      applyStyle(alone, source, 'body', contested, merged.pStyle as string);
      expect(merged).toEqual(shown(alone, 'body')[contested]);
    } finally {
      left.destroy();
      right.destroy();
      alone.destroy();
    }
  });

  it('reads a cell’s table style from the session when the editor holds only the host document', async () => {
    const bytes = fixture(WORD_BODY, WORD_STYLES);
    const session = await createYrsSession({ clientId: 82010 });
    let reopened: Awaited<ReturnType<typeof reopen>> | undefined;
    try {
      // As Capy loads it: openDocx, whose host document carries styles but no body.
      const host = session.openDocx(bytes, true);
      expect(host.document.package.document.content).toEqual([]);
      const base = session.materializeDocx()!;
      const styles = createStyleResolver(host.document.package.styles);
      const cells = stories(session).filter((id) => id !== 'body');
      const before = cells.map((story) => shown(session, story)[0]);
      for (const story of cells) {
        const { paraId } = session.paragraphs(story)[0]!;
        applyStyleValues(
          session,
          { story, start: { paraId, offset: 0 }, end: { paraId, offset: 0 } },
          'Normal',
          (id, at) => styleParagraphValues(styles, id, { cell: cellParagraphFormatting(session, styles, at) })
        );
      }
      expect(cells.map((story) => shown(session, story)[0])).toEqual(before.map((paragraph) => ({ ...paragraph, pStyle: 'Normal' })));
      reopened = await reopen(await save(session, base), 82011);
      const saved = paragraphs(reopened.document).slice(WORD_LOOKS.length * 2 + 1, WORD_LOOKS.length * 2 + 4);
      expect(saved.map((paragraph) => paragraph.formatting)).toEqual(cells.map(() => ({ styleId: 'Normal' })));
    } finally {
      session.destroy();
      reopened?.session.destroy();
    }
  });

  it('compares a table or row made in the session against its own table style', async () => {
    const bytes = fixture(WORD_BODY, WORD_STYLES);
    const source = await parseDocx(bytes.buffer, { preloadFonts: false });
    const session = await createYrsSession({ clientId: 82012 });
    const tableSession = await createYrsSession({ clientId: 82013 });
    let reopened: Awaited<ReturnType<typeof reopen>> | undefined;
    const typed = (document: Document) =>
      paragraphs(document).find((paragraph) =>
        paragraph.content.some((run) => run.type === 'run' && run.content.some((item) => item.type === 'text' && item.text.startsWith('typed ')))
      )?.formatting;
    const centre = (target: YrsSession, story: string, alignment: 'center' | 'right') => {
      const { paraId } = target.paragraphs(story)[0]!;
      target.insertText({ story, paraId, offset: 0 }, 'typed ');
      target.setParagraphAttrs({ story, start: { paraId, offset: 0 }, end: { paraId, offset: 0 } }, { alignment });
    };
    try {
      // The last row goes and a row above the header takes its story id; the user right-aligns it.
      session.seedFromDocx(bytes);
      const lastRow = { story: 'body', tableIndex: 0, row: 2, column: 0 };
      session.deleteRow({ anchor: lastRow, head: lastRow });
      session.insertRow({ story: 'body', tableIndex: 0, row: 0, column: 0 }, 'above');
      const top = stories(session)[1]!;
      centre(session, top, 'right');
      reopened = await reopen(await save(session, source), 82014);
      expect(typed(reopened.document)).toMatchObject({ alignment: 'right' });
      reopened.session.destroy();

      // The table goes and a new one takes its id; the user centres its paragraph.
      tableSession.seedFromDocx(bytes);
      tableSession.deleteTable({ story: 'body', tableIndex: 0 });
      const before = tableSession.paragraphs('body')[WORD_LOOKS.length * 2]!;
      tableSession.insertTable({ story: 'body', paraId: before.paraId, offset: before.text.length }, 1, 1);
      const cell = stories(tableSession).find((id) => id !== 'body')!;
      centre(tableSession, cell, 'center');
      reopened = await reopen(await save(tableSession, source), 82015);
      expect(typed(reopened.document)).toMatchObject({ alignment: 'center' });
    } finally {
      session.destroy();
      tableSession.destroy();
      reopened?.session.destroy();
    }
  });

  it('gives a row added in the session its table style when a style is applied in it', async () => {
    const bytes = fixture(WORD_BODY, WORD_STYLES);
    const source = await parseDocx(bytes.buffer, { preloadFonts: false });
    const session = await createYrsSession({ clientId: 82016 });
    let reopened: Awaited<ReturnType<typeof reopen>> | undefined;
    try {
      session.seedFromDocx(bytes);
      const before = new Set(stories(session));
      session.insertRow({ story: 'body', tableIndex: 0, row: 1, column: 0 }, 'below');
      const added = stories(session).find((id) => !before.has(id))!;
      applyStyle(session, source, added, 0, 'Normal');
      expect(shown(session, added)[0]).toMatchObject({ pStyle: 'Normal', spaceAfter: 0, lineSpacing: 240 });
      const editor = shownStories(session);
      reopened = await reopen(await save(session, source), 82017);
      expect(shownStories(reopened.session)).toEqual(editor);
    } finally {
      session.destroy();
      reopened?.session.destroy();
    }
  });

  it('takes and leaves a style’s numbering, keeping numbering set on the paragraph', async () => {
    const bytes = fixture(WORD_BODY, WORD_STYLES);
    const source = await parseDocx(bytes.buffer, { preloadFonts: false });
    const session = await createYrsSession({ clientId: 82018 });
    let reopened: Awaited<ReturnType<typeof reopen>> | undefined;
    try {
      session.seedFromDocx(bytes);
      applyStyle(session, source, 'body', NUMBERED, 'Normal');
      applyStyle(session, source, 'body', NUMBERED + 1, 'ListNumber');
      applyStyle(session, source, 'body', NUMBERED + 2, 'Heading1');
      const editor = shown(session, 'body');
      expect(editor[NUMBERED]!.numPr).toBeUndefined();
      expect(editor[NUMBERED + 1]!.numPr).toEqual({ numId: 1 });
      expect(session.paragraphs('body')[NUMBERED + 1]!.properties).toMatchObject({ listNumFmt: 'decimal' });
      expect(editor[NUMBERED + 2]!.numPr).toEqual({ numId: 1, ilvl: 0 });
      reopened = await reopen(await save(session, source), 82019);
      expect(shown(reopened.session, 'body')).toEqual(editor);
    } finally {
      session.destroy();
      reopened?.session.destroy();
    }
  });

  it('keeps direct paragraph formatting and the mark’s own size when a style is applied', async () => {
    const bytes = fixture(WORD_BODY, WORD_STYLES);
    const source = await parseDocx(bytes.buffer, { preloadFonts: false });
    const session = await createYrsSession({ clientId: 82020 });
    let reopened: Awaited<ReturnType<typeof reopen>> | undefined;
    const mark = (target: YrsSession) => target.paragraphs('body')[DIRECT + 5]!.properties.defaultTextFormatting;
    try {
      session.seedFromDocx(bytes);
      const markBefore = mark(session) as { fontSize?: unknown };
      for (let index = DIRECT; index <= DIRECT + 5; index += 1) applyStyle(session, source, 'body', index, 'Heading1');
      const editor = shown(session, 'body');
      expect(editor[DIRECT]).toMatchObject({ pStyle: 'Heading1', bidi: true, alignment: 'right', keepNext: true, outlineLevel: 0 });
      expect(editor[DIRECT + 1]).toMatchObject({ shading: { fill: { rgb: 'FFFF00' } }, spaceBefore: 240 });
      expect(editor[DIRECT + 2]).toMatchObject({ tabs: [{ position: 8000, alignment: 'right', leader: 'dot' }] });
      expect(editor[DIRECT + 3]).toMatchObject({ alignment: 'center', indentLeft: 1440, spaceAfter: 0 });
      expect(editor[DIRECT + 4]).toMatchObject({ spaceBefore: 80, spaceBeforeLines: 100, beforeAutospacing: true });
      expect((mark(session) as { fontSize?: unknown }).fontSize).toEqual(markBefore.fontSize);
      reopened = await reopen(await save(session, source), 82021);
      expect(shown(reopened.session, 'body')).toEqual(editor);
      expect(mark(reopened.session)).toEqual(mark(session));
      // The file keeps each paragraph's own pPr under the new style, as Word writes it
      // (document order counts the table's three cell paragraphs before these).
      const before = paragraphs(source);
      const after = paragraphs(reopened.document);
      for (let index = DIRECT + 3; index <= DIRECT + 8; index += 1) {
        expect(after[index]!.formatting).toEqual({ ...before[index]!.formatting, styleId: 'Heading1' });
      }
    } finally {
      session.destroy();
      reopened?.session.destroy();
    }
  });

  it('saves no list indents for a style’s list, applied or taken away', async () => {
    const bytes = fixture(WORD_BODY, WORD_STYLES);
    const source = await parseDocx(bytes.buffer, { preloadFonts: false });
    // Body paragraphs after the table sit three cell paragraphs later in document order.
    const inFile = (index: number) => index + 3;
    for (const [index, styles, clientId] of [
      // A plain paragraph, and one with its own centring and indent (no source paragraph has its key).
      [NUMBERED + 1, ['ListNumber'], 82024],
      [DIRECT + 3, ['ListNumber'], 82026],
      [NUMBERED + 1, ['ListNumber', 'Normal'], 82028],
      // Numbering of its own stays, with its level's indents unwritten.
      [NUMBERED + 2, ['Heading1'], 82032],
      [NUMBERED + 2, ['ListNumber'], 82034],
    ] as const) {
      const session = await createYrsSession({ clientId });
      let reopened: Awaited<ReturnType<typeof reopen>> | undefined;
      try {
        session.seedFromDocx(bytes);
        for (const styleId of styles) applyStyle(session, source, 'body', index, styleId);
        const editor = shown(session, 'body');
        reopened = await reopen(await save(session, source), clientId + 1);
        expect(shown(reopened.session, 'body')).toEqual(editor);
        const { numPrFromStyle, ...saved } = paragraphs(reopened.document)[inFile(index)]!.formatting!;
        // The parser lists a style's numbering on the paragraph; the file holds only pStyle.
        if (numPrFromStyle) {
          expect(saved.numPr).toEqual(numPrFromStyle);
          delete saved.numPr;
        }
        expect(saved).toEqual({ ...paragraphs(source)[inFile(index)]!.formatting, styleId: styles.at(-1) });
      } finally {
        session.destroy();
        reopened?.session.destroy();
      }
    }
  });

  it('keeps a value the paragraph set that equals its old style’s', async () => {
    const bytes = fixture(WORD_BODY, WORD_STYLES);
    const source = await parseDocx(bytes.buffer, { preloadFonts: false });
    const session = await createYrsSession({ clientId: 82030 });
    let reopened: Awaited<ReturnType<typeof reopen>> | undefined;
    const heading = DIRECT + 6;
    try {
      session.seedFromDocx(bytes);
      applyStyle(session, source, 'body', heading, 'Normal');
      const editor = shown(session, 'body');
      expect(editor[heading]).toMatchObject({ pStyle: 'Normal', spaceBefore: 240 });
      reopened = await reopen(await save(session, source), 82031);
      expect(shown(reopened.session, 'body')).toEqual(editor);
      expect(paragraphs(reopened.document)[heading + 3]!.formatting).toEqual({
        ...paragraphs(source)[heading + 3]!.formatting,
        styleId: 'Normal',
      });
    } finally {
      session.destroy();
      reopened?.session.destroy();
    }
  });

  it('stores 0 when the ruler clears the indents a list level gives', async () => {
    const bytes = fixture(WORD_BODY, WORD_STYLES);
    const source = await parseDocx(bytes.buffer, { preloadFonts: false });
    const session = await createYrsSession({ clientId: 82022 });
    let reopened: Awaited<ReturnType<typeof reopen>> | undefined;
    try {
      session.seedFromDocx(bytes);
      const item = NUMBERED + 2;
      expect(shown(session, 'body')[item]).toMatchObject({ indentLeft: 720, hangingIndent: true });
      const style = styleParagraphValues(createStyleResolver(source.package.styles), null);
      session.setParagraphAttrs(
        range(session, item),
        explicitParagraphAttrs({ indentLeft: null, indentFirstLine: null, hangingIndent: false }, style, true)
      );
      const editor = shown(session, 'body');
      expect(editor[item]).toMatchObject({ indentLeft: 0, indentFirstLine: 0 });
      expect(editor[item]!.hangingIndent).toBeUndefined();
      reopened = await reopen(await save(session, source), 82023);
      expect(shown(reopened.session, 'body')).toEqual(editor);
    } finally {
      session.destroy();
      reopened?.session.destroy();
    }
  });

  it('removes a tab stop the ruler holds in the old shape', async () => {
    const bytes = fixture(WORD_BODY, WORD_STYLES);
    const source = await parseDocx(bytes.buffer, { preloadFonts: false });
    const session = await createYrsSession({ clientId: 82009 });
    try {
      session.seedFromDocx(bytes);
      // A Header paragraph whose stops an older op stored as {pos, val}; the ruler drops 4680.
      const { paraId } = session.paragraphs('body')[2]!;
      session.setParagraphAttr(paraId, 'tabs', [{ pos: 4680, val: 'center' }, { pos: 9360, val: 'right' }]);
      const style = styleParagraphValues(createStyleResolver(source.package.styles), 'Header');
      session.setParagraphAttrs(
        { story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 0 } },
        explicitParagraphAttrs({ tabs: [{ pos: 9360, val: 'right' }] as never }, style)
      );
      expect(session.paragraphs('body')[2]!.properties.tabs).toEqual([
        { position: 9360, alignment: 'right' },
        { position: 4680, alignment: 'clear' },
      ]);
    } finally {
      session.destroy();
    }
  });
});
