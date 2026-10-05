import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import {
  buildMirrorPage,
  buildMirrorTextPage,
  type DisplayPage,
} from '@betteroffice/docx/layout/render';
import { configureDefaultFonts, initWasm, openDocumentViewer } from '@betteroffice/docx/viewer';
import { preloadOpcWasm, rezipContainer } from '@betteroffice/docx/wasm/opc';
import { markerText, selectAround, selectStory, textLayerText, useTextLayer } from './textLayer';

const { renderHook } = await import('@testing-library/react');

afterEach(() => {
  document.body.replaceChildren();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

/** One run per primitive; `line` moves the baseline down a line. */
function run(text: string, docStart: number | undefined, x: number, extra: Record<string, unknown> = {}) {
  const line = (extra.line as number | undefined) ?? 0;
  return {
    kind: 'text',
    text,
    x,
    baselineY: 100 + line * 20,
    width: text.length * 8,
    font: '400 16px Calibri',
    color: '#000000',
    ...(docStart === undefined ? {} : { docStart, docEnd: docStart + text.length }),
    lineIndex: line,
    ...extra,
  };
}

const cell = (row: number, col: number, start: number, text: string, extra: Record<string, unknown> = {}) =>
  run(text, start, 100 + col * 200, {
    blockKey: `body:t0:r${row}c${col}:p0`,
    fragmentDocStart: start - 1,
    cell: { row, col, rowSpan: 1, colSpan: 1 },
    table: { tableId: 't0' },
    line: 4 + row,
    ...extra,
  });

// Header, list number, tab, line break, wrapped line, a 2x2 table with a run
// clipped out of its row, a paragraph starting with a tab, and a run after
// positions that paint nothing (hidden text).
const page = {
  pageIndex: 0,
  width: 800,
  height: 1000,
  header: {
    rId: 'rId1',
    kind: 'header',
    y: 20,
    height: 30,
    primitives: [run('Header', 1, 100, { blockKey: 'header:p0', baselineY: 40 })],
  },
  primitives: [
    run('1.', undefined, 80, { blockKey: 'body:p0', listMarker: true }),
    run('Name', 1, 100, { blockKey: 'body:p0', fragmentDocStart: 0 }),
    run('Value', 6, 200, { blockKey: 'body:p0', tabsBefore: 1, breaksAfter: 1 }),
    run('Second', 12, 100, { blockKey: 'body:p0', line: 1 }),
    run('Wrapped ', 20, 100, { blockKey: 'body:p1', fragmentDocStart: 19, line: 2 }),
    run('line', 28, 100, { blockKey: 'body:p1', line: 3 }),
    cell(0, 0, 40, 'A1'),
    cell(0, 1, 44, 'B1'),
    cell(1, 0, 48, 'A2'),
    cell(1, 1, 52, 'B2'),
    cell(1, 1, 54, 'cut', { clipGroup: { clip: { x: 0, y: 0, w: 800, h: 50 } } }),
    run('Indented', 62, 140, { blockKey: 'body:p2', fragmentDocStart: 60, tabsBefore: 1, line: 7 }),
    run('More', 76, 204, { blockKey: 'body:p2', line: 7 }),
  ],
} as unknown as DisplayPage;

// A body paragraph with a footnote reference, and the footnote.
const noted = {
  pageIndex: 0,
  width: 800,
  height: 1000,
  primitives: [
    run('Body', 1, 100, { blockKey: 'body:p0' }),
    run('1', 5, 132, { blockKey: 'body:p0', noteRef: { id: 1, kind: 'footnote' } }),
    run('More', 8, 100, { blockKey: 'body:p1', line: 1 }),
  ],
  noteAreas: [
    {
      kind: 'footnote',
      y: 900,
      height: 60,
      noteIds: [1],
      notes: [{ id: 1, label: '1' }],
      primitives: [
        run('1', undefined, 100, { blockKey: 'fn1:p0', groupId: 'footnote-1', baselineY: 920 }),
        run('Note text', 1, 112, { blockKey: 'fn1:p0', groupId: 'footnote-1', baselineY: 920 }),
      ],
    },
  ],
} as unknown as DisplayPage;

const ALL = '1.\tName\tValue\nSecond\nWrapped line\nA1\tB1\nA2\tB2\n\tIndentedMore';

function mount(...mirrors: HTMLElement[]): HTMLElement {
  const root = document.createElement('div');
  root.append(...mirrors);
  document.body.append(root);
  return root;
}

function textNode(root: HTMLElement, text: string): Text {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    if ((node as Text).data === text) return node as Text;
  }
  throw new Error(`no text node ${text}`);
}

function rangeOver(root: HTMLElement, from: [string, number], to: [string, number]): Range {
  const range = document.createRange();
  range.setStart(textNode(root, from[0]), from[1]);
  range.setEnd(textNode(root, to[0]), to[1]);
  return range;
}

function copyAll(root: HTMLElement, pages: readonly DisplayPage[]): string {
  const range = document.createRange();
  range.selectNodeContents(root);
  return textLayerText(range, root, pages);
}

describe('text layer copy', () => {
  test('copies a whole page as the editor would, with list numbers as shown', () => {
    expect(copyAll(mount(buildMirrorPage(page)), [page])).toBe(ALL);
  });

  test('copies the selected part of the first and last runs', () => {
    const root = mount(buildMirrorPage(page));
    const range = rangeOver(root, ['Value', 2], ['A2', 1]);
    expect(textLayerText(range, root, [page])).toBe('lue\nSecond\nWrapped line\nA1\tB1\nA');
  });

  test('reads a far page from its positioned form at the same offsets', () => {
    const root = mount(buildMirrorTextPage(page));
    expect(copyAll(root, [page])).toBe(ALL);
    const part = rangeOver(root, ['Value', 2], ['A2', 1]);
    expect(textLayerText(part, root, [page])).toBe('lue\nSecond\nWrapped line\nA1\tB1\nA');
  });

  test('a row split across pages copies as two rows, without leading tabs', () => {
    const before = {
      pageIndex: 0,
      width: 800,
      height: 1000,
      primitives: [cell(0, 0, 80, 'ab'), cell(0, 1, 84, 'cd')],
    } as unknown as DisplayPage;
    const after = {
      pageIndex: 1,
      width: 800,
      height: 1000,
      primitives: [cell(0, 1, 86, 'ef', { fragmentDocStart: 83 })],
    } as unknown as DisplayPage;
    const root = mount(buildMirrorPage(before), buildMirrorPage(after));
    expect(copyAll(root, [before, after])).toBe('ab\tcd\nef');
  });

  test('a table header row repeated on the next page copies once', () => {
    const header = (col: number, text: string, start: number, repeatedHeader?: boolean) =>
      cell(0, col, start, text, {
        cell: { row: 0, col, rowSpan: 1, colSpan: 1, isHeader: true, repeatedHeader },
      });
    const first = {
      pageIndex: 0,
      width: 800,
      height: 1000,
      primitives: [header(0, 'Head A', 2), header(1, 'Head B', 10), cell(1, 0, 20, 'r1')],
    } as unknown as DisplayPage;
    const second = {
      pageIndex: 1,
      width: 800,
      height: 1000,
      primitives: [header(0, 'Head A', 2, true), header(1, 'Head B', 10, true), cell(2, 0, 30, 'r2')],
    } as unknown as DisplayPage;
    const root = mount(buildMirrorPage(first), buildMirrorPage(second));
    expect(copyAll(root, [first, second])).toBe('Head A\tHead B\nr1\nr2');
  });

  test('footnotes are their own story: left out of a body selection, copied from inside', () => {
    for (const mirror of [buildMirrorPage, buildMirrorTextPage]) {
      const root = mount(mirror(noted));
      expect(copyAll(root, [noted])).toBe('Body1\nMore');
      const fromNote = document.createRange();
      fromNote.setStart(textNode(root, 'Note text'), 0);
      fromNote.setEndAfter(root.lastChild!);
      expect(textLayerText(fromNote, root, [noted])).toBe('Note text');
      document.body.replaceChildren();
    }
  });

  test('footnotes and endnotes are separate stories, and select all keeps to the notes', () => {
    const area = (kind: string, id: number, text: string) => ({
      kind,
      y: kind === 'footnote' ? 880 : 940,
      height: 40,
      noteIds: [id],
      notes: [{ id, label: String(id) }],
      primitives: [
        run(text, 1, 112, { blockKey: `${kind}${id}:p0`, groupId: `${kind}-${id}`, baselineY: 900 }),
      ],
    });
    const pageOf = (pageIndex: number, body: string, notes: unknown[]) =>
      ({
        pageIndex,
        width: 800,
        height: 1000,
        primitives: [run(body, 1 + pageIndex * 10, 100, { blockKey: `body:p${pageIndex}` })],
        noteAreas: notes,
      }) as unknown as DisplayPage;
    const pages = [
      pageOf(0, 'One', [area('footnote', 1, 'Foot one.'), area('endnote', 1, 'End one.')]),
      pageOf(1, 'Two', [area('footnote', 2, 'Foot two.')]),
    ];
    const root = mount(...pages.map((page) => buildMirrorPage(page)));
    expect(copyAll(root, pages)).toBe('One\nTwo');
    // a drag from a footnote on into the endnote copies the footnote only
    const range = document.createRange();
    range.setStart(textNode(root, 'Foot one.'), 0);
    range.setEnd(textNode(root, 'End one.'), 3);
    expect(textLayerText(range, root, pages)).toBe('Foot one.');

    const selection = document.getSelection()!;
    selection.collapse(textNode(root, 'Foot two.'), 2);
    selectStory(root);
    expect(textLayerText(selection.getRangeAt(0), root, pages, 'footnotes')).toBe(
      'Foot one.\nFoot two.'
    );
    selection.collapse(textNode(root, 'End one.'), 1);
    selectStory(root);
    expect(textLayerText(selection.getRangeAt(0), root, pages, 'endnotes')).toBe('End one.');
    selection.collapse(textNode(root, 'Two'), 1);
    selectStory(root);
    expect(textLayerText(selection.getRangeAt(0), root, pages, 'body')).toBe('One\nTwo');
  });

  test('symbol-font list markers copy as their Unicode character, otherwise a bullet', () => {
    expect(markerText('\uf0fc', 'Wingdings')).toBe('\u2713');
    expect(markerText('\uf0d8', 'Wingdings')).toBe('\u27a2');
    expect(markerText('\uf0a7', 'Wingdings')).toBe('\u25aa');
    expect(markerText('\uf041', 'Wingdings')).toBe('\u2022');
    expect(markerText('\uf0b7', 'Symbol')).toBe('\u2022');
    expect(markerText('\u00b7', 'Symbol, serif')).toBe('\u2022');
    expect(markerText('\uf0b7', 'Calibri')).toBe('\u2022');
    expect(markerText('o', 'Courier New')).toBe('o');
    expect(markerText('1.', 'Calibri')).toBe('1.');
  });

  test('the viewer copies from the story the selection began in, and a press on the page clears it', () => {
    const root = mount(buildMirrorPage(noted));
    const canvas = document.createElement('canvas');
    root.prepend(canvas);
    const { unmount } = renderHook(() => useTextLayer({ current: root }, [noted]));
    const selection = document.getSelection()!;
    const copied = (): string => {
      const event = new ClipboardEvent('copy', {
        bubbles: true,
        cancelable: true,
        clipboardData: new DataTransfer(),
      });
      document.dispatchEvent(event);
      return event.clipboardData!.getData('text/plain');
    };
    const press = (target: Element) =>
      target.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0, detail: 1 }));

    // dragged up from the note into the body: the note is what is highlighted
    selection.setBaseAndExtent(textNode(root, 'Note text'), 4, textNode(root, 'More'), 0);
    expect(copied()).toBe('1Note');
    selection.setBaseAndExtent(textNode(root, 'More'), 0, textNode(root, 'Note text'), 4);
    expect(copied()).toBe('More');

    // a press on empty page area clears the selection, whatever its story
    press(canvas);
    expect(selection.rangeCount === 0 || selection.isCollapsed).toBe(true);
    expect(copied()).toBe('');
    unmount();
  });

  test('double and triple click select the word and the paragraph', () => {
    const root = mount(buildMirrorPage(page));
    const value = textNode(root, 'Value').parentElement!;
    selectAround(value, 2, 'word');
    const selection = document.getSelection()!;
    expect(textLayerText(selection.getRangeAt(0), root, [page])).toBe('Value');
    selectAround(value, 2, 'paragraph');
    expect(textLayerText(selection.getRangeAt(0), root, [page])).toBe('1.\tName\tValue\nSecond');
  });
});

describe('text layer over a laid-out document', () => {
  const W =
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const type = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
  const p = (runs: string, ppr = '') => `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${runs}</w:p>`;
  const t = (text: string, rpr = '') =>
    `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
  const tab = '<w:r><w:tab/></w:r>';
  const br = '<w:r><w:br/></w:r>';
  const tc = (content: string) => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${p(content)}</w:tc>`;
  const sdt = (content: string) =>
    `<w:sdt><w:sdtPr><w:id w:val="9"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;

  function docx(body: string, footnotes = ''): Uint8Array {
    const parts: Record<string, string> = {
      '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${type}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${type}.styles+xml"/><Override PartName="/word/footnotes.xml" ContentType="${type}.footnotes+xml"/></Types>`,
      '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
      'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSt" Type="${R}/styles" Target="styles.xml"/><Relationship Id="rIdFn" Type="${R}/footnotes" Target="footnotes.xml"/></Relationships>`,
      'word/styles.xml': `<w:styles ${W}><w:style w:type="paragraph" w:default="1" w:styleId="Body"><w:name w:val="Body"/></w:style></w:styles>`,
      'word/footnotes.xml': `<w:footnotes ${W}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>${footnotes}</w:footnotes>`,
      'word/document.xml': `<w:document ${W}><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
    };
    return rezipContainer(
      Object.fromEntries(Object.entries(parts).map(([path, xml]) => [path, new TextEncoder().encode(xml)]))
    );
  }

  /** The whole document's copy, from positioned pages and from far (plain-text) pages alike. */
  async function copied(body: string, footnotes = ''): Promise<string> {
    const viewer = await openDocumentViewer(docx(body, footnotes));
    try {
      const { pages } = viewer.displayList();
      const positioned = copyAll(mount(...pages.map((page) => buildMirrorPage(page))), pages);
      document.body.replaceChildren();
      const far = copyAll(mount(...pages.map((page) => buildMirrorTextPage(page))), pages);
      expect(far).toBe(positioned);
      return positioned;
    } finally {
      viewer.dispose();
    }
  }

  beforeAll(async () => {
    const generated = resolve(import.meta.dir, '../../../docx/src/wasm/generated');
    await initWasm(await readFile(resolve(generated, 'viewer/docx_view_wasm_bg.wasm')));
    await preloadOpcWasm(await readFile(resolve(generated, 'opc/ooxml_opc_bg.wasm')));
    // Real metrics, so runs are per glyph cluster and lines wrap as in Capy.
    const font = await readFile(
      resolve(import.meta.dir, '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')
    );
    const load = () =>
      Promise.resolve(font.buffer.slice(font.byteOffset, font.byteOffset + font.byteLength));
    configureDefaultFonts({
      fonts: { createFontProvider: () => ({ resolve: () => load, resolveLastResort: () => load }) },
    });
  });
  afterAll(() => {
    configureDefaultFonts({});
  });

  test('copies tabs, breaks, fields and tables as the source text', async () => {
    const body = [
      p(`${t('Name')}${tab}${t('Value')}${br}${t('Second line')}`),
      p(`${t('Page ')}<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>${t(' of the plan')}`),
      `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${tc(t('A1'))}${tc(t('B1'))}</w:tr><w:tr><w:tc>${p(t('A2'))}${p(t('A2 more'))}</w:tc>${tc(t('B2'))}</w:tr></w:tbl>`,
      p(t('After')),
    ].join('');
    expect(await copied(body)).toBe(
      'Name\tValue\nSecond line\nPage 1 of the plan\nA1\tB1\nA2\nA2 more\tB2\nAfter'
    );
  });

  test('copies only real tabs: not content-control edges or hidden text', async () => {
    const body = [
      `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${tc(t('Name'))}${tc(sdt(t('John')))}</w:tr></w:tbl>`,
      p(`${t('Name')}${tab}${t('Secret', '<w:vanish/>')}${t('Value')}`),
      p(`${t('Secret ', '<w:vanish/>')}${t('Visible')}`),
      p(`${t('Lead')}${tab}${sdt(t('Inside'))}`),
    ].join('');
    expect(await copied(body)).toBe('Name\tJohn\nName\tValue\nVisible\nLead\tInside');
  });

  test('keeps tabs in right-to-left text, around line breaks, at the start and at the end', async () => {
    const body = [
      p(`${tab}${t('Indented')}`),
      p(`${t('שלום', '<w:rtl/>')}${tab}${t('עולם', '<w:rtl/>')}`, '<w:bidi/>'),
      p(`${t('A')}${br}${tab}${t('B')}`),
      p(`${t('C')}${br}${br}${t('D')}`),
      p(`${t('Name')}${tab}`),
      p(`${t('E')}${tab}${br}${tab}${t('F')}`),
      p(t('Next')),
    ].join('');
    expect(await copied(body)).toBe(
      '\tIndented\nשלום\tעולם\nA\n\tB\nC\n\nD\nName\t\nE\t\n\tF\nNext'
    );
  });

  test('copies the text around a page or column break inside a paragraph once, in order', async () => {
    for (const kind of ['page', 'column']) {
      const body = p(`${t('Aa')}<w:r><w:br w:type="${kind}"/></w:r>${t('Bb')}`) + p(t('Cc'));
      expect(await copied(p(t('Lead')) + body)).toBe('Lead\nAa\nBb\nCc');
    }
  });

  test('copies a repeated table header row once', async () => {
    const rows = Array.from({ length: 60 }, (_, i) => `<w:tr>${tc(t(`r${i}a`))}${tc(t(`r${i}b`))}</w:tr>`).join('');
    const text = await copied(
      `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:trPr><w:tblHeader/></w:trPr>${tc(t('Head A'))}${tc(t('Head B'))}</w:tr>${rows}</w:tbl>${p(t('End'))}`
    );
    expect(text.split('Head A')).toHaveLength(2);
    expect(text).toStartWith('Head A\tHead B\nr0a\tr0b\n');
    expect(text).toEndWith('r59a\tr59b\nEnd');
    // two pages of cluster runs, laid out and mirrored twice: slow under load
  }, 30_000);

  test('opens a document with a footnote and leaves the note out of a body copy', async () => {
    const text = await copied(
      p(`${t('Body with a note')}<w:r><w:footnoteReference w:id="1"/></w:r>`) + p(t('Next')),
      `<w:footnote w:id="1">${p(t('The footnote text.'))}</w:footnote>`
    );
    expect(text).toStartWith('Body with a note');
    expect(text).toEndWith('\nNext');
    expect(text).not.toContain('footnote text');
  });
});
