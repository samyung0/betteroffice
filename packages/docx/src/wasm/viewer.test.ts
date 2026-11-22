import { beforeAll, describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { buildResidentRegionLayoutRequest } from '../editor/computeLayout';
import type { DisplayList, DisplayPage } from '../layout/render';
import { configureDefaultFonts, type DocxPageSize, initWasm, openDocumentViewer } from '../viewer';
import { createYrsSession } from '../yrs';
import { preloadEditWasm } from './edit';
import { preloadOpcWasm, rezipContainer } from './opc';
import { DocxViewDocument } from './generated/viewer/docx_view_wasm.js';
import { openViewDocument } from './viewer';

const root = resolve(import.meta.dir, '../../../..');
const fixture = resolve(root, 'poc/fixtures/feature-rich.docx');
const viewerWasm = resolve(
  import.meta.dir,
  'generated/viewer/docx_view_wasm_bg.wasm'
);

beforeAll(async () => {
  await Promise.all([
    initWasm(await readFile(viewerWasm)),
    preloadEditWasm(
      await readFile(
        resolve(import.meta.dir, 'generated/edit/docx_edit_bg.wasm')
      )
    ),
    preloadOpcWasm(
      await readFile(
        resolve(import.meta.dir, 'generated/opc/ooxml_opc_bg.wasm')
      )
    ),
  ]);
});

/** Two sections, a footnote, a declared default paragraph style and TOC styles by id and by name. */
function sectionedPackage(): Uint8Array {
  const W =
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const R =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const type = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
  return rezipContainer(
    Object.fromEntries(
      Object.entries({
        '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${type}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${type}.styles+xml"/><Override PartName="/word/footnotes.xml" ContentType="${type}.footnotes+xml"/></Types>`,
        '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
        'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${R}/footnotes" Target="footnotes.xml"/></Relationships>`,
        'word/styles.xml': `<w:styles ${W}><w:style w:type="paragraph" w:default="1" w:styleId="Body"><w:name w:val="Body"/></w:style><w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/></w:style><w:style w:type="paragraph" w:styleId="Contents2"><w:name w:val="TOC 2"/></w:style><w:style w:type="character" w:styleId="TOC3"><w:name w:val="toc 3"/></w:style></w:styles>`,
        'word/footnotes.xml': `<w:footnotes ${W}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:footnote></w:footnotes>`,
        'word/document.xml': `<w:document ${W}><w:body><w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr><w:r><w:t>One</w:t></w:r></w:p><w:p><w:r><w:t>Two</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p><w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/></w:sectPr></w:body></w:document>`,
      }).map(([path, xml]) => [path, new TextEncoder().encode(xml)])
    )
  );
}

/** A page's text: forty pages hold the text a body needs before the viewer
 * lays its first pages out alone (Q822 A). */
const PAGE_TEXT =
  'The committee reviews every application against the published criteria before the spring deadline, and late submissions wait for the autumn round. '.repeat(
    8
  );

/** `count` pages, each a paragraph ending in a page break. */
function pagedPackage(count: number): Uint8Array {
  const W =
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const R =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const page = (index: number) =>
    `<w:p><w:r><w:t>Page ${index}. ${PAGE_TEXT}</w:t></w:r>${index < count ? '<w:r><w:br w:type="page"/></w:r>' : ''}</w:p>`;
  return rezipContainer(
    Object.fromEntries(
      Object.entries({
        '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
        '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
        'word/document.xml': `<w:document ${W}><w:body>${Array.from({ length: count }, (_, index) => page(index + 1)).join('')}<w:sectPr/></w:body></w:document>`,
      }).map(([path, xml]) => [path, new TextEncoder().encode(xml)])
    )
  );
}

/** `count` pages as `pagedPackage` gives them, under a footer whose page
 * count (NUMPAGES) saves a stale 99. */
function countedPackage(count: number): Uint8Array {
  const W =
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const R =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const type = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
  const page = (index: number) =>
    `<w:p><w:r><w:t>Page ${index}. ${PAGE_TEXT}</w:t></w:r>${index < count ? '<w:r><w:br w:type="page"/></w:r>' : ''}</w:p>`;
  return rezipContainer(
    Object.fromEntries(
      Object.entries({
        '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${type}.document.main+xml"/><Override PartName="/word/footer1.xml" ContentType="${type}.footer+xml"/></Types>`,
        '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
        'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdFooter" Type="${R}/footer" Target="footer1.xml"/></Relationships>`,
        'word/footer1.xml': `<w:ftr ${W}><w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve"> of </w:t></w:r><w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>99</w:t></w:r></w:fldSimple></w:p></w:ftr>`,
        'word/document.xml': `<w:document ${W}><w:body>${Array.from({ length: count }, (_, index) => page(index + 1)).join('')}<w:sectPr><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr></w:body></w:document>`,
      }).map(([path, xml]) => [path, new TextEncoder().encode(xml)])
    )
  );
}

/** The text of the page-count (NUMPAGES) fields on a page, its header and footer included. */
function pageCounts(page: DisplayPage): string[] {
  return [page.primitives, page.header?.primitives ?? [], page.footer?.primitives ?? []]
    .flat()
    .flatMap((primitive) => {
      const field = primitive as { field?: { category?: string }; text?: string };
      return field.field?.category === 'NUMPAGES' ? [field.text ?? ''] : [];
    });
}

/** Glyph runs past the first line of their block: measured text that wrapped. */
function wrappedGlyphRuns(displayList: DisplayList): number {
  return displayList.pages
    .flatMap((page) => page.primitives)
    .filter((primitive) => primitive.kind === 'glyphRun' && (primitive.lineIndex ?? 0) > 0).length;
}

describe('DOCX viewer wasm', () => {
  test('rejects bytes that are not a DOCX package', async () => {
    await expect(
      openDocumentViewer(Uint8Array.of(0x00, 0x01, 0x02, 0x03))
    ).rejects.toThrow();
  });

  test('renders the feature-rich fixture without exposing editing methods', async () => {
    const viewer = await openDocumentViewer(
      new Uint8Array(await readFile(fixture))
    );
    try {
      const displayList = viewer.displayList();
      expect(displayList.pages.length).toBeGreaterThan(0);
      expect(displayList.pages.some((page) => page.primitives.length > 0)).toBe(
        true
      );
      expect('save' in viewer).toBe(false);
      expect('applyInput' in viewer).toBe(false);
      expect('encodeStateAsUpdate' in viewer).toBe(false);
    } finally {
      viewer.dispose();
    }
  });

  test('builds the display list at most eight pages at a time, then frees the engine', async () => {
    const ranges: [number, number][] = [];
    let frees = 0;
    const build = DocxViewDocument.prototype.displayPagesJson;
    const free = DocxViewDocument.prototype.free;
    DocxViewDocument.prototype.displayPagesJson = function (start, end) {
      ranges.push([start, end]);
      return build.call(this, start, end);
    };
    DocxViewDocument.prototype.free = function () {
      frees += 1;
      free.call(this);
    };
    try {
      const viewer = await openDocumentViewer(pagedPackage(20));
      expect(frees).toBe(1);
      const { pages } = viewer.displayList();
      viewer.dispose();
      expect(frees).toBe(1);
      expect(pages).toHaveLength(20);
      expect(pages.map((page) => page.pageIndex)).toEqual(pages.map((_, index) => index));
      expect(ranges).toEqual(
        Array.from({ length: Math.ceil(pages.length / 8) }, (_, batch) => [
          batch * 8,
          Math.min(pages.length, batch * 8 + 8),
        ])
      );
    } finally {
      DocxViewDocument.prototype.displayPagesJson = build;
      DocxViewDocument.prototype.free = free;
    }
  });

  test('a body too short to cut hands over the first pages before building the rest, then the rest in batches', async () => {
    const steps: string[] = [];
    const build = DocxViewDocument.prototype.displayPagesJson;
    DocxViewDocument.prototype.displayPagesJson = function (start, end) {
      steps.push(`build ${start}-${end}`);
      return build.call(this, start, end);
    };
    try {
      let sizes: DocxPageSize[] = [];
      const viewer = await openDocumentViewer(pagedPackage(12), {
        progressive: {
          firstPages(all, width) {
            sizes = [...all];
            expect(width).toBe(Math.max(...all.map((size) => size.width)));
            return 2;
          },
          onPages({ pages }) {
            steps.push(`pages ${pages[0]!.pageIndex}-${pages.at(-1)!.pageIndex + 1}`);
          },
        },
      });
      expect(steps).toEqual([
        'build 0-2',
        'pages 0-2',
        'build 2-10',
        'pages 2-10',
        'build 10-12',
        'pages 10-12',
      ]);
      const { pages } = viewer.displayList();
      viewer.dispose();
      expect(sizes).toEqual(pages.map(({ width, height }) => ({ width, height })));
    } finally {
      DocxViewDocument.prototype.displayPagesJson = build;
    }
  });

  // Q801 B: the first screen's pages are laid out alone and shown while the rest is laid out.
  test('lays out and hands over the first pages before laying out the rest, ending as a whole open', async () => {
    const bytes = pagedPackage(40);
    const whole = (await openDocumentViewer(bytes)).displayList().pages;
    const steps: string[] = [];
    const layoutFirst = DocxViewDocument.prototype.layoutFirst;
    const layout = DocxViewDocument.prototype.layout;
    DocxViewDocument.prototype.layoutFirst = function (request, pages) {
      const first = layoutFirst.call(this, request, pages);
      steps.push(`layout first ${pages}: ${first}`);
      return first;
    };
    DocxViewDocument.prototype.layout = function (request) {
      steps.push('layout');
      return layout.call(this, request);
    };
    try {
      const received: DisplayPage[] = [];
      let width = 0;
      const viewer = await openDocumentViewer(bytes, {
        progressive: {
          firstPages(sizes) {
            steps.push(`first pages of ${sizes.length}`);
            // The first screen wants a fourth page: it is laid out too.
            return 4;
          },
          onPages({ pages, width: widest }) {
            width = widest;
            steps.push(`pages ${pages[0]!.pageIndex}-${pages.at(-1)!.pageIndex + 1}`);
            for (const page of pages) received[page.pageIndex] = page;
          },
        },
      });
      expect(steps.slice(0, 7)).toEqual([
        'layout first 3: 3',
        'first pages of 3',
        'layout first 4: 4',
        'first pages of 4',
        'pages 0-4',
        'layout',
        'pages 4-12',
      ]);
      expect(width).toBe(whole[0]!.width);
      expect(received).toEqual(whole);
      expect(viewer.displayList().pages).toEqual(whole);
    } finally {
      DocxViewDocument.prototype.layoutFirst = layoutFirst;
      DocxViewDocument.prototype.layout = layout;
    }
  });

  // Q823 A: a citation or a print waits for every page, so it skips the first pages' layout.
  test('lays the whole body out first for an open that waits for every page', async () => {
    const steps: string[] = [];
    const layoutFirst = DocxViewDocument.prototype.layoutFirst;
    const layout = DocxViewDocument.prototype.layout;
    DocxViewDocument.prototype.layoutFirst = function (request, pages) {
      steps.push('layout first');
      return layoutFirst.call(this, request, pages);
    };
    DocxViewDocument.prototype.layout = function (request) {
      steps.push('layout');
      return layout.call(this, request);
    };
    try {
      const viewer = await openDocumentViewer(pagedPackage(40), {
        progressive: {
          firstPages: () => 2,
          onPages({ pages }) {
            steps.push(`pages ${pages[0]!.pageIndex}-${pages.at(-1)!.pageIndex + 1}`);
          },
          whole: true,
        },
      });
      expect(steps.slice(0, 3)).toEqual(['layout', 'pages 0-2', 'pages 2-10']);
      expect(viewer.displayList().pages).toHaveLength(40);
    } finally {
      DocxViewDocument.prototype.layoutFirst = layoutFirst;
      DocxViewDocument.prototype.layout = layout;
    }
  });

  test('shows the saved page count on the first pages, then sends them again with the real one', async () => {
    const sent: { counts: string[]; index: number }[] = [];
    const viewer = await openDocumentViewer(countedPackage(40), {
      progressive: {
        firstPages: () => 3,
        onPages({ pages }) {
          for (const page of pages) sent.push({ counts: pageCounts(page), index: page.pageIndex });
        },
      },
    });
    const pages = viewer.displayList().pages;
    expect(pages).toHaveLength(40);
    expect(pages.map(pageCounts)).toEqual(pages.map(() => ['40']));
    expect(sent.slice(0, 6)).toEqual([
      { counts: ['99'], index: 0 },
      { counts: ['99'], index: 1 },
      { counts: ['99'], index: 2 },
      { counts: ['40'], index: 0 },
      { counts: ['40'], index: 1 },
      { counts: ['40'], index: 2 },
    ]);
  });

  test('builds the layout request the editor builds', async () => {
    for (const bytes of [
      new Uint8Array(await readFile(fixture)),
      sectionedPackage(),
    ]) {
      const view = openViewDocument(bytes);
      const session = await createYrsSession();
      try {
        session.openDocx(bytes, true);
        const document = session.materializeDocx();
        if (!document) throw new Error('missing package');
        const editor = buildResidentRegionLayoutRequest(document, 24, {
          themeColors: { ...document.package.theme?.colorScheme },
          defaultTabStopTwips:
            document.package.settings?.defaultTabStop ?? null,
          numericIds: {},
        });
        // bodyFontSize sizes the viewer's reading layout; the editor never reads it.
        const { bodyFontSize, ...request } = JSON.parse(view.layoutRequestJson());
        expect(bodyFontSize).toBeGreaterThan(0);
        expect(request).toEqual(JSON.parse(JSON.stringify(editor)));
      } finally {
        view.free();
        session.destroy();
      }
    }
  });

  // Fallback metrics wrap too since upstream #1766 (one em per character, kept by Q604), so the
  // line count no longer tells them apart: measured text is drawn as glyph runs. Since #1931
  // the viewer measures with the bundled fonts by default, so there is no unmeasured baseline.
  test('measures with the configured fonts, so long paragraphs wrap at their glyph widths', async () => {
    const bytes = new Uint8Array(await readFile(fixture));
    const font = await readFile(
      resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')
    );
    const load = () =>
      Promise.resolve(
        font.buffer.slice(font.byteOffset, font.byteOffset + font.byteLength)
      );
    configureDefaultFonts({
      fonts: {
        createFontProvider: () => ({
          resolve: () => load,
          resolveLastResort: () => load,
        }),
      },
    });
    try {
      const viewer = await openDocumentViewer(bytes);
      try {
        expect(wrappedGlyphRuns(viewer.displayList())).toBeGreaterThan(0);
      } finally {
        viewer.dispose();
      }
    } finally {
      configureDefaultFonts({});
    }
  });

  // An empty REF shows nothing, so "B" follows "A"; an empty PAGE field
  // keeps a digit's width, as each page fills it in.
  test('a field that shows nothing takes no width', async () => {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const field = (instr: string) =>
      `<w:p><w:r><w:t>A</w:t></w:r><w:fldSimple w:instr="${instr}"/><w:r><w:t>B</w:t></w:r></w:p>`;
    const bytes = rezipContainer(
      Object.fromEntries(
        Object.entries({
          '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
          '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
          'word/document.xml': `<w:document ${W}><w:body>${field(' REF x ')}${field(' PAGE ')}<w:sectPr/></w:body></w:document>`,
        }).map(([path, xml]) => [path, new TextEncoder().encode(xml)])
      )
    );
    const font = await readFile(
      resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')
    );
    const load = () =>
      Promise.resolve(
        font.buffer.slice(font.byteOffset, font.byteOffset + font.byteLength)
      );
    configureDefaultFonts({
      fonts: {
        createFontProvider: () => ({
          resolve: () => load,
          resolveLastResort: () => load,
        }),
      },
    });
    try {
      const viewer = await openDocumentViewer(bytes);
      try {
        const glyphs = viewer
          .displayList()
          .pages[0]!.primitives.flatMap((primitive) =>
            primitive.kind === 'glyphRun'
              ? primitive.glyphs.map((glyph) => ({
                  ...glyph,
                  text: primitive.text[glyph.cluster],
                }))
              : []
          );
        const gap = (line: number) => {
          const [a, b] = ['A', 'B'].map(
            (text) => glyphs.filter((glyph) => glyph.text === text)[line]!
          );
          return b!.x - (a!.x + a!.advance!);
        };
        expect(gap(0)).toBeCloseTo(0, 3);
        expect(gap(1)).toBeGreaterThan(5);
      } finally {
        viewer.dispose();
      }
    } finally {
      configureDefaultFonts({});
    }
  });
});
