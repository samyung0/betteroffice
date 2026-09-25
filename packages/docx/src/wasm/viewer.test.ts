import { beforeAll, describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { buildResidentRegionLayoutRequest } from '../editor/computeLayout';
import type { DisplayList } from '../layout/render';
import { configureDefaultFonts, initWasm, openDocumentViewer } from '../viewer';
import { createYrsSession } from '../yrs';
import { preloadEditWasm } from './edit';
import { preloadOpcWasm, rezipContainer } from './opc';
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

/** Laid-out text lines across all pages. */
function lineCount(displayList: DisplayList): number {
  const lines = new Set<string>();
  displayList.pages.forEach((page, index) => {
    for (const primitive of page.primitives)
      if (primitive.kind === 'text')
        lines.add(`${index}:${primitive.blockKey}:${primitive.lineIndex}`);
  });
  return lines.size;
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
        expect(JSON.parse(view.layoutRequestJson())).toEqual(
          JSON.parse(JSON.stringify(editor))
        );
      } finally {
        view.free();
        session.destroy();
      }
    }
  });

  test('measures with the configured fonts, so long paragraphs wrap', async () => {
    const bytes = new Uint8Array(await readFile(fixture));
    const unwrapped = await openDocumentViewer(bytes);
    const synthetic = lineCount(unwrapped.displayList());
    unwrapped.dispose();
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
        expect(lineCount(viewer.displayList())).toBeGreaterThan(synthetic);
      } finally {
        viewer.dispose();
      }
    } finally {
      configureDefaultFonts({});
    }
  });
});
