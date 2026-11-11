import { afterEach, beforeAll, expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { zstdDecompressSync } from 'node:zlib';
import * as fonts from '@betteroffice/fonts';
import type { BundledFontLoader, BundledFontProvider } from '../layout/measure/fontRegistry';
import type { DisplayList } from '../layout/render';
import { configureDefaultFonts, initWasm, openDocumentViewer } from '../viewer';
import { preloadOpcWasm, rezipContainer } from './opc';

const root = resolve(import.meta.dir, '../../../..');
const assets = resolve(root, 'packages/fonts-cjk/assets');

beforeAll(async () => {
  await Promise.all([
    initWasm(await readFile(resolve(import.meta.dir, 'generated/viewer/docx_view_wasm_bg.wasm'))),
    preloadOpcWasm(await readFile(resolve(import.meta.dir, 'generated/opc/ooxml_opc_bg.wasm'))),
  ]);
});

afterEach(() => configureDefaultFonts({}));

type Location = fonts.SliceLocation;

/**
 * The fonts package's provider with its loaders' slices kept (`sliced`) or
 * dropped, recording the slice files and whole faces it reads; `swap` serves
 * another file's bytes for a slice file.
 */
function recordingFonts(sliced: boolean, swap: Record<string, string> = {}) {
  const sliceFiles: string[] = [];
  const wholeFaces: string[] = [];
  // Nothing kept between opens, as a host keeps nothing: each open reads its files.
  const provider = fonts.createFontProvider(undefined, { retain: false }) as unknown as BundledFontProvider & Record<string, unknown>;
  const wrapped = new WeakMap<BundledFontLoader, BundledFontLoader>();
  const wrap = (loader: BundledFontLoader | undefined): BundledFontLoader | undefined => {
    if (!loader) return undefined;
    let next = wrapped.get(loader);
    if (!next) {
      const { slices, ...rest } = loader;
      const location = slices as Location | undefined;
      next = Object.assign(() => {
        if (location) wholeFaces.push(location.ref.index);
        return loader();
      }, rest, location && sliced ? {
        slices: {
          ...location,
          url: (file: string) => {
            sliceFiles.push(file);
            return location.url(swap[file] ?? file);
          },
        },
      } : {}) as BundledFontLoader;
      if (loader.readOnlyBytes) next.readOnlyBytes = () => next!();
      wrapped.set(loader, next);
    }
    return next;
  };
  const methods = ['resolve', 'resolveFamily', 'resolveScriptFallback', 'resolveGlyphFallback', 'resolveLastResort', 'resolveBaseLatin'];
  const recording = Object.fromEntries(methods.map((name) => [name,
    (...args: unknown[]) => wrap((provider[name] as (...values: unknown[]) => BundledFontLoader | undefined)(...args))]));
  configureDefaultFonts({ fonts: { createFontProvider: () => recording as unknown as BundledFontProvider } });
  return { sliceFiles, wholeFaces };
}

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

/** Japanese text in MS Gothic: `body` in the body, `header` in the default header. */
function japanesePackage(body: string, header: string): Uint8Array {
  const run = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
  return rezipContainer(Object.fromEntries(Object.entries({
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${TYPE}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${TYPE}.styles+xml"/><Override PartName="/word/header1.xml" ContentType="${TYPE}.header+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${R}/header" Target="header1.xml"/></Relationships>`,
    'word/styles.xml': `<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic" w:eastAsia="MS Gothic"/><w:lang w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/header1.xml': `<w:hdr ${W}>${run(header)}</w:hdr>`,
    'word/document.xml': `<w:document ${W} xmlns:r="${R}"><w:body>${run(body)}<w:sectPr><w:headerReference w:type="default" r:id="rId2"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440" w:header="720"/></w:sectPr></w:body></w:document>`,
  }).map(([path, xml]) => [path, new TextEncoder().encode(xml)])));
}

/** The slice chunk files `text`'s characters map to in a face's skeleton, the core included. */
async function chunkFiles(face: string, text: string): Promise<string[]> {
  const { parseSkeleton } = await import('../../../fonts/src/slices');
  const index = JSON.parse(await readFile(resolve(assets, `${face}/slices.json`), 'utf8')) as {
    skeleton: { file: string }; core: { file: string }; chunks: Array<{ id: number; file: string }>;
  };
  const skeleton = parseSkeleton(new Uint8Array(zstdDecompressSync(await readFile(resolve(assets, index.skeleton.file)))));
  const files = new Set<string>();
  for (const character of text) {
    const chunk = skeleton.chunkOfCodepoint(character.codePointAt(0)!);
    if (chunk === 0) files.add(index.core.file);
    else if (chunk !== undefined) files.add(index.chunks.find(({ id }) => id === chunk)!.file);
  }
  return [...files].sort();
}

const BODY = '鬱蒼とした森𩸽';
const HEADER = '齟';

test('a sliced face loads its skeleton and the chunks of exactly the characters the pages draw, headers and supplementary characters included', async () => {
  const { sliceFiles, wholeFaces } = recordingFonts(true);
  const handle = await openDocumentViewer(japanesePackage(BODY, HEADER));
  const face = 'NotoSansJP-Regular';
  const chunks = await chunkFiles(face, BODY + HEADER);
  // The header's and the supplementary character's chunks are their own.
  expect(chunks).toEqual([`${face}/c017.zst`, `${face}/c023.zst`, `${face}/c095.zst`, `${face}/c100.zst`, `${face}/core.zst`]);
  expect([...new Set(sliceFiles)].filter((file) => file.endsWith('.zst') && !file.endsWith('skeleton.zst')).sort()).toEqual(chunks);
  expect(new Set(sliceFiles).has(`${face}/skeleton.zst`)).toBe(true);
  expect(wholeFaces).toEqual([]);
  const [sliced] = handle.slicedFaces();
  expect(handle.slicedFaces()).toHaveLength(1);
  expect((sliced!.slices as Location).ref.index).toBe(`${face}/slices.json`);
  // The painted face holds those outlines: it is larger than the skeleton's.
  const skeletonLength = (JSON.parse(await readFile(resolve(assets, `${face}/slices.json`), 'utf8')) as { skeleton: { byteLength: number } }).skeleton.byteLength;
  expect(sliced!.bytes.byteLength).toBeGreaterThan(skeletonLength);
});

test('a slice chunk that fails to load fails the open instead of painting without it', async () => {
  recordingFonts(true, { 'NotoSansJP-Regular/c095.zst': 'NotoSansJP-Regular/c094.zst' });
  await expect(openDocumentViewer(japanesePackage(BODY, HEADER))).rejects.toThrow('NotoSansJP-Regular/c095.zst');
});

test('a sliced face whose skeleton fails to load fails the open instead of loading it whole', async () => {
  const { wholeFaces } = recordingFonts(true, { 'NotoSansJP-Regular/skeleton.zst': 'NotoSansKR-Regular/skeleton.zst' });
  await expect(openDocumentViewer(japanesePackage(BODY, HEADER))).rejects.toThrow('NotoSansJP-Regular/skeleton.zst');
  expect(wholeFaces).toEqual([]);
});

/** One Calibri paragraph of `text`. */
function latinPackage(text: string): Uint8Array {
  return rezipContainer(Object.fromEntries(Object.entries({
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${TYPE}.document.main+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/document.xml': `<w:document ${W}><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr><w:t>${text}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
  }).map(([path, xml]) => [path, new TextEncoder().encode(xml)])));
}

test('the viewer loads every face with the fonts package decoder, the base Latin backstop included, and fails when the backstop fails (Q734)', async () => {
  const create = spyOn(fonts, 'createFontProvider');
  configureDefaultFonts({});
  try {
    await openDocumentViewer(latinPackage('Arrows ⇄ ⇆'));
    expect(create.mock.calls.length).toBeGreaterThan(0);
    for (const [, runtime] of create.mock.calls) expect(runtime?.decoder).toBeUndefined();
  } finally {
    create.mockRestore();
  }
  const provider = fonts.createFontProvider(undefined, { retain: false });
  configureDefaultFonts({ fonts: { createFontProvider: () => ({
    ...provider,
    resolveBaseLatin: () => async () => { throw new Error('backstop offline'); },
  }) as unknown as BundledFontProvider } });
  await expect(openDocumentViewer(latinPackage('Arrows ⇄ ⇆'))).rejects.toThrow('backstop offline');
});

/** `list` with each font id renumbered by first use, so two runs that registered fonts in another order compare. */
function canonical(list: DisplayList): string {
  const ids = new Map<number, number>();
  return JSON.stringify(list, (key, value: unknown) => {
    if (key !== 'fontId' || typeof value !== 'number') return value;
    if (!ids.has(value)) ids.set(value, ids.size);
    return ids.get(value);
  });
}

for (const name of ['exchange-plan.docx', 'long-handbook.docx', 'japanese.docx']) {
  test(`${name} lays out and draws the same glyphs with sliced faces as with whole faces`, async () => {
    const bytes = new Uint8Array(await readFile(resolve(root, 'poc/fixtures', name)));
    const whole = recordingFonts(false);
    const wholeList = (await openDocumentViewer(bytes)).displayList();
    expect(whole.wholeFaces.length).toBeGreaterThan(0);
    const sliced = recordingFonts(true);
    const slicedHandle = await openDocumentViewer(bytes);
    expect(sliced.wholeFaces).toEqual([]);
    expect(slicedHandle.slicedFaces().length).toBe(whole.wholeFaces.length);
    const slicedList = slicedHandle.displayList();
    expect(slicedList.pages.length).toBe(wholeList.pages.length);
    expect(canonical(slicedList)).toBe(canonical(wholeList));
  }, 60_000);
}
