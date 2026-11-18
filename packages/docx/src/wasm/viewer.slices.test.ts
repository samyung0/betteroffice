import { afterEach, beforeAll, expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { zstdDecompressSync } from 'node:zlib';
import * as fonts from '@betteroffice/fonts';
import type { BundledFontLoader, BundledFontProvider } from '../layout/measure/fontRegistry';
import type { DisplayList } from '../layout/render';
import { configureDefaultFonts, initWasm, openDocumentViewer } from '../viewer';
import { preloadOpcWasm, rezipContainer } from './opc';
import * as viewerModule from './viewer';

const root = resolve(import.meta.dir, '../../../..');
const assets = resolve(root, 'packages/fonts-cjk/assets');
// Spied before the engine starts: starting it also starts the font decoder.
const decoderPreload = spyOn(fonts, 'preloadFontDecoder');

beforeAll(async () => {
  await Promise.all([
    initWasm(await readFile(resolve(import.meta.dir, 'generated/viewer/docx_view_wasm_bg.wasm'))),
    preloadOpcWasm(await readFile(resolve(import.meta.dir, 'generated/opc/ooxml_opc_bg.wasm'))),
  ]);
});

afterEach(() => configureDefaultFonts({}));

test('starting the view engine also starts the fonts package decoder its faces decode with', () => {
  expect(decoderPreload).toHaveBeenCalledTimes(1);
  expect(decoderPreload.mock.calls[0]).toEqual([undefined]);
});

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
          // Noted once its URL resolves, when the request goes out.
          url: (file: string) => location.url(swap[file] ?? file).then((url) => {
            sliceFiles.push(file);
            return url;
          }),
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
  // An editor opening the face takes these decoded files instead of fetching them again.
  expect([...sliced!.files.keys()].sort()).toEqual([...chunks, `${face}/skeleton.zst`, `${face}/slices.json`].sort());
  // The painted face holds those outlines: it is larger than the skeleton's.
  const skeletonLength = (JSON.parse(await readFile(resolve(assets, `${face}/slices.json`), 'utf8')) as { skeleton: { byteLength: number } }).skeleton.byteLength;
  expect(sliced!.bytes.byteLength).toBeGreaterThan(skeletonLength);
});

/**
 * MS Gothic Japanese with `body` in a list item whose number reads `鰯%1`
 * (or the bullet `bullet`), in `markerFont` when given, so the marker draws
 * characters no run shapes.
 */
function markerPackage(body: string, markerFont?: string, bullet?: string): Uint8Array {
  const level = bullet ? `<w:numFmt w:val="bullet"/><w:lvlText w:val="${bullet}"/>` : '<w:numFmt w:val="decimal"/><w:lvlText w:val="鰯%1"/>';
  const markerRun = markerFont ? `<w:rPr><w:rFonts w:ascii="${markerFont}" w:hAnsi="${markerFont}" w:eastAsia="${markerFont}"/></w:rPr>` : '';
  return rezipContainer(Object.fromEntries(Object.entries({
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${TYPE}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${TYPE}.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="${TYPE}.numbering+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/><Relationship Id="rId3" Type="${R}/numbering" Target="numbering.xml"/></Relationships>`,
    'word/styles.xml': `<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic" w:eastAsia="MS Gothic"/><w:lang w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/numbering.xml': `<w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/>${level}<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="840" w:hanging="840"/></w:pPr>${markerRun}</w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
    'word/document.xml': `<w:document ${W}><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>${body}</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>`,
  }).map(([path, xml]) => [path, new TextEncoder().encode(xml)])));
}

test('a list marker loads the chunks of characters only it draws', async () => {
  const { sliceFiles } = recordingFonts(true);
  const handle = await openDocumentViewer(markerPackage('本文'));
  const face = 'NotoSansJP-Regular';
  expect(JSON.stringify(handle.displayList().pages)).toContain('"text":"鰯1"');
  const bodyChunks = await chunkFiles(face, '本文');
  const own = (await chunkFiles(face, '鰯')).filter((file) => !bodyChunks.includes(file));
  expect(own.length).toBeGreaterThan(0);
  const loaded = new Set(sliceFiles);
  for (const file of own) expect(loaded.has(file)).toBe(true);
});

test('a list marker in a Latin family paints its CJK with the sliced face, named after the family', async () => {
  const { sliceFiles } = recordingFonts(true);
  const handle = await openDocumentViewer(markerPackage('本文', 'Calibri'));
  const [sliced] = handle.slicedFaces();
  expect(sliced!.family).toMatch(/^BetterOffice Sliced \d+\.0$/);
  const marker = handle.displayList().pages[0]!.primitives.find((primitive) => primitive.kind === 'text' && primitive.text === '鰯1');
  expect(marker && 'font' in marker && marker.font).toContain(`"Calibri", "${sliced!.family}"`);
  const loaded = new Set(sliceFiles);
  for (const file of await chunkFiles('NotoSansJP-Regular', '鰯')) expect(loaded.has(file)).toBe(true);
});

test('a ① bullet in Century, a Japanese document Latin font, paints with the sliced face after Century and loads its chunk', async () => {
  const { sliceFiles } = recordingFonts(true);
  const handle = await openDocumentViewer(markerPackage('本文', 'Century', '①'));
  const [sliced] = handle.slicedFaces();
  const marker = handle.displayList().pages[0]!.primitives.find((primitive) => primitive.kind === 'text' && primitive.text === '①');
  expect(marker && 'font' in marker && marker.font).toContain(`"Century", "${sliced!.family}"`);
  const loaded = new Set(sliceFiles);
  for (const file of await chunkFiles('NotoSansJP-Regular', '①')) expect(loaded.has(file)).toBe(true);
});

/** Notes in `events` when the engine is ready (it opens the document) and where its layout starts and ends. */
function markLayout(events: string[]) {
  const open = viewerModule.openViewDocument;
  return spyOn(viewerModule, 'openViewDocument').mockImplementation((bytes: Uint8Array) => {
    events.push('engine ready');
    const document = open(bytes);
    const layout = document.layout.bind(document);
    return Object.assign(document, {
      layout: (json: string) => {
        events.push('layout start');
        const pages = layout(json);
        events.push('layout end');
        return pages;
      },
    });
  });
}

/** One paragraph of `text` with Latin in Calibri and East Asian text in SimSun, in Simplified Chinese. */
function simSunPackage(text: string): Uint8Array {
  const fonts = '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="SimSun"/><w:lang w:eastAsia="zh-CN"/>';
  return rezipContainer(Object.fromEntries(Object.entries({
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${TYPE}.document.main+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/document.xml': `<w:document ${W}><w:body><w:p><w:r><w:rPr>${fonts}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
  }).map(([path, xml]) => [path, new TextEncoder().encode(xml)])));
}

/** The slice files an open of `bytes` requested, with when its engine was ready and where its layout started and ended. */
async function openMarked(bytes: Uint8Array): Promise<string[]> {
  const { sliceFiles } = recordingFonts(true);
  const marks = markLayout(sliceFiles);
  try {
    await openDocumentViewer(bytes);
  } finally {
    marks.mockRestore();
  }
  return sliceFiles;
}

/** Where `files` requested `file`, checked to come after the engine and every skeleton and before layout ends. */
function startedAsLayoutRuns(files: string[], file: string): number {
  const at = files.indexOf(file);
  expect(at).toBeGreaterThan(files.indexOf('engine ready'));
  for (const [index, name] of files.entries()) if (name.endsWith('/skeleton.zst')) expect(at).toBeGreaterThan(index);
  expect(at).toBeLessThan(files.indexOf('layout end'));
  return at;
}

test('the face that paints East Asian text first requests its core as layout starts; one that fills only what it lacks waits (Q758 B, Q769 A)', async () => {
  // long-handbook's Han paints with Noto Sans TC first; Noto Sans SC draws only what TC lacks.
  const handbook = await openMarked(new Uint8Array(await readFile(resolve(root, 'poc/fixtures/long-handbook.docx'))));
  startedAsLayoutRuns(handbook, 'NotoSansTC-Regular/core.zst');
  expect(handbook.indexOf('NotoSansSC-Regular/core.zst')).toBeGreaterThan(handbook.indexOf('layout end'));
}, 60_000);

test('the Han of a Calibri run in SimSun starts the core of the face Calibri paints it with, never that of SimSun (Q762 A)', async () => {
  const files = await openMarked(simSunPackage('Hello 你好世界'));
  startedAsLayoutRuns(files, 'NotoSansSC-Regular/core.zst');
  expect(files).toContain('NotoSerifSC-Regular/slices.json');
  expect(files).not.toContain('NotoSerifSC-Regular/core.zst');
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

/**
 * `list` with each font id renumbered by first use, so two runs that
 * registered fonts in another order compare, and without the private
 * families a sliced open names its faces by in markers' and watermarks' fonts.
 */
function canonical(list: DisplayList): string {
  const ids = new Map<number, number>();
  return JSON.stringify(list, (key, value: unknown) => {
    if (key === 'font' && typeof value === 'string') {
      return value.replace(/, "BetterOffice Sliced [\d.]+"/g, '').replace(/"([^"]*)"/g, '$1');
    }
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
