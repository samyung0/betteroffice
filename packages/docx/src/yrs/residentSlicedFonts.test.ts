import { beforeAll, expect, test } from 'bun:test';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { zstdDecompressSync } from 'node:zlib';
import * as fonts from '@betteroffice/fonts';
import { buildResidentRegionLayoutRequest } from '../editor/computeLayout';
import type { BundledFontLoader, BundledFontProvider } from '../layout/measure/fontRegistry';
import { createRustMeasureSource, type ResidentFontRequirement } from '../layout/measure/rustMeasureSource';
import type { DisplayList } from '../layout/render';
import { applyFrameDeltaOwned, decodeFrameDelta, type RetainedFrame } from '../layout/render/frameDelta';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, rezipContainer } from '../wasm/opc';
import { preloadParseWasm } from '../wasm/parse';
import { createYrsSession } from './index';
import { drawnByFont, FontSliceUnavailableError } from './residentProgressiveFonts';

beforeAll(() =>
  Promise.all([
    preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))),
    preloadOpcWasm(),
    preloadParseWasm(),
  ])
);

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

/** One paragraph of Japanese text in MS Gothic, tagged ja-JP. */
const documentBytes = rezipContainer(Object.fromEntries(Object.entries({
  '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${TYPE}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${TYPE}.styles+xml"/></Types>`,
  '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/></Relationships>`,
  'word/styles.xml': `<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic" w:eastAsia="MS Gothic"/><w:lang w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
  'word/document.xml': `<w:document ${W}><w:body><w:p><w:r><w:t>鬱蒼とした森</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>`,
}).map(([path, xml]) => [path, new TextEncoder().encode(xml)])));

const assets = resolve(import.meta.dir, '../../../fonts-cjk/assets');
/** Every slice file by the SHA-256 of its decoded bytes. */
const sliceFiles = new Map<string, string>();
for (const face of ['NotoSansJP-Regular', 'NotoSansKR-Regular', 'NotoSansSC-Regular', 'NotoSansTC-Regular', 'NotoSerifSC-Regular']) {
  const index = JSON.parse(readFileSync(resolve(assets, face, 'slices.json'), 'utf8')) as {
    skeleton: { file: string; sha256: string }; core: { file: string; sha256: string }; chunks: Array<{ file: string; sha256: string }>;
  };
  for (const entry of [index.skeleton, index.core, ...index.chunks]) sliceFiles.set(entry.sha256, entry.file);
}

/** The fonts package's provider, its loaders' slices kept or dropped, recording the slice files decoded. */
function provider(sliced: boolean, reads: string[]): BundledFontProvider {
  const source = fonts.createFontProvider(undefined, { retain: false }) as unknown as Record<string, (...args: unknown[]) => BundledFontLoader | undefined>;
  const wrapped = new WeakMap<BundledFontLoader, BundledFontLoader>();
  const wrap = (loader: BundledFontLoader | undefined) => {
    if (!loader) return undefined;
    let next = wrapped.get(loader);
    if (!next) {
      const { slices, ...rest } = loader;
      const location = slices as fonts.SliceLocation | undefined;
      next = Object.assign(() => loader(), rest, location && sliced
        ? { slices: { ...location, decode: async (encoded: ArrayBuffer, maxOutput: number) => {
          const decoded = await location.decode(encoded, maxOutput);
          reads.push(sliceFiles.get(createHash('sha256').update(new Uint8Array(decoded)).digest('hex')) ?? 'unknown');
          return decoded;
        } } }
        : {}) as BundledFontLoader;
      if (loader.readOnlyBytes) next.readOnlyBytes = () => next!();
      wrapped.set(loader, next);
    }
    return next;
  };
  return Object.fromEntries(['resolve', 'resolveFamily', 'resolveScriptFallback', 'resolveGlyphFallback', 'resolveLastResort', 'resolveBaseLatin']
    .map((name) => [name, (...args: unknown[]) => wrap(source[name]!(...args))])) as unknown as BundledFontProvider;
}

/** An editor session over `bytes`, laid out and displayed as the editor does after each edit. */
async function editor(
  sliced: boolean,
  bytes = documentBytes,
  adjust: (request: ReturnType<typeof buildResidentRegionLayoutRequest>) => void = () => {}
) {
  const reads: string[] = [];
  const errors: Error[] = [];
  const session = await createYrsSession({ clientId: sliced ? 71_001 : 71_002 });
  session.openDocx(bytes, true);
  const source = createRustMeasureSource({ engine: session, bundled: provider(sliced, reads), progressive: sliced, onError: (error) => errors.push(error) });
  let frame: RetainedFrame | null = null;
  let chains: Record<string, number[]> = {};
  const display = async (): Promise<DisplayList> => {
    const request = buildResidentRegionLayoutRequest(session.materializeDocx()!, 24, {});
    adjust(request);
    const requirements = JSON.parse(session.layoutFontRequirementsJson(JSON.stringify(request))) as ResidentFontRequirement[];
    await source.prepareFontRequirements(requirements);
    request.measurement = source.measurementConfigForRequirements(requirements)!;
    chains = request.measurement.fontChains;
    session.layoutDocumentWithRegionsRetainedJson(JSON.stringify(request));
    const extras = JSON.stringify({ fontChains: request.measurement.fontChains });
    frame = applyFrameDeltaOwned(frame, decodeFrameDelta(session.buildDisplayListFrame(extras, frame?.frameEpoch ?? 0)));
    return frame.displayList;
  };
  const type = (text: string) => {
    const paragraph = session.paragraphs('body')[0]!;
    session.insertText({ story: 'body', paraId: paragraph.paraId, offset: paragraph.text.length }, text);
  };
  return { session, source, reads, errors, display, type, chains: () => chains };
}

/** The glyph ids `list` draws with `fontId`. */
function glyphsOf(list: DisplayList, fontId: number): Set<number> {
  const drawn = new Map();
  for (const page of list.pages) drawnByFont(page, drawn);
  return drawn.get(fontId)?.glyphs ?? new Set();
}

/** Every font id `list` draws with, in drawing order. */
function fontIds(list: DisplayList): number[] {
  const ids: number[] = [];
  JSON.stringify(list, (key, value: unknown) => {
    if (key === 'fontId' && typeof value === 'number') ids.push(value);
    return value;
  });
  return ids;
}

const zip = <A, B>(a: A[], b: B[]): Array<[A, B]> => a.map((item, index) => [item, b[index]!]);

/** `list` with each font id renumbered by first use, so two sessions that registered fonts in another order compare. */
function canonical(list: DisplayList): string {
  const ids = new Map<number, number>();
  return JSON.stringify(list, (key, value: unknown) => {
    if (key !== 'fontId' || typeof value !== 'number') return value;
    if (!ids.has(value)) ids.set(value, ids.size);
    return ids.get(value);
  });
}

test('a typed character whose chunk is not loaded loads it, its outline installs, and the layout equals the whole face\'s', async () => {
  const sliced = await editor(true);
  try {
    const before = await sliced.display();
    // The id the runs draw with: MS Gothic's measurement view of the sliced face.
    const fontId = [...sliced.source.progressiveFaces().keys()].find((id) => glyphsOf(before, id).size > 0);
    expect(fontId).toBeDefined();
    const outline = (glyph: number) => (JSON.parse(sliced.session.outlineGlyphJson(fontId!, glyph)) as { cmds: unknown[] }).cmds.length;
    expect(outline([...glyphsOf(before, fontId!)][0]!)).toBe(0);
    for (const page of before.pages) await sliced.session.prepareSlicedPaint(page, {});
    const opened = new Set(sliced.reads);
    expect(opened.has('NotoSansJP-Regular/c093.zst')).toBe(false);
    for (const glyph of glyphsOf(before, fontId!)) expect(outline(glyph)).toBeGreaterThan(0);

    // 鰯 is in chunk 93, which nothing on the page needed.
    sliced.type('鰯');
    const after = await sliced.display();
    const added = [...glyphsOf(after, fontId!)].filter((glyph) => !glyphsOf(before, fontId!).has(glyph));
    expect(added).toHaveLength(1);
    expect(outline(added[0]!)).toBe(0);
    const changed = new Set<number>();
    for (const page of after.pages) for (const id of (await sliced.session.prepareSlicedPaint(page, {})).changed) changed.add(id);
    expect(changed.has(fontId!)).toBe(true);
    expect(sliced.reads.filter((file) => !opened.has(file))).toEqual(['NotoSansJP-Regular/c093.zst']);
    expect(outline(added[0]!)).toBeGreaterThan(0);
    expect(sliced.errors).toEqual([]);
    // Read before the second session opens: sessions in one module share its font store.
    const slicedOutlines = [...glyphsOf(after, fontId!)].map((glyph) => [glyph, sliced.session.outlineGlyphJson(fontId!, glyph)] as const);

    const whole = await editor(false);
    try {
      await whole.display();
      whole.type('鰯');
      const wholeAfter = await whole.display();
      expect(canonical(after)).toBe(canonical(wholeAfter));
      // The pages paint from these outlines, so they match the whole face's glyph for glyph.
      const wholeId = new Map(zip(fontIds(after), fontIds(wholeAfter))).get(fontId!)!;
      expect(slicedOutlines.length).toBeGreaterThan(0);
      for (const [glyph, outline] of slicedOutlines) expect(outline).toBe(whole.session.outlineGlyphJson(wholeId, glyph));
    } finally {
      whole.session.destroy();
    }
  } finally {
    sliced.session.destroy();
  }
}, 60_000);

test('a worker snapshot names the sliced faces with their files, for the worker to open them', async () => {
  const sliced = await editor(true);
  try {
    await sliced.display();
    const snapshot = sliced.session.residentWorkerSnapshot()!;
    const [fontId] = [...sliced.source.progressiveFaces().keys()];
    expect(snapshot.progressiveFonts?.map((spec) => spec.slot)).toEqual([fontId!]);
    const { source } = snapshot.progressiveFonts![0]!;
    expect(source.ref.index).toBe('NotoSansJP-Regular/slices.json');
    expect(Object.keys(source.urls)).toContain('NotoSansJP-Regular/c093.zst');
    // The page's decoded files go along, so the worker decodes none of them again.
    expect([...snapshot.progressiveFonts![0]!.files!.keys()]).toContain('NotoSansJP-Regular/skeleton.zst');
    expect(sliced.session.residentWorkerSnapshot({ knownFontsRevision: snapshot.fontsRevision })!.progressiveFonts).toBeUndefined();
  } finally {
    sliced.session.destroy();
  }
}, 60_000);

test('a slice the worker could not load is reported once, not fetched here, and asked for again after retryFontSlices', async () => {
  const sliced = await editor(true);
  try {
    const before = await sliced.display();
    for (const page of before.pages) await sliced.session.prepareSlicedPaint(page, {});
    const fontId = [...sliced.source.progressiveFaces().keys()].find((id) => glyphsOf(before, id).size > 0)!;
    const outline = (glyph: number) => (JSON.parse(sliced.session.outlineGlyphJson(fontId, glyph)) as { cmds: unknown[] }).cmds.length;
    sliced.type('鰯');
    const after = await sliced.display();
    const [added] = [...glyphsOf(after, fontId)].filter((glyph) => !glyphsOf(before, fontId).has(glyph));
    // The pages were presented; then the engine worker could not reach the slice (offline).
    sliced.session.markPagesPresented();
    const [{ slot }] = sliced.session.residentWorkerSnapshot()!.progressiveFonts!;
    const file = 'NotoSansJP-Regular/c093.zst';
    sliced.session.reportResidentFontErrors([{ slot: slot!, message: 'Failed to fetch', unavailable: { file, reason: 'network' } }]);
    sliced.session.reportResidentFontErrors([{ slot: slot!, message: 'Failed to fetch', unavailable: { file, reason: 'network' } }]);
    // This page's own paint neither fetches it nor reports it again.
    for (const page of after.pages) await sliced.session.prepareSlicedPaint(page, {});
    expect(sliced.reads).not.toContain(file);
    expect(outline(added!)).toBe(0);
    expect(sliced.errors).toHaveLength(1);
    expect(sliced.errors[0]).toBeInstanceOf(FontSliceUnavailableError);
    expect(sliced.errors[0]).toMatchObject({ file, reason: 'network' });

    // Back online: the worker is told, the pages repaint, and the slice loads.
    let retried = 0;
    sliced.session.onFontSlicesRetry(() => {
      retried += 1;
    });
    const repaints: Array<Set<number>> = [];
    sliced.session.onGlyphOutlines((fontIds) => repaints.push(fontIds));
    sliced.session.retryFontSlices();
    expect(retried).toBe(1);
    expect(repaints).toEqual([new Set([slot!])]);
    for (const page of after.pages) await sliced.session.prepareSlicedPaint(page, {});
    for (let wait = 0; wait < 100 && outline(added!) === 0; wait += 1) await new Promise((resolve) => setTimeout(resolve, 20));
    expect(sliced.reads.filter((read) => read === file)).toHaveLength(1);
    expect(outline(added!)).toBeGreaterThan(0);
    expect(sliced.errors).toHaveLength(1);
  } finally {
    sliced.session.destroy();
  }
}, 60_000);

test('a slice that fails while the document opens fails the open; after its pages are presented it is a notice', async () => {
  const sliced = await editor(true);
  try {
    await sliced.display();
    const [{ slot }] = sliced.session.residentWorkerSnapshot()!.progressiveFonts!;
    const file = 'NotoSansJP-Regular/c093.zst';
    const fail = () => sliced.session.reportResidentFontErrors([{ slot: slot!, message: 'HTTP 503', unavailable: { file, reason: 'server', status: 503 } }]);
    fail();
    expect(sliced.errors).toHaveLength(1);
    expect(sliced.errors[0]).not.toBeInstanceOf(FontSliceUnavailableError);
    expect(sliced.errors[0]!.message).toContain(file);
    sliced.session.markPagesPresented();
    fail();
    expect(sliced.errors).toHaveLength(2);
    expect(sliced.errors[1]).toBeInstanceOf(FontSliceUnavailableError);
    expect(sliced.errors[1]).toMatchObject({ file, reason: 'server', status: 503 });
  } finally {
    sliced.session.destroy();
  }
}, 60_000);

test('retryFontSlices clears a failure the worker reported for a face no page here painted', async () => {
  const sliced = await editor(true);
  try {
    await sliced.display();
    sliced.type('鰯');
    const after = await sliced.display();
    const fontId = [...sliced.source.progressiveFaces().keys()].find((id) => glyphsOf(after, id).size > 0)!;
    const outline = (glyph: number) => (JSON.parse(sliced.session.outlineGlyphJson(fontId, glyph)) as { cmds: unknown[] }).cmds.length;
    // The worker painted and could not reach the slice; nothing painted here.
    const [{ slot }] = sliced.session.residentWorkerSnapshot()!.progressiveFonts!;
    const file = 'NotoSansJP-Regular/c093.zst';
    sliced.session.reportResidentFontErrors([{ slot: slot!, message: 'Failed to fetch', unavailable: { file, reason: 'network' } }]);
    // A worker opened again (new fonts) takes the failure along and does not fetch it either.
    expect(sliced.session.residentWorkerSnapshot()!.progressiveFonts![0]!.failed).toEqual([{ file, reason: 'network' }]);
    sliced.session.retryFontSlices();
    expect(sliced.session.residentWorkerSnapshot()!.progressiveFonts![0]!.failed).toBeUndefined();
    // Back online, this realm's first paint (a fallback or print) asks for it again.
    for (const page of after.pages) await sliced.session.prepareSlicedPaint(page, sliced.chains());
    const drawn = [...glyphsOf(after, fontId)];
    for (let wait = 0; wait < 100 && drawn.some((glyph) => outline(glyph) === 0); wait += 1) await new Promise((resolve) => setTimeout(resolve, 20));
    expect(sliced.reads.filter((read) => read === file)).toHaveLength(1);
    expect(drawn.length).toBeGreaterThan(0);
    for (const glyph of drawn) expect(outline(glyph)).toBeGreaterThan(0);
  } finally {
    sliced.session.destroy();
  }
}, 60_000);

/** A package of `parts` (styles, settings, numbering) when given, and `body`'s paragraphs. */
function packageOf(parts: { styles?: string; settings?: string; numbering?: string }, body: string): Uint8Array {
  const named = Object.entries(parts).filter((entry): entry is [string, string] => entry[1] !== undefined);
  return rezipContainer(Object.fromEntries(Object.entries({
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${TYPE}.document.main+xml"/>${named.map(([part]) => `<Override PartName="/word/${part}.xml" ContentType="${TYPE}.${part}+xml"/>`).join('')}</Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${named.map(([part], index) => `<Relationship Id="rId${index + 1}" Type="${R}/${part}" Target="${part}.xml"/>`).join('')}</Relationships>`,
    ...Object.fromEntries(named.map(([part, xml]) => [`word/${part}.xml`, xml])),
    'word/document.xml': `<w:document ${W}><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>`,
  }).map(([path, xml]) => [path, new TextEncoder().encode(xml)])));
}

/** The slice files of `face` that `text`'s characters map to. */
async function chunkFilesOf(face: string, text: string): Promise<string[]> {
  const { parseSkeleton } = await import('../../../fonts/src/slices');
  const index = JSON.parse(readFileSync(resolve(assets, face, 'slices.json'), 'utf8')) as {
    skeleton: { file: string }; core: { file: string }; chunks: Array<{ id: number; file: string }>;
  };
  const skeleton = parseSkeleton(new Uint8Array(zstdDecompressSync(readFileSync(resolve(assets, index.skeleton.file)))));
  const files = new Set<string>();
  for (const character of text) {
    const chunk = skeleton.chunkOfCodepoint(character.codePointAt(0)!);
    if (chunk === 0) files.add(index.core.file);
    else if (chunk !== undefined) files.add(index.chunks.find(({ id }) => id === chunk)!.file);
  }
  return [...files].sort();
}

/** Runs `body` with a FontFace constructor and font set in this realm, recording the faces added. */
async function withRealmFonts<T>(body: (added: Array<{ family: string; source: ArrayBuffer }>) => Promise<T>): Promise<T> {
  const realm = globalThis as { FontFace?: unknown; fonts?: unknown };
  const added: Array<{ family: string; source: ArrayBuffer }> = [];
  realm.FontFace = class {
    constructor(readonly family: string, readonly source: ArrayBuffer) {}
    async load() {
      return this;
    }
  };
  realm.fonts = { add: (face: { family: string; source: ArrayBuffer }) => added.push(face), delete: () => {} };
  try {
    return await body(added);
  } finally {
    delete realm.FontFace;
    delete realm.fonts;
  }
}

test('an editor page paints its list marker and watermark with the sliced face, registered once their chunks have loaded', async () => {
  // The marker reads 鰯1 (鰯 is in chunk 93) and the watermark 齟 (its own chunk): characters no run shapes.
  const markerDocument = packageOf({
    styles: `<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic" w:eastAsia="MS Gothic"/><w:lang w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    numbering: `<w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="鰯%1"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="840" w:hanging="840"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
  }, '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>本文</w:t></w:r></w:p>');
  await withRealmFonts(async (added) => {
    // The watermark as the host's request carries it.
    const sliced = await editor(true, markerDocument, (request) => {
      request.regions.watermark = { kind: 'text', text: '齟', font: 'MS Gothic' } as typeof request.regions.watermark;
    });
    try {
      const [page] = (await sliced.display()).pages;
      expect(JSON.stringify(page)).toContain('"text":"鰯1"');
      expect(JSON.stringify(page)).toContain('"text":"齟"');
      const opened = new Set(sliced.reads);
      const paint = await sliced.session.prepareSlicedPaint(page, sliced.chains());
      // MS Gothic's chain starts with the sliced face, which takes its place.
      const name = paint.fontFamilies.get('MS Gothic') as string;
      expect(name).toMatch(/^BetterOffice Sliced \d+$/);
      expect(sliced.source.progressiveFaces().has(Number(name!.split(' ').at(-1)))).toBe(true);
      expect(added.map((face) => face.family)).toEqual([name!]);
      for (const text of ['鰯', '齟']) {
        const own = (await chunkFilesOf('NotoSansJP-Regular', text)).filter((file) => !opened.has(file));
        expect(own.length).toBeGreaterThan(0);
        expect(sliced.reads.filter((file) => !opened.has(file))).toEqual(expect.arrayContaining(own));
      }
      expect([...paint.outlineOnlyFontIds]).toEqual(expect.arrayContaining([...sliced.source.progressiveFaces().keys()]));
    } finally {
      sliced.session.destroy();
    }
  });
}, 60_000);

/**
 * Kanji-only text naming no language, in a family no face stands in for (so
 * its Han draws with the script slot's face), in a document whose theme
 * fonts are for `theme` (`w:themeFontLang`).
 */
const untaggedHan = (theme?: string) => packageOf({
  styles: `<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Han Test Face" w:hAnsi="Han Test Face" w:eastAsia="Han Test Face"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
  settings: `<w:settings ${W}>${theme ? `<w:themeFontLang w:val="en-US" w:eastAsia="${theme}"/>` : ''}</w:settings>`,
}, '<w:p><w:r><w:t>骨直角</w:t></w:r></w:p>');

for (const [theme, face] of [['ja-JP', 'NotoSansJP'], ['zh-TW', 'NotoSansTC'], [undefined, 'NotoSansSC']] as const) {
  test(`untagged Han ${theme ? `in a document whose theme fonts are for ${theme}` : 'naming no language anywhere'} draws with ${face}`, async () => {
    const sliced = await editor(true, untaggedHan(theme));
    try {
      const list = await sliced.display();
      const faces = sliced.source.progressiveFaces();
      const drawing = [...new Set(fontIds(list))].filter((id) => faces.has(id) && glyphsOf(list, id).size > 0);
      expect(drawing.map((id) => (faces.get(id)!.slices as fonts.SliceLocation).ref.index)).toEqual([`${face}-Regular/slices.json`]);
      expect(sliced.errors).toEqual([]);
    } finally {
      sliced.session.destroy();
    }
  }, 60_000);
}

test('a remote update that brings a character from a chunk not loaded loads it, and its outline installs', async () => {
  const sliced = await editor(true);
  const peer = await createYrsSession({ clientId: 71_003 });
  try {
    const before = await sliced.display();
    for (const page of before.pages) await sliced.session.prepareSlicedPaint(page, sliced.chains());
    const fontId = [...sliced.source.progressiveFaces().keys()].find((id) => glyphsOf(before, id).size > 0)!;
    const outline = (glyph: number) => (JSON.parse(sliced.session.outlineGlyphJson(fontId, glyph)) as { cmds: unknown[] }).cmds.length;
    const opened = new Set(sliced.reads);

    // A collaborator types 鰯 (chunk 93); its update arrives here.
    peer.applyUpdate(sliced.session.encodeStateAsUpdate());
    const paragraph = peer.paragraphs('body')[0]!;
    peer.insertText({ story: 'body', paraId: paragraph.paraId, offset: paragraph.text.length }, '鰯');
    sliced.session.applyUpdate(peer.encodeStateAsUpdate(sliced.session.encodeStateVector()));

    const after = await sliced.display();
    const [added] = [...glyphsOf(after, fontId)].filter((glyph) => !glyphsOf(before, fontId).has(glyph));
    expect(added).toBeDefined();
    const changed = new Set<number>();
    for (const page of after.pages) for (const id of (await sliced.session.prepareSlicedPaint(page, sliced.chains())).changed) changed.add(id);
    expect(changed.has(fontId)).toBe(true);
    expect(sliced.reads.filter((file) => !opened.has(file))).toEqual(['NotoSansJP-Regular/c093.zst']);
    expect(outline(added!)).toBeGreaterThan(0);
  } finally {
    peer.destroy();
    sliced.session.destroy();
  }
}, 60_000);
