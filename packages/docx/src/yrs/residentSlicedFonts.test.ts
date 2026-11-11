import { beforeAll, expect, test } from 'bun:test';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
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
import { drawnByFont } from './residentProgressiveFonts';

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

/** An editor session over the document, laid out and displayed as the editor does after each edit. */
async function editor(sliced: boolean) {
  const reads: string[] = [];
  const errors: Error[] = [];
  const session = await createYrsSession({ clientId: sliced ? 71_001 : 71_002 });
  session.openDocx(documentBytes, true);
  const source = createRustMeasureSource({ engine: session, bundled: provider(sliced, reads), progressive: sliced, onError: (error) => errors.push(error) });
  let frame: RetainedFrame | null = null;
  const display = async (): Promise<DisplayList> => {
    const request = buildResidentRegionLayoutRequest(session.materializeDocx()!, 24, {});
    const requirements = JSON.parse(session.layoutFontRequirementsJson(JSON.stringify(request))) as ResidentFontRequirement[];
    await source.prepareFontRequirements(requirements);
    request.measurement = source.measurementConfigForRequirements(requirements)!;
    session.layoutDocumentWithRegionsRetainedJson(JSON.stringify(request));
    const extras = JSON.stringify({ fontChains: request.measurement.fontChains });
    frame = applyFrameDeltaOwned(frame, decodeFrameDelta(session.buildDisplayListFrame(extras, frame?.frameEpoch ?? 0)));
    return frame.displayList;
  };
  const type = (text: string) => {
    const paragraph = session.paragraphs('body')[0]!;
    session.insertText({ story: 'body', paraId: paragraph.paraId, offset: paragraph.text.length }, text);
  };
  return { session, source, reads, errors, display, type };
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
    for (const page of before.pages) await sliced.session.ensureGlyphOutlines(page);
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
    for (const page of after.pages) for (const id of await sliced.session.ensureGlyphOutlines(page)) changed.add(id);
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
