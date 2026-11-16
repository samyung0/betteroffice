import { afterAll, beforeEach, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { decodeFontFace } from '../../../fonts/src/decode';
import { REQUEST_TIMEOUT_MS } from './residentEngineWorkerClient';
import { FontSliceUnavailableError, glyphsByFont, PROGRESSIVE_OUTLINE_WAIT_MS, ResidentProgressiveFonts, sliceFaceOpener, type ProgressiveFontSpec } from './residentProgressiveFonts';

const assets = new URL('../../../fonts-cjk/assets/', import.meta.url);
const manifest = JSON.parse(readFileSync(new URL('manifest.json', assets), 'utf8')) as {
  faces: Array<{ file: string; slices?: ProgressiveFontSpec['source']['ref'] }>;
};
const ref = manifest.faces.find((face) => face.file === 'NotoSansJP-Regular.otf.zst')!.slices!;
const index = JSON.parse(readFileSync(new URL(ref.index, assets), 'utf8')) as {
  skeleton: { file: string };
  core: { file: string };
  chunks: Array<{ id: number; file: string }>;
};
const chunkFile = (id: number) => index.chunks.find((chunk) => chunk.id === id)!.file;

const hits = new Map<string, number>();
const failing = new Set<string>();
let gate: Promise<void> | null = null;
const server = Bun.serve({
  port: 0,
  async fetch(request) {
    const file = decodeURIComponent(new URL(request.url).pathname.slice(1));
    hits.set(file, (hits.get(file) ?? 0) + 1);
    if (failing.delete(file)) return new Response('unavailable', { status: 503 });
    if (gate && index.chunks.some((chunk) => chunk.file === file)) await gate;
    return new Response(Bun.file(new URL(file, assets)));
  },
});
afterAll(() => server.stop(true));
beforeEach(() => {
  hits.clear();
  failing.clear();
  gate = null;
});

const files = [ref.index, index.skeleton.file, index.core.file, ...index.chunks.map((chunk) => chunk.file)];
const spec: ProgressiveFontSpec = {
  slot: 0,
  source: { ref, urls: Object.fromEntries(files.map((file) => [file, new URL(file, server.url).href])), timeoutMs: 5_000 },
};
/** A wait no load here outlasts, for tests about what installs rather than when. */
const UNTIL_INSTALLED = 10_000;
const decode = async (encoded: ArrayBuffer, maxOutput: number) => (await decodeFontFace(encoded, undefined, { maxOutput })).font;

const skeleton = new DataView(await decode(await Bun.file(new URL(index.skeleton.file, assets)).arrayBuffer(), 64 << 20));
const glyphIn = (chunk: number) => {
  const glyphs = skeleton.getUint32(8);
  for (let glyph = 0; glyph < glyphs; glyph += 1) if (skeleton.getUint16(32 + glyph * 2) === chunk) return glyph;
  throw new Error(`no glyph in chunk ${chunk}`);
};

/** The sliced JP face as font 10, with font 11 a measurement view of it. */
function fonts(engine: (fontId: number, bytes: Uint8Array) => void = () => {}) {
  const logs: string[] = [];
  const errors: Error[] = [];
  const updates: Array<{ fontId: number; length: number }> = [];
  const resident = new ResidentProgressiveFonts({
    updateFontOutlines(fontId, bytes) {
      engine(fontId, bytes);
      updates.push({ fontId, length: bytes.byteLength });
    },
  }, (message) => logs.push(message));
  resident.add(10, sliceFaceOpener(spec.source, decode), (error) => errors.push(error));
  resident.addView(11, 10);
  return { resident, logs, errors, updates };
}

test('a page walk collects the glyph ids of every run, including nested regions', () => {
  const page = {
    items: [{ kind: 'glyphRun', fontId: 10, glyphs: [{ id: 5 }, { id: 6 }] }],
    regions: [{ notes: [{ kind: 'glyphRun', fontId: 11, glyphs: [{ id: 7 }] }] }],
  };
  expect(glyphsByFont(page)).toEqual(new Map([[10, new Set([5, 6])], [11, new Set([7])]]));
});

test('outlines a frame draws install before it paints, one fetch per chunk across concurrent frames and views', async () => {
  const { resident, updates } = fonts();
  const glyph = glyphIn(41);
  const [first, second] = await Promise.all([
    resident.ensure(new Map([[10, new Set([glyph])]]), UNTIL_INSTALLED),
    resident.ensure(new Map([[11, new Set([glyph])]]), UNTIL_INSTALLED),
  ]);
  expect(new Set([...first, ...second])).toEqual(new Set([10, 11]));
  expect(updates.map((update) => update.fontId)).toEqual([10]);
  expect(hits.get(chunkFile(41))).toBe(1);
  // The core loads like any chunk, when a glyph needs it.
  expect([ref.index, index.skeleton.file, index.core.file].map((file) => hits.get(file))).toEqual([1, 1, undefined]);

  expect(await resident.ensure(new Map([[10, new Set([glyph])]]))).toEqual(new Set());
  expect(updates).toHaveLength(1);
  expect(await resident.ensure(new Map([[10, new Set([glyphIn(0)])]]), UNTIL_INSTALLED)).toEqual(new Set([10, 11]));
  expect(hits.get(index.core.file)).toBe(1);
  expect(await resident.ensure(new Map([[99, new Set([glyph])]]))).toEqual(new Set());
});

test('an opener handed the files the sender decoded fetches only the index and the chunks it lacks', async () => {
  const decoded = async (file: string) => decode(await Bun.file(new URL(file, assets)).arrayBuffer(), 64 << 20);
  const files = new Map([[index.skeleton.file, await decoded(index.skeleton.file)], [chunkFile(42), await decoded(chunkFile(42))]]);
  const resident = new ResidentProgressiveFonts({ updateFontOutlines() {} });
  resident.add(10, sliceFaceOpener(spec.source, decode, files), () => {});
  expect(await resident.ensure(new Map([[10, new Set([glyphIn(42), glyphIn(43)])]]))).toEqual(new Set([10]));
  expect([ref.index, index.skeleton.file, chunkFile(42), chunkFile(43)].map((file) => hits.get(file))).toEqual([1, undefined, undefined, 1]);
});

test('a frame stops waiting at the guard, logs once, and repaints when the outlines arrive', async () => {
  const { resident, logs, updates } = fonts();
  let release!: () => void;
  gate = new Promise((resolve) => {
    release = resolve;
  });
  const late: Array<Set<number>> = [];
  resident.onLateInstall((fontIds) => late.push(fontIds));
  expect(await resident.ensure(new Map([[10, new Set([glyphIn(50)])]]), 30)).toEqual(new Set());
  expect(await resident.ensure(new Map([[10, new Set([glyphIn(51)])]]), 30)).toEqual(new Set());
  expect(logs.filter((message) => message.includes('still loading after 30 ms'))).toHaveLength(1);
  expect(updates).toHaveLength(0);
  release();
  await resident.whenInstalled();
  expect(late.length).toBeGreaterThanOrEqual(1);
  expect(late[0]).toEqual(new Set([10, 11]));
  expect(resident.lateInstalls).toBe(late.length);
  expect([chunkFile(50), chunkFile(51)].map((file) => hits.get(file))).toEqual([1, 1]);
});

test('a chunk slower than the wait yields the frame within the wait and repaints once it lands, inside every frame request budget', async () => {
  const { resident, updates } = fonts();
  // The skeleton is open before the slow request, as for any later frame.
  await resident.ensure(new Map([[10, new Set([glyphIn(0)])]]), UNTIL_INSTALLED);
  let release!: () => void;
  gate = new Promise((resolve) => {
    release = resolve;
  });
  const late: Array<Set<number>> = [];
  resident.onLateInstall((fontIds) => late.push(fontIds));
  const started = performance.now();
  expect(await resident.ensure(new Map([[10, new Set([glyphIn(54)])]]))).toEqual(new Set());
  const waited = performance.now() - started;
  expect(waited).toBeGreaterThanOrEqual(PROGRESSIVE_OUTLINE_WAIT_MS - 20);
  expect(waited).toBeLessThan(PROGRESSIVE_OUTLINE_WAIT_MS + 500);
  for (const type of ['buildFrame', 'applyInput', 'applyDelete'] as const) {
    expect(PROGRESSIVE_OUTLINE_WAIT_MS * 10).toBeLessThanOrEqual(REQUEST_TIMEOUT_MS[type]);
  }
  release();
  await resident.whenInstalled();
  expect(updates).toHaveLength(2);
  expect(late).toEqual([new Set([10, 11])]);
});

test('a chunk that fails is reported once per face, not fetched again until a retry, and the rest of the face still installs', async () => {
  const { resident, updates, errors } = fonts();
  failing.add(chunkFile(60));
  expect(await resident.ensure(new Map([[10, new Set([glyphIn(60)])]]), UNTIL_INSTALLED)).toEqual(new Set());
  expect(await resident.ensure(new Map([[11, new Set([glyphIn(60), glyphIn(61)])]]), UNTIL_INSTALLED)).toEqual(new Set([10, 11]));
  expect(errors).toHaveLength(1);
  expect(errors[0]).toBeInstanceOf(FontSliceUnavailableError);
  expect(errors[0]).toMatchObject({ file: chunkFile(60), reason: 'server', status: 503 });
  expect(hits.get(chunkFile(60))).toBe(1);
  expect(hits.get(chunkFile(61))).toBe(1);
  expect(updates).toHaveLength(1);

  // Back online: asked for again on the next frame, and it installs.
  resident.retryUnavailable();
  expect(await resident.ensure(new Map([[10, new Set([glyphIn(60)])]]), UNTIL_INSTALLED)).toEqual(new Set([10, 11]));
  expect(hits.get(chunkFile(60))).toBe(2);
  expect(errors).toHaveLength(1);
});

test('an engine that refuses the outlines is reported once and never sent outlines again', async () => {
  let calls = 0;
  const { resident, errors } = fonts(() => {
    calls += 1;
    throw new Error('tables differ');
  });
  expect(await resident.ensure(new Map([[10, new Set([glyphIn(70)])]]))).toEqual(new Set());
  expect(await resident.ensure(new Map([[11, new Set([glyphIn(71)])]]))).toEqual(new Set());
  expect(calls).toBe(1);
  expect(errors.map((error) => error.message)).toEqual(['tables differ']);
  expect(hits.get(chunkFile(71))).toBeUndefined();
});

test('fonts that were replaced never install outlines that arrive later', async () => {
  const { resident, updates } = fonts();
  let release!: () => void;
  gate = new Promise((resolve) => {
    release = resolve;
  });
  let repaints = 0;
  resident.onLateInstall(() => {
    repaints += 1;
  });
  expect(await resident.ensure(new Map([[10, new Set([glyphIn(52)])]]), 20)).toEqual(new Set());
  resident.dispose();
  release();
  await resident.whenInstalled();
  expect(updates).toEqual([]);
  expect(repaints).toBe(0);
  expect(await resident.ensure(new Map([[10, new Set([glyphIn(53)])]]))).toEqual(new Set());
  expect(hits.get(chunkFile(53))).toBeUndefined();
});

test('list markers and watermarks on a sliced face load their characters\' chunks and paint with it under a private family, blank until the chunks land', async () => {
  const realm = globalThis as { FontFace?: unknown; fonts?: unknown };
  class RealmFontFace {
    constructor(readonly family: string, readonly source: ArrayBuffer) {}
    async load() {
      return this;
    }
  }
  const registered: RealmFontFace[] = [];
  const removed: RealmFontFace[] = [];
  realm.FontFace = RealmFontFace;
  realm.fonts = { add: (face: RealmFontFace) => registered.push(face), delete: (face: RealmFontFace) => removed.push(face) };
  try {
    const { resident } = fonts();
    const late: Array<Set<number>> = [];
    resident.onLateInstall((fontIds) => late.push(fontIds));
    // A circled-number marker and a watermark in MS Mincho, whose chain starts with font 11 (a view of the sliced face 10), and a Latin marker.
    const page = {
      items: [
        { kind: 'text', font: '14px "MS Mincho", serif', text: '①' },
        { kind: 'text', font: 'bold 48px "MS Mincho"', text: '社外秘' },
        { kind: 'text', font: '11px Calibri', text: '1.' },
      ],
    };
    const chains = { 'ms mincho|0|0': [11, 3], 'calibri|0|0': [2] };
    let release!: () => void;
    gate = new Promise((resolve) => {
      release = resolve;
    });
    const families = await resident.textFamilies(page, chains, 30);
    expect(families).toEqual(new Map([['MS Mincho', 'BetterOffice Sliced 10']]));
    expect(registered.map((face) => face.family)).toEqual(['BetterOffice Sliced 10']);
    const chunkHits = () => index.chunks.filter((chunk) => hits.has(chunk.file)).length;
    expect(chunkHits()).toBeGreaterThan(0);

    release();
    await resident.whenInstalled();
    expect(late).toEqual([new Set([10, 11])]);
    // The repaint registers the face with the landed outlines in place of the blank one.
    expect(await resident.textFamilies(page, chains, 30)).toEqual(families);
    expect(registered).toHaveLength(2);
    expect(registered[1]!.source.byteLength).toBeGreaterThan(registered[0]!.source.byteLength);
    expect(removed).toEqual([registered[0]!]);
    resident.dispose();
    expect(removed).toEqual([registered[0]!, registered[1]!]);
  } finally {
    delete realm.FontFace;
    delete realm.fonts;
  }
});
