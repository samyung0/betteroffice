import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import type { SliceLocation } from '@betteroffice/fonts';
import { TextMeasureFontRegistry, type BundledFontLoader, type ProgressiveFontFace } from './fontRegistry';

// A real face's bytes stand in for a skeleton: the registry only registers them.
const bytes = Uint8Array.from(readFileSync(new URL('../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url))).buffer;

function fakeFace(): ProgressiveFontFace {
  return {
    revision: 1,
    bytes: () => bytes.slice(0),
    ensureGlyphs: async () => 1,
    ensureText: async () => 1,
    prefetchCore: () => {},
    files: () => new Map(),
    retryFailed: () => {},
    markFailed: () => {},
    failures: () => new Map(),
  };
}

/** A loader whose face opens sliced, from `slices`. */
function slicedLoader(index: string): BundledFontLoader {
  return Object.assign(async () => bytes, { slices: { ref: { format: 1, index, indexSha256: '' } } as unknown as SliceLocation });
}

function deferred() {
  let resolve!: () => void;
  const promise = new Promise<void>((done) => {
    resolve = done;
  });
  return { promise, resolve };
}

test('a sliced face whose open outlasts setEmbeddedFaces still registers its outlines with the engine', async () => {
  const started = deferred();
  const opening = deferred();
  const face = fakeFace();
  const loader = slicedLoader('NotoSansJP-Regular/slices.json');
  const sliced: number[] = [];
  let next = 0;
  const registry = new TextMeasureFontRegistry(
    { registerFont: () => next++, registerSlicedFont: (id) => sliced.push(id) },
    {
      bundled: { resolve: () => loader },
      progressive: async () => {
        started.resolve();
        await opening.promise;
        return face;
      },
    },
  );
  const first = registry.getFontIdChain('MS Gothic', false, false);
  await started.promise;
  // The document's embedded faces change while the skeleton downloads: the engine's store stays.
  registry.setEmbeddedFaces([]);
  opening.resolve();
  await first;
  const [id] = await registry.getFontIdChain('MS Gothic', false, false);
  expect(sliced).toEqual([id!]);
  expect(registry.progressiveFaces().get(id!)?.face).toBe(face);
});

test('a chain that wants the core of its face starts it although a script fallback opened the face first without it (Q758 B)', async () => {
  const cores: string[] = [];
  const opens: boolean[] = [];
  const face = { ...fakeFace(), prefetchCore: () => cores.push('NotoSansJP') };
  // One loader for the face, whether a family or a script reaches it.
  const loader = slicedLoader('NotoSansJP-Regular/slices.json');
  let next = 0;
  const registry = new TextMeasureFontRegistry({ registerFont: () => next++ }, {
    bundled: { resolve: () => loader, resolveScriptFallback: () => loader },
    progressive: async (_slices, { core }) => {
      opens.push(core);
      return face;
    },
  });
  await registry.getScriptFallbackIds(['cjk-jp']);
  await registry.getFontIdChain('MS Gothic', false, false, true);
  expect(opens).toEqual([false]);
  expect(cores).toEqual(['NotoSansJP']);
});

test('a sliced face whose skeleton fails is reported once and not fetched again on later layout passes', async () => {
  const failure = new Error('Failed to fetch bundled font NotoSansJP-Regular/skeleton.zst: HTTP 503');
  let opens = 0;
  const errors: Error[] = [];
  let next = 0;
  const registry = new TextMeasureFontRegistry({ registerFont: () => next++ }, {
    // A new loader on every resolve, as a host's provider may hand out.
    bundled: { resolve: () => slicedLoader('NotoSansJP-Regular/slices.json') },
    progressive: async () => {
      opens += 1;
      throw failure;
    },
    onError: (error) => errors.push(error),
  });
  await registry.getFontIdChain('MS Gothic', false, false);
  // Each later pass resolves the chain again, since it failed.
  await registry.getFontIdChain('MS Gothic', false, false);
  registry.setEmbeddedFaces([]);
  await registry.getFontIdChain('MS Gothic', false, false);
  expect(opens).toBe(1);
  expect(errors).toEqual([failure]);
});

test('every face of a chain starts loading at once, and the chain keeps its order whatever lands first', async () => {
  const started: string[] = [];
  const gates = new Map<string, () => void>();
  // Each face's bytes differ in length, which names it once registered.
  const lengths = new Map([['primary', 0], ['last resort', 1], ['coverage', 2]]);
  const gated = (name: string) => () => {
    started.push(name);
    return new Promise<ArrayBuffer>((resolve) => gates.set(name, () => resolve(bytes.slice(0, bytes.byteLength - lengths.get(name)!))));
  };
  const registered = new Map<number, string>();
  let next = 0;
  const registry = new TextMeasureFontRegistry({
    registerFont: (data) => {
      registered.set(next, [...lengths].find(([, cut]) => data.byteLength === bytes.byteLength - cut)![0]);
      return next++;
    },
  }, {
    bundled: {
      resolve: () => gated('primary'),
      resolveLastResort: () => gated('last resort'),
      resolveGlyphFallback: () => gated('coverage'),
    },
  });
  const chain = registry.getFontIdChain('Calibri', false, false);
  for (let turn = 0; turn < 50 && started.length < 3; turn += 1) await new Promise((resolve) => setTimeout(resolve, 1));
  expect(started.sort()).toEqual(['coverage', 'last resort', 'primary']);
  for (const name of ['coverage', 'last resort', 'primary']) gates.get(name)!();
  expect((await chain).map((id) => registered.get(id))).toEqual(['primary', 'last resort', 'coverage']);
});
