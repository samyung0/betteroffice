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
    files: () => new Map(),
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
