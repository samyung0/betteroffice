import { expect, mock, spyOn, test } from 'bun:test';
import type { SliceLocation } from '@betteroffice/fonts';
import type { BundledFontLoader } from './fontRegistry';
import { createRustMeasureSource, type RustTextEngine } from './rustMeasureSource';

// Offline before anything opened a sliced face: the module that opens them cannot load.
mock.module('@betteroffice/fonts/progressive', () => {
  throw new TypeError('Failed to fetch dynamically imported module');
});

test('the fonts module failing to load fails the face as its index failing to load would (Q763 A)', async () => {
  const korean: BundledFontLoader = Object.assign(async () => new ArrayBuffer(0), {
    slices: { ref: { format: 1, index: 'NotoSansKR-Regular/slices.json', indexSha256: '0'.repeat(64) }, url: async (file: string) => new URL(file, 'http://127.0.0.1:9/') } as unknown as SliceLocation,
  });
  const reported: Error[] = [];
  const engine: RustTextEngine = {
    registerFont: () => 1,
    registerSlicedFont() {},
    clearFonts() {},
    reportSlicedFaceError: (error) => reported.push(error),
  };
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  try {
    const source = createRustMeasureSource({
      engine,
      bundled: { resolve: () => undefined, resolveScriptFallback: (script) => (script === 'cjk-kr' ? korean : undefined) },
      progressive: true,
      onError: () => {},
    });
    await source.prepareFontRequirements([{ key: 'calibri|0|0', family: 'Calibri', bold: false, italic: false, scripts: ['cjk-kr'] }]);
    expect(reported).toHaveLength(1);
    expect(reported[0]).toMatchObject({ file: 'NotoSansKR-Regular/slices.json', reason: 'network' });
  } finally {
    warn.mockRestore();
  }
});
