import { expect, test } from 'bun:test';
import type { GlyphCache } from '@betteroffice/docx/layout/render';
import type { SlicedPaint } from '@betteroffice/docx/yrs';
import { printPageOptions } from './printPaint';

test('print and thumbnails ready every outline a page draws first, then paint its glyph runs from outlines and its markers in the registered face', async () => {
  const invalidated: Array<Set<number>> = [];
  const glyphCache = { invalidateFonts: (fontIds: Set<number>) => invalidated.push(fontIds) } as unknown as GlyphCache;
  const page = { pageIndex: 0 };
  const readied: unknown[][] = [];
  const sliced: SlicedPaint = {
    changed: new Set([7]),
    outlineOnlyFontIds: new Set([7, 8]),
    fontFamilies: new Map([['MS Gothic', 'BetterOffice Sliced 7']]),
  };
  const options = await printPageOptions(page, {
    glyphCache,
    prepareSlicedPaint: async (...args) => {
      readied.push(args);
      return sliced;
    },
  });
  // Complete: no cut-off, as the screen's 300 ms.
  expect(readied).toEqual([[page, true]]);
  expect(invalidated).toEqual([new Set([7])]);
  expect(options).toEqual({
    resolveImage: undefined,
    glyphCache,
    outlineOnlyFontIds: sliced.outlineOnlyFontIds,
    fontFamilies: sliced.fontFamilies,
  });
});

test('a page whose font slice failed fails print with an explicit error, never blank characters', async () => {
  const failure = Object.assign(new Error('Failed to fetch bundled font NotoSansJP-Regular/c093.zst: HTTP 503'), { file: 'NotoSansJP-Regular/c093.zst' });
  const printing = printPageOptions({ pageIndex: 0 }, { prepareSlicedPaint: async () => Promise.reject(failure) });
  await expect(printing).rejects.toThrow('Cannot render the page for print or a thumbnail: font outlines it draws did not load (Failed to fetch bundled font NotoSansJP-Regular/c093.zst: HTTP 503)');
  await expect(printing).rejects.toMatchObject({ cause: failure });
});
