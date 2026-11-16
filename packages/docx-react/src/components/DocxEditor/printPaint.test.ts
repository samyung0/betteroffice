import { expect, test } from 'bun:test';
import type { GlyphCache } from '@betteroffice/docx/layout/render';
import type { SlicedPaint } from '@betteroffice/docx/yrs';
import { printPageOptions } from './printPaint';

test('print and thumbnails ready a page\'s sliced faces first, then paint its glyph runs from outlines and its markers in the registered face', async () => {
  const invalidated: Array<Set<number>> = [];
  const glyphCache = { invalidateFonts: (fontIds: Set<number>) => invalidated.push(fontIds) } as unknown as GlyphCache;
  const page = { pageIndex: 0 };
  const readied: unknown[] = [];
  const sliced: SlicedPaint = {
    changed: new Set([7]),
    outlineOnlyFontIds: new Set([7, 8]),
    fontFamilies: new Map([['MS Gothic', 'BetterOffice Sliced 7']]),
  };
  const options = await printPageOptions(page, {
    glyphCache,
    prepareSlicedPaint: async (ready) => {
      readied.push(ready);
      return sliced;
    },
  });
  expect(readied).toEqual([page]);
  expect(invalidated).toEqual([new Set([7])]);
  expect(options).toEqual({
    resolveImage: undefined,
    glyphCache,
    outlineOnlyFontIds: sliced.outlineOnlyFontIds,
    fontFamilies: sliced.fontFamilies,
  });
});
