import type { DrawPageOptions, GlyphCache, ImageResolver } from '@betteroffice/docx/layout/render';
import type { SlicedPaint } from '@betteroffice/docx/yrs';

/**
 * How print and thumbnails (`renderPages`) paint a page: glyph runs from the
 * engine's outlines, as the screen does, with the page's sliced faces readied
 * first (their glyph runs outline-only, their markers in the registered
 * face), so CJK text never paints in a system font.
 */
export async function printPageOptions(
  page: unknown,
  {
    resolveImage,
    glyphCache,
    prepareSlicedPaint,
  }: {
    resolveImage?: ImageResolver;
    glyphCache?: GlyphCache;
    prepareSlicedPaint?: ((page: unknown) => Promise<SlicedPaint>) | null;
  }
): Promise<DrawPageOptions> {
  const sliced = prepareSlicedPaint ? await prepareSlicedPaint(page) : null;
  if (sliced && sliced.changed.size > 0) glyphCache?.invalidateFonts(sliced.changed);
  return {
    resolveImage,
    glyphCache,
    ...(sliced && { outlineOnlyFontIds: sliced.outlineOnlyFontIds, fontFamilies: sliced.fontFamilies }),
  };
}
