import type { DrawPageOptions, GlyphCache, ImageResolver } from '@betteroffice/docx/layout/render';
import type { SlicedPaint } from '@betteroffice/docx/yrs';

/**
 * How print and thumbnails (`renderPages`) paint a page: glyph runs from the
 * engine's outlines, as the screen does, with the page's sliced faces readied
 * first (their glyph runs outline-only, their markers in the registered
 * face), so CJK text never paints in a system font. Print waits for every
 * outline the page draws, with no cut-off; a slice that failed fails the page
 * with an explicit error, so nothing prints with characters missing.
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
    prepareSlicedPaint?: ((page: unknown, complete?: boolean) => Promise<SlicedPaint>) | null;
  }
): Promise<DrawPageOptions> {
  let sliced: SlicedPaint | null = null;
  if (prepareSlicedPaint) {
    try {
      sliced = await prepareSlicedPaint(page, true);
    } catch (error) {
      throw new Error(
        `Cannot render the page for print or a thumbnail: font outlines it draws did not load (${error instanceof Error ? error.message : String(error)})`,
        { cause: error }
      );
    }
  }
  if (sliced && sliced.changed.size > 0) glyphCache?.invalidateFonts(sliced.changed);
  return {
    resolveImage,
    glyphCache,
    ...(sliced && { outlineOnlyFontIds: sliced.outlineOnlyFontIds, fontFamilies: sliced.fontFamilies }),
  };
}
