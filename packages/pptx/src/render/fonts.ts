import type { PptxFontFace } from '../types';

/**
 * Adds a deck's embedded faces to the page's fonts, so canvas text paints in
 * the faces layout measured. A face the browser refuses is left out with a
 * warning: its text paints in the page's fallback, as before embedded fonts.
 * Resolves to the faces added, for {@link removeFontFaces} when the deck
 * closes. Does nothing outside a DOM.
 */
export async function installEmbeddedFonts(
  faces: ReadonlyArray<PptxFontFace>
): Promise<FontFace[]> {
  if (typeof FontFace === 'undefined' || typeof document === 'undefined') return [];
  const loaded = await Promise.all(
    faces.map(async (face) => {
      try {
        const font = await new FontFace(face.family, face.bytes.slice().buffer as ArrayBuffer, {
          style: face.italic ? 'italic' : 'normal',
          weight: face.bold ? '700' : '400',
        }).load();
        document.fonts.add(font);
        return font;
      } catch (error) {
        console.warn(`[pptx] the embedded face "${face.family}" did not load`, error);
        return null;
      }
    })
  );
  return loaded.filter((font): font is FontFace => font !== null);
}

/** Removes faces {@link installEmbeddedFonts} added. */
export function removeFontFaces(faces: ReadonlyArray<FontFace>): void {
  if (typeof document === 'undefined') return;
  for (const face of faces) document.fonts.delete(face);
}
