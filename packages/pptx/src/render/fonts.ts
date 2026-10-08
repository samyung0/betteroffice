import type { PptxEmbeddedFontFace } from '../types';

/** A deck's handle, as far as {@link installEmbeddedFonts} needs it. */
export interface EmbeddedFontSource {
  embeddedFonts(): PptxEmbeddedFontFace[];
  refuseEmbeddedFont(fontId: number): void;
}

/**
 * Adds a deck's embedded faces to the page's fonts under their per-deck
 * aliases, so canvas text paints in the faces layout measured and nothing
 * else on the page changes. A face the browser refuses is left out with a
 * warning and dropped from the deck's layout too, so its text falls back in
 * both. Resolves to the faces added, for {@link removeFontFaces} when the
 * deck closes. Does nothing outside a DOM.
 */
export async function installEmbeddedFonts(deck: EmbeddedFontSource): Promise<FontFace[]> {
  if (typeof FontFace === 'undefined' || typeof document === 'undefined') return [];
  const loaded = await Promise.all(
    deck.embeddedFonts().map(async (face) => {
      try {
        const font = await new FontFace(face.family, face.bytes.slice().buffer as ArrayBuffer, {
          style: face.italic ? 'italic' : 'normal',
          weight: face.bold ? '700' : '400',
        }).load();
        document.fonts.add(font);
        return font;
      } catch (error) {
        console.warn(`[pptx] the embedded face "${face.typeface}" did not load`, error);
        try {
          deck.refuseEmbeddedFont(face.fontId);
        } catch {
          // The deck closed while the face loaded.
        }
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
