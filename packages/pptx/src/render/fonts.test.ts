import { describe, expect, spyOn, test } from 'bun:test';
import type { PptxEmbeddedFontFace } from '../types';
import { installEmbeddedFonts, removeFontFaces } from './fonts';

class FakeFontFace {
  constructor(
    readonly family: string,
    readonly source: ArrayBuffer,
    readonly descriptors: FontFaceDescriptors
  ) {}

  load(): Promise<this> {
    return this.family === 'bo-embedded-refused'
      ? Promise.reject(new Error('OTS parsing error'))
      : Promise.resolve(this);
  }
}

const face = (family: string, fontId: number, bold = false, italic = false): PptxEmbeddedFontFace => ({
  family,
  typeface: 'Lato',
  fontId,
  bold,
  italic,
  bytes: Uint8Array.of(0, 1, 0, 0),
});

describe('installEmbeddedFonts', () => {
  // Security regression (REVIEW1 S2): a refused face is dropped from layout too.
  test('adds the faces the page loads under their aliases and drops refused ones from layout', async () => {
    const globals = globalThis as Record<string, unknown>;
    const saved = { FontFace: globals.FontFace, document: globals.document };
    const fonts = new Set<unknown>();
    globals.FontFace = FakeFontFace;
    globals.document = { fonts };
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    const refused: number[] = [];
    try {
      const added = await installEmbeddedFonts({
        embeddedFonts: () => [face('bo-embedded-a', 3, true, true), face('bo-embedded-refused', 4)],
        refuseEmbeddedFont: (fontId) => refused.push(fontId),
      });
      const faces = added as unknown as FakeFontFace[];
      expect(faces.map((each) => each.family)).toEqual(['bo-embedded-a']);
      expect(faces[0].descriptors).toEqual({ style: 'italic', weight: '700' });
      expect(faces[0].source.byteLength).toBe(4);
      expect([...fonts]).toEqual(faces);
      expect(refused).toEqual([4]);
      expect(warn).toHaveBeenCalledTimes(1);
      removeFontFaces(added);
      expect(fonts.size).toBe(0);
    } finally {
      warn.mockRestore();
      globals.FontFace = saved.FontFace;
      globals.document = saved.document;
    }
  });

  test('does nothing outside a DOM', async () => {
    expect(typeof FontFace).toBe('undefined');
    const refused: number[] = [];
    const added = await installEmbeddedFonts({
      embeddedFonts: () => [face('bo-embedded-a', 1)],
      refuseEmbeddedFont: (fontId) => refused.push(fontId),
    });
    expect(added).toEqual([]);
    expect(refused).toEqual([]);
  });
});
