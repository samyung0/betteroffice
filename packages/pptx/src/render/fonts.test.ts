import { describe, expect, spyOn, test } from 'bun:test';
import { installEmbeddedFonts, removeFontFaces } from './fonts';

class FakeFontFace {
  constructor(
    readonly family: string,
    readonly source: ArrayBuffer,
    readonly descriptors: FontFaceDescriptors
  ) {}

  load(): Promise<this> {
    return this.family === 'Refused'
      ? Promise.reject(new Error('OTS parsing error'))
      : Promise.resolve(this);
  }
}

describe('installEmbeddedFonts', () => {
  test('adds the faces the page loads, leaves out refused ones and removes them again', async () => {
    const globals = globalThis as Record<string, unknown>;
    const saved = { FontFace: globals.FontFace, document: globals.document };
    const fonts = new Set<unknown>();
    globals.FontFace = FakeFontFace;
    globals.document = { fonts };
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const bytes = Uint8Array.of(0, 1, 0, 0);
      const added = await installEmbeddedFonts([
        { family: 'Lato', bold: true, italic: true, bytes },
        { family: 'Refused', bytes },
      ]);
      const faces = added as unknown as FakeFontFace[];
      expect(faces.map((face) => face.family)).toEqual(['Lato']);
      expect(faces[0].descriptors).toEqual({ style: 'italic', weight: '700' });
      expect(faces[0].source.byteLength).toBe(4);
      expect([...fonts]).toEqual(faces);
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
    expect(await installEmbeddedFonts([{ family: 'Lato', bytes: Uint8Array.of(0) }])).toEqual([]);
  });
});
