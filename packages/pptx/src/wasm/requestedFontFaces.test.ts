import { beforeAll, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { initWasm, openPresentation } from '../index';

const root = resolve(import.meta.dir, '../../../..');
let deck: Uint8Array;

beforeAll(async () => {
  const [wasm, pptx] = await Promise.all([
    readFile(resolve(import.meta.dir, 'generated/pptx_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
  ]);
  await initWasm(wasm);
  deck = pptx;
});

test('the deck lists every face its slides request, before any layout', () => {
  const handle = openPresentation(deck, { clientId: 9301 });
  try {
    const faces = handle.requestedFontFaces();
    expect(faces.length).toBeGreaterThan(0);
    const key = (face: { family: string; bold: boolean; italic: boolean }) =>
      `${face.family}|${face.bold}|${face.italic}`;
    const deckKeys = new Set(faces.map(key));
    for (const slide of handle.snapshot().slides) {
      for (const face of handle.requestedFontFaces(slide.id)) expect(deckKeys.has(key(face))).toBe(true);
    }
    expect(() => handle.requestedFontFaces('missing')).toThrow();
  } finally {
    handle.dispose();
  }
});
