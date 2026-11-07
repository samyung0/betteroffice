import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { expectFontFace, fontCases, fontErrorCases } from '../../../../bindings/font-fixtures';
import { decodeFontFace } from '../core';
import { preloadEditWasm } from './edit';

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(new URL('./generated/edit/docx_edit_bg.wasm', import.meta.url)))), 30_000);

test('the root font decoder decodes packs with the DOCX edit engine, returning ArrayBuffers', async () => {
  for (const entry of fontCases()) {
    const backing = new Uint8Array(entry.bytes.length + 2);
    backing.set(entry.bytes, 1);
    const decoded = await decodeFontFace(backing.subarray(1, -1), entry);
    expect(decoded.font).toBeInstanceOf(ArrayBuffer);
    if (decoded.payload !== undefined) expect(decoded.payload).toBeInstanceOf(ArrayBuffer);
    expectFontFace({ font: new Uint8Array(decoded.font),
      ...(decoded.payload === undefined ? {} : { payload: new Uint8Array(decoded.payload) }) }, entry);
    expect(backing.subarray(1, -1)).toEqual(new Uint8Array(entry.bytes));
  }
  for (const entry of fontErrorCases()) {
    await expect(decodeFontFace(entry.bytes.slice(), entry)).rejects.toThrow(new Error(entry.message));
  }
});
