import { beforeAll, describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import {
  initWasm as initEditorWasm,
  openPresentation as openEditorPresentation,
} from './loader';
import {
  analyzePresentation,
  initWasm as initViewerWasm,
  openPresentation as openViewerPresentation,
} from './viewer';

const root = resolve(import.meta.dir, '../../../..');
const fixture = resolve(root, 'apps/demo/public/betteroffice-demo.pptx');
const font = resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf');
const editorWasm = resolve(import.meta.dir, 'generated/pptx_wasm_bg.wasm');
const viewerWasm = resolve(import.meta.dir, 'generated/viewer/pptx_view_wasm_bg.wasm');

beforeAll(async () => {
  const [editor, viewer] = await Promise.all([readFile(editorWasm), readFile(viewerWasm)]);
  await Promise.all([initEditorWasm(editor), initViewerWasm(viewer)]);
});

describe('PPTX viewer wasm', () => {
  test('rejects bytes that are not a PPTX package', () => {
    expect(() => openViewerPresentation(Uint8Array.of(0x00, 0x01, 0x02, 0x03))).toThrow();
  });

  test('matches the unedited editor snapshot and display output', async () => {
    const [deck, fontBytes] = await Promise.all([readFile(fixture), readFile(font)]);
    const fonts = [{ family: 'Liberation Sans', bytes: new Uint8Array(fontBytes) }];
    const viewer = openViewerPresentation(new Uint8Array(deck), { fonts });
    const editor = openEditorPresentation(new Uint8Array(deck), { clientId: 811, fonts });
    try {
      expect(viewer.snapshot()).toEqual(editor.snapshot());
      expect(viewer.layoutSlide(0)).toEqual(editor.layoutSlide(0));
      expect('save' in viewer).toBe(false);
      expect('insertText' in viewer).toBe(false);
      expect('encodeStateAsUpdate' in viewer).toBe(false);
    } finally {
      viewer.dispose();
      editor.dispose();
    }
  });

  test('analyzes presentation metadata without retaining a document handle', async () => {
    const bytes = new Uint8Array(await readFile(fixture));
    const analysis = analyzePresentation(bytes);
    expect(analysis.format).toBe('pptx');
    expect(analysis.slideCount).toBeGreaterThan(0);
    expect(analysis.widthEmu).toBeGreaterThan(0);
    expect(analysis.heightEmu).toBeGreaterThan(0);
    expect(analysis.textCharacterCount).toBeGreaterThan(0);
  });
});
