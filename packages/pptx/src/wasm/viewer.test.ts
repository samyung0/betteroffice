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
const embeddedLato = resolve(root, 'crates/pptx-render/tests/fixtures/embedded-lato.pptx');
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

  test('lays out in the embedded faces of the deck and hands them to the page', async () => {
    const [deck, fontBytes, plain] = await Promise.all([
      readFile(embeddedLato),
      readFile(font),
      readFile(fixture),
    ]);
    const fonts = [{ family: 'Arial', bytes: new Uint8Array(fontBytes) }];
    const viewer = openViewerPresentation(new Uint8Array(deck), { fonts });
    const editor = openEditorPresentation(new Uint8Array(deck), { clientId: 812, fonts });
    const other = openViewerPresentation(new Uint8Array(plain), { fonts });
    try {
      const faces = viewer.embeddedFonts();
      for (const handle of [viewer, editor]) {
        const own = handle.embeddedFonts();
        expect(own.map(({ typeface, bold, italic }) => ({ typeface, bold, italic }))).toEqual([
          { typeface: 'Lato', bold: false, italic: false },
          { typeface: 'Lato', bold: true, italic: false },
        ]);
        // Named by an alias of the bytes, the same for both handles.
        expect(own.map(({ family }) => family)).toEqual(faces.map(({ family }) => family));
        for (const face of own) expect(face.family).toMatch(/^bo-embedded-[0-9a-f]{16}$/);
        // Decoded from MicroType Express to a TrueType sfnt the page can load.
        for (const face of own) expect([...face.bytes.subarray(0, 4)]).toEqual([0, 1, 0, 0]);
      }
      const frame = viewer.layoutSlide(0);
      expect(frame).toEqual(editor.layoutSlide(0));
      const box = frame.primitives.find(
        (primitive) => primitive.kind === 'textBox' && primitive.objectId === 4
      );
      if (box?.kind !== 'textBox') throw new Error('shape 4 has no text box');
      expect(box.lines.map((line) => line.runs.map((run) => run.text).join(''))).toEqual([
        'Embedded as text ',
        'it wraps later',
        'Second line',
      ]);
      expect(box.lines[0].runs[0].fontFamily).toBe(faces[0].family);
      expect(other.embeddedFonts()).toEqual([]);
      // Faces the page refused lay out in the host face again.
      for (const face of faces) viewer.refuseEmbeddedFont(face.fontId);
      const refused = viewer.layoutSlide(0);
      const fallback = refused.primitives.find(
        (primitive) => primitive.kind === 'textBox' && primitive.objectId === 4
      );
      if (fallback?.kind !== 'textBox') throw new Error('shape 4 has no text box');
      expect(fallback.lines[0].runs[0].fontFamily).toBe('Arial');
      expect(fallback.lines.map((line) => line.runs.map((run) => run.text).join(''))).toEqual([
        'Embedded as ',
        'text it wraps later',
        'Second line',
      ]);
    } finally {
      viewer.dispose();
      editor.dispose();
      other.dispose();
    }
  });
});
