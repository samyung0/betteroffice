import { beforeAll, describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { initWasm, openDocumentViewer } from '../viewer';

const root = resolve(import.meta.dir, '../../../..');
const fixture = resolve(root, 'poc/fixtures/feature-rich.docx');
const viewerWasm = resolve(
  import.meta.dir,
  'generated/viewer/docx_view_wasm_bg.wasm'
);

beforeAll(async () => {
  await initWasm(await readFile(viewerWasm));
});

describe('DOCX viewer wasm', () => {
  test('rejects bytes that are not a DOCX package', async () => {
    await expect(
      openDocumentViewer(Uint8Array.of(0x00, 0x01, 0x02, 0x03))
    ).rejects.toThrow();
  });

  test('renders the feature-rich fixture without exposing editing methods', async () => {
    const viewer = await openDocumentViewer(
      new Uint8Array(await readFile(fixture))
    );
    try {
      const displayList = viewer.displayList();
      expect(displayList.pages.length).toBeGreaterThan(0);
      expect(displayList.pages.some((page) => page.primitives.length > 0)).toBe(
        true
      );
      expect('save' in viewer).toBe(false);
      expect('applyInput' in viewer).toBe(false);
      expect('encodeStateAsUpdate' in viewer).toBe(false);
    } finally {
      viewer.dispose();
    }
  });
});
