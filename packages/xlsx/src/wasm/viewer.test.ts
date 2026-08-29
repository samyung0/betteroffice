import { beforeAll, describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import {
  initWasm as initEditorWasm,
  openWorkbook as openEditorWorkbook,
} from './loader';
import {
  analyzeWorkbook,
  initWasm as initViewerWasm,
  openWorkbook as openViewerWorkbook,
} from './viewer';

const fixture = resolve(import.meta.dir, '../../test-fixtures/sample.xlsx');
const chartFixture = resolve(import.meta.dir, '../../test-fixtures/charts.xlsx');
const editorWasm = resolve(import.meta.dir, 'generated/xlsx_wasm_bg.wasm');
const viewerWasm = resolve(import.meta.dir, 'generated/viewer/xlsx_view_wasm_bg.wasm');

beforeAll(async () => {
  const [editor, viewer] = await Promise.all([readFile(editorWasm), readFile(viewerWasm)]);
  await Promise.all([initEditorWasm(editor), initViewerWasm(viewer)]);
});

describe('XLSX viewer wasm', () => {
  test('rejects bytes that are not an XLSX package', () => {
    expect(() => openViewerWorkbook(Uint8Array.of(0x00, 0x01, 0x02, 0x03))).toThrow();
  });

  test('matches editor metadata and display output without edit methods', async () => {
    const bytes = new Uint8Array(await readFile(fixture));
    const viewer = openViewerWorkbook(bytes);
    const editor = openEditorWorkbook(bytes);
    try {
      expect(viewer.sheetInfo()).toEqual(editor.sheetInfo());
      expect(viewer.displayList({ x: 0, y: 0, width: 500, height: 300 })).toEqual(
        editor.displayList({ x: 0, y: 0, width: 500, height: 300 })
      );
      expect(viewer.cellText(0, 0, 0)).toBe('Quarterly Budget Report');
      expect('save' in viewer).toBe(false);
      expect('editCell' in viewer).toBe(false);
      expect('encodeStateAsUpdate' in viewer).toBe(false);
    } finally {
      viewer.dispose();
      editor.dispose();
    }
  });

  test('analyzes workbook metadata without retaining a document handle', async () => {
    const bytes = new Uint8Array(await readFile(fixture));
    expect(analyzeWorkbook(bytes)).toEqual({
      format: 'xlsx',
      sheetCount: 3,
      sheetNames: ['Budget', 'Summary', 'Styled'],
      contentWidth: 499.05078,
      contentHeight: 1260,
    });
  });

  test('renders and hit-tests charts through the read-only package', async () => {
    const viewer = openViewerWorkbook(new Uint8Array(await readFile(chartFixture)));
    try {
      const viewport = { x: 0, y: 0, width: 800, height: 800 };
      const frame = viewer.displayList(viewport);
      expect(frame.charts?.length).toBe(4);
      const chart = frame.charts![0];
      expect(
        viewer.chartAtPoint(
          viewport,
          chart.clip.x + chart.clip.w / 2,
          chart.clip.y + chart.clip.h / 2
        )?.id
      ).toBe(chart.id);
    } finally {
      viewer.dispose();
    }
  });
});
