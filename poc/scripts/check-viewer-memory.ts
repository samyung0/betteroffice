import { readFile } from "node:fs/promises";
import path from "node:path";

import initXlsxEditor, {
  XlsxDocument,
} from "../../packages/xlsx/src/wasm/generated/xlsx_wasm.js";
import initXlsxViewer, {
  XlsxViewDocument,
} from "../../packages/xlsx/src/wasm/generated/viewer/xlsx_view_wasm.js";
import initPptxEditor, {
  PptxDocument,
  PptxRenderer,
} from "../../packages/pptx/src/wasm/generated/pptx_wasm.js";
import initPptxViewer, {
  PptxViewDocument,
  PptxViewRenderer,
} from "../../packages/pptx/src/wasm/generated/viewer/pptx_view_wasm.js";

const root = path.resolve(import.meta.dir, "../..");
const [xlsxBytes, pptxBytes, fontBytes] = await Promise.all([
  readFile(path.join(root, "poc/fixtures/feature-rich.xlsx")),
  readFile(path.join(root, "poc/fixtures/feature-rich.pptx")),
  readFile(path.join(root, "crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf")),
]);

const xlsxEditorModule = await initXlsxEditor({
  module_or_path: await readFile(
    path.join(root, "packages/xlsx/src/wasm/generated/xlsx_wasm_bg.wasm"),
  ),
});
const xlsxViewerModule = await initXlsxViewer({
  module_or_path: await readFile(
    path.join(root, "packages/xlsx/src/wasm/generated/viewer/xlsx_view_wasm_bg.wasm"),
  ),
});
const pptxEditorModule = await initPptxEditor({
  module_or_path: await readFile(
    path.join(root, "packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm"),
  ),
});
const pptxViewerModule = await initPptxViewer({
  module_or_path: await readFile(
    path.join(root, "packages/pptx/src/wasm/generated/viewer/pptx_view_wasm_bg.wasm"),
  ),
});

const viewport = JSON.stringify({ x: 0, y: 0, width: 1400, height: 900 });
const xlsxEditor = XlsxDocument.open(xlsxBytes);
const xlsxEditorOpen = xlsxEditorModule.memory.buffer.byteLength;
xlsxEditor.displayListJson(viewport);
const xlsxEditorRendered = xlsxEditorModule.memory.buffer.byteLength;

const xlsxViewer = XlsxViewDocument.open(xlsxBytes);
const xlsxViewerOpen = xlsxViewerModule.memory.buffer.byteLength;
xlsxViewer.displayListJson(viewport);
const xlsxViewerRendered = xlsxViewerModule.memory.buffer.byteLength;

const pptxEditor = PptxDocument.openCollaborative(pptxBytes, 9021);
const pptxEditorRenderer = new PptxRenderer();
pptxEditorRenderer.registerFont("Liberation Sans", false, false, fontBytes);
const pptxEditorOpen = pptxEditorModule.memory.buffer.byteLength;
pptxEditorRenderer.layoutSlideJson(pptxEditor, 0);
const pptxEditorRendered = pptxEditorModule.memory.buffer.byteLength;

const pptxViewer = PptxViewDocument.open(pptxBytes);
const pptxViewerRenderer = new PptxViewRenderer();
pptxViewerRenderer.registerFont("Liberation Sans", false, false, fontBytes);
const pptxViewerOpen = pptxViewerModule.memory.buffer.byteLength;
pptxViewerRenderer.layoutSlideJson(pptxViewer, 0);
const pptxViewerRendered = pptxViewerModule.memory.buffer.byteLength;

const rows = [
  memoryRow("XLSX", xlsxEditorOpen, xlsxEditorRendered, xlsxViewerOpen, xlsxViewerRendered),
  memoryRow("PPTX", pptxEditorOpen, pptxEditorRendered, pptxViewerOpen, pptxViewerRendered),
];
console.table(rows);

for (const row of rows) {
  if (row.viewerRenderedBytes >= row.editorRenderedBytes) {
    throw new Error(`${row.format} viewer did not reduce wasm linear memory after rendering`);
  }
}

xlsxEditor.free();
xlsxViewer.free();
pptxEditorRenderer.free();
pptxEditor.free();
pptxViewerRenderer.free();
pptxViewer.free();

function memoryRow(
  format: string,
  editorOpenBytes: number,
  editorRenderedBytes: number,
  viewerOpenBytes: number,
  viewerRenderedBytes: number,
) {
  return {
    format,
    editorOpenBytes,
    viewerOpenBytes,
    editorRenderedBytes,
    viewerRenderedBytes,
    viewerRenderedPercent: Number(
      ((viewerRenderedBytes / editorRenderedBytes) * 100).toFixed(1),
    ),
  };
}

