import { brotliCompressSync, constants } from "node:zlib";
import { readFile, stat } from "node:fs/promises";
import path from "node:path";

const root = path.resolve(import.meta.dir, "../..");

interface Split {
  format: string;
  editor: string;
  viewer: string;
  maxRatio: number;
  forbiddenViewerDeclarations: RegExp[];
}

const splits: Split[] = [
  {
    format: "DOCX",
    editor: "packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm",
    viewer: "packages/docx/src/wasm/generated/viewer/docx_view_wasm_bg.wasm",
    // The view-only Cargo profile keeps the transient OOXML/Yrs lowering
    // bridge size-optimized while the editor remains tuned for interaction.
    maxRatio: 0.8,
    forbiddenViewerDeclarations: [
      /save|serialize|write_docx/,
      /applyInput|applyDelete|replace|formatRange/,
      /encodeState|applyUpdate|startUpdateObservation/,
    ],
  },
  {
    format: "XLSX",
    editor: "packages/xlsx/src/wasm/generated/xlsx_wasm_bg.wasm",
    viewer: "packages/xlsx/src/wasm/generated/viewer/xlsx_view_wasm_bg.wasm",
    maxRatio: 0.35,
    forbiddenViewerDeclarations: [
      /saveBytes|save\(/,
      /editCell|applyOps|patchRangeStyle/,
      /encodeState|applyUpdate|startUpdateObservation/,
    ],
  },
  {
    format: "PPTX",
    editor: "packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm",
    viewer: "packages/pptx/src/wasm/generated/viewer/pptx_view_wasm_bg.wasm",
    maxRatio: 0.7,
    forbiddenViewerDeclarations: [
      /saveBytes|save\(/,
      /insertText|addShape|moveShape/,
      /encodeState|applyUpdate|startUpdateObservation/,
    ],
  },
];

const rows = [];
for (const split of splits) {
  const editorPath = path.join(root, split.editor);
  const viewerPath = path.join(root, split.viewer);
  const [editor, viewer, declarations] = await Promise.all([
    readFile(editorPath),
    readFile(viewerPath),
    readFile(viewerPath.replace(/_bg\.wasm$/, ".d.ts"), "utf8"),
  ]);
  const ratio = viewer.byteLength / editor.byteLength;
  if (ratio > split.maxRatio) {
    throw new Error(
      `${split.format} viewer is ${(ratio * 100).toFixed(
        1
      )}% of editor; limit is ${split.maxRatio * 100}%`
    );
  }
  for (const forbidden of split.forbiddenViewerDeclarations) {
    if (forbidden.test(declarations)) {
      throw new Error(
        `${split.format} viewer exports editor API matching ${forbidden}`
      );
    }
  }
  rows.push({
    format: split.format,
    editorBytes: (await stat(editorPath)).size,
    viewerBytes: (await stat(viewerPath)).size,
    viewerPercent: Number((ratio * 100).toFixed(1)),
    editorBrotliBytes: brotli(editor),
    viewerBrotliBytes: brotli(viewer),
  });
}

console.table(rows);

function brotli(bytes: Uint8Array): number {
  return brotliCompressSync(bytes, {
    params: { [constants.BROTLI_PARAM_QUALITY]: 11 },
  }).byteLength;
}
