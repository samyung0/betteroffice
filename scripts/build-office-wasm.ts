import { buildWasmModules } from "./wasm.ts";

await buildWasmModules([
  {
    crate: "ooxml-opc",
    name: "ooxml_opc",
    generated: "packages/docx/src/wasm/generated/opc",
    cargoArgs: ["--locked", "--features", "wasm"],
  },
  // DOCX features match scripts/build-docx-wasm.ts: the checkpoint bundle imports the TIFF decoder.
  {
    crate: "docx-edit",
    name: "docx_edit",
    generated: "packages/docx/src/wasm/generated/edit",
    cargoArgs: ["--locked", "--features", "wasm,tiff"],
  },
  {
    crate: "docx-parse",
    name: "docx_parse",
    generated: "packages/docx/src/wasm/generated/parse",
    cargoArgs: ["--locked", "--features", "wasm,tiff"],
  },
  {
    crate: "xlsx-wasm",
    name: "xlsx_wasm",
    generated: "packages/xlsx/src/wasm/generated",
  },
  {
    crate: "pptx-wasm",
    name: "pptx_wasm",
    generated: "packages/pptx/src/wasm/generated",
  },
]);
await import("./build-office-checkpoint.ts");
