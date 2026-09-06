import { createHash } from "node:crypto";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
const root = resolve(import.meta.dir, "..");
const result = await Bun.build({
  entrypoints: [resolve(root, "shared/office-checkpoint.ts")],
  outdir: resolve(root, "shared"),
  naming: "office-checkpoint.mjs",
  target: "node",
  minify: false,
});
if (!result.success)
  throw new AggregateError(result.logs, "Office checkpoint bundle failed");
const sourcePaths = {
  docx: "packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm",
  parse: "packages/docx/src/wasm/generated/parse/docx_parse_bg.wasm",
  opc: "packages/docx/src/wasm/generated/opc/ooxml_opc_bg.wasm",
  xlsx: "packages/xlsx/src/wasm/generated/xlsx_wasm_bg.wasm",
  pptx: "packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm",
};
const output = resolve(root, "shared/office-runtime");
await mkdir(output, { recursive: true });
const hashes = Object.fromEntries(
  await Promise.all(
    Object.entries(sourcePaths).map(async ([name, path]) => {
      const bytes = await readFile(resolve(root, path));
      await writeFile(resolve(output, `${name}.wasm`), bytes);
      return [name, createHash("sha256").update(bytes).digest("hex")];
    })
  )
);
await writeFile(
  resolve(output, "manifest.json"),
  `${JSON.stringify(hashes, null, 2)}\n`
);
console.log("Built shared/office-checkpoint.mjs and shared/office-runtime/");
