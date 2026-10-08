import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { rezipContainer } from "../packages/docx/src/wasm/opc";
import { createYrsSession } from "../packages/docx/src/yrs";

// Every editor and viewer module grows linear memory in large steps
// (crates/wasm-alloc), counted by the allocator's own export.
const MODULES = [
  "docx/src/wasm/generated/edit/docx_edit",
  "docx/src/wasm/generated/viewer/docx_view_wasm",
  "xlsx/src/wasm/generated/xlsx_wasm",
  "xlsx/src/wasm/generated/viewer/xlsx_view_wasm",
  "pptx/src/wasm/generated/pptx_wasm",
  "pptx/src/wasm/generated/viewer/pptx_view_wasm",
];

interface RawExports {
  __wbindgen_free(ptr: number, size: number, align: number): void;
  __wbindgen_malloc(size: number, align: number): number;
  betteroffice_memory_grows(): number;
  memory: WebAssembly.Memory;
}

/** A module's raw exports (the loader's instance when it made one). */
async function rawExports(path: string): Promise<RawExports> {
  const glue = await import(`../packages/${path}.js`);
  return glue.initSync({
    module: readFileSync(new URL(`../packages/${path}_bg.wasm`, import.meta.url)),
  });
}

function longDocx(paragraphs: number): Uint8Array {
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const text = "The quick brown fox jumps over the lazy dog. ".repeat(4);
  const body = Array.from(
    { length: paragraphs },
    (_, i) => `<w:p><w:r><w:t xml:space="preserve">${i} ${text}</w:t></w:r></w:p>`,
  ).join("");
  const parts: Record<string, string> = {
    "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
    "word/document.xml": `<w:document ${W}><w:body>${body}</w:body></w:document>`,
  };
  return rezipContainer(
    Object.fromEntries(
      Object.entries(parts).map(([path, xml]) => [path, new TextEncoder().encode(xml)]),
    ),
  );
}

// First, while the module's memory is still small.
test("opening a 6,000-paragraph DOCX grows the editing engine's memory a few times", async () => {
  const session = await createYrsSession();
  const wasm = await rawExports(MODULES[0]!);
  const bytes = longDocx(6000);
  const grows = wasm.betteroffice_memory_grows();
  const before = wasm.memory.buffer.byteLength;
  session.openDocx(bytes, true);
  // Large enough that dlmalloc alone grew memory hundreds of times.
  expect(wasm.memory.buffer.byteLength - before).toBeGreaterThan(32 << 20);
  expect(wasm.betteroffice_memory_grows() - grows).toBeLessThanOrEqual(16);
  session.destroy();
});

// The editing engine's memory is already large from the open above.
for (const path of MODULES.slice(1))
  test(`${path.split("/").at(-1)}: 64 MiB in 16 KiB allocations grows memory a few times`, async () => {
    const wasm = await rawExports(path);
    const grows = wasm.betteroffice_memory_grows();
    const before = wasm.memory.buffer.byteLength;
    const size = 16 << 10;
    const blocks = Array.from({ length: 4096 }, () => wasm.__wbindgen_malloc(size, 8));
    expect(wasm.memory.buffer.byteLength - before).toBeGreaterThanOrEqual(48 << 20);
    // dlmalloc alone grew it once per 64 KiB: 1,024 times.
    expect(wasm.betteroffice_memory_grows() - grows).toBeLessThanOrEqual(16);
    for (const block of blocks) wasm.__wbindgen_free(block, size, 8);
  });
