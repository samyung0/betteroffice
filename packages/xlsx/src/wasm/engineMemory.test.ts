/**
 * Upstream #1730: an engine that runs out of memory fails as a workbook too
 * large, not a bare trap. Its own file: the engine it exhausts stays broken.
 */

import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { engineOutOfMemory, initWasm, openWorkbook } from './loader';

const FIXTURE = resolve(import.meta.dir, '../../test-fixtures/sample.xlsx');
const WASM = resolve(import.meta.dir, './generated/xlsx_wasm_bg.wasm');

test('a trap after a failed allocation is named as the workbook being too large', async () => {
  // keep the instance's memory, to grow it to the wasm32 ceiling from outside.
  const memories: WebAssembly.Memory[] = [];
  const instantiate = WebAssembly.instantiate;
  WebAssembly.instantiate = (async (...args: Parameters<typeof WebAssembly.instantiate>) => {
    const result = await (instantiate as (...input: unknown[]) => Promise<unknown>).apply(
      WebAssembly,
      args
    );
    const instance =
      result instanceof WebAssembly.Instance
        ? result
        : (result as { instance: WebAssembly.Instance }).instance;
    for (const value of Object.values(instance.exports))
      if (value instanceof WebAssembly.Memory) memories.push(value);
    return result;
  }) as typeof WebAssembly.instantiate;
  try {
    await initWasm(new Uint8Array(readFileSync(WASM)));
  } finally {
    WebAssembly.instantiate = instantiate;
  }
  const workbook = openWorkbook(new Uint8Array(readFileSync(FIXTURE)));
  expect(engineOutOfMemory(new Error('plain'))).toBeNull();
  for (const memory of memories)
    for (let pages = 1 << 15; pages >= 1; pages >>= 1)
      for (;;) {
        try {
          memory.grow(pages);
        } catch {
          break;
        }
      }
  // larger than the free heap the allocator keeps (one step, at most 64 MiB).
  let caught: unknown;
  try {
    workbook.editCells(0, [{ row: 0, col: 0, input: 'x'.repeat(96 << 20) }]);
  } catch (error) {
    caught = error;
  }
  expect(caught).toBeInstanceOf(WebAssembly.RuntimeError);
  expect(engineOutOfMemory(caught)?.message).toContain('The workbook is too large');
});
