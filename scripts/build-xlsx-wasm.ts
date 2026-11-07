import { buildWasmModules } from './wasm.ts';

await buildWasmModules([
  {
    crate: 'xlsx-view-wasm',
    name: 'xlsx_view_wasm',
    generated: 'packages/xlsx/src/wasm/generated/viewer',
  },
  {
    crate: 'xlsx-wasm',
    name: 'xlsx_wasm',
    generated: 'packages/xlsx/src/wasm/generated',
  },
]);
