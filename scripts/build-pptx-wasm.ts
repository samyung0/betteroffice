import { buildWasmModules } from './wasm.ts';

await buildWasmModules([
  {
    crate: 'pptx-view-wasm',
    name: 'pptx_view_wasm',
    generated: 'packages/pptx/src/wasm/generated/viewer',
  },
  {
    crate: 'pptx-wasm',
    name: 'pptx_wasm',
    generated: 'packages/pptx/src/wasm/generated',
  },
]);
