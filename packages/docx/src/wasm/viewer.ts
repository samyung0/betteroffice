/** Read-only DOCX engine. Reached only through the viewer entry. */

import wasmInit, {
  DocxViewDocument,
  initSync,
} from './generated/viewer/docx_view_wasm.js';
import { createWasmModuleState, type WasmAsyncInput } from './loadWasmAsset';

const state = createWasmModuleState({
  label: 'docx-view',
  preloadName: 'preloadViewWasm',
  assetUrl: () =>
    new URL('./generated/viewer/docx_view_wasm_bg.wasm', import.meta.url),
  initAsync: wasmInit,
  initSync,
});

export function preloadViewWasm(input?: WasmAsyncInput): Promise<void> {
  return state.preload(input);
}

export function openViewDocument(bytes: Uint8Array): DocxViewDocument {
  state.ensure();
  return DocxViewDocument.open(bytes);
}

export type { DocxViewDocument };
