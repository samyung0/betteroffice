/** Read-only DOCX engine. Reached only through the viewer entry. */

import type { RustTextEngine } from '../layout/measure/rustMeasureSource';
import wasmInit, {
  DocxViewDocument,
  clear_measure_fonts,
  initSync,
  register_measure_font,
  register_substitute_measure_font,
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

/** The viewer module's measurement font store. */
export const viewTextEngine: RustTextEngine = {
  registerFont: register_measure_font,
  registerSubstituteFont: register_substitute_measure_font,
  clearFonts: clear_measure_fonts,
};
