import { expect, mock, test } from 'bun:test';

// A wasm trap mid-layout leaves wasm-bindgen's borrow flag set, so the
// document's free() then throws a borrow error of its own.
const trap = new WebAssembly.RuntimeError('unreachable');
mock.module('./wasm/viewer', () => ({
  preloadViewWasm: async () => {},
  viewTextEngine: { registerFont: () => 0, clearFonts: () => {} },
  openViewDocument: () => ({
    layoutRequestJson: () => '{"options":{"pageGap":24},"regions":{}}',
    fontRequirementsJson: () => '[]',
    layout: () => {
      throw trap;
    },
    free: () => {
      throw new Error('attempted to take ownership of Rust value while it was borrowed');
    },
  }),
}));

const { openDocumentViewer } = await import('./viewer');

test('a trap during layout is the error reported, not the failed free()', async () => {
  await expect(openDocumentViewer(new Uint8Array())).rejects.toBe(trap);
});
