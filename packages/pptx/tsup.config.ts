import { copyFile, mkdir } from 'node:fs/promises';
import { defineConfig } from 'tsup';

export default defineConfig({
  entry: { index: 'src/index.ts', editor: 'src/editor.ts', viewer: 'src/viewer.ts' },
  format: ['esm'],
  dts: true,
  splitting: true,
  sourcemap: false,
  clean: true,
  treeshake: true,
  minify: true,
  onSuccess: async () => {
    await mkdir('dist/generated', { recursive: true });
    await copyFile('src/wasm/generated/pptx_wasm_bg.wasm', 'dist/generated/pptx_wasm_bg.wasm');
    await copyFile(
      'src/wasm/generated/viewer/pptx_view_wasm_bg.wasm',
      'dist/generated/pptx_view_wasm_bg.wasm'
    );
  },
});
