import { expect, mock, test } from 'bun:test';
import { createRustMeasureSource, type RustTextEngine } from './rustMeasureSource';

// Counts the loads of the module that opens sliced faces.
let loads = 0;
mock.module('@betteroffice/fonts/progressive', () => {
  loads += 1;
  return {};
});

test('the editor loads the module that opens sliced faces as it opens, before any face needs it; View does not (Q771 A)', async () => {
  const view: RustTextEngine = { registerFont: () => 1, clearFonts() {} };
  createRustMeasureSource({ engine: view, progressive: true });
  await Bun.sleep(20);
  expect(loads).toBe(0);
  // The editor's engine paints sliced faces in its worker.
  createRustMeasureSource({ engine: { ...view, registerSlicedFont() {} }, progressive: true });
  for (let wait = 0; wait < 50 && loads === 0; wait += 1) await Bun.sleep(10);
  expect(loads).toBe(1);
});
