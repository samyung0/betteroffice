import { expect, test } from 'bun:test';
import { resolve } from 'node:path';

const ROOT = resolve(import.meta.dir, '..');
const WASM = 'wasm32-unknown-unknown';

// Capy's wasm artifacts that read JSON, with the features their build scripts pass.
const ARTIFACTS: Array<[name: string, dir: string, args: string[]]> = [
  ['docx layout', 'crates/docx-layout', []],
  ['docx view', 'crates/docx-view-wasm', []],
  ['docx edit', 'crates/docx-edit', ['--features', 'wasm,tiff,metafile']],
  ['docx parse', 'crates/docx-parse', ['--features', 'wasm,tiff,metafile']],
  ['pptx view', 'crates/pptx-view-wasm', []],
  ['pptx', 'crates/pptx-wasm', []],
  ['xlsx view', 'crates/xlsx-view-wasm', []],
  ['xlsx', 'crates/xlsx-wasm', []],
  ['vsdx', 'crates/vsdx-wasm', []],
];

for (const [name, dir, args] of ARTIFACTS) {
  test(`${name} reads JSON numbers back exactly`, () => {
    const run = Bun.spawnSync([
      'cargo', 'tree', '--manifest-path', `${dir}/Cargo.toml`, '--target', WASM,
      '-e', 'normal', '--prefix', 'none', '-f', '{p} [{f}]', ...args,
    ], { cwd: ROOT });
    expect(run.exitCode, run.stderr.toString()).toBe(0);
    const lines = run.stdout.toString().split('\n').filter((line) => line.startsWith('serde_json v'));
    expect(lines.length).toBeGreaterThan(0);
    for (const line of lines) expect(line).toContain('float_roundtrip');
  }, 120_000);
}
