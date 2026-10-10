import { expect, test } from 'bun:test';
import { sortRawExports } from './wasm.ts';

test('raw wasm export lists are sorted so Windows and Linux builds vendor the same declarations', () => {
  const glue = [
    'export class View {',
    '  readonly width: number;',
    '  free(): void;',
    '}',
    'export interface InitOutput {',
    '    readonly memory: WebAssembly.Memory;',
    '    readonly install_panic_hook: () => void;',
    '    readonly close_display_list: (a: number) => void;',
    '}',
    '',
  ].join('\n');
  expect(sortRawExports(glue)).toBe(
    glue.replace(
      '    readonly memory: WebAssembly.Memory;\n    readonly install_panic_hook: () => void;\n    readonly close_display_list: (a: number) => void;',
      '    readonly close_display_list: (a: number) => void;\n    readonly install_panic_hook: () => void;\n    readonly memory: WebAssembly.Memory;'
    )
  );
  const bg = '/* eslint-disable */\nexport const memory: WebAssembly.Memory;\nexport const b: () => void;\nexport const a: () => void;\n';
  expect(sortRawExports(bg)).toBe('/* eslint-disable */\nexport const a: () => void;\nexport const b: () => void;\nexport const memory: WebAssembly.Memory;\n');
});
