import { expect, spyOn, test } from 'bun:test';
import { readFontBytes } from './bytes';

test('package assets read from disk when getBuiltinModule exists', async () => {
  const fetch = spyOn(globalThis, 'fetch');
  try {
    const bytes = await readFontBytes('index.ts', new URL('./index.ts', import.meta.url));
    expect(bytes.byteLength).toBeGreaterThan(0);
    expect(fetch).not.toHaveBeenCalled();
  } finally {
    fetch.mockRestore();
  }
});
