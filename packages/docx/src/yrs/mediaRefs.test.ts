import { beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { createEditSession, preloadEditWasm } from '../wasm/edit';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const SOURCE = resolve(import.meta.dir, '../../../../poc/fixtures/exchange-plan.docx');

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

function imageSources(session: ReturnType<typeof createEditSession>): string[] {
  const sources: string[] = [];
  const visit = (value: unknown): void => {
    if (Array.isArray(value)) value.forEach(visit);
    else if (value && typeof value === 'object') {
      const record = value as Record<string, unknown>;
      if (record.kind === 'image' && typeof record.src === 'string') sources.push(record.src);
      Object.values(record).forEach(visit);
    }
  };
  visit(JSON.parse(session.yrs_blocks_for_story('body', '{}')));
  return sources;
}

it('a replica without the source resolves image references once it gets the media', () => {
  const main = createEditSession(1);
  const replica = createEditSession(2);
  try {
    main.open_docx(new Uint8Array(readFileSync(SOURCE)), true);
    replica.load(main.encode_state());
    const resolved = imageSources(main);
    expect(resolved.length).toBe(3);
    expect(resolved.every((src) => src.startsWith('data:image/png;base64,'))).toBe(true);
    expect(imageSources(replica).every((src) => src.startsWith('media:word/media/'))).toBe(true);
    replica.set_media_json(main.media_json());
    expect(imageSources(replica)).toEqual(resolved);
  } finally {
    main.free();
    replica.free();
  }
});
