import { expect, mock, test } from 'bun:test';
import { paintSlide } from '@betteroffice/pptx';
import { imageCache } from './imageCache';

/** A PNG header declaring 4096 by 4096 pixels: 64 MiB decoded. */
function png(): Uint8Array {
  const bytes = new Uint8Array(24);
  bytes.set([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82]);
  const view = new DataView(bytes.buffer);
  view.setUint32(16, 4096);
  view.setUint32(20, 4096);
  return bytes;
}

function bitmaps(count: number) {
  const made = Array.from({ length: count }, () => ({ width: 4096, height: 4096, close: mock(() => {}) }));
  let decoded = 0;
  const original = globalThis.createImageBitmap;
  globalThis.createImageBitmap = (() => Promise.resolve(made[decoded++] as unknown as ImageBitmap)) as typeof createImageBitmap;
  return { made, restore: () => { globalThis.createImageBitmap = original; } };
}

test('does not read a cached picture again on a second paint', async () => {
  const { restore } = bitmaps(1);
  const images = imageCache();
  const read = mock(() => png());
  const painted = mock(() => {});
  const ctx = new Proxy({} as CanvasRenderingContext2D, {
    get: (_, key) => key === 'drawImage' ? painted : () => {},
    set: () => true,
  });
  const display = {
    contractVersion: 1, width: 960, height: 540,
    primitives: [{ kind: 'image' as const, objectId: 1, name: 'Picture', assetId: 'a', x: 0, y: 0, w: 960, h: 540 }],
  };
  try {
    const resolveImage = images.resolver(read, 'undecodable');
    await paintSlide(ctx, display, 1, 1, { resolveImage });
    await paintSlide(ctx, display, 1, 1, { resolveImage });
    expect(read).toHaveBeenCalledTimes(1);
    expect(painted).toHaveBeenCalledTimes(2);
  } finally {
    images.clear();
    restore();
  }
});

test('evicts decoded bytes in least recently used order', async () => {
  const { made, restore } = bitmaps(4);
  const images = imageCache();
  const resolver = images.resolver(() => png(), 'undecodable');
  const paint = async (id: string) => {
    const resolve = resolver.acquire!();
    try { return await resolve(id); } finally { resolve.release!(); }
  };
  try {
    expect(await paint('a')).toBe(made[0]);
    expect(await paint('b')).toBe(made[1]);
    expect(await paint('a')).toBe(made[0]);
    expect(await paint('c')).toBe(made[2]);
    expect(made[1].close).toHaveBeenCalledTimes(1);
    expect(made[0].close).not.toHaveBeenCalled();
    expect(await paint('b')).toBe(made[3]);
    expect(made[0].close).toHaveBeenCalledTimes(1);
  } finally {
    images.clear();
    restore();
  }
});

test('holds an evicted bitmap until every paint holding it releases it', async () => {
  const { made, restore } = bitmaps(3);
  const images = imageCache();
  const resolver = images.resolver(() => png(), 'undecodable');
  const first = resolver.acquire!();
  const second = resolver.acquire!();
  try {
    await first('a');
    await second('a');
    for (const id of ['b', 'c']) {
      const paint = resolver.acquire!();
      await paint(id);
      paint.release!();
    }
    expect(made[0].close).not.toHaveBeenCalled();
    first.release!();
    expect(made[0].close).not.toHaveBeenCalled();
    second.release!();
    expect(made[0].close).toHaveBeenCalledTimes(1);
  } finally {
    images.clear();
    restore();
  }
});
