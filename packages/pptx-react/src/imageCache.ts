import { decodePresentationImage, type CanvasImageResolver } from '@betteroffice/pptx';

type HeldImages = CanvasImageResolver & { release(): void };

interface CachedImage {
  promise: Promise<CanvasImageSource | null>;
  source: CanvasImageSource | null;
  references: number;
  retained: boolean;
  closed: boolean;
  bytes: number;
}

const MAX_IMAGES = 25;

/**
 * Decoded slide pictures, bounded by count and by decoded bytes (128 MiB by
 * default), least recently used first out. A paint holds what it resolves
 * until it finishes, so an evicted bitmap is closed only once no paint holds it.
 */
export function imageCache(options: { maxDecodedBytes?: number } = {}): {
  resolver(read: (id: string) => Uint8Array | undefined, errorMessage: string): CanvasImageResolver;
  clear(): void;
} {
  const budget = options.maxDecodedBytes ?? 128 * 1024 * 1024;
  if (!Number.isSafeInteger(budget) || budget < 0) throw new Error('invalid decoded image budget');
  const cache = new Map<string, CachedImage>();
  let retainedBytes = 0;
  const close = (image: CachedImage) => {
    if (image.retained || image.references || image.closed || !image.source) return;
    image.closed = true;
    if ('close' in image.source && typeof image.source.close === 'function') image.source.close();
    image.source = null;
  };
  const trim = () => {
    while (cache.size > MAX_IMAGES || retainedBytes > budget) {
      const oldest = cache.entries().next().value;
      if (!oldest) break;
      cache.delete(oldest[0]);
      retainedBytes -= oldest[1].bytes;
      oldest[1].retained = false;
      close(oldest[1]);
    }
  };
  const acquire = (read: (id: string) => Uint8Array | undefined, errorMessage: string): HeldImages => {
    const held = new Map<string, CachedImage>();
    let released = false;
    const resolve: CanvasImageResolver = (id) => {
      if (released) return Promise.resolve(null);
      let image = held.get(id) ?? cache.get(id);
      if (!image) {
        const bytes = read(id);
        if (!bytes) return Promise.resolve(null);
        const entry: CachedImage = {
          promise: Promise.resolve(null), source: null, references: 0, retained: true, closed: false, bytes: 0,
        };
        entry.promise = decodePresentationImage(bytes, errorMessage).then((source) => {
          entry.source = source;
          const size = source as { naturalWidth?: number; naturalHeight?: number; width?: number; height?: number };
          const decoded = (size.naturalWidth ?? size.width ?? 0) * (size.naturalHeight ?? size.height ?? 0) * 4;
          entry.bytes = Number.isSafeInteger(decoded) && decoded >= 0 ? decoded : budget + 1;
          if (entry.retained) retainedBytes += entry.bytes;
          trim();
          close(entry);
          return source;
        });
        image = entry;
        cache.set(id, image);
      } else if (cache.get(id) === image) {
        cache.delete(id);
        cache.set(id, image);
      }
      if (!held.has(id)) {
        held.set(id, image);
        image.references += 1;
      }
      trim();
      return image.promise;
    };
    return Object.assign(resolve, {
      release() {
        if (released) return;
        released = true;
        for (const image of held.values()) {
          image.references -= 1;
          close(image);
        }
        held.clear();
      },
    });
  };
  return {
    resolver: (read, errorMessage) => Object.assign(async (id: string) => {
      const resolve = acquire(read, errorMessage);
      try { return await resolve(id); } finally { resolve.release(); }
    }, { acquire: () => acquire(read, errorMessage) }),
    clear() {
      for (const image of cache.values()) {
        image.retained = false;
        close(image);
      }
      cache.clear();
      retainedBytes = 0;
    },
  };
}
