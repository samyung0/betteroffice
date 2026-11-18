import { readFontAsset } from './asset';
import type { FontDecoder } from './decode';
import { openProgressiveFace, type ProgressiveFontFace, type SliceLocation } from './progressive';

export { readFontAsset } from './asset';
export {
  FontSliceError,
  openProgressiveFace,
  sliceSourceOf,
  type ProgressiveFontFace,
  type SliceLocation,
  type SliceRef,
  type SliceSource,
  type SliceTransport,
} from './progressive';

const faces = new Map<string, Promise<ProgressiveFontFace>>();
const engineFaces = new WeakMap<FontDecoder, Map<string, Promise<ProgressiveFontFace>>>();

/**
 * Opens the sliced face a loader's `slices` locates. One face is shared per
 * sliced face, asset base and decoder; a failed open is dropped so a later
 * call retries. With `core`, its core chunk starts loading at once.
 * @public
 */
export function openSlicedFace(location: SliceLocation, options: { core?: boolean } = {}): Promise<ProgressiveFontFace> {
  const { ref, base, url, timeoutMs, decode, decoder, retain, files } = location;
  let cache = faces;
  if (decoder !== undefined) {
    cache = engineFaces.get(decoder) ?? new Map();
    engineFaces.set(decoder, cache);
  }
  const key = `${ref.indexSha256}|${base}|${timeoutMs ?? ''}`;
  let pending = cache.get(key);
  if (pending === undefined) {
    pending = openProgressiveFace(ref, {
      read: async (file) => readFontAsset(file, await url(file), timeoutMs),
      url,
      timeoutMs,
      decode,
      ...(files && { held: (file: string) => files.get(file) }),
    }, options);
    cache.set(key, pending);
    const current = pending;
    const forget = () => {
      if (cache.get(key) === current) cache.delete(key);
    };
    pending.then(retain === false ? forget : undefined, forget);
  } else if (options.core) {
    void pending.then((face) => face.prefetchCore(), () => {});
  }
  return pending;
}
