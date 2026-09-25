import type { Document } from '../types/document';

/** Resolves `media:<part>` image sources to the package's data URLs, in place, for comparing the byte seeder with the document seeder. */
export function resolveImageRefs<T>(value: T, media: Document['package']['media']): T {
  if (Array.isArray(value)) value.forEach((item) => resolveImageRefs(item, media));
  else if (value instanceof Map) value.forEach((item) => resolveImageRefs(item, media));
  else if (value && typeof value === 'object') {
    const record = value as Record<string, unknown>;
    if (typeof record.src === 'string' && record.src.startsWith('media:'))
      record.src = media?.get(record.src.slice('media:'.length))?.dataUrl;
    Object.values(record).forEach((item) => resolveImageRefs(item, media));
  }
  return value;
}
