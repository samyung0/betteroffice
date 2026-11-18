import type { FontDecoder } from './decode';
import { chunkEntries, parseSkeleton, type SliceSkeleton } from './slices';
import { sha256 } from './sha256';

/**
 * A sliced face that loads its outlines as they are needed. Its layout tables
 * are complete from the start, so shaping, metrics and line breaks equal the
 * whole face's; a glyph whose outline chunk is not loaded yet draws nothing.
 * @public
 */
export interface ProgressiveFontFace {
  /** Increases each time outlines are added. */
  readonly revision: number;
  /** The face with every outline loaded so far, in a new buffer on each call. */
  bytes(): ArrayBuffer;
  /** Loads the outlines of these glyph ids; resolves with a revision that includes them. */
  ensureGlyphs(glyphIds: Iterable<number>): Promise<number>;
  /** Starts loading the outlines of these characters' glyphs; `ensureGlyphs` is authoritative. */
  prefetch(text: string): void;
  /** Loads the outlines of every glyph these characters can shape to; resolves with a revision that includes them. */
  ensureText(text: string): Promise<number>;
  /** Calls `listener` with each new revision; returns an unsubscribe. */
  onUpdate(listener: (revision: number) => void): () => void;
  /** Starts loading the core chunk (the outlines of the face's most used characters) before anything asks for it. */
  prefetchCore(): void;
  /** Whether every outline is loaded. */
  complete(): boolean;
  /**
   * A chunk that failed to load fails again at once, without a fetch, until
   * this is called; then the next request for it asks for it again.
   */
  retryFailed(): void;
  /** `file`, a chunk of this face, failed in another realm: this one doesn't fetch it either until `retryFailed`. */
  markFailed(file: string, error: Error): void;
  /** The chunks that failed and are not fetched again until `retryFailed`, by file, with their errors. */
  failures(): Map<string, unknown>;
  /**
   * Copies of the files loaded so far (the index, the skeleton and each
   * outline chunk, decoded) by manifest path: what another realm's open of this face takes as
   * `SliceLocation.files` instead of reading and decoding them again.
   */
  files(): Map<string, ArrayBuffer>;
}

/**
 * Where a sliced face's files are and how to decode them: what a font loader
 * carries as `slices`, which `openSlicedFace` from `@betteroffice/fonts/progressive` opens.
 * @public
 */
export interface SliceLocation {
  ref: SliceRef;
  /** The asset base the files resolve against; part of the shared face's identity. */
  base: string;
  /** The URL of a pack file by its manifest path. */
  url(file: string): Promise<URL>;
  timeoutMs?: number;
  /** Decompresses a zstd file to at most `maxOutput` bytes. */
  decode(encoded: ArrayBuffer, maxOutput: number): Promise<ArrayBuffer>;
  /** The host's decoder behind `decode`, when it has one; part of the shared face's identity. */
  decoder?: FontDecoder;
  /** @internal Drop the opened face from the shared cache once it settles, as `retain: false` does for whole faces. */
  retain?: boolean;
  /**
   * Decoded files of this face another of our realms loaded (its `files()`),
   * used in place of reading and decoding them; one that differs from its
   * index entry fails the load.
   */
  files?: ReadonlyMap<string, ArrayBuffer>;
}

/**
 * A slice file that could not be loaded: unreachable or timed out
 * (`network`), or answered with an error status or the wrong bytes
 * (`server`, with the HTTP `status` when there was one).
 * @public
 */
export class FontSliceError extends Error {
  override name = 'FontSliceError';
  constructor(
    message: string,
    readonly file: string,
    readonly reason: 'network' | 'server',
    readonly status?: number,
    options?: ErrorOptions,
  ) {
    super(message, options);
  }
}

/** Reads `file`; a failure says whether the network or the server failed. */
async function fetchSlice(transport: SliceTransport, file: string): Promise<ArrayBuffer> {
  try {
    return await transport.read(file);
  } catch (error) {
    const status = (error as { status?: unknown } | null)?.status;
    const message = error instanceof Error ? error.message : String(error);
    throw typeof status === 'number'
      ? new FontSliceError(message, file, 'server', status, { cause: error })
      : new FontSliceError(message, file, 'network', undefined, { cause: error });
  }
}

/** @internal A face row's pointer to its slice index. */
export interface SliceRef {
  format: number;
  index: string;
  /** Pins the index, which pins every file it lists by length and SHA-256. */
  indexSha256: string;
}

interface SliceFile {
  id?: number;
  file: string;
  encodedLength: number;
  byteLength: number;
  sha256: string;
}

interface SliceIndex {
  format: number;
  skeleton: SliceFile;
  core: SliceFile;
  chunks: SliceFile[];
  byteLength: number;
  sha256: string;
}

/** @internal */
export interface SliceTransport {
  /** The bytes of a pack file by its manifest path. */
  read(file: string): Promise<ArrayBuffer>;
  /** Where `read` fetches a pack file from, for another realm to fetch the same files. */
  url?(file: string): Promise<URL>;
  /** The fetch timeout `read` uses, which another realm reading the same files keeps. */
  timeoutMs?: number;
  /** Decompresses a zstd file to at most `maxOutput` bytes. */
  decode(encoded: ArrayBuffer, maxOutput: number): Promise<ArrayBuffer>;
  /**
   * A file's decoded bytes another of our realms handed over, used instead of
   * reading it; one that differs from its index entry fails the load.
   */
  held?(file: string): ArrayBuffer | undefined;
}

/** @internal Where a progressive face's files are, so a worker can open the same face. */
export interface SliceSource {
  ref: SliceRef;
  /** Absolute URL of the index, skeleton, core and every chunk by manifest path. */
  urls: Record<string, string>;
  /** The fetch timeout of the realm that opened the face. */
  timeoutMs?: number;
}

const opened = new WeakMap<ProgressiveFontFace, [SliceRef, SliceIndex, SliceTransport]>();
const sources = new WeakMap<ProgressiveFontFace, Promise<SliceSource>>();

/** @internal The files of a face `openProgressiveFace` opened with a transport that names URLs. */
export function sliceSourceOf(face: ProgressiveFontFace): Promise<SliceSource> | undefined {
  const entry = opened.get(face);
  const url = entry?.[2].url;
  if (!entry || !url) return undefined;
  const [ref, index, { timeoutMs }] = entry;
  let source = sources.get(face);
  if (source === undefined) {
    source = (async () => {
      const urls: Record<string, string> = {};
      for (const { file } of [{ file: ref.index }, index.skeleton, index.core, ...index.chunks]) urls[file] = (await url(file)).href;
      return { ref, urls, ...(timeoutMs === undefined ? {} : { timeoutMs }) };
    })();
    sources.set(face, source);
  }
  return source;
}

async function readFile(transport: SliceTransport, entry: SliceFile): Promise<Uint8Array> {
  const held = transport.held?.(entry.file);
  if (held !== undefined) {
    // Only our own realms hand files over, so a mismatch is our bug: fail, don't fetch around it.
    if (held.byteLength !== entry.byteLength || await sha256(held) !== entry.sha256) {
      throw new Error(`Font slice ${entry.file} handed over by another realm does not match its index entry`);
    }
    return new Uint8Array(held);
  }
  const encoded = await fetchSlice(transport, entry.file);
  const served = (message: string) => new FontSliceError(message, entry.file, 'server');
  if (encoded.byteLength !== entry.encodedLength) {
    throw served(`Font slice ${entry.file} has ${encoded.byteLength} encoded bytes; expected ${entry.encodedLength}`);
  }
  let decoded: ArrayBuffer;
  try {
    decoded = await transport.decode(encoded, entry.byteLength);
  } catch (error) {
    throw new FontSliceError(`Font slice ${entry.file} does not decode: ${error instanceof Error ? error.message : String(error)}`, entry.file, 'server', undefined, { cause: error });
  }
  if (decoded.byteLength !== entry.byteLength) {
    throw served(`Font slice ${entry.file} has ${decoded.byteLength} bytes; expected ${entry.byteLength}`);
  }
  if (await sha256(decoded) !== entry.sha256) throw served(`Font slice ${entry.file} SHA-256 mismatch`);
  return new Uint8Array(decoded);
}

/**
 * @internal Loads a sliced face's index and skeleton; every outline chunk, the
 * core (chunk 0) included, loads when a glyph or character needs it. With
 * `core`, the core's request goes out beside the skeleton's.
 */
export async function openProgressiveFace(
  ref: SliceRef,
  transport: SliceTransport,
  options: { core?: boolean } = {},
): Promise<ProgressiveFontFace> {
  if (ref.format !== 1) throw new Error(`Unsupported font slice format ${ref.format}`);
  const heldIndex = transport.held?.(ref.index);
  const indexBytes = heldIndex ?? await fetchSlice(transport, ref.index);
  if (await sha256(indexBytes) !== ref.indexSha256) {
    throw heldIndex
      ? new Error(`Font slice index ${ref.index} handed over by another realm does not match its pointer`)
      : new FontSliceError(`Font slice index ${ref.index} SHA-256 mismatch`, ref.index, 'server');
  }
  const index = JSON.parse(new TextDecoder().decode(indexBytes)) as SliceIndex;
  if (index.format !== 1) throw new Error(`Unsupported font slice index format in ${ref.index}`);
  const decodedFiles = new Map<string, Uint8Array>([[ref.index, new Uint8Array(indexBytes)]]);
  const loaded = new Set<number>();
  // Chunks that failed: remembered, so nothing fetches them again until retryFailed().
  const failed = new Map<number, unknown>();
  const pending = new Map<number, Promise<void>>();
  // Chunks read before anything asked for them. They install when first asked
  // for, so the face's bytes change only through `ensure`; another realm's
  // open takes them (`files()`) once read.
  const early = new Map<number, Promise<Uint8Array>>();
  const prefetchCore = () => {
    if (early.has(0) || loaded.has(0) || failed.has(0) || pending.has(0)) return;
    const reading = readFile(transport, index.core);
    early.set(0, reading);
    reading.then((bytes) => decodedFiles.set(index.core.file, bytes), (error: unknown) => {
      if (early.get(0) !== reading) return;
      early.delete(0);
      failed.set(0, error);
    });
  };
  if (options.core) prefetchCore();
  const skeletonBytes = await readFile(transport, index.skeleton);
  const skeleton: SliceSkeleton = parseSkeleton(skeletonBytes);
  decodedFiles.set(index.skeleton.file, skeletonBytes);
  const files = new Map<number, SliceFile>([[0, index.core], ...index.chunks.map((chunk): [number, SliceFile] => [chunk.id!, chunk])]);
  const outlines = new Map<number, Uint8Array>();
  const listeners = new Set<(revision: number) => void>();
  let revision = 1;
  let dirty = false;
  const install = (bytes: Uint8Array) => {
    for (const [glyph, outline] of chunkEntries(bytes)) outlines.set(glyph, outline);
  };

  const load = (chunk: number): Promise<void> => {
    if (loaded.has(chunk)) return Promise.resolve();
    if (failed.has(chunk)) return Promise.reject(failed.get(chunk));
    let promise = pending.get(chunk);
    if (promise === undefined) {
      const entry = files.get(chunk);
      if (entry === undefined) return Promise.reject(new Error(`Unknown font slice chunk ${chunk}`));
      const reading = early.get(chunk) ?? readFile(transport, entry);
      early.delete(chunk);
      promise = reading.then((bytes) => {
        install(bytes);
        decodedFiles.set(entry.file, bytes);
        loaded.add(chunk);
        pending.delete(chunk);
        dirty = true;
      }, (error: unknown) => {
        pending.delete(chunk);
        failed.set(chunk, error);
        throw error;
      });
      pending.set(chunk, promise);
    }
    return promise;
  };

  const publish = (): number => {
    if (!dirty) return revision;
    dirty = false;
    revision += 1;
    for (const listener of [...listeners]) {
      try {
        listener(revision);
      } catch {}
    }
    return revision;
  };

  const ensure = async (chunks: Iterable<number>): Promise<number> => {
    const wanted = [...new Set(chunks)].filter((chunk) => !loaded.has(chunk));
    if (wanted.length === 0) return publish();
    const results = await Promise.allSettled(wanted.map(load));
    const failure = results.find((result) => result.status === 'rejected');
    const next = publish();
    if (failure) throw (failure as PromiseRejectedResult).reason;
    return next;
  };

  // A chunk holds the GSUB closure of its codepoints, so these chunks hold
  // every glyph a shaper can draw for `text`.
  const ensureText = (text: string): Promise<number> => {
    const chunks: number[] = [];
    for (const character of text) {
      const chunk = skeleton.chunkOfCodepoint(character.codePointAt(0)!);
      if (chunk !== undefined) chunks.push(chunk);
    }
    return ensure(chunks);
  };

  const face: ProgressiveFontFace = {
    get revision() {
      return revision;
    },
    bytes() {
      return skeleton.reassemble(outlines);
    },
    ensureGlyphs(glyphIds) {
      const chunks: number[] = [];
      for (const glyph of glyphIds) {
        const chunk = skeleton.chunkOfGlyph(glyph);
        if (chunk !== undefined) chunks.push(chunk);
      }
      return ensure(chunks);
    },
    prefetch(text) {
      void ensureText(text).catch(() => {});
    },
    ensureText,
    onUpdate(listener) {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
    prefetchCore,
    complete() {
      return loaded.size === files.size;
    },
    retryFailed() {
      failed.clear();
    },
    markFailed(file, error) {
      for (const [chunk, entry] of files) {
        if (entry.file === file && !loaded.has(chunk)) failed.set(chunk, error);
      }
    },
    failures() {
      return new Map([...failed].map(([chunk, error]) => [files.get(chunk)!.file, error]));
    },
    files() {
      return new Map([...decodedFiles].map(([file, bytes]) => [file, bytes.slice().buffer]));
    },
  };
  opened.set(face, [ref, index, transport]);
  return face;
}
