/** A sliced face's pack files, as `sliceSourceOf` from `@betteroffice/fonts/progressive` describes them. */
export interface ProgressiveFontSource {
  ref: { format: number; index: string; indexSha256: string };
  /** Absolute URL of the index, skeleton, core and every chunk by manifest path. */
  urls: Record<string, string>;
  timeoutMs?: number;
}

/** A font id registered from a sliced face's skeleton, and the face's files. */
export interface ProgressiveFontSpec {
  slot: number;
  source: ProgressiveFontSource;
  /** Decoded files the sender already loaded, by manifest path, which the receiver's open reads instead. */
  files?: ReadonlyMap<string, ArrayBuffer>;
}

/** What the engines use of a face `@betteroffice/fonts/progressive` opens. */
interface OutlineFace {
  readonly revision: number;
  bytes(): ArrayBuffer;
  ensureGlyphs(glyphIds: Iterable<number>): Promise<number>;
}

/**
 * How long a page waits for outlines still loading before it paints those
 * glyphs blank in their space and repaints once they land (Q750 B). The page
 * raster and the engine worker both use it, so a slow slice holds no request
 * near the worker's request budgets and no other page longer than this.
 */
export const PROGRESSIVE_OUTLINE_WAIT_MS = 300;

interface Owner {
  open: () => Promise<OutlineFace>;
  onError: (error: Error) => void;
  opening?: Promise<OutlineFace>;
  face?: OutlineFace;
  /** The face revision whose outlines the engine holds. */
  installed: number;
  failed: boolean;
}

interface OutlineEngine {
  updateFontOutlines(fontId: number, bytes: Uint8Array): void;
}

/** The glyph ids and characters a font draws. */
export interface DrawnGlyphs {
  glyphs: Set<number>;
  characters: Set<string>;
}

/**
 * Glyph ids and characters per font id drawn by a display page: its own runs
 * and those of its regions (headers, footers, text boxes) and note areas.
 */
export function drawnByFont(page: unknown, into = new Map<number, DrawnGlyphs>()): Map<number, DrawnGlyphs> {
  const visit = (node: unknown, depth: number): void => {
    if (depth > 12 || node === null || typeof node !== 'object') return;
    if (Array.isArray(node)) {
      for (const child of node) visit(child, depth + 1);
      return;
    }
    const value = node as { kind?: unknown; fontId?: unknown; glyphs?: unknown; text?: unknown; paintText?: unknown };
    if (value.kind === 'glyphRun' && typeof value.fontId === 'number' && Array.isArray(value.glyphs)) {
      let drawn = into.get(value.fontId);
      if (!drawn) into.set(value.fontId, drawn = { glyphs: new Set(), characters: new Set() });
      for (const glyph of value.glyphs as Array<{ id?: unknown }>) {
        if (typeof glyph.id === 'number') drawn.glyphs.add(glyph.id);
      }
      for (const text of [value.text, value.paintText]) {
        if (typeof text === 'string') for (const character of text) drawn.characters.add(character);
      }
      return;
    }
    for (const child of Object.values(node)) visit(child, depth + 1);
  };
  visit(page, 0);
  return into;
}

/** Glyph ids per font id drawn by a display page: its own runs and those of its regions and note areas. */
export function glyphsByFont(page: unknown): Map<number, Set<number>> {
  return new Map([...drawnByFont(page)].map(([fontId, { glyphs }]) => [fontId, glyphs]));
}

/** What a realm uses of `@betteroffice/fonts/progressive` to open a face from its files. */
interface ProgressiveModule {
  openProgressiveFace(ref: ProgressiveFontSource['ref'], transport: {
    read(file: string): Promise<ArrayBuffer>;
    decode(encoded: ArrayBuffer, maxOutput: number): Promise<ArrayBuffer>;
    held?(file: string): ArrayBuffer | undefined;
  }): Promise<OutlineFace>;
  readFontAsset(file: string, url: URL, timeoutMs?: number): Promise<ArrayBuffer>;
}

/**
 * Opens the face `source` names from its files, for a realm that did not open
 * it (an engine worker); `files` the sender decoded are checked and used
 * instead of fetched and decoded again.
 */
export function sliceFaceOpener(
  source: ProgressiveFontSource,
  decode: (encoded: ArrayBuffer, maxOutput: number) => Promise<ArrayBuffer>,
  files?: ReadonlyMap<string, ArrayBuffer>,
  loadModule: () => Promise<ProgressiveModule> = () => import('@betteroffice/fonts/progressive'),
): () => Promise<OutlineFace> {
  return () => loadModule().then(({ openProgressiveFace, readFontAsset }) => openProgressiveFace(source.ref, {
    read: (file) => {
      const url = source.urls[file];
      if (url === undefined) return Promise.reject(new Error(`Font slice ${file} has no URL`));
      return readFontAsset(file, new URL(url), source.timeoutMs);
    },
    decode,
    ...(files && { held: (file: string) => files.get(file) }),
  }));
}

/**
 * Font ids registered from sliced CJK faces' skeletons. Before a page is
 * rasterized, the outlines of the glyphs it draws are loaded (each chunk
 * once) and installed in the engine; layout never changes. A face whose
 * chunks fail to load is reported once and installs nothing more: its
 * unloaded glyphs stay blank.
 */
export class ResidentProgressiveFonts {
  private readonly owners = new Map<number, Owner>();
  /** Every font id drawing with an owner's bytes: the owners and their measurement views. */
  private readonly ownerOf = new Map<number, number>();
  private readonly pending = new Set<Promise<unknown>>();
  private readonly lateListeners = new Set<(fontIds: Set<number>) => void>();
  private warned = false;
  private late = 0;
  private disposed = false;

  constructor(
    private readonly engine: OutlineEngine,
    private readonly log: (message: string) => void = (message) => console.warn(message),
  ) {}

  /**
   * Font `id` draws with the face `open` gives, whose outlines up to
   * `installed` the engine holds; `onError` hears once that its outlines
   * failed to load or install.
   */
  add(id: number, open: () => Promise<OutlineFace>, onError: (error: Error) => void, installed = 1): void {
    if (this.owners.has(id)) return;
    this.owners.set(id, { open, onError, installed, failed: false });
    this.ownerOf.set(id, id);
  }

  /** `view`, a measurement view of `base`, draws with `base`'s outlines. */
  addView(view: number, base: number): void {
    const owner = this.ownerOf.get(base);
    if (owner !== undefined) this.ownerOf.set(view, owner);
  }

  get size(): number {
    return this.owners.size;
  }

  /** Outline installs that arrived after the frame that needed them had painted. */
  get lateInstalls(): number {
    return this.late;
  }

  /** Stops installing outlines, for fonts that were cleared or replaced; pending loads finish unused. */
  dispose(): void {
    this.disposed = true;
    this.lateListeners.clear();
  }

  /** Calls `listener` with the font ids whose outlines a frame stopped waiting for, once they install. */
  onLateInstall(listener: (fontIds: Set<number>) => void): () => void {
    this.lateListeners.add(listener);
    return () => {
      this.lateListeners.delete(listener);
    };
  }

  /**
   * Installs the outlines of `glyphs` (per font id) before a raster. Resolves
   * with the font ids, views included, whose outlines changed once they are
   * installed, or after `waitMs` with those installed so far; the rest
   * install later and reach `onLateInstall`.
   */
  async ensure(glyphs: Map<number, Set<number>>, waitMs = PROGRESSIVE_OUTLINE_WAIT_MS): Promise<Set<number>> {
    const wanted = new Map<number, Set<number>>();
    for (const [fontId, ids] of glyphs) {
      const owner = this.ownerOf.get(fontId);
      if (owner === undefined || ids.size === 0) continue;
      const set = wanted.get(owner) ?? new Set<number>();
      for (const id of ids) set.add(id);
      wanted.set(owner, set);
    }
    if (wanted.size === 0 || this.disposed) return new Set();
    const changed = new Set<number>();
    let waiting = true;
    const work = Promise.all([...wanted].map(async ([owner, ids]) => {
      if (!(await this.load(owner, ids)) || !this.install(owner)) return;
      if (waiting) {
        changed.add(owner);
        return;
      }
      this.late += 1;
      const fontIds = this.fontIdsOf(new Set([owner]));
      for (const listener of [...this.lateListeners]) listener(fontIds);
    }));
    this.track(work);
    let timer: ReturnType<typeof setTimeout> | undefined;
    const timedOut = await Promise.race([
      work.then(() => false),
      new Promise<boolean>((resolve) => {
        timer = setTimeout(() => resolve(true), waitMs);
      }),
    ]).finally(() => clearTimeout(timer));
    waiting = false;
    if (timedOut && !this.warned) {
      this.warned = true;
      this.log(`[docx] font outlines still loading after ${waitMs} ms; painting with the outlines installed and repainting when the rest arrive`);
    }
    return this.fontIdsOf(changed);
  }

  /** Settles once every outline requested so far is installed or has failed. */
  async whenInstalled(): Promise<void> {
    while (this.pending.size) await Promise.allSettled([...this.pending]);
  }

  private fontIdsOf(owners: ReadonlySet<number>): Set<number> {
    const ids = new Set<number>();
    for (const [id, owner] of this.ownerOf) if (owners.has(owner)) ids.add(id);
    return ids;
  }

  private track(work: Promise<unknown>): void {
    const settled = work.then(() => {}, () => {});
    this.pending.add(settled);
    void settled.finally(() => this.pending.delete(settled));
  }

  /** Loads the chunks of `ids`; false when the face failed, now or before. */
  private async load(owner: number, ids: Set<number>): Promise<boolean> {
    const state = this.owners.get(owner)!;
    if (state.failed || this.disposed) return false;
    try {
      state.face = await (state.opening ??= state.open());
      await state.face.ensureGlyphs(ids);
      return true;
    } catch (error) {
      this.fail(state, error);
      return false;
    }
  }

  /** Installs the face's current outlines; true when they changed. */
  private install(owner: number): boolean {
    const state = this.owners.get(owner)!;
    const face = state.face;
    if (this.disposed || state.failed || !face || face.revision <= state.installed) return false;
    try {
      this.engine.updateFontOutlines(owner, new Uint8Array(face.bytes()));
      state.installed = face.revision;
      return true;
    } catch (error) {
      if (error instanceof WebAssembly.RuntimeError) throw error;
      this.fail(state, error);
      return false;
    }
  }

  private fail(state: Owner, error: unknown): void {
    if (state.failed) return;
    state.failed = true;
    state.onError(error instanceof Error ? error : new Error(String(error)));
  }
}
