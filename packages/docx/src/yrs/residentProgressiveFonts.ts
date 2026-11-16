/** A sliced face's pack files, as `sliceSourceOf` from `@betteroffice/fonts/progressive` describes them. */
export interface ProgressiveFontSource {
  ref: { format: number; index: string; indexSha256: string };
  /** Absolute URL of the index, skeleton, core and every chunk by manifest path. */
  urls: Record<string, string>;
  timeoutMs?: number;
}

/** A sliced face whose outlines failed in the engine worker, by the page's font id. */
export interface ResidentFontError {
  slot: number;
  message: string;
  /** Set when a slice failed (not fatal): which file, and why. */
  unavailable?: { file: string; reason: 'network' | 'server'; status?: number };
}

/** A font id registered from a sliced face's skeleton, and the face's files. */
export interface ProgressiveFontSpec {
  slot: number;
  source: ProgressiveFontSource;
  /** Decoded files the sender already loaded, by manifest path, which the receiver's open reads instead. */
  files?: ReadonlyMap<string, ArrayBuffer>;
  /** Chunks that failed for the sender: the receiver's open does not fetch them either until a retry. */
  failed?: SliceFailure[];
}

/** A chunk that failed, as `FontSliceUnavailableError` describes it. */
export interface SliceFailure {
  file: string;
  reason: 'network' | 'server';
  status?: number;
}

/** What the engines use of a face `@betteroffice/fonts/progressive` opens. */
interface OutlineFace {
  readonly revision: number;
  bytes(): ArrayBuffer;
  ensureGlyphs(glyphIds: Iterable<number>): Promise<number>;
  ensureText(text: string): Promise<number>;
  retryFailed(): void;
  markFailed(file: string, error: Error): void;
}

/**
 * Some characters cannot be drawn: an outline slice of their face did not
 * load (Q752 B). Not fatal: the text stays in the document with those glyphs
 * blank, and nothing asks for the slice again until the host calls
 * `retryFontSlices()` (on the browser's `online` event). The editor's
 * `onError` hears it once per face until then. `reason` is `network`
 * (offline, unreachable, timed out) or `server` (an error `status`, or bytes
 * that do not match the index).
 */
export class FontSliceUnavailableError extends Error {
  override name = 'FontSliceUnavailableError';
  constructor(
    readonly file: string,
    readonly reason: 'network' | 'server',
    readonly status?: number,
    options?: ErrorOptions,
  ) {
    super(
      `Some characters cannot be drawn until the font slice ${file} loads (${reason}${status === undefined ? '' : ` ${status}`})`,
      options,
    );
  }

  /** From the `FontSliceError` a slice load failed with. */
  static from(error: unknown): FontSliceUnavailableError {
    const { file, reason, status } = (error ?? {}) as { file?: unknown; reason?: unknown; status?: unknown };
    return new FontSliceUnavailableError(
      typeof file === 'string' ? file : 'unknown',
      reason === 'network' ? 'network' : 'server',
      typeof status === 'number' ? status : undefined,
      { cause: error },
    );
  }
}

/**
 * How long a page waits for outlines still loading before it paints those
 * glyphs blank in their space and repaints once they land (Q750 B). The page
 * raster and the engine worker both use it, so a slow slice holds no request
 * near the worker's request budgets and no other page longer than this.
 */
export const PROGRESSIVE_OUTLINE_WAIT_MS = 300;

/** The wait print and thumbnails use: every outline installed, or an explicit failure. */
export const UNTIL_COMPLETE = Infinity;

interface Owner {
  open: () => Promise<OutlineFace>;
  onError: (error: Error) => void;
  opening?: Promise<OutlineFace>;
  face?: OutlineFace;
  /** The face revision whose outlines the engine holds. */
  installed: number;
  /** The face could not open, or the engine refused its outlines: it installs nothing more. */
  failed: boolean;
  /** Why it failed, for a print that needs the face. */
  failure?: unknown;
  /** A slice failed and was reported; reported again only after `retryUnavailable`. */
  unavailable: boolean;
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

/**
 * The text a display page paints with the browser's shaper, `text`
 * primitives (list markers, watermarks) and their leaders: each one's CSS
 * font and the characters it draws.
 */
export function textRunsOf(page: unknown): Array<{ font: string; text: string }> {
  const runs: Array<{ font: string; text: string }> = [];
  forEachTextPrimitive(page, (value) => {
    const painted = typeof value.paintText === 'string' ? value.paintText : typeof value.text === 'string' ? value.text : '';
    const text = value.allCaps === true ? painted.toUpperCase() : painted;
    if (text) runs.push({ font: value.font, text });
    const leader = value.leaderGlyphs;
    if (typeof leader?.font === 'string' && typeof leader.glyph === 'string') runs.push({ font: leader.font, text: leader.glyph });
  });
  return runs;
}

type TextPrimitiveNode = { font: string; text?: unknown; paintText?: unknown; allCaps?: unknown; leaderGlyphs?: { font?: unknown; glyph?: unknown } };

/** Calls `visit` with every `text` primitive of a display page, its regions' and note areas' included. */
export function forEachTextPrimitive(page: unknown, visit: (primitive: TextPrimitiveNode) => void): void {
  const walk = (node: unknown, depth: number): void => {
    if (depth > 12 || node === null || typeof node !== 'object') return;
    if (Array.isArray(node)) {
      for (const child of node) walk(child, depth + 1);
      return;
    }
    const value = node as { kind?: unknown; font?: unknown };
    if (value.kind === 'text' && typeof value.font === 'string') visit(value as TextPrimitiveNode);
    else for (const child of Object.values(node)) walk(child, depth + 1);
  };
  walk(page, 0);
}

const CSS_FONT = /(?:^|\s)\d+(?:\.\d+)?px(?:\/\S+)?\s+(.+)$/;

/** The family list and style of a CSS font shorthand, or null for one it cannot read. */
export function cssFontFamilies(font: string): { families: string[]; bold: boolean; italic: boolean } | null {
  const match = CSS_FONT.exec(font);
  if (!match) return null;
  const style = font.slice(0, match.index);
  return {
    families: match[1]!.split(',').map((family) => family.trim().replace(/^["']|["']$/g, '')).filter(Boolean),
    bold: /\b(?:bold|[6-9]00)\b/.test(style),
    italic: /\b(?:italic|oblique)\b/.test(style),
  };
}

// Characters no Latin face draws: a Latin-family marker or watermark takes them from a sliced CJK face.
const CJK = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Bopomofo}①-⓿　-〿＀-￯]/u;

/** By sliced font id: the characters `page`'s `text` primitives draw with it and the families it stands for. */
export type TextSlicedFaces = Map<number, { characters: Set<string>; replaced: Set<string>; followed: Set<string> }>;

/**
 * The sliced face each `text` primitive of `page` (list markers, watermarks)
 * paints with, as the engine measured it: the face heading its family's chain
 * (`chains`, by `family|bold|italic` key), which replaces the family; else,
 * for CJK characters, the chain's first sliced face after its Latin head or,
 * for a family no chain names (a watermark in a font no text uses), the
 * first sliced face of any chain, which follows the family.
 */
export function textSlicedFaces(
  page: unknown,
  chains: Readonly<Record<string, number[]>>,
  isSliced: (fontId: number) => boolean
): TextSlicedFaces {
  const faces: TextSlicedFaces = new Map();
  const anySliced = Object.values(chains).flat().find(isSliced);
  for (const { font, text } of textRunsOf(page)) {
    const css = cssFontFamilies(font);
    const family = css?.families[0];
    if (!css || !family) continue;
    const key = family.toLowerCase();
    const chain = chains[`${key}|${css.bold ? 1 : 0}|${css.italic ? 1 : 0}`] ?? chains[`${key}|0|0`];
    const heads = chain !== undefined && isSliced(chain[0]!);
    const id = heads ? chain[0] : CJK.test(text) ? (chain?.find(isSliced) ?? anySliced) : undefined;
    if (id === undefined) continue;
    let entry = faces.get(id);
    if (!entry) faces.set(id, entry = { characters: new Set(), replaced: new Set(), followed: new Set() });
    (heads ? entry.replaced : entry.followed).add(family);
    for (const character of text) entry.characters.add(character);
  }
  return faces;
}

/** The CSS families `faces` paint with, each face under `nameOf` its id: a heading face replaces its family, another follows it. */
export function slicedFontFamilies(
  faces: TextSlicedFaces,
  nameOf: (fontId: number) => string | undefined,
  into = new Map<string, string | string[]>()
): Map<string, string | string[]> {
  for (const [id, { replaced, followed }] of faces) {
    const name = nameOf(id);
    if (name === undefined) continue;
    for (const family of replaced) into.set(family, name);
    for (const family of followed) into.set(family, [family, name]);
  }
  return into;
}

/** What a page's painter needs from its sliced faces once they are ready. */
export interface SlicedPaint {
  /** Font ids whose outlines changed, for the glyph cache to drop. */
  changed: Set<number>;
  /** Font ids drawing with sliced faces: their glyph runs paint outlines only, never browser text. */
  outlineOnlyFontIds: ReadonlySet<number>;
  /** The private family each document family's `text` primitives paint with, where its face is sliced. */
  fontFamilies: ReadonlyMap<string, string | readonly string[]>;
}

/** This realm's font set, where sliced faces register for the browser's shaper. */
function realmFonts(): FontFaceSet | undefined {
  if (typeof FontFace !== 'function') return undefined;
  return typeof document !== 'undefined' ? document.fonts : (globalThis as { fonts?: FontFaceSet }).fonts;
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
 * instead of fetched and decoded again, and chunks that `failed` for it are
 * not fetched until a retry.
 */
export function sliceFaceOpener(
  source: ProgressiveFontSource,
  decode: (encoded: ArrayBuffer, maxOutput: number) => Promise<ArrayBuffer>,
  files?: ReadonlyMap<string, ArrayBuffer>,
  failed: readonly SliceFailure[] = [],
  loadModule: () => Promise<ProgressiveModule> = () => import('@betteroffice/fonts/progressive'),
): () => Promise<OutlineFace> {
  return async () => {
    const { openProgressiveFace, readFontAsset } = await loadModule();
    const face = await openProgressiveFace(source.ref, {
      read: (file) => {
        const url = source.urls[file];
        if (url === undefined) return Promise.reject(new Error(`Font slice ${file} has no URL`));
        return readFontAsset(file, new URL(url), source.timeoutMs);
      },
      decode,
      ...(files && { held: (file: string) => files.get(file) }),
    });
    for (const { file, reason, status } of failed) face.markFailed(file, new FontSliceUnavailableError(file, reason, status));
    return face;
  };
}

/** The chunks `face` remembers as failed, for another realm's open of it (`ProgressiveFontSpec.failed`). */
export function sliceFailures(face: { failures(): Map<string, unknown> }): SliceFailure[] {
  return [...face.failures()].map(([file, error]) => {
    const { reason, status } = FontSliceUnavailableError.from(error);
    return { file, reason, ...(status === undefined ? {} : { status }) };
  });
}

/**
 * Font ids registered from sliced CJK faces' skeletons. Before a page is
 * rasterized, the outlines of the glyphs it draws are loaded (each chunk
 * once) and installed in the engine; layout never changes. A chunk that
 * fails is reported once per face as a `FontSliceUnavailableError` and not
 * fetched again until `retryUnavailable`; its glyphs stay blank while the
 * face keeps installing the rest. A face that cannot open, or whose outlines
 * the engine refuses, is reported once and installs nothing more.
 */
export class ResidentProgressiveFonts {
  private readonly owners = new Map<number, Owner>();
  /** Every font id drawing with an owner's bytes: the owners and their measurement views. */
  private readonly ownerOf = new Map<number, number>();
  private readonly pending = new Set<Promise<unknown>>();
  private readonly lateListeners = new Set<(fontIds: Set<number>) => void>();
  /** Sliced faces registered as FontFaces for `text` primitives, by owner, at the revision registered. */
  private readonly textFaces = new Map<number, { name: string; revision: number; face: FontFace }>();
  private warned = false;
  private late = 0;
  private disposed = false;

  constructor(
    private readonly engine: OutlineEngine,
    // A page painting before its outlines land is expected (Q750 B): debug, not a warning.
    private readonly log: (message: string) => void = (message) => console.debug(message),
  ) {}

  /**
   * Font `id` draws with the face `open` gives, whose outlines up to
   * `installed` the engine holds; `onError` hears once that its outlines
   * failed to load or install.
   */
  add(id: number, open: () => Promise<OutlineFace>, onError: (error: Error) => void, installed = 1): void {
    if (this.owners.has(id)) return;
    this.owners.set(id, { open, onError, installed, failed: false, unavailable: false });
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
    const fonts = realmFonts();
    for (const { face } of this.textFaces.values()) fonts?.delete(face);
    this.textFaces.clear();
  }

  /** Every font id drawing with a sliced face, measurement views included. */
  fontIds(): Set<number> {
    return new Set(this.ownerOf.keys());
  }

  /**
   * Readies the `text` primitives of `page` (list markers, watermarks) that
   * paint with a sliced face (`textSlicedFaces`, by the measurement's
   * `chains`): loads the chunks of their characters, waiting up to `waitMs`,
   * and registers the face as a FontFace in this realm under a private
   * family. Resolves with each family's CSS names (`slicedFontFamilies`), so
   * the painter draws them with the face, never a system one; characters
   * still loading draw blank and the page repaints once they land.
   */
  async textFamilies(
    page: unknown,
    chains: Readonly<Record<string, number[]>>,
    waitMs = PROGRESSIVE_OUTLINE_WAIT_MS
  ): Promise<Map<string, string | string[]>> {
    const faces = textSlicedFaces(page, chains, (id) => this.ownerOf.has(id));
    const characters = new Map<number, Set<string>>();
    for (const [id, entry] of faces) {
      const owner = this.ownerOf.get(id)!;
      let wanted = characters.get(owner);
      if (!wanted) characters.set(owner, wanted = new Set());
      for (const character of entry.characters) wanted.add(character);
    }
    const fonts = realmFonts();
    if (characters.size === 0 || !fonts || this.disposed) return new Map();
    await Promise.all([...characters].map(([owner, wanted]) => this.registerText(owner, [...wanted].join(''), waitMs, fonts)));
    return slicedFontFamilies(faces, (id) => this.textFaces.get(this.ownerOf.get(id)!)?.name);
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
   * install later and reach `onLateInstall`. `UNTIL_COMPLETE` (print) waits
   * for every outline and rejects when a slice or the face fails.
   */
  async ensure(glyphs: Map<number, Set<number>>, waitMs = PROGRESSIVE_OUTLINE_WAIT_MS): Promise<Set<number>> {
    const complete = waitMs === UNTIL_COMPLETE;
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
      const installed = (await this.load(owner, ids, complete)) && this.install(owner);
      const state = this.owners.get(owner)!;
      if (complete && state.failed) throw state.failure;
      if (!installed) return;
      if (waiting) {
        changed.add(owner);
        return;
      }
      this.late += 1;
      const fontIds = this.fontIdsOf(new Set([owner]));
      for (const listener of [...this.lateListeners]) listener(fontIds);
    }));
    this.track(work);
    if (complete) {
      await work;
      return this.fontIdsOf(changed);
    }
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

  /** Asks again, on the next frame, for every slice that failed; a failure is reported again. */
  retryUnavailable(): void {
    for (const owner of this.owners.values()) {
      owner.face?.retryFailed();
      owner.unavailable = false;
    }
  }

  /** Settles once every outline requested so far is installed or has failed. */
  async whenInstalled(): Promise<void> {
    while (this.pending.size) await Promise.allSettled([...this.pending]);
  }

  /**
   * Waits for the owner's face to open (its skeleton is held from
   * measurement), loads `text`'s chunks of it up to `waitMs`, and registers
   * the face as it stands; a face the browser cannot load is reported once.
   */
  private async registerText(owner: number, text: string, waitMs: number, fonts: FontFaceSet): Promise<void> {
    const state = this.owners.get(owner)!;
    const complete = waitMs === UNTIL_COMPLETE;
    if (state.failed) {
      if (complete) throw state.failure;
      return;
    }
    let face: OutlineFace;
    try {
      face = state.face = await (state.opening ??= state.open());
    } catch (error) {
      this.fail(state, error);
      if (complete) throw error;
      return;
    }
    let waiting = true;
    const loading = face.ensureText(text).then(
      (revision) => {
        // Characters a page stopped waiting for repaint once they land.
        if (!waiting && !this.disposed && revision > (this.textFaces.get(owner)?.revision ?? 0)) {
          for (const listener of [...this.lateListeners]) listener(this.fontIdsOf(new Set([owner])));
        }
      },
      (error: unknown) => {
        this.chunkFailed(state, error);
        if (complete) throw error;
      }
    );
    this.track(loading);
    if (complete) {
      await loading;
    } else {
      let timer: ReturnType<typeof setTimeout> | undefined;
      await Promise.race([loading, new Promise<void>((resolve) => {
        timer = setTimeout(resolve, waitMs);
      })]).finally(() => clearTimeout(timer));
    }
    waiting = false;
    const revision = face.revision;
    if (this.disposed || (this.textFaces.get(owner)?.revision ?? 0) >= revision) return;
    const name = `BetterOffice Sliced ${owner}`;
    const fontFace = new FontFace(name, face.bytes());
    try {
      await fontFace.load();
    } catch (error) {
      this.fail(state, error);
      if (complete) throw error;
      return;
    }
    // Another page may have registered the same or a newer revision meanwhile.
    const known = this.textFaces.get(owner);
    if (this.disposed || (known?.revision ?? 0) >= revision) return;
    if (known) fonts.delete(known.face);
    fonts.add(fontFace);
    this.textFaces.set(owner, { name, revision, face: fontFace });
  }

  /**
   * A chunk that could not be fetched (network or server) is reported once
   * until `retryUnavailable`; any other failure, such as a handed-over chunk
   * that does not match its index entry, is our bug and fails the face.
   */
  private chunkFailed(state: Owner, error: unknown): void {
    const { file, reason } = (error ?? {}) as { file?: unknown; reason?: unknown };
    if (typeof file === 'string' && (reason === 'network' || reason === 'server')) this.unavailable(state, error);
    else this.fail(state, error);
  }

  /** Reports a slice that failed for this face, once until `retryUnavailable`. */
  private unavailable(state: Owner, error: unknown): void {
    if (state.unavailable || this.disposed) return;
    state.unavailable = true;
    state.onError(FontSliceUnavailableError.from(error));
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

  /**
   * Loads the chunks of `ids`; false when the face failed, now or before. A
   * chunk that fails leaves the rest to install.
   */
  private async load(owner: number, ids: Set<number>, complete = false): Promise<boolean> {
    const state = this.owners.get(owner)!;
    if (state.failed || this.disposed) return false;
    try {
      state.face = await (state.opening ??= state.open());
    } catch (error) {
      this.fail(state, error);
      return false;
    }
    try {
      await state.face.ensureGlyphs(ids);
    } catch (error) {
      this.chunkFailed(state, error);
      if (complete) throw error;
    }
    return !state.failed;
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
    state.failure = error;
    state.onError(error instanceof Error ? error : new Error(String(error)));
  }
}
