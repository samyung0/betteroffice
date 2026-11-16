/** Where a sliced face's files are, as `@betteroffice/fonts` loaders carry it; opaque to the registry. */
export type SliceLocation = object;

/** What the layout uses of a face `@betteroffice/fonts/progressive` opens. */
export interface ProgressiveFontFace {
  /** Increases each time outlines are added. */
  readonly revision: number;
  /** The face with every outline loaded so far. */
  bytes(): ArrayBuffer;
  ensureGlyphs(glyphIds: Iterable<number>): Promise<number>;
  ensureText(text: string): Promise<number>;
  /** Copies of the decoded files loaded so far, which another realm's open of the face can take. */
  files(): Map<string, ArrayBuffer>;
  /** Asks again for the chunks that failed; until then they fail without a fetch. */
  retryFailed(): void;
  /** `file`, a chunk of this face, failed in another realm: not fetched here either until `retryFailed`. */
  markFailed(file: string, error: Error): void;
  /** The chunks that failed and are not fetched again until `retryFailed`, by file, with their errors. */
  failures(): Map<string, unknown>;
}

/** A sliced face registered from its skeleton, and where it came from. */
export interface SlicedFontFace {
  face: ProgressiveFontFace;
  slices: SliceLocation;
}

import { baseLatinAlias, baseLatinCandidate, coverageFromBits, symbolFamily, wordsCoverage, type FontCoverage } from './fontCoverage';
import { ImplicitFontAbsentError } from './defaultFontProvider';

async function parsedCoverage(data: Uint8Array): Promise<FontCoverage | undefined> {
  try {
    const words = (await import('./cmapCoverage')).fontCoverageWords(data);
    return words && wordsCoverage(words);
  } catch {
    return undefined;
  }
}

/** A lazy face loader with measurement and fallback options. */
export type BundledFontLoader = (() => Promise<ArrayBuffer>) & {
  /** Word measurement family; defaults to the requested family. */
  measureAs?: string;
  /** Only stands in for a primary face that failed to load. */
  onlyIfPrimaryFails?: boolean;
  /** Locates a sliced face whose outlines load as they are needed. */
  slices?: SliceLocation;
  /** @internal Precomputed base Latin coverage bits of the face. */
  coverage?: string;
  /** @internal The face's decoded bytes, shared with the font cache: read them, never change or transfer them. */
  readOnlyBytes?: () => Promise<ArrayBuffer>;
};

/** Byte sink of the wasm text engine (`crates/docx-text` FontStore). */
export interface TextEngineFontSink {
  /**
   * Register raw sfnt bytes with the engine; returns its font id.
   * Throws on unparseable bytes.
   */
  registerFont(bytes: Uint8Array, face?: { family: string; bold: boolean; italic: boolean }): number;
  /**
   * Register a measurement view of `id` carrying the vertical metrics and
   * advance pitch Word measures `family` with, and return its id — the
   * engine's answer to a face this host had to substitute. Returns `id` for a
   * family whose metrics the engine does not know. Optional so partial sinks
   * (tests, older hosts) keep working; without it a substitute measures with
   * its own metrics.
   */
  registerSubstituteFont?(id: number, family: string): number;
  /**
   * Called once `id` is registered from a sliced face's skeleton:
   * its glyphs measure exactly, and their outlines install with
   * `face.ensureGlyphs` and the engine's outline update before they paint.
   */
  registerSlicedFont?(id: number, face: ProgressiveFontFace): void;
}

/**
 * Script bucket a fallback face provides glyph coverage for. Han text is
 * region-resolved by the caller (SC/TC/JP/KR — see `rustMeasureSource.ts`);
 * kana, Hangul, Arabic and Hebrew detect directly.
 *
 * @public
 */
export type FontScript = 'cjk-sc' | 'cjk-tc' | 'cjk-jp' | 'cjk-kr' | 'arabic' | 'hebrew';

/**
 * Resolver for the bundled metric-compatible set (Carlito↔Calibri,
 * Liberation↔Arial/Times/Courier, …). Returns a lazy byte loader so the
 * (lazily fetched, same-origin) binary is only downloaded when a document
 * actually needs the face, or `undefined` when the family has no bundled
 * substitute.
 *
 * @public
 */
export interface BundledFontProvider {
  /** @internal */
  resolveBaseLatin?(bold: boolean, italic: boolean): BundledFontLoader | undefined;
  /** Resolve a Word family to bundled metric-compatible face byte loaders, or undefined. */
  resolve(family: string, bold: boolean, italic: boolean): BundledFontLoader | undefined;
  /** Like {@link BundledFontProvider.resolve}, but also accepts bundled family names such as `"Gelasio"`, and leaves missing styles for browser synthesis. */
  resolveFamily?(
    family: string,
    bold: boolean,
    italic: boolean
  ): (() => Promise<ArrayBuffer>) | undefined;
  /**
   * Optional per-script coverage fallback (Noto CJK/RTL faces). Same loader
   * contract as {@link BundledFontProvider.resolve}; providers without
   * script faces simply omit the method. These faces are coverage fallbacks
   * first, metric approximations second — CJK metric compatibility is much
   * weaker than the Carlito/Calibri class of substitutes.
   */
  resolveScriptFallback?(
    script: FontScript,
    bold: boolean,
    italic: boolean
  ): (() => Promise<ArrayBuffer>) | undefined;
  /** Coverage for characters absent from a bundled substitute. */
  resolveGlyphFallback?(
    family: string,
    bold: boolean,
    italic: boolean,
    office?: 'word' | 'powerpoint'
  ): (() => Promise<ArrayBuffer>) | undefined;
  /**
   * The always-available last-resort base face (broad-coverage Latin —
   * Liberation Sans/Serif per serif-ness). Unlike {@link
   * BundledFontProvider.resolve}, a conforming provider (e.g.
   * `@betteroffice/fonts`'s `resolveLastResortFace`) returns a loader for
   * EVERY family, so the chain this registry builds is guaranteed non-empty
   * and a run never routes to the browser measurer for want of font bytes. The
   * face's metrics are the base font's, not the requested family's — the
   * accepted divergence for a truly-unknown font, in exchange for staying on
   * the native measurement path.
   *
   * Optional so mock/partial providers can omit it; when absent (or returning
   * undefined) an unmapped family still yields an empty chain and the caller
   * browser-falls-back that run — the pre-policy behavior. Same lazy loader
   * contract as {@link BundledFontProvider.resolve}. This registry passes
   * `office: 'word'` so the pick follows Word's substitution.
   */
  resolveLastResort?(
    family: string,
    bold: boolean,
    italic: boolean,
    office?: 'word' | 'powerpoint'
  ): BundledFontLoader | undefined;
}

/**
 * One embedded face extracted from the open document. Structurally identical
 * to `EmbeddedFontFace` (`utils/embeddedFonts.ts`), so the output of
 * `getEmbeddedFontFaces` feeds straight in; duplicated here so this module
 * stays decoupled from the loader.
 *
 * @public
 */
export interface EmbeddedFaceInput {
  /** Word font name the face is registered under (attacker-controlled). */
  family: string;
  /** `'bold'` for the embedBold/embedBoldItalic slots, else `'normal'`. */
  weight: 'normal' | 'bold';
  /** `'italic'` for the embedItalic/embedBoldItalic slots, else `'normal'`. */
  style: 'normal' | 'italic';
  /** De-obfuscated OpenType/TrueType bytes. */
  data: ArrayBuffer;
  /** Whether the source face was subsetted (`w:subsetted`). */
  subsetted: boolean;
}

/**
 * Family keys are matched case-insensitively and whitespace-trimmed —
 * the same normalization `resolveFontFamily` (`utils/fontResolver.ts`)
 * applies to Word font names.
 */
function familyKey(family: string): string {
  return family.trim().toLowerCase();
}

function chainKey(family: string, bold: boolean, italic: boolean): string {
  return `${familyKey(family)}|${bold ? 1 : 0}|${italic ? 1 : 0}`;
}

/** A provider or async factory; `undefined` means no bundled bytes. @public */
export type BundledFontProviderSource =
  | BundledFontProvider
  | (() => Promise<BundledFontProvider | undefined>);

interface ChainResolution {
  ids: number[];
  retryable: boolean;
}

interface ScriptResolution {
  id: number | null;
  retryable: boolean;
}

interface PromiseMemo<K, V> {
  get(key: K): Promise<V> | undefined;
  delete(key: K): boolean;
}

function evictFailedPromise<K, V>(
  memo: PromiseMemo<K, V>,
  key: K,
  pending: Promise<V>,
  failed: (value: V) => boolean
): void {
  pending.then(
    (value) => {
      if (failed(value) && memo.get(key) === pending) memo.delete(key);
    },
    () => {
      if (memo.get(key) === pending) memo.delete(key);
    }
  );
}

export class TextMeasureFontRegistry {
  private coverage = new Map<number, FontCoverage>();
  /** Ids whose coverage is a known lower bound, with the loader that yields their exact coverage. */
  private inexact = new Map<number, () => Promise<FontCoverage | undefined>>();
  private baseLatinResults = new Map<string, { ids: number[]; characters: Set<string>; earlier: string; failed?: Set<string> }>();
  private readonly baseLatinSource: BundledFontProviderSource | undefined;
  private baseLatinLoaders = new WeakMap<BundledFontLoader, BundledFontLoader>();
  private readonly sink: TextEngineFontSink;
  private readonly bundledSource: BundledFontProviderSource | undefined;
  private bundledPromise: Promise<BundledFontProvider | undefined> | undefined;
  /** Per-registry so misconfigured editors warn independently. */
  private warnedFallback = false;

  /** Normalized family → embedded faces of the current document. */
  private facesByFamily = new Map<string, EmbeddedFaceInput[]>();
  private appFacesByFamily = new Map<string, EmbeddedFaceInput[]>();
  /**
   * Per-face registration memo, keyed by face object identity so re-feeding
   * the same faces (or sharing one face across several chains) never
   * re-registers its bytes. Failed registrations are evicted. Weak so faces of
   * a closed document can be collected.
   */
  private faceIds = new WeakMap<EmbeddedFaceInput, Promise<number | null>>();
  private bundledIds = new Map<string, Promise<number | null>>();
  /**
   * Last-resort base-face registration memo, keyed by chain key. Kept separate
   * from `bundledIds` so a family's metric-compat face and its (possibly
   * different) last-resort base face don't clobber each other's registration.
   * Document-independent like `bundledIds` — survives `setEmbeddedFaces`, reset
   * by `clear()`. Failed registrations are evicted.
   */
  private lastResortIds = new Map<string, Promise<number | null>>();
  /**
   * Byte-identity dedupe across every registration path: concurrent loads of
   * one face share a buffer, so the same face requested under several chain
   * keys (or as both a family face and a script fallback) registers once. A
   * later load of a face hands over new bytes, which the engine itself
   * matches to the face's first id. Rejections are evicted. Weak so buffers
   * of a cleared registry can be collected.
   */
  private bufferIds = new WeakMap<ArrayBuffer, Promise<number>>();
  private loaderIds = new WeakMap<() => Promise<ArrayBuffer>, Promise<number>>();
  private scriptIds = new Map<FontScript, Promise<ScriptResolution>>();
  /** Settled per-script results for the synchronous view. */
  private scriptResults = new Map<FontScript, ScriptResolution>();
  /** Chain memo — concurrent `getFontIdChain` calls share one resolution. */
  private chains = new Map<string, Promise<ChainResolution>>();
  /** Settled chains for the synchronous view. */
  private chainResults = new Map<string, ChainResolution>();
  /** Bumped on invalidation so stale in-flight resolutions can't repopulate caches. */
  private generation = 0;
  /**
   * Bumped only by `clear()`, when the engine's font store is replaced: a
   * face opened before then must not register into the new store, while one
   * that outlives a `setEmbeddedFaces` (same store) still must.
   */
  private storeGeneration = 0;

  /** Sliced faces registered from their skeleton, by font id and measurement view id. */
  private progressiveIds = new Map<number, SlicedFontFace>();
  /**
   * Sliced faces whose index or skeleton failed, by slice index: layout
   * needs them, so the failure stands (reported once, not fetched again on
   * later passes) until `clear()`.
   */
  private failedSlices = new Map<string, unknown>();
  /** Failures already reported, so a remembered one reaches `onError` once. */
  private reported = new WeakSet<object>();
  private readonly progressive?: (slices: SliceLocation) => Promise<ProgressiveFontFace>;
  private readonly onError?: (error: Error) => void;

  /**
   * `progressive` opens sliced faces, which register from their skeleton.
   * `onError` hears every face that fails to load (the chain then
   * measures without it), so a host can show the failure rather than a
   * fallback layout.
   */
  constructor(sink: TextEngineFontSink, opts?: {
    bundled?: BundledFontProviderSource;
    baseLatin?: BundledFontProviderSource;
    progressive?: (slices: SliceLocation) => Promise<ProgressiveFontFace>;
    onError?: (error: Error) => void;
  }) {
    this.sink = sink;
    this.bundledSource = opts?.bundled;
    this.baseLatinSource = opts?.baseLatin;
    this.progressive = opts?.progressive;
    this.onError = opts?.onError;
  }

  /** The sliced faces registered from their skeleton, by font id and measurement view id. */
  progressiveFaces(): ReadonlyMap<number, SlicedFontFace> {
    return this.progressiveIds;
  }

  /** Warns once and reports a face that failed to load; an absent optional add-on is not a failure. */
  private failed(message: string, error: unknown): void {
    this.warnFailure(`${message}: ${error instanceof Error ? error.message : String(error)}`);
    if (error instanceof ImplicitFontAbsentError) return;
    if (typeof error === 'object' && error !== null) {
      if (this.reported.has(error)) return;
      this.reported.add(error);
    }
    this.onError?.(error instanceof Error ? error : new Error(String(error)));
  }

  /** Shares in-flight resolution but evicts misses for retry. */
  private bundled(): Promise<BundledFontProvider | undefined> {
    if (this.bundledPromise === undefined) {
      const source = this.bundledSource;
      const promise = Promise.resolve()
        .then(() => (typeof source === 'function' ? source() : source))
        .catch(() => undefined);
      promise.then((provider) => {
        if (provider === undefined && this.bundledPromise === promise) {
          this.bundledPromise = undefined;
        }
      });
      this.bundledPromise = promise;
    }
    return this.bundledPromise;
  }

  /**
   * Feed the embedded faces extracted from the open document, replacing any
   * previous set and invalidating every cached chain. Font ids already issued
   * by the sink stay valid (the engine keeps the bytes); faces passed again
   * by object identity keep their registration. Script-fallback
   * registrations are document-independent and survive.
   */
  setEmbeddedFaces(faces: EmbeddedFaceInput[], appFaces: EmbeddedFaceInput[] = []): void {
    this.baseLatinResults.clear();
    this.generation++;
    this.facesByFamily = new Map();
    this.appFacesByFamily = new Map();
    for (const [source, byFamily] of [[faces, this.facesByFamily], [appFaces, this.appFacesByFamily]] as const) {
      for (const face of source) {
        const key = familyKey(face.family);
        const list = byFamily.get(key);
        if (list) list.push(face);
        else byFamily.set(key, [face]);
      }
    }
    this.chains = new Map();
    this.chainResults = new Map();
  }

  /**
   * Lazily resolves and registers a font chain. Concurrent calls share work;
   * failed registrations remain retryable.
   */
  getFontIdChain(family: string, bold: boolean, italic: boolean): Promise<number[]> {
    const key = chainKey(family, bold, italic);
    let chain = this.chains.get(key);
    if (!chain) {
      const chains = this.chains;
      const results = this.chainResults;
      const generation = this.generation;
      const pending = this.resolveChain(family, bold, italic);
      pending.then(
        (resolution) => {
          if (generation !== this.generation) return;
          results.set(key, resolution);
          if (resolution.retryable && chains.get(key) === pending) chains.delete(key);
        },
        () => {
          if (chains.get(key) === pending) chains.delete(key);
        }
      );
      chains.set(key, pending);
      chain = pending;
    }
    return chain.then((resolution) => resolution.ids);
  }

  /** Synchronous settled view; retryable results are evicted after one read. */
  getCachedFontIdChain(
    family: string,
    bold: boolean,
    italic: boolean,
    consumeRetryable = true
  ): readonly number[] | undefined {
    const key = chainKey(family, bold, italic);
    const resolution = this.chainResults.get(key);
    if (consumeRetryable && resolution?.retryable) {
      const results = this.chainResults;
      queueMicrotask(() => {
        if (results.get(key) === resolution) results.delete(key);
      });
    }
    return resolution?.ids;
  }

  coverageOf(id: number): FontCoverage | undefined {
    return this.coverage.get(id);
  }

  /** Characters the base Latin tier settled for `key`, without those whose face failed to load. */
  settledBaseLatinCharacters(key: string): Set<string> {
    const result = this.baseLatinResults.get(key);
    return new Set([...(result?.characters ?? [])].filter(character => !result?.failed?.has(character)));
  }

  getCachedBaseLatinIds(
    family: string,
    bold: boolean,
    italic: boolean,
    characters?: readonly string[],
    earlier?: readonly number[],
    key = chainKey(family, bold, italic),
  ): readonly number[] | undefined {
    if (earlier?.length === 0 || earlier?.some(id => !this.coverage.has(id))) return [];
    const result = this.baseLatinResults.get(key);
    if (characters?.length) {
      if (result && earlier !== undefined && result.earlier !== earlier.join(',')) return undefined;
      const ids = [...(earlier ?? []), ...(result?.ids ?? [])];
      for (const character of characters) {
        if (result?.characters.has(character)) continue;
        if (baseLatinCandidate(character, false) && (!baseLatinAlias(character) || symbolFamily(family)) &&
          !ids.some(id => this.coverage.get(id)?.(character))) return undefined;
        result?.characters.add(character);
      }
    }
    return result?.ids ?? [];
  }

  async prepareBaseLatin(
    family: string,
    bold: boolean,
    italic: boolean,
    characters: readonly string[],
    earlier: readonly number[],
    key = chainKey(family, bold, italic),
  ): Promise<boolean> {
    const generation = this.generation;
    const aliases = symbolFamily(family);
    const uncovered = () => earlier.length === 0 || earlier.some(id => !this.coverage.has(id)) ? [] : characters.filter((character) =>
      baseLatinCandidate(character, false) && (!baseLatinAlias(character) || aliases) &&
      !earlier.some((id) => this.coverage.get(id)?.(character)));
    let missing = uncovered();
    if (missing.length > 0 && earlier.some((id) => this.inexact.has(id))) {
      await Promise.all(earlier.map((id) => this.exactCoverage(id)));
      missing = uncovered();
    }
    const ids: number[] = [];
    let failed: Set<string> | undefined;
    if (missing.length > 0) {
      try {
        const { loadBaseLatinFaces } = await import('./baseLatinFaces');
        await loadBaseLatinFaces({
          bold, italic, missing, ids,
          provider: async () => {
            const provider = await this.bundled();
            if (provider?.resolveBaseLatin) return provider;
            const source = this.baseLatinSource;
            return typeof source === 'function' ? source() : source;
          },
          wrapped: this.baseLatinLoaders,
          register: (loader, face) => this.registerLoader(loader, face),
          coverage: (id) => this.coverage.get(id),
          candidate: baseLatinCandidate,
        });
      } catch (error) {
        failed = new Set(missing);
        this.failed(`[fontRegistry] base Latin face for "${family}" failed to load; falling back`, error);
      }
    }
    if (generation !== this.generation) return false;
    const previous = this.baseLatinResults.get(key)?.ids ?? [];
    this.baseLatinResults.set(key, { ids, characters: new Set(characters), earlier: earlier.join(','), failed });
    return previous.length !== ids.length || previous.some((id, index) => ids[index] !== id);
  }

  /** Resolves unique script fallbacks in order; stable misses cache and failures retry. */
  getScriptFallbackIds(scripts: FontScript[]): Promise<number[]> {
    const unique = [...new Set(scripts)];
    return Promise.all(unique.map((script) => this.resolveScript(script))).then((resolutions) => {
      const out: number[] = [];
      for (const { id } of resolutions) {
        if (id !== null && !out.includes(id)) out.push(id);
      }
      return out;
    });
  }

  /** Synchronous settled view; retryable results are evicted after one pass. */
  getCachedScriptFallbackIds(
    scripts: FontScript[],
    consumeRetryable = true
  ): readonly number[] | undefined {
    const results = this.scriptResults;
    const resolutions: Array<[FontScript, ScriptResolution]> = [];
    for (const script of scripts) {
      const resolution = results.get(script);
      if (!resolution) return undefined;
      resolutions.push([script, resolution]);
    }
    const out: number[] = [];
    for (const [script, resolution] of resolutions) {
      const { id } = resolution;
      if (id !== null && !out.includes(id)) out.push(id);
      if (consumeRetryable && resolution.retryable) {
        queueMicrotask(() => {
          if (results.get(script) === resolution) results.delete(script);
        });
      }
    }
    return out;
  }

  /**
   * Forget every cached chain, face registration, bundled registration and
   * script-fallback registration. Call when the engine/FontStore is
   * recreated — previously issued font ids are invalid then. The embedded
   * face set is retained; pass a new set via `setEmbeddedFaces` when the
   * document changes.
   */
  clear(): void {
    this.coverage.clear();
    this.inexact.clear();
    this.baseLatinResults.clear();
    this.baseLatinLoaders = new WeakMap();
    this.generation++;
    this.storeGeneration++;
    this.bundledPromise = undefined;
    this.faceIds = new WeakMap();
    this.bundledIds = new Map();
    this.lastResortIds = new Map();
    this.bufferIds = new WeakMap();
    this.loaderIds = new WeakMap();
    this.progressiveIds = new Map();
    this.failedSlices = new Map();
    this.scriptIds = new Map();
    this.scriptResults = new Map();
    this.chains = new Map();
    this.chainResults = new Map();
  }

  private warnFailure(message: string): void {
    if (this.warnedFallback) return;
    this.warnedFallback = true;
    console.warn(message);
  }

  private async resolveChain(
    family: string,
    bold: boolean,
    italic: boolean
  ): Promise<ChainResolution> {
    const key = chainKey(family, bold, italic);
    const ids: number[] = [];
    let retryable = false;

    // Every face the chain may use starts loading at once, so a document's
    // faces download in one batch; the ids still go in chain order. Only a
    // last resort that stands in for a failed primary waits for it.
    const embedded = [this.facesByFamily, this.appFacesByFamily].map((byFamily) => {
      const face = this.pickEmbeddedFace(family, bold, italic, byFamily);
      return face && this.registerFace(face);
    });

    const bundled = await this.bundled();
    if (bundled === undefined && typeof this.bundledSource === 'function') retryable = true;
    let loader: BundledFontLoader | undefined;
    try {
      loader = bundled?.resolve(family, bold, italic);
    } catch (error) {
      retryable = true;
      this.warnFailure(
        `[fontRegistry] bundled face resolver for "${family}" failed; falling back: ` +
          `${error instanceof Error ? error.message : String(error)}`
      );
    }
    // Terminal link: the always-available last-resort base face. Appended after
    // the embedded + metric-compatible faces so the chain NEVER ends empty — a
    // run whose family has no embedded/bundled match still has real (Liberation)
    // bytes to measure with, keeping the block on the native path instead of the
    // browser fallback. Deduped by engine id (a family whose metric-compat IS
    // the base face, e.g. Arial→Liberation Sans, contributes one id).
    let lastResort: BundledFontLoader | undefined;
    try {
      lastResort = bundled?.resolveLastResort?.(family, bold, italic, 'word');
    } catch (error) {
      retryable = true;
      this.warnFailure(
        `[fontRegistry] last-resort resolver for "${family}" failed; falling back: ` +
          `${error instanceof Error ? error.message : String(error)}`
      );
    }
    let coverage: (() => Promise<ArrayBuffer>) | undefined;
    try {
      coverage = bundled?.resolveGlyphFallback?.(family, bold, italic, 'word');
    } catch {
      retryable = true;
    }

    const primary = loader && this.registerBundled(key, loader, family);
    const eagerLastResort = lastResort && !lastResort.onlyIfPrimaryFails
      ? this.registerLastResort(key, lastResort, family) : undefined;
    const glyphFallback = coverage && this.registerLastResort(`${key}|coverage`, coverage, family);

    for (const pending of embedded) {
      if (!pending) continue;
      const id = await pending;
      if (id !== null && !ids.includes(id)) ids.push(id);
      else if (id === null) retryable = true;
    }
    let primaryRegistered = false;
    if (primary) {
      const id = await primary;
      if (id !== null && !ids.includes(id)) ids.push(id);
      if (id === null) retryable = true;
      primaryRegistered = id !== null;
    }
    const lastResortId = eagerLastResort ??
      (lastResort && !primaryRegistered ? this.registerLastResort(key, lastResort, family) : undefined);
    for (const pending of [lastResortId, glyphFallback]) {
      if (!pending) continue;
      const id = await pending;
      if (id !== null && !ids.includes(id)) ids.push(id);
      if (id === null) retryable = true;
    }

    if (ids.length === 0) {
      this.warnFallback(
        family,
        bundled !== undefined,
        loader !== undefined || lastResort !== undefined
      );
    }

    return { ids, retryable };
  }

  /**
   * An empty chain forces browser measurement. Distinguish absent, incomplete,
   * and broken providers so the one warning gives applicable remediation.
   */
  private warnFallback(family: string, providerResolved: boolean, loaderResolved: boolean): void {
    if (this.warnedFallback) return;
    this.warnedFallback = true;
    const consequence =
      'Text is measured with approximate fallback metrics, so pagination can vary. ' +
      'Reported once per registry.';
    console.warn(
      loaderResolved
        ? `[fontRegistry] no font bytes for "${family}": the font provider resolved but every ` +
            `matching face failed to load. ${consequence} If a bundler inlined ` +
            '@betteroffice/fonts, mark it external so its asset URLs resolve.'
        : providerResolved
          ? `[fontRegistry] no font bytes for "${family}": the configured font provider has ` +
            `no matching face. Add coverage for this family. ${consequence}`
          : `[fontRegistry] no font bytes for "${family}": no font provider is configured. ` +
            'Install @betteroffice/fonts and pass it with configureDefaultFonts({ fonts }). ' +
            consequence
    );
  }

  /**
   * Exact (weight, style) match first; otherwise the family's regular face —
   * the engine synthesizes bold/italic from regular outlines when asked for a
   * style the document did not embed.
   */
  private pickEmbeddedFace(
    family: string,
    bold: boolean,
    italic: boolean,
    byFamily: ReadonlyMap<string, EmbeddedFaceInput[]>,
  ): EmbeddedFaceInput | undefined {
    const faces = byFamily.get(familyKey(family));
    if (!faces) return undefined;
    const weight = bold ? 'bold' : 'normal';
    const style = italic ? 'italic' : 'normal';
    return (
      faces.find((f) => f.weight === weight && f.style === style) ??
      faces.find((f) => f.weight === 'normal' && f.style === 'normal')
    );
  }

  /**
   * The id that measures `base`'s bytes the way Word measures `family` — what
   * a chain head must be for a substituted face, since line metrics come from
   * the head. Falls back to `base` when the sink or the engine has no metrics
   * for the family, which is how every face measured before this existed.
   */
  private measuredAs(base: number, family: string): number {
    const substitute = this.sink.registerSubstituteFont;
    if (!substitute) return base;
    try {
      const id = substitute.call(this.sink, base, family);
      const coverage = this.coverage.get(base);
      if (coverage) this.coverage.set(id, coverage);
      const exact = this.inexact.get(base);
      if (exact) this.inexact.set(id, exact);
      const sliced = this.progressiveIds.get(base);
      if (sliced) this.progressiveIds.set(id, sliced);
      return id;
    } catch (error) {
      this.warnFailure(
        `[fontRegistry] measurement view of the substitute for "${family}" failed; ` +
          `measuring with the substitute's own metrics: ` +
          `${error instanceof Error ? error.message : String(error)}`
      );
      return base;
    }
  }

  /** Register raw bytes exactly once per buffer identity (see `bufferIds`). */
  private registerBuffer(
    bytes: ArrayBuffer,
    face?: { family: string; bold: boolean; italic: boolean },
    known?: BundledFontLoader,
  ): Promise<number> {
    let pending = this.bufferIds.get(bytes);
    if (!pending) {
      const ids = this.bufferIds;
      pending = Promise.resolve().then(async () => {
        const data = new Uint8Array(bytes);
        const id = face ? this.sink.registerFont(data, face) : this.sink.registerFont(data);
        const bits = known?.coverage;
        const coverage = bits === undefined ? await parsedCoverage(data) : coverageFromBits(bits);
        if (coverage) {
          this.coverage.set(id, coverage);
          if (bits !== undefined) {
            let exact: Promise<FontCoverage | undefined> | undefined;
            this.inexact.set(id, () => (exact ??= (known!.readOnlyBytes ?? known!)().then((bytes) => parsedCoverage(new Uint8Array(bytes)))));
          }
        }
        return id;
      });
      evictFailedPromise(ids, bytes, pending, () => false);
      ids.set(bytes, pending);
    }
    return pending;
  }

  /** Replaces a known lower bound by the face's exact coverage. */
  private async exactCoverage(id: number): Promise<void> {
    const exact = this.inexact.get(id);
    if (!exact) return;
    let coverage: FontCoverage | undefined;
    try { coverage = await exact(); } catch {}
    if (this.inexact.get(id) !== exact) return;
    this.inexact.delete(id);
    if (coverage) this.coverage.set(id, coverage);
  }

  /**
   * Registers a sliced face's skeleton when progressive loading is
   * on, else the whole face. A sliced face that fails to open fails: only a
   * host without the optional CJK add-on loads the face whole, which stands in
   * a Latin face.
   */
  private registerFaceLoader(loader: BundledFontLoader): Promise<number> {
    const open = this.progressive;
    const slices = loader.slices;
    if (open === undefined || slices === undefined) return this.registerLoader(loader);
    let pending = this.loaderIds.get(loader);
    if (!pending) {
      const ids = this.loaderIds;
      const generation = this.storeGeneration;
      const failures = this.failedSlices;
      // Every slice location names its index (`SliceLocation` is opaque here).
      const key = (slices as { ref: { index: string } }).ref.index;
      pending = (async () => {
        if (failures.has(key)) throw failures.get(key);
        let face: ProgressiveFontFace;
        try {
          face = await open(slices);
        } catch (error) {
          if ((error as { name?: unknown } | null)?.name !== 'MissingCjkAddonError') {
            failures.set(key, error);
            throw error;
          }
          return this.registerBuffer(await (loader.readOnlyBytes ?? loader)(), undefined, loader);
        }
        const known = loader.coverage === undefined ? undefined
          : Object.assign(async () => face.bytes(), { coverage: loader.coverage });
        const id = await this.registerBuffer(face.bytes(), undefined, known);
        // Identical skeletons register once; the first face opened keeps the id.
        if (generation === this.storeGeneration && !this.progressiveIds.has(id)) {
          this.progressiveIds.set(id, { face, slices });
          this.sink.registerSlicedFont?.(id, face);
        }
        return id;
      })();
      evictFailedPromise(ids, loader, pending, () => false);
      ids.set(loader, pending);
    }
    return pending;
  }

  private registerLoader(loader: BundledFontLoader, face?: { family: string; bold: boolean; italic: boolean }): Promise<number> {
    let pending = this.loaderIds.get(loader);
    if (!pending) {
      const ids = this.loaderIds;
      pending = Promise.resolve().then(() => (loader.readOnlyBytes ?? loader)()).then((bytes) => this.registerBuffer(bytes, face, loader));
      evictFailedPromise(ids, loader, pending, () => false);
      ids.set(loader, pending);
    }
    return pending;
  }

  private registerFace(face: EmbeddedFaceInput): Promise<number | null> {
    let pending = this.faceIds.get(face);
    if (!pending) {
      const ids = this.faceIds;
      pending = (async () => {
        try {
          return await this.registerBuffer(face.data);
        } catch {
          // A corrupt embedded face (attacker-controlled bytes the engine
          // rejected) must not take down the chain — drop it and let the
          // bundled/browser fallbacks cover the run. console.warn matches
          // the sibling loader's convention (see loadEmbeddedFonts).
          this.warnFailure(
            `[fontRegistry] embedded face "${face.family}" (${face.weight} ${face.style}) ` +
              'was rejected by the text engine; falling back'
          );
          return null;
        }
      })();
      evictFailedPromise(ids, face, pending, (id) => id === null);
      ids.set(face, pending);
    }
    return pending;
  }

  private registerBundled(
    key: string,
    loader: BundledFontLoader,
    family: string
  ): Promise<number | null> {
    let pending = this.bundledIds.get(key);
    if (!pending) {
      const ids = this.bundledIds;
      pending = (async () => {
        try {
          const base = await this.registerFaceLoader(loader);
          return this.measuredAs(base, loader.measureAs ?? family);
        } catch (error) {
          this.failed(`[fontRegistry] bundled face for "${family}" failed to load or register; falling back`, error);
          return null;
        }
      })();
      evictFailedPromise(ids, key, pending, (id) => id === null);
      ids.set(key, pending);
    }
    return pending;
  }

  /**
   * Register the last-resort base face. Mirrors {@link registerBundled} but
   * uses its own memo so it never collides with metric-compatible registration.
   */
  private registerLastResort(
    key: string,
    loader: () => Promise<ArrayBuffer>,
    family: string
  ): Promise<number | null> {
    let pending = this.lastResortIds.get(key);
    if (!pending) {
      const ids = this.lastResortIds;
      pending = (async () => {
        try {
          return this.measuredAs(await this.registerLoader(loader), family);
        } catch (error) {
          this.failed(`[fontRegistry] last-resort base face for "${family}" failed to load or register; falling back`, error);
          return null;
        }
      })();
      evictFailedPromise(ids, key, pending, (id) => id === null);
      ids.set(key, pending);
    }
    return pending;
  }

  private resolveScript(script: FontScript): Promise<ScriptResolution> {
    let pending = this.scriptIds.get(script);
    if (!pending) {
      const ids = this.scriptIds;
      const results = this.scriptResults;
      pending = (async () => {
        let id: number | null = null;
        let retryable = false;
        const bundled = await this.bundled();
        if (bundled === undefined && typeof this.bundledSource === 'function') retryable = true;
        let loader: (() => Promise<ArrayBuffer>) | undefined;
        try {
          loader = bundled?.resolveScriptFallback?.(script, false, false);
        } catch (error) {
          retryable = true;
          this.warnFailure(
            `[fontRegistry] script-fallback resolver for "${script}" failed; falling back: ` +
              `${error instanceof Error ? error.message : String(error)}`
          );
        }
        if (loader) {
          try {
            id = await this.registerFaceLoader(loader);
          } catch (error) {
            retryable = true;
            this.failed(`[fontRegistry] script-fallback face for "${script}" failed to load or register; falling back`, error);
          }
        }
        const resolution = { id, retryable };
        results.set(script, resolution);
        return resolution;
      })();
      evictFailedPromise(ids, script, pending, (resolution) => resolution.retryable);
      ids.set(script, pending);
    }
    return pending;
  }
}
