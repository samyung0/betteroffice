import { FontSliceUnavailableError, type ProgressiveFontSource } from '../../yrs/residentProgressiveFonts';
import type { CompatibilityFlags } from '../../docx/settingsParser';
import type { FontDecoder, SliceLocation } from '@betteroffice/fonts';
import { baseLatinAlias, baseLatinCandidate, charactersFromBits, symbolFamily, type FontCoverage } from './fontCoverage';
import {
  excludeHostedFontFamilies,
  hostedFontFamilies,
  IMPLICIT_FONT_COVERAGE,
  isImplicitFontProvider,
  resolveDefaultFontProvider,
  baseLatinFontState,
} from './defaultFontProvider';
import {
  TextMeasureFontRegistry,
  type BundledFontProvider,
  type EmbeddedFaceInput,
  type FontScript,
  type ProgressiveFontFace,
  type SlicedFontFace,
} from './fontRegistry';

export interface RustTextEngine {
  registerFont(bytes: Uint8Array, face?: { family: string; bold: boolean; italic: boolean }): number;
  /**
   * Measurement view of `id` carrying the vertical metrics and advance pitch
   * Word measures `family` with; `id` when the engine knows none. Optional so
   * a host that predates it still measures, with the substitute's own
   * metrics.
   */
  registerSubstituteFont?(id: number, family: string): number;
  /**
   * `id` was registered from a sliced face's skeleton: the engine
   * installs the face's outlines as its pages need them, and reports a face
   * whose outlines fail to load to `onError`. `source` names its files, for
   * an engine worker to open the same face.
   */
  registerSlicedFont?(
    id: number,
    face: ProgressiveFontFace,
    source: ProgressiveFontSource,
    onError: (error: Error) => void,
  ): void;
  clearFonts(): void;
  /**
   * Reports a sliced face that failed to open (its index, skeleton or the
   * fonts module) by the rule its slices' failures follow; without it the
   * failure reaches `onError` as it is.
   */
  reportSlicedFaceError?(error: Error, onError: (error: Error) => void): void;
  /** Calls `listener` when the host asks for failed font slices again; returns an unsubscribe. */
  onFontSlicesRetry?(listener: () => void): () => void;
}

export interface ResidentFontRequirement {
  key: string;
  family: string;
  bold: boolean;
  italic: boolean;
  scripts?: FontScript[];
  characters?: string[];
  characterBits?: string;
  /** Some run paints East Asian characters along this family's chain. */
  eastAsian?: boolean;
}

/** @internal */
export function fontRequirementCharacters(requirement: ResidentFontRequirement): string[] {
  return [...new Set([...(requirement.characters ?? []), ...charactersFromBits(requirement.characterBits)])];
}

function mergeRequirements(requirements: ResidentFontRequirement[]): ResidentFontRequirement[] {
  const merged = new Map<string, ResidentFontRequirement>();
  for (const requirement of requirements) {
    const characters = fontRequirementCharacters(requirement);
    const known = merged.get(requirement.key);
    merged.set(requirement.key, known ? {
      ...known,
      scripts: [...new Set([...(known.scripts ?? []), ...(requirement.scripts ?? [])])],
      characters: [...new Set([...(known.characters ?? []), ...characters])],
      ...(requirement.eastAsian && { eastAsian: true }),
    } : { ...requirement, characters });
  }
  return [...merged.values()];
}

export interface ResidentMeasurementConfig {
  fontChains: Record<string, number[]>;
  defaults: { fontSize: number; fontFamily: string };
  compat: {
    compatibilityMode?: number;
    noLeading: boolean;
    doNotExpandShiftReturn: boolean;
  };
  authoritativeShaping: true;
}

export interface RustMeasureSource {
  setEmbeddedFaces(faces: EmbeddedFaceInput[], appFaces?: EmbeddedFaceInput[]): void;
  setCompat(flags: CompatibilityFlags | undefined): void;
  prepareFontRequirements(
    requirements: ResidentFontRequirement[],
  ): Promise<boolean>;
  measurementConfigForRequirements(
    requirements: ResidentFontRequirement[],
  ): ResidentMeasurementConfig | undefined;
  /** @internal Whether later requirements need no face the prepared chains lack (characters the base Latin tier settled count as covered); holds no font bytes, so it outlives `clear()`. */
  coverageCheck(
    requirements: ResidentFontRequirement[],
  ): (next: ResidentFontRequirement[]) => boolean;
  clear(): void;
  /** The sliced faces registered from their skeleton, by font id and measurement view id. */
  progressiveFaces(): ReadonlyMap<number, SlicedFontFace>;
  /**
   * Starts the core chunk of each prepared sliced face that East Asian text in
   * `requirements` paints with first (those flagged `eastAsian`). View calls
   * it as layout starts (Q758 B, Q769 A); a face that only fills characters an
   * earlier face lacks loads its core when the pages draw. Resolves whether
   * it started any.
   */
  startEastAsianCores(requirements: ResidentFontRequirement[]): Promise<boolean>;
}

let enginePromise: Promise<RustTextEngine> | null = null;

export function getRustTextEngine(): Promise<RustTextEngine> {
  enginePromise ??= import('../wasm/index').then(async (module) => {
    await module.preloadLayoutWasm();
    return {
      registerFont: module.registerMeasureFont,
      registerSubstituteFont: module.registerSubstituteMeasureFont,
      clearFonts: module.clearMeasureFonts,
    };
  });
  return enginePromise;
}

/** Builds a source with an explicit or default font provider. */
export function createRustMeasureSource(options: {
  engine: RustTextEngine;
  bundled?: BundledFontProvider;
  runtime?: { decoder?: FontDecoder };
  /**
   * Register sliced faces (the CJK faces whose loaders carry `slices`) from
   * their skeleton (every table but the outlines); their outlines load as glyphs need them.
   */
  progressive?: boolean;
  /** Hears every face that fails to load, sliced faces and the base Latin backstop included. */
  onError?: (error: Error) => void;
}): RustMeasureSource {
  const engine = options.engine;
  const hosted = hostedFontFamilies(options.bundled);
  let implicit = false;
  let generation = 0;
  let appFamilies = new Set<string>();
  const hostedCoverageUnknown = (family: string) => implicit &&
    hosted?.has(family.trim().toLowerCase()) === true && !appFamilies.has(family.trim().toLowerCase());
  // The files of each sliced face, which an engine worker opens the face from.
  const sources = new WeakMap<ProgressiveFontFace, ProgressiveFontSource>();
  const registry = new TextMeasureFontRegistry(
    {
      registerFont: (bytes, face) => face ? engine.registerFont(bytes, face) : engine.registerFont(bytes),
      ...(engine.registerSubstituteFont && {
        registerSubstituteFont: (id, family) =>
          engine.registerSubstituteFont!(id, family),
      }),
      ...(engine.registerSlicedFont && {
        registerSlicedFont: (id, face) => engine.registerSlicedFont!(id, face, sources.get(face)!,
          options.onError ?? ((error) => console.warn('[rustMeasureSource] font outlines failed to load', error))),
      }),
    },
    { baseLatin: () => import('./baseLatinFaces').then(tier => tier.resolveBaseLatinProvider(baseLatinFontState(), options.runtime)), bundled: (hosted ? undefined : options.bundled) ?? (async () => {
      const current = generation;
      const provider = excludeHostedFontFamilies(await resolveDefaultFontProvider(options.runtime), hosted);
      if (current === generation) implicit = isImplicitFontProvider(provider);
      return provider;
    }), progressive: options.progressive
      ? async (slices) => {
        const location = slices as SliceLocation;
        // The index's URL resolves beside the module that opens the face, so
        // its request goes out with the other faces'.
        void location.url(location.ref.index).catch(() => {});
        const { openSlicedFace, sliceSourceOf } = await import('@betteroffice/fonts/progressive').catch((error: unknown) => {
          // Offline, the module fails the face as its index would (Q763 A).
          throw new FontSliceUnavailableError(location.ref.index, 'network', undefined, { cause: error });
        });
        const face = await openSlicedFace(location);
        if (engine.registerSlicedFont) {
          const source = await sliceSourceOf(face);
          if (!source) throw new Error('A sliced face opened without file URLs cannot load its outlines in a worker');
          sources.set(face, source);
        }
        return face;
      }
      : undefined,
    onError: options.onError && ((error) => {
      // A sliced face that failed to open fails as its slices do (Q755 B).
      const sliced = error instanceof FontSliceUnavailableError || error.name === 'FontSliceError';
      if (sliced && engine.reportSlicedFaceError) engine.reportSlicedFaceError(error, options.onError!);
      else options.onError!(error);
    }) },
  );
  // Back online, faces that failed to open are opened again by the next layout.
  engine.onFontSlicesRetry?.(() => registry.retryFailedSlices());
  let compat: CompatibilityFlags | undefined;

  return {
    progressiveFaces: () => registry.progressiveFaces(),

    async startEastAsianCores(requirements): Promise<boolean> {
      const eastAsian = mergeRequirements(requirements).filter((requirement) => requirement.eastAsian);
      const faces = await registry.eastAsianFaces(eastAsian);
      for (const face of faces) face.prefetchCore();
      return faces.size > 0;
    },

    setEmbeddedFaces(faces, appFaces): void {
      appFamilies = new Set((appFaces ?? []).map(face => face.family.trim().toLowerCase()));
      registry.setEmbeddedFaces(faces, appFaces);
    },

    setCompat(flags): void {
      compat = flags;
    },

    async prepareFontRequirements(requirements): Promise<boolean> {
      requirements = mergeRequirements(requirements);
      const settled = await Promise.all(
        requirements.flatMap((requirement) => [
          (async () => {
            try {
              const cached = registry.getCachedFontIdChain(
                requirement.family,
                requirement.bold,
                requirement.italic,
                false,
              );
              if (cached !== undefined) return false;
              await registry.getFontIdChain(
                requirement.family,
                requirement.bold,
                requirement.italic,
              );
              return true;
            } catch {
              return false;
            }
          })(),
          ...Array.from(new Set(requirement.scripts ?? []), async (script) => {
            try {
              if (hosted?.has(requirement.family.trim().toLowerCase())) {
                await registry.getFontIdChain(requirement.family, requirement.bold, requirement.italic);
                if (implicit) return false;
              }
              if (
                registry.getCachedScriptFallbackIds([script], false) !==
                undefined
              )
                return false;
              await registry.getScriptFallbackIds([script]);
              return true;
            } catch {
              return false;
            }
          }),
        ]),
      );
      const backstops = await Promise.all(requirements.map(async (requirement) => {
        if (hostedCoverageUnknown(requirement.family)) return false;
        const earlier = [
          ...(registry.getCachedFontIdChain(requirement.family, requirement.bold, requirement.italic, false) ?? []),
          ...(implicit && hosted?.has(requirement.family.trim().toLowerCase()) ? [] :
            registry.getCachedScriptFallbackIds(requirement.scripts ?? [], false) ?? []),
        ];
        return registry.prepareBaseLatin(requirement.family, requirement.bold, requirement.italic,
          requirement.characters ?? [], earlier, requirement.key);
      }));
      return settled.some(Boolean) || backstops.some(Boolean);
    },

    measurementConfigForRequirements(
      requirements,
    ): ResidentMeasurementConfig | undefined {
      requirements = mergeRequirements(requirements);
      const settled = requirements.map((requirement) => ({
        requirement,
        familyIds: registry.getCachedFontIdChain(
          requirement.family,
          requirement.bold,
          requirement.italic,
          false,
        ),
        scriptIds: implicit && hosted?.has(requirement.family.trim().toLowerCase()) ? [] : registry.getCachedScriptFallbackIds(
          requirement.scripts ?? [],
          false,
        ),
      }));
      if (
        settled.some(
          ({ requirement, familyIds, scriptIds }) =>
            familyIds === undefined || scriptIds === undefined ||
            (!hostedCoverageUnknown(requirement.family) &&
              registry.getCachedBaseLatinIds(requirement.family, requirement.bold, requirement.italic,
                requirement.characters, [...familyIds, ...scriptIds], requirement.key) === undefined),
        )
      )
        return undefined;
      const fontChains: Record<string, number[]> = {};
      for (const { requirement } of settled) {
        const familyIds = registry.getCachedFontIdChain(
          requirement.family,
          requirement.bold,
          requirement.italic,
        );
        const scriptIds = implicit && hosted?.has(requirement.family.trim().toLowerCase()) ? [] : registry.getCachedScriptFallbackIds(
          requirement.scripts ?? [],
        );
        if (familyIds === undefined || scriptIds === undefined)
          return undefined;
        const backstopIds = hostedCoverageUnknown(requirement.family) ? [] :
          registry.getCachedBaseLatinIds(requirement.family, requirement.bold, requirement.italic,
            requirement.characters, [...familyIds, ...scriptIds], requirement.key);
        if (backstopIds === undefined) return undefined;
        if (implicit && familyIds.length === 0 && backstopIds.length === 0) continue;
        const chain = Array.from(familyIds);
        for (const id of scriptIds) if (!chain.includes(id)) chain.push(id);
        for (const id of backstopIds) {
          if (!chain.includes(id)) chain.push(id);
        }
        if (chain.length > 0) fontChains[requirement.key] = chain;
      }
      if (implicit) {
        const ids = [...new Set(Object.values(fontChains).flat())];
        if (ids.length > 0) fontChains[IMPLICIT_FONT_COVERAGE] = ids;
      }
      return {
        fontChains,
        defaults: { fontSize: 11, fontFamily: 'Calibri' },
        compat: {
          compatibilityMode: compat?.compatibilityMode ?? 12,
          noLeading: compat?.noLeading ?? false,
          doNotExpandShiftReturn: compat?.doNotExpandShiftReturn ?? false,
        },
        authoritativeShaping: true,
      };
    },

    coverageCheck(requirements) {
      const chains = new Map<string, FontCoverage[] | null>();
      const settled = new Map<string, Set<string>>();
      for (const requirement of mergeRequirements(requirements)) {
        settled.set(requirement.key, registry.settledBaseLatinCharacters(requirement.key));
        const familyIds = registry.getCachedFontIdChain(requirement.family, requirement.bold, requirement.italic, false);
        const scriptIds = implicit && hosted?.has(requirement.family.trim().toLowerCase()) ? [] :
          registry.getCachedScriptFallbackIds(requirement.scripts ?? [], false);
        if (familyIds === undefined || scriptIds === undefined) continue;
        const earlier = [...familyIds, ...scriptIds];
        const coverage = earlier.map(id => registry.coverageOf(id));
        if (hostedCoverageUnknown(requirement.family) || earlier.length === 0 || coverage.some(covers => !covers)) {
          chains.set(requirement.key, null);
          continue;
        }
        const backstopIds = registry.getCachedBaseLatinIds(requirement.family, requirement.bold, requirement.italic,
          requirement.characters, earlier, requirement.key);
        if (backstopIds === undefined) continue;
        chains.set(requirement.key, [...coverage, ...backstopIds.map(id => registry.coverageOf(id))]
          .filter((covers): covers is FontCoverage => covers !== undefined));
      }
      return next => mergeRequirements(next).every(requirement => {
        const coverage = chains.get(requirement.key);
        if (coverage === undefined) return false;
        const aliases = symbolFamily(requirement.family);
        const known = settled.get(requirement.key);
        return coverage === null || (requirement.characters ?? []).every(character => known?.has(character) ||
          !baseLatinCandidate(character, false) || (baseLatinAlias(character) && !aliases) ||
          coverage.some(covers => covers(character)));
      });
    },

    clear(): void {
      generation++;
      implicit = false;
      registry.clear();
    },
  };
}
