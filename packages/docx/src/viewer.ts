import type { CompatibilityFlags } from './docx/settingsParser';
import type { ProgressiveFontFace, SliceLocation, SlicedFontFace } from './layout/measure/fontRegistry';
import {
  createRustMeasureSource,
  type ResidentFontRequirement,
} from './layout/measure/rustMeasureSource';
import type { DisplayList } from './layout/render';
import { withFontFamilies } from './layout/render/canvasBackend';
import { drawnByFont, forEachTextPrimitive, slicedFontFamilies, textSlicedFaces } from './yrs/residentProgressiveFonts';

import { preloadDefaultFontDecoder } from './layout/measure/defaultFontProvider';

export { configureDefaultFonts } from './layout/measure/defaultFontProvider';

/** A sliced face the layout measured with, and its bytes with every outline the pages draw. */
export interface DocxViewerSlicedFace {
  /** The `slices` of the font loader that located the face. */
  slices: SliceLocation;
  /**
   * The private family the display list's list markers and watermarks name
   * the face by: a painter registers `bytes` under it (unique to this open).
   */
  family: string;
  bytes: ArrayBuffer;
  /**
   * The decoded slice files behind `bytes`: an editor opening the same face
   * takes them as its slice location's `files` and decodes none of them again.
   */
  files: Map<string, ArrayBuffer>;
}

export interface DocxViewerHandle {
  displayList(pageGap?: number): DisplayList;
  /**
   * The sliced (CJK) faces the layout measured with. A painter registers
   * these bytes as the face: they hold the outline of every glyph the pages
   * draw, and the same tables the layout measured.
   */
  slicedFaces(): DocxViewerSlicedFace[];
  dispose(): void;
}

/** A reading layout: the text reflowed into one column `width` CSS px wide,
 * `padding` from each side, with no page furniture. */
export interface DocxReadingLayout {
  width: number;
  padding: number;
}

export interface DocxViewerOptions {
  reading?: DocxReadingLayout;
}

export interface DocxViewerAnalysis {
  format: 'docx';
  pageCount: number;
}

let modulePromise: Promise<typeof import('./wasm/viewer')> | null = null;
let initialization: Promise<void> | null = null;

/** Packed faces decode with the fonts package's own decoder, never the edit engine's. */
const viewerRuntime = {};

function loadModule(): Promise<typeof import('./wasm/viewer')> {
  modulePromise ??= import('./wasm/viewer');
  return modulePromise;
}

/** Loads the view engine, and beside it the fonts package's decoder its packed faces decode with. */
export function initWasm(
  input?: Parameters<Awaited<ReturnType<typeof loadModule>>['preloadViewWasm']>[0]
): Promise<void> {
  if (!initialization) preloadDefaultFontDecoder();
  initialization ??= loadModule().then((module) => module.preloadViewWasm(input));
  return initialization;
}

/** Pages per `displayPagesJson` call: the list is never built in one piece. */
const PAGE_BATCH = 8;

const TWIPS_PER_PX = 15;
/** A reading layout's page: tall, so few lines meet a page join, but well
 * inside what a canvas can hold. */
const READING_PAGE_HEIGHT_PX = 2400;
/** The body's reading size in pt: 16px at 100%. */
const READING_TEXT_SIZE = 12;

interface ViewRequest {
  /** The document's body size in pt. */
  bodyFontSize: number;
  renderEnv: Record<string, unknown>;
  options: { pageGap: number; continuousFlow?: boolean };
  regions: {
    sections: { properties?: Record<string, unknown> }[];
    settings?: { compatibilityFlags?: CompatibilityFlags };
  };
}

/** Section properties a reading layout has no use for: orientation, the
 * authored columns, headers and footers, line numbers, borders, centring. */
const READING_DROPPED_KEYS = [
  'orientation',
  'columns',
  'headerReferences',
  'footerReferences',
  'titlePg',
  'evenAndOddHeaders',
  'lineNumbers',
  'pageBorders',
  'verticalAlign',
];

/** Every section as one column of the reading width, its pages abutting and
 * no break starting one; lowering eases sizes and indents and fits tables and
 * pictures to it. */
function readingRequest(request: ViewRequest, { width, padding }: DocxReadingLayout): ViewRequest {
  const sections = request.regions.sections.map((section) => {
    const properties: Record<string, unknown> = { ...section.properties };
    for (const key of READING_DROPPED_KEYS) delete properties[key];
    return {
      ...section,
      properties: {
        ...properties,
        pageWidth: width * TWIPS_PER_PX,
        pageHeight: READING_PAGE_HEIGHT_PX * TWIPS_PER_PX,
        marginTop: 0,
        marginBottom: 0,
        marginLeft: padding * TWIPS_PER_PX,
        marginRight: padding * TWIPS_PER_PX,
        headerDistance: 0,
        footerDistance: 0,
        gutter: 0,
        columnCount: 1,
      },
    };
  });
  return {
    ...request,
    options: { ...request.options, pageGap: 0, continuousFlow: true },
    regions: { ...request.regions, sections },
    renderEnv: {
      ...request.renderEnv,
      reading: {
        bodySize: request.bodyFontSize,
        textSize: READING_TEXT_SIZE,
        contentWidth: width - 2 * padding,
      },
    },
  };
}

/**
 * The faces of `faces` with the outline of every glyph `list` draws: chunks by
 * the glyph ids the layout shaped, and by the characters, whose chunks hold
 * every glyph another shaper (the painter's) can draw for them. Text the
 * painter shapes itself (list markers, watermarks) counts for the face it
 * paints with (`textSlicedFaces`, by `chains`), and its CSS font is rewritten
 * to name that face by its private family.
 */
// Opens so far: each names its sliced faces' private families apart from another open's.
let opens = 0;

async function outlineSlicedFaces(
  faces: ReadonlyMap<number, SlicedFontFace>,
  list: DisplayList,
  chains: Readonly<Record<string, number[]>>
): Promise<DocxViewerSlicedFace[]> {
  if (faces.size === 0) return [];
  const drawn = new Map<number, { glyphs: Set<number>; characters: Set<string> }>();
  for (const page of list.pages) drawnByFont(page, drawn);
  const wanted = new Map<ProgressiveFontFace, { slices: SliceLocation; glyphs: Set<number>; characters: Set<string> }>();
  for (const [id, { face, slices }] of faces) {
    let entry = wanted.get(face);
    if (!entry) wanted.set(face, entry = { slices, glyphs: new Set(), characters: new Set() });
    const used = drawn.get(id);
    for (const glyph of used?.glyphs ?? []) entry.glyphs.add(glyph);
    for (const character of used?.characters ?? []) entry.characters.add(character);
  }
  opens += 1;
  const names = new Map([...wanted.keys()].map((face, index) => [face, `BetterOffice Sliced ${opens}.${index}`]));
  const families = new Map<string, string | string[]>();
  for (const page of list.pages) {
    const painted = textSlicedFaces(page, chains, (id) => faces.has(id));
    for (const [id, { characters }] of painted) {
      const entry = wanted.get(faces.get(id)!.face)!;
      for (const character of characters) entry.characters.add(character);
    }
    slicedFontFamilies(painted, (id) => names.get(faces.get(id)!.face), families);
  }
  if (families.size > 0) {
    for (const page of list.pages) {
      forEachTextPrimitive(page, (text) => {
        text.font = withFontFamilies(text.font, families);
        const leader = text.leaderGlyphs;
        if (typeof leader?.font === 'string') leader.font = withFontFamilies(leader.font, families);
      });
    }
  }
  await Promise.all([...wanted].map(([face, { glyphs, characters }]) =>
    Promise.all([face.ensureGlyphs(glyphs), face.ensureText([...characters].join(''))])));
  return [...wanted].map(([face, { slices }]) => ({ slices, family: names.get(face)!, bytes: face.bytes(), files: face.files() }));
}

/**
 * Measures with the faces `configureDefaultFonts` provides; without them text
 * gets synthetic metrics and paragraphs do not wrap. A face that fails to
 * load, an outline chunk of a sliced face included, fails the open.
 */
export async function openDocumentViewer(
  bytes: Uint8Array,
  options: DocxViewerOptions = {}
): Promise<DocxViewerHandle> {
  await initWasm();
  const module = await loadModule();
  const document = module.openViewDocument(bytes);
  let pageGap: number;
  let list: DisplayList | null;
  let fontChains: Record<string, number[]> = {};
  // A registry per open: its failures and sliced faces are this open's. The
  // engine answers bytes it already holds with their first id.
  const failures: Error[] = [];
  const source = createRustMeasureSource({
    engine: module.viewTextEngine,
    runtime: viewerRuntime,
    progressive: true,
    onError: (error) => failures.push(error),
  });
  try {
    const parsed = JSON.parse(document.layoutRequestJson()) as ViewRequest;
    const request = options.reading ? readingRequest(parsed, options.reading) : parsed;
    const requestJson = JSON.stringify(request);
    pageGap = request.options.pageGap;
    const requirements = JSON.parse(
      document.fontRequirementsJson(requestJson)
    ) as ResidentFontRequirement[];
    await source.prepareFontRequirements(requirements);
    if (failures.length > 0) throw failures[0];
    source.setCompat(request.regions.settings?.compatibilityFlags);
    const measurement = source.measurementConfigForRequirements(requirements);
    if (!measurement) throw new Error('DOCX viewer fonts did not load');
    fontChains = measurement.fontChains;
    // The face that paints East Asian text first loads its core while layout
    // runs, no sooner, so it never shares the link with the engine or the
    // skeletons layout waits for (Q769 A).
    await source.startEastAsianCores(requirements);
    const pageCount = document.layout(JSON.stringify({ ...request, measurement }));
    const displayList = JSON.parse(
      document.displayPagesJson(0, Math.min(PAGE_BATCH, pageCount))
    ) as DisplayList;
    for (let start = PAGE_BATCH; start < pageCount; start += PAGE_BATCH) {
      const batch = JSON.parse(
        document.displayPagesJson(start, Math.min(start + PAGE_BATCH, pageCount))
      ) as DisplayList;
      displayList.pages.push(...batch.pages);
    }
    list = displayList;
  } catch (error) {
    // A trap (out of memory) leaves the document borrowed, so free() throws
    // too; the trap's error is the one to report.
    try {
      document.free();
    } catch {}
    throw error;
  }
  // Every page is built: the engine goes now, not on dispose.
  document.free();
  let sliced = await outlineSlicedFaces(source.progressiveFaces(), list!, fontChains);
  return {
    displayList(gap) {
      if (!list) throw new Error('DOCX viewer handle is disposed');
      if (gap !== undefined && gap !== pageGap) {
        throw new Error('DOCX view page gap is fixed when the document opens');
      }
      return list;
    },
    slicedFaces() {
      if (!list) throw new Error('DOCX viewer handle is disposed');
      return sliced;
    },
    dispose() {
      list = null;
      sliced = [];
    },
  };
}

export function analyzeOpenDocument(handle: DocxViewerHandle): DocxViewerAnalysis {
  return {
    format: 'docx',
    pageCount: handle.displayList().pages.length,
  };
}
