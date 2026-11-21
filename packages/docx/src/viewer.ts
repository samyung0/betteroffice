import type { CompatibilityFlags } from './docx/settingsParser';
import type { ProgressiveFontFace, SliceLocation, SlicedFontFace } from './layout/measure/fontRegistry';
import {
  createRustMeasureSource,
  type ResidentFontRequirement,
} from './layout/measure/rustMeasureSource';
import type { DisplayList, DisplayPage } from './layout/render';
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

/** A page's size in CSS px, known from the layout before its page is built. */
export interface DocxPageSize {
  width: number;
  height: number;
}

/** Pages as they are built (`DocxViewerOptions.progressive`). */
export interface DocxViewerPages {
  /**
   * In order, each at its `pageIndex`; a page that comes again replaces the
   * one before (a later page's list markers renamed its sliced face).
   */
  pages: DisplayPage[];
  /**
   * The sliced faces whose outlines grew for these pages (every one with the
   * first pages), each holding every outline the pages so far draw.
   */
  slicedFaces: Omit<DocxViewerSlicedFace, 'files'>[];
}

export interface DocxViewerOptions {
  reading?: DocxReadingLayout;
  /**
   * Shows pages before the whole document is built: with every page's size
   * from the layout, `firstPages` says how many to build first; those, then
   * the rest `PAGE_BATCH` at a time, go to `onPages` once their sliced faces
   * hold their outlines. The handle resolves after the last.
   */
  progressive?: {
    firstPages(sizes: readonly DocxPageSize[]): number;
    onPages(batch: DocxViewerPages): void;
  };
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

// Opens so far: each names its sliced faces' private families apart from another open's.
let opens = 0;

/**
 * Readies `faces` for the pages added to it, a batch at a time: each face gets
 * the outline of every glyph the pages draw, by the glyph ids the layout
 * shaped and by the characters, whose chunks hold every glyph another shaper
 * (the painter's) can draw for them. Text the painter shapes itself (list
 * markers, watermarks) counts for the face it paints with (`textSlicedFaces`,
 * by `chains`), and its CSS font is rewritten to name that face by its private
 * family, as the families of the pages so far give it; `finish` rewrites the
 * earlier pages a later one's text renamed.
 */
function slicedOutlines(faces: ReadonlyMap<number, SlicedFontFace>, chains: Readonly<Record<string, number[]>>) {
  // Each face once, with the revision a batch last returned it at.
  const wanted = new Map<ProgressiveFontFace, { slices: SliceLocation; sent: number }>();
  for (const { face, slices } of faces.values()) if (!wanted.has(face)) wanted.set(face, { slices, sent: 0 });
  if (wanted.size > 0) opens += 1;
  const names = new Map([...wanted.keys()].map((face, index) => [face, `BetterOffice Sliced ${opens}.${index}`]));
  const families = new Map<string, string | string[]>();
  // Every rewritable font as the layout wrote it, to rewrite again under the final families.
  const written: Array<{ page: DisplayPage; node: { font: string }; font: string }> = [];
  const rewrite = (node: { font: string }, font: string) => {
    const next = withFontFamilies(font, families);
    if (next === node.font) return false;
    node.font = next;
    return true;
  };
  return {
    /** Names `pages`' fonts at once and resolves once their outlines are in. */
    async add(pages: readonly DisplayPage[]): Promise<void> {
      if (wanted.size === 0) return;
      const drawn = new Map<number, { glyphs: Set<number>; characters: Set<string> }>();
      for (const page of pages) drawnByFont(page, drawn);
      const needs = new Map([...wanted.keys()].map((face) => [face, { glyphs: new Set<number>(), characters: new Set<string>() }]));
      for (const [id, { face }] of faces) {
        const need = needs.get(face)!;
        for (const glyph of drawn.get(id)?.glyphs ?? []) need.glyphs.add(glyph);
        for (const character of drawn.get(id)?.characters ?? []) need.characters.add(character);
      }
      for (const page of pages) {
        const painted = textSlicedFaces(page, chains, (id) => faces.has(id));
        for (const [id, { characters }] of painted) {
          const need = needs.get(faces.get(id)!.face)!;
          for (const character of characters) need.characters.add(character);
        }
        slicedFontFamilies(painted, (id) => names.get(faces.get(id)!.face), families);
      }
      const from = written.length;
      for (const page of pages) {
        forEachTextPrimitive(page, (text) => {
          written.push({ page, node: text, font: text.font });
          const leader = text.leaderGlyphs;
          if (typeof leader?.font === 'string') written.push({ page, node: leader as { font: string }, font: leader.font });
        });
      }
      for (const { node, font } of written.slice(from)) rewrite(node, font);
      await Promise.all([...needs].map(([face, { glyphs, characters }]) =>
        Promise.all([face.ensureGlyphs(glyphs), face.ensureText([...characters].join(''))])));
    },
    /** The faces whose outlines grew since the last call, each with every outline loaded so far. */
    grown(): Omit<DocxViewerSlicedFace, 'files'>[] {
      return [...wanted].flatMap(([face, entry]) => {
        if (entry.sent === face.revision) return [];
        entry.sent = face.revision;
        return [{ slices: entry.slices, family: names.get(face)!, bytes: face.bytes() }];
      });
    },
    /** The pages added before the families took their final names whose fonts those names change, rewritten. */
    finish(): DisplayPage[] {
      const changed = new Set<DisplayPage>();
      for (const { page, node, font } of written) if (rewrite(node, font)) changed.add(page);
      return [...changed];
    },
    /** Every face with all its outlines so far. */
    faces(): DocxViewerSlicedFace[] {
      return [...wanted].map(([face, { slices }]) => ({ slices, family: names.get(face)!, bytes: face.bytes(), files: face.files() }));
    },
  };
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
  let list: DisplayList | null = null;
  let outlines!: ReturnType<typeof slicedOutlines>;
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
    // The face that paints East Asian text first loads its core while layout
    // runs, no sooner, so it never shares the link with the engine or the
    // skeletons layout waits for (Q769 A). A task boundary lets the requests
    // go out before layout holds the thread.
    if (await source.startEastAsianCores(requirements)) await new Promise((resolve) => setTimeout(resolve, 0));
    const pageCount = document.layout(JSON.stringify({ ...request, measurement }));
    const progressive = options.progressive;
    outlines = slicedOutlines(source.progressiveFaces(), measurement.fontChains);
    let next = PAGE_BATCH;
    if (progressive) {
      const sizes = (JSON.parse(document.pageSizesJson()) as [number, number][]).map(([width, height]) => ({ width, height }));
      next = Math.max(1, Math.floor(progressive.firstPages(sizes)));
    }
    // Batches go to `onPages` in order, each once its outlines are in. The
    // first goes before anything else is built; a later one's chunk fetches
    // overlap the next batch's build (Q804 A).
    let delivered = Promise.resolve();
    // At least one call, so a document of no pages still gets its list.
    for (let start = 0; start < pageCount || !list; start = next, next += PAGE_BATCH) {
      const batch = JSON.parse(document.displayPagesJson(start, Math.min(next, pageCount))) as DisplayList;
      if (list) list.pages.push(...batch.pages);
      else list = { ...batch, pages: [...batch.pages] };
      if (!progressive) continue;
      const loaded = outlines.add(batch.pages);
      // Reported through `delivered`, in order.
      loaded.catch(() => {});
      delivered = delivered
        .then(() => loaded)
        .then(() => progressive.onPages({ pages: batch.pages, slicedFaces: outlines.grown() }));
      // A batch whose outlines are loaded goes before the next is built; one
      // still fetching lets the next build meanwhile.
      await (start === 0 ? delivered : Promise.race([delivered, new Promise((resolve) => setTimeout(resolve, 0))]));
    }
    await delivered;
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
  if (!options.progressive) await outlines.add(list!.pages);
  else {
    const renamed = outlines.finish();
    if (renamed.length > 0) options.progressive.onPages({ pages: renamed, slicedFaces: outlines.grown() });
  }
  let sliced = outlines.faces();
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
