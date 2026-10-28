import type { CompatibilityFlags } from './docx/settingsParser';
import {
  createRustMeasureSource,
  type ResidentFontRequirement,
  type RustMeasureSource,
} from './layout/measure/rustMeasureSource';
import type { DisplayList } from './layout/render';

export { configureDefaultFonts } from './layout/measure/defaultFontProvider';

export interface DocxViewerHandle {
  displayList(pageGap?: number): DisplayList;
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
let measureSource: RustMeasureSource | null = null;

function loadModule(): Promise<typeof import('./wasm/viewer')> {
  modulePromise ??= import('./wasm/viewer');
  return modulePromise;
}

export function initWasm(
  input?: Parameters<Awaited<ReturnType<typeof loadModule>>['preloadViewWasm']>[0]
): Promise<void> {
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
 * Measures with the faces `configureDefaultFonts` provides; without them text
 * gets synthetic metrics and paragraphs do not wrap.
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
  try {
    const parsed = JSON.parse(document.layoutRequestJson()) as ViewRequest;
    const request = options.reading ? readingRequest(parsed, options.reading) : parsed;
    const requestJson = JSON.stringify(request);
    pageGap = request.options.pageGap;
    const requirements = JSON.parse(
      document.fontRequirementsJson(requestJson)
    ) as ResidentFontRequirement[];
    const source = (measureSource ??= createRustMeasureSource({ engine: module.viewTextEngine }));
    await source.prepareFontRequirements(requirements);
    source.setCompat(request.regions.settings?.compatibilityFlags);
    const measurement = source.measurementConfigForRequirements(requirements);
    if (!measurement) throw new Error('DOCX viewer fonts did not load');
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
  return {
    displayList(gap) {
      if (!list) throw new Error('DOCX viewer handle is disposed');
      if (gap !== undefined && gap !== pageGap) {
        throw new Error('DOCX view page gap is fixed when the document opens');
      }
      return list;
    },
    dispose() {
      list = null;
    },
  };
}

export function analyzeOpenDocument(handle: DocxViewerHandle): DocxViewerAnalysis {
  return {
    format: 'docx',
    pageCount: handle.displayList().pages.length,
  };
}
