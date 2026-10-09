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

/**
 * Measures with the faces `configureDefaultFonts` provides; without them text
 * gets synthetic metrics and paragraphs do not wrap.
 */
export async function openDocumentViewer(bytes: Uint8Array): Promise<DocxViewerHandle> {
  await initWasm();
  const module = await loadModule();
  const document = module.openViewDocument(bytes);
  let pageGap: number;
  let list: DisplayList | null;
  try {
    const requestJson = document.layoutRequestJson();
    const request = JSON.parse(requestJson) as {
      options: { pageGap: number };
      regions: { settings?: { compatibilityFlags?: CompatibilityFlags } };
    };
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
