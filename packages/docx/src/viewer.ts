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

/**
 * Measures with the faces `configureDefaultFonts` provides; without them text
 * gets synthetic metrics and paragraphs do not wrap.
 */
export async function openDocumentViewer(bytes: Uint8Array): Promise<DocxViewerHandle> {
  await initWasm();
  const module = await loadModule();
  const document = module.openViewDocument(bytes);
  try {
    const requestJson = document.layoutRequestJson();
    const request = JSON.parse(requestJson) as {
      regions: { settings?: { compatibilityFlags?: CompatibilityFlags } };
    };
    const requirements = JSON.parse(
      document.fontRequirementsJson(requestJson)
    ) as ResidentFontRequirement[];
    const source = (measureSource ??= createRustMeasureSource({ engine: module.viewTextEngine }));
    await source.prepareFontRequirements(requirements);
    source.setCompat(request.regions.settings?.compatibilityFlags);
    const measurement = source.measurementConfigForRequirements(requirements);
    if (!measurement) throw new Error('DOCX viewer fonts did not load');
    document.layout(JSON.stringify({ ...request, measurement }));
  } catch (error) {
    document.free();
    throw error;
  }
  let disposed = false;
  return {
    displayList(pageGap) {
      if (disposed) throw new Error('DOCX viewer handle is disposed');
      return JSON.parse(document.displayListJson(pageGap)) as DisplayList;
    },
    dispose() {
      if (disposed) return;
      disposed = true;
      document.free();
    },
  };
}

export function analyzeOpenDocument(handle: DocxViewerHandle): DocxViewerAnalysis {
  return {
    format: 'docx',
    pageCount: handle.displayList().pages.length,
  };
}
