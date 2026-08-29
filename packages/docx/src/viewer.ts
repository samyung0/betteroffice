import type { DisplayList } from './layout/render';

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

export async function openDocumentViewer(bytes: Uint8Array): Promise<DocxViewerHandle> {
  await initWasm();
  const module = await loadModule();
  const document = module.openViewDocument(bytes);
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
