import initWasmModule, {
  PptxViewDocument,
  PptxViewRenderer,
} from './generated/viewer/pptx_view_wasm.js';
import type { InitInput } from './generated/viewer/pptx_view_wasm.js';
import type {
  DeckSnapshot,
  HitTestResult,
  PptxFontFace,
  SlideDisplayList,
} from '../types';

export type WasmInitInput = InitInput | Promise<InitInput>;

export interface PresentationViewerHandle {
  snapshot(): DeckSnapshot;
  registerFont(face: PptxFontFace): number;
  layoutSlide(slideIndex: number): SlideDisplayList;
  hitTest(x: number, y: number): HitTestResult | null;
  mediaBytes(partPath: string): Uint8Array;
  dispose(): void;
}

export interface PresentationAnalysis {
  format: 'pptx';
  slideCount: number;
  widthEmu: number;
  heightEmu: number;
  textCharacterCount: number;
}

let initialized = false;
let initialization: Promise<void> | undefined;

export function initWasm(
  input: WasmInitInput = new URL('./generated/viewer/pptx_view_wasm_bg.wasm', import.meta.url)
): Promise<void> {
  if (initialized) return Promise.resolve();
  if (initialization) return initialization;
  initialization = initWasmModule({ module_or_path: input }).then(
    () => {
      initialized = true;
    },
    (error: unknown) => {
      initialization = undefined;
      throw toError(error);
    }
  );
  return initialization;
}

export function isWasmAvailable(): boolean {
  return typeof WebAssembly === 'object';
}

export function openPresentation(
  bytes: Uint8Array,
  options: { fonts?: ReadonlyArray<PptxFontFace> } = {}
): PresentationViewerHandle {
  if (!initialized) throw new Error('pptx viewer wasm is not initialized; call initWasm() first');
  const document = call(() => PptxViewDocument.open(bytes));
  const renderer = createRenderer(document, options.fonts ?? []);
  let disposed = false;

  const read = <T,>(operation: () => T): T => {
    if (disposed) throw new Error('presentation viewer handle is disposed');
    return call(operation);
  };

  return {
    snapshot: () => parseJson(() => read(() => document.snapshotJson())),
    registerFont: (face) => read(() => registerFont(renderer, face)),
    layoutSlide: (slideIndex) =>
      parseJson(() => read(() => renderer.layoutSlideJson(document, slideIndex))),
    hitTest: (x, y) => parseJson(() => read(() => renderer.hitTestJson(x, y))),
    mediaBytes: (partPath) => read(() => document.mediaBytes(partPath)),
    dispose(): void {
      if (disposed) return;
      disposed = true;
      let failure: unknown;
      try {
        renderer.free();
      } catch (error) {
        failure = error;
      }
      try {
        document.free();
      } catch (error) {
        failure ??= error;
      }
      if (failure !== undefined) throw toError(failure);
    },
  };
}

export function analyzePresentation(bytes: Uint8Array): PresentationAnalysis {
  const presentation = openPresentation(bytes);
  try {
    return analyzeOpenPresentation(presentation);
  } finally {
    presentation.dispose();
  }
}

/** Read upload metadata from an already-open viewer without reparsing bytes. */
export function analyzeOpenPresentation(
  presentation: PresentationViewerHandle
): PresentationAnalysis {
  const snapshot = presentation.snapshot();
  let textCharacterCount = 0;
  for (const slide of snapshot.slides) {
    for (const shape of slide.shapes) {
      textCharacterCount += countShapeTextCharacters(shape);
    }
  }
  return {
    format: 'pptx',
    slideCount: snapshot.slides.length,
    widthEmu: snapshot.widthEmu,
    heightEmu: snapshot.heightEmu,
    textCharacterCount,
  };
}

export function wasmVersion(): string {
  if (!initialized) throw new Error('pptx viewer wasm is not initialized; call initWasm() first');
  return PptxViewDocument.version();
}

function registerFont(renderer: PptxViewRenderer, face: PptxFontFace): number {
  return renderer.registerFont(face.family, face.bold ?? false, face.italic ?? false, face.bytes);
}

function countShapeTextCharacters(shape: DeckSnapshot['slides'][number]['shapes'][number]): number {
  let count = 0;
  for (const story of shape.textStories) {
    for (const paragraph of story.paragraphs) {
      for (const run of paragraph.runs) count += run.text.length;
    }
  }
  for (const child of shape.children) count += countShapeTextCharacters(child);
  return count;
}

function createRenderer(
  document: PptxViewDocument,
  fonts: ReadonlyArray<PptxFontFace>
): PptxViewRenderer {
  try {
    const renderer = call(() => new PptxViewRenderer());
    try {
      for (const face of fonts) registerFont(renderer, face);
      return renderer;
    } catch (error) {
      renderer.free();
      throw error;
    }
  } catch (error) {
    document.free();
    throw toError(error);
  }
}

function parseJson<T>(operation: () => string): T {
  return JSON.parse(operation()) as T;
}

function call<T>(operation: () => T): T {
  try {
    return operation();
  } catch (error) {
    throw toError(error);
  }
}

function toError(error: unknown): Error {
  if (error instanceof Error) return error;
  return new Error(typeof error === 'string' ? error : String(error));
}
