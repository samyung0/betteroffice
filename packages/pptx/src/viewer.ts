export { paintSlide, sizeCanvasForSlide } from './render/canvas';
export type { CanvasImageResolver, PaintSlideOptions, SlideCanvasLike } from './render/canvas';
export {
  analyzeOpenPresentation,
  analyzePresentation,
  initWasm,
  isWasmAvailable,
  openPresentation,
  wasmVersion,
} from './wasm/viewer';
export type {
  PresentationAnalysis,
  PresentationViewerHandle,
  WasmInitInput,
} from './wasm/viewer';
export type {
  DeckSnapshot,
  HitTestResult,
  PptxFontFace,
  SlideDisplayList,
  SlidePrimitive,
} from './types';
