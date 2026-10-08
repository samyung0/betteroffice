export { paintSlide, sizeCanvasForSlide } from './render/canvas';
export { installEmbeddedFonts, removeFontFaces } from './render/fonts';
export type { EmbeddedFontSource } from './render/fonts';
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
  PptxEmbeddedFontFace,
  PptxFontFace,
  SlideDisplayList,
  SlidePrimitive,
} from './types';
