/**
 * Pure viewport seam: virtualization math and its types. DOM-free by
 * construction (see the seam-purity test).
 */

export type { TrackOffsets, ViewportState, VisibleRange, VisibleCells } from './types';
export type { ScrollBox } from './math';
export {
  trackCount,
  totalExtent,
  visibleRange,
  visibleCells,
  clampScroll,
  uniformOffsets,
  zoomedViewport,
} from './math';
