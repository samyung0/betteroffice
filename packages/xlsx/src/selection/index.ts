/**
 * Pure selection seam: the anchor/focus model, its transforms, and the keyboard
 * reducer. DOM-free by construction (see the seam-purity test).
 */

export type {
  CellAddr,
  Selection,
  CellRange,
  Direction,
  SelectionLimits,
  MoveOpts,
  KeyInput,
  SelectionAction,
  SelectionData,
  SelectionHidden,
  SelectionMerges,
} from './types';

export {
  activeCell,
  cellAddress,
  columnLetter,
  selectionAt,
  normalizeRange,
  rangeContains,
  cellCount,
  extendTo,
  mergeAnchor,
  moveFocus,
} from './model';

export { selectAll, selectionDataRequest, selectionKeyReducer } from './keyboard';
