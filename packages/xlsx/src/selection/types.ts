/**
 * Pure selection types. Plain data only — no DOM, no framework — so the model
 * is golden-testable. The chrome layer interprets the
 * declarative {@link SelectionAction}s the reducer emits.
 */

/**
 * A single cell coordinate. `row`/`col` are zero-based sheet indices, not px.
 */
export interface CellAddr {
  row: number;
  col: number;
}

/**
 * An active selection as an anchor/focus pair. `anchor` is the fixed corner
 * where the selection started and the active cell (the name box, the formula
 * bar and typing go there, as in Excel); `focus` is the moving corner a drag,
 * Shift+click or Shift+arrow moves. A collapsed selection has `anchor` equal to
 * `focus`.
 */
export interface Selection {
  anchor: CellAddr;
  focus: CellAddr;
  /** The active cell when it is not the anchor: select all keeps it where it was. */
  active?: CellAddr;
}

/**
 * A normalized, inclusive rectangle of cells. `top <= bottom`, `left <= right`.
 * Derived from a {@link Selection} by {@link normalizeRange}; the natural shape
 * for hit-testing, overlay rects, and range reads.
 */
export interface CellRange {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

/**
 * A cardinal move direction for {@link moveFocus} and arrow-key handling.
 */
export type Direction = 'up' | 'down' | 'left' | 'right';

/**
 * Grid bounds: a move stays inside `[0, rows)` × `[0, cols)` (the whole sheet in
 * the editor, as in Excel), Ctrl/Cmd+arrows, End, Ctrl/Cmd+End and select all
 * stop at the used range `[0, usedRows)` × `[0, usedCols)`, and `rowsPerPage`
 * is the viewport height in whole rows for PageUp/PageDown steps.
 */
export interface SelectionLimits {
  rows: number;
  cols: number;
  usedRows: number;
  usedCols: number;
  rowsPerPage: number;
}

/**
 * Options for {@link moveFocus}: whether to extend (keep the anchor) or collapse
 * to a single cell, how many cells to step, and the grid bounds to clamp into.
 */
export interface MoveOpts {
  extend?: boolean;
  step?: number;
  limits: SelectionLimits;
  /** Hidden rows and columns a move steps over. */
  hidden?: SelectionHidden;
  /** Merged ranges a move treats as one cell. */
  merges?: SelectionMerges;
}

/**
 * What the engine answers for one keyboard jump (`Workbook::selection_data`):
 * the data edge a Ctrl/Cmd+arrow reaches, the row's last data cell for End or
 * the used range's last cell for Ctrl/Cmd+End (`null` when there is none),
 * and the merged range the jump lands in.
 */
export interface SelectionData {
  merge?: CellRange | null;
  edge?: CellAddr | null;
  usedRange?: CellRange | null;
}

/** The sheet's hidden rows and columns, which keyboard moves step over. */
export interface SelectionHidden {
  isRowHidden(row: number): boolean;
  isColHidden(col: number): boolean;
}

/** Merged ranges, so keyboard moves treat each one as a single cell. */
export interface SelectionMerges {
  mergeAt(row: number, col: number): CellRange | null;
}

/**
 * A keyboard event reduced to the fields the reducer reads. `metaKey`/`ctrlKey`
 * are both accepted so a caller need not normalize mac cmd vs win/linux ctrl.
 */
export interface KeyInput {
  key: string;
  shiftKey?: boolean;
  metaKey?: boolean;
  ctrlKey?: boolean;
  altKey?: boolean;
}

/**
 * The declarative outcome of a keystroke. The chrome interprets it: `move`
 * replaces the selection, `startEdit` opens the in-cell editor (seeding it with
 * `initialInput` when a printable key triggered it), `clear` deletes the
 * selected cells' contents, and `none` means the key was not handled here.
 */
export type SelectionAction =
  | { type: 'move'; selection: Selection }
  | { type: 'startEdit'; initialInput?: string }
  | { type: 'clear' }
  | { type: 'none' };
