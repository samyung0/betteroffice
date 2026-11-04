/**
 * Selection model: pure transforms over {@link Selection} data. Every function
 * returns a fresh value and never mutates its input, so state lives with the
 * chrome and this layer stays a deterministic reducer core.
 */

import type {
  CellAddr,
  CellRange,
  Direction,
  MoveOpts,
  Selection,
  SelectionLimits,
  SelectionMerges,
} from './types';

function clamp(value: number, lo: number, hi: number): number {
  return Math.max(lo, Math.min(hi, value));
}

/** Bijective base-26 column letters: 0 is A, 25 is Z, 26 is AA. */
export function columnLetter(col: number): string {
  let n = col;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}

/** A cell's A1 address, zero-based row and column. */
export function cellAddress(row: number, col: number): string {
  return `${columnLetter(col)}${row + 1}`;
}

/** The selection's active cell: `active` when set, else the anchor. */
export function activeCell(sel: Selection): CellAddr {
  return sel.active ?? sel.anchor;
}

/**
 * A collapsed selection at a single cell.
 */
export function selectionAt(addr: CellAddr): Selection {
  return { anchor: { ...addr }, focus: { ...addr } };
}

/**
 * The normalized inclusive rectangle spanned by a selection's anchor and focus.
 */
export function normalizeRange(sel: Selection): CellRange {
  return {
    top: Math.min(sel.anchor.row, sel.focus.row),
    bottom: Math.max(sel.anchor.row, sel.focus.row),
    left: Math.min(sel.anchor.col, sel.focus.col),
    right: Math.max(sel.anchor.col, sel.focus.col),
  };
}

/**
 * Whether a cell falls inside a selection's normalized range.
 */
export function rangeContains(sel: Selection, addr: CellAddr): boolean {
  const r = normalizeRange(sel);
  return addr.row >= r.top && addr.row <= r.bottom && addr.col >= r.left && addr.col <= r.right;
}

/**
 * Number of cells covered by a selection.
 */
export function cellCount(sel: Selection): number {
  const r = normalizeRange(sel);
  return (r.bottom - r.top + 1) * (r.right - r.left + 1);
}

/**
 * Extend the selection to a new focus, holding the anchor (and the active cell)
 * fixed: shift-click / shift-arrow semantics. The address is clamped into the
 * grid when limits pass.
 */
export function extendTo(sel: Selection, addr: CellAddr, limits?: SelectionLimits): Selection {
  const focus = limits
    ? { row: clamp(addr.row, 0, limits.rows - 1), col: clamp(addr.col, 0, limits.cols - 1) }
    : { ...addr };
  return sel.active
    ? { anchor: { ...sel.anchor }, focus, active: { ...sel.active } }
    : { anchor: { ...sel.anchor }, focus };
}

/** The top-left cell of the merged range holding `addr`, else `addr`. */
export function mergeAnchor(addr: CellAddr, merges?: SelectionMerges): CellAddr {
  const merged = merges?.mergeAt(addr.row, addr.col);
  return merged ? { row: merged.top, col: merged.left } : { ...addr };
}

/**
 * Step one (or `step`) cells in a direction. When `extend` is set the anchor
 * holds and the focus moves, growing the range; otherwise the selection
 * collapses onto the cell that far from the active cell. Hidden rows and
 * columns are stepped over, and a merged range counts as one cell: a move
 * leaves it from its far edge and a collapsed move landing in one lands on its
 * top-left cell, as in Excel. The result stays inside the grid bounds in
 * `opts.limits`.
 */
export function moveFocus(sel: Selection, direction: Direction, opts: MoveOpts): Selection {
  const { limits, hidden, merges } = opts;
  const vertical = direction === 'up' || direction === 'down';
  const back = direction === 'up' || direction === 'left';
  const from = opts.extend ? sel.focus : activeCell(sel);
  const merged = merges?.mergeAt(from.row, from.col);
  const start = merged
    ? vertical
      ? back
        ? merged.top
        : merged.bottom
      : back
        ? merged.left
        : merged.right
    : vertical
      ? from.row
      : from.col;
  const limit = vertical ? limits.rows : limits.cols;
  const isHidden = vertical ? hidden?.isRowHidden : hidden?.isColHidden;
  const step = back ? -1 : 1;
  let position = clamp(start, 0, limit - 1);
  for (let remaining = opts.step ?? 1; remaining > 0; remaining--) {
    let next = position + step;
    while (next >= 0 && next < limit && isHidden?.(next)) next += step;
    if (next < 0 || next >= limit) break;
    position = next;
  }
  const moved = position !== start;
  const landed = !moved
    ? { row: clamp(from.row, 0, limits.rows - 1), col: clamp(from.col, 0, limits.cols - 1) }
    : vertical
      ? { row: position, col: from.col }
      : { row: from.row, col: position };
  return opts.extend ? extendTo(sel, landed) : selectionAt(mergeAnchor(landed, merges));
}
