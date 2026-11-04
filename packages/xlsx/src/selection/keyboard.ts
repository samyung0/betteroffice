/**
 * Keyboard reducer: maps a keystroke over the current selection to a
 * declarative {@link SelectionAction}. Pure — it neither mutates state nor
 * touches the DOM; the chrome owns the keydown listener and interprets the
 * action. Excel navigation semantics, with two deliberate choices documented at
 * the F2/Enter handlers below.
 */

import { activeCell, extendTo, mergeAnchor, moveFocus, selectionAt } from './model';
import type {
  CellAddr,
  Direction,
  KeyInput,
  Selection,
  SelectionAction,
  SelectionData,
  SelectionHidden,
  SelectionLimits,
  SelectionMerges,
} from './types';

const ARROWS: Record<string, Direction> = {
  ArrowUp: 'up',
  ArrowDown: 'down',
  ArrowLeft: 'left',
  ArrowRight: 'right',
};

function mod(key: KeyInput): boolean {
  return Boolean(key.metaKey || key.ctrlKey);
}

// the used-range edge in a direction (ctrl/cmd+arrow without the engine's data
// edge); past it the focus stays.
function edgeInDirection(focus: CellAddr, direction: Direction, limits: SelectionLimits): CellAddr {
  switch (direction) {
    case 'up':
      return { row: 0, col: focus.col };
    case 'down':
      return { row: Math.max(focus.row, limits.usedRows - 1), col: focus.col };
    case 'left':
      return { row: focus.row, col: 0 };
    case 'right':
      return { row: focus.row, col: Math.max(focus.col, limits.usedCols - 1) };
  }
}

// the first index from `start` stepping by `step` that is not hidden, if any.
function visibleIndex(
  start: number,
  step: number,
  limit: number,
  hidden?: (index: number) => boolean
): number | null {
  let index = start;
  while (index >= 0 && index < limit && hidden?.(index)) index += step;
  return index >= 0 && index < limit ? index : null;
}

// `edge`, or the nearest visible cell before it on the way from `from`; `from`
// itself when there is none.
function visibleEdge(
  from: CellAddr,
  edge: CellAddr,
  direction: Direction,
  limits: SelectionLimits,
  hidden?: SelectionHidden
): CellAddr {
  const vertical = direction === 'up' || direction === 'down';
  const isHidden = vertical ? hidden?.isRowHidden : hidden?.isColHidden;
  const target = vertical ? edge.row : edge.col;
  if (!isHidden?.(target)) return edge;
  const back = direction === 'up' || direction === 'left';
  const origin = vertical ? from.row : from.col;
  const index = visibleIndex(target, back ? 1 : -1, vertical ? limits.rows : limits.cols, isHidden);
  if (index === null || (back ? index >= origin : index <= origin)) return from;
  return vertical ? { row: index, col: edge.col } : { row: edge.row, col: index };
}

function moveAction(selection: Selection): SelectionAction {
  return { type: 'move', selection };
}

// a lone printable character (no ctrl/cmd/alt) opens the editor seeded with it.
function isPrintable(key: KeyInput): boolean {
  return key.key.length === 1 && !key.ctrlKey && !key.metaKey && !key.altKey;
}

/**
 * The whole used range, keeping the selection's active cell as Excel does; an
 * active cell outside that range moves to its first cell, so the selection
 * always holds it.
 */
export function selectAll(state: Selection | null, limits: SelectionLimits): Selection {
  const anchor = { row: 0, col: 0 };
  const focus = { row: limits.usedRows - 1, col: limits.usedCols - 1 };
  const active = state ? activeCell(state) : anchor;
  const inside = active.row <= focus.row && active.col <= focus.col;
  return { anchor, focus: { ...focus }, active: { ...(inside ? active : anchor) } };
}

/**
 * What a key needs from the engine before {@link selectionKeyReducer} can
 * place it: the data edge in a direction (Ctrl/Cmd+arrow), the row's last data
 * cell (End) or the used range's last cell (Ctrl/Cmd+End), asked from `from`.
 */
export function selectionDataRequest(
  state: Selection,
  key: KeyInput
): { from: CellAddr; direction: Direction | 'rowEnd' | 'usedRange' } | null {
  const direction = ARROWS[key.key];
  if (direction && mod(key)) {
    return { from: key.shiftKey ? state.focus : activeCell(state), direction };
  }
  if (key.key === 'End') {
    return { from: activeCell(state), direction: mod(key) ? 'usedRange' : 'rowEnd' };
  }
  return null;
}

/**
 * Reduce a keystroke to a selection action against the current selection and
 * grid limits. `data` is the engine's answer to {@link selectionDataRequest}
 * for this key; without it Ctrl/Cmd+arrow, End and Ctrl/Cmd+End stop at the
 * used range. `hidden` rows and columns are stepped over and `merges` count
 * as one cell. Returns `{ type: 'none' }` for keys this layer does not own so
 * the chrome can fall through to its own handling.
 */
export function selectionKeyReducer(
  state: Selection,
  key: KeyInput,
  limits: SelectionLimits,
  data?: SelectionData,
  hidden?: SelectionHidden,
  merges?: SelectionMerges
): SelectionAction {
  const land = (addr: CellAddr) => selectionAt(mergeAnchor(addr, merges));
  const options = { limits, hidden, merges };
  const direction = ARROWS[key.key];
  if (direction) {
    if (mod(key)) {
      // the engine's data edge (Excel), else the used range's edge.
      const from = key.shiftKey ? state.focus : activeCell(state);
      const edge = visibleEdge(
        from,
        data?.edge ?? edgeInDirection(from, direction, limits),
        direction,
        limits,
        hidden
      );
      return moveAction(key.shiftKey ? extendTo(state, edge, limits) : land(edge));
    }
    return moveAction(moveFocus(state, direction, { ...options, extend: key.shiftKey }));
  }

  switch (key.key) {
    // tab/shift+tab step right/left like the arrows and stop at the sheet's
    // edges, as excel does for one cell; they always collapse (excel commits it).
    case 'Tab':
      return moveAction(moveFocus(state, key.shiftKey ? 'left' : 'right', { ...options, extend: false }));

    // enter navigates (down/up); it never opens the editor — F2 does. this is
    // the excel split: type to overwrite, F2 to edit in place, enter to commit.
    case 'Enter':
      return moveAction(moveFocus(state, key.shiftKey ? 'up' : 'down', { ...options, extend: false }));

    case 'Home': {
      const col = visibleIndex(0, 1, limits.cols, hidden?.isColHidden) ?? 0;
      const row = mod(key)
        ? (visibleIndex(0, 1, limits.rows, hidden?.isRowHidden) ?? 0)
        : activeCell(state).row;
      return moveAction(land({ row, col }));
    }

    // end: the row's last data cell (else the used range's last column), and
    // ctrl/cmd+end the used range's last cell; shift extends to it.
    case 'End': {
      const active = activeCell(state);
      const used = mod(key)
        ? { row: limits.usedRows - 1, col: limits.usedCols - 1 }
        : visibleEdge(active, { row: active.row, col: limits.usedCols - 1 }, 'right', limits, hidden);
      const target = data?.edge ?? used;
      return moveAction(key.shiftKey ? extendTo(state, target, limits) : land(target));
    }

    case 'PageDown':
      return moveAction(
        moveFocus(state, 'down', { ...options, extend: key.shiftKey, step: limits.rowsPerPage })
      );

    case 'PageUp':
      return moveAction(
        moveFocus(state, 'up', { ...options, extend: key.shiftKey, step: limits.rowsPerPage })
      );

    // F2 opens the editor on the current cell without seeding input.
    case 'F2':
      return { type: 'startEdit' };

    case 'Delete':
    case 'Backspace':
      return { type: 'clear' };
  }

  // cmd/ctrl+A selects the whole used range, the active cell staying put.
  if ((key.key === 'a' || key.key === 'A') && mod(key)) {
    return moveAction(selectAll(state, limits));
  }

  if (isPrintable(key)) {
    return { type: 'startEdit', initialInput: key.key };
  }

  return { type: 'none' };
}
