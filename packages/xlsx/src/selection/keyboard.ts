/**
 * Keyboard reducer: maps a keystroke over the current selection to a
 * declarative {@link SelectionAction}. Pure — it neither mutates state nor
 * touches the DOM; the chrome owns the keydown listener and interprets the
 * action. Excel navigation semantics, with two deliberate choices documented at
 * the F2/Enter handlers below.
 */

import { activeCell, extendTo, moveFocus, selectionAt } from './model';
import type {
  CellAddr,
  Direction,
  KeyInput,
  Selection,
  SelectionAction,
  SelectionLimits,
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

// jump to the used-range edge in a direction (ctrl/cmd+arrow) — we have no data
// map here, so the boundary is the used range itself; past it the focus stays.
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

function moveAction(selection: Selection): SelectionAction {
  return { type: 'move', selection };
}

// a lone printable character (no ctrl/cmd/alt) opens the editor seeded with it.
function isPrintable(key: KeyInput): boolean {
  return key.key.length === 1 && !key.ctrlKey && !key.metaKey && !key.altKey;
}

/** The whole used range, keeping the selection's active cell, as Excel does. */
export function selectAll(state: Selection | null, limits: SelectionLimits): Selection {
  return {
    anchor: { row: 0, col: 0 },
    focus: { row: limits.usedRows - 1, col: limits.usedCols - 1 },
    active: { ...(state ? activeCell(state) : { row: 0, col: 0 }) },
  };
}

/**
 * Reduce a keystroke to a selection action against the current selection and
 * grid limits. Returns `{ type: 'none' }` for keys this layer does not own so
 * the chrome can fall through to its own handling.
 */
export function selectionKeyReducer(
  state: Selection,
  key: KeyInput,
  limits: SelectionLimits
): SelectionAction {
  const direction = ARROWS[key.key];
  if (direction) {
    if (mod(key)) {
      return moveAction(
        key.shiftKey
          ? extendTo(state, edgeInDirection(state.focus, direction, limits), limits)
          : selectionAt(edgeInDirection(activeCell(state), direction, limits))
      );
    }
    return moveAction(moveFocus(state, direction, { extend: key.shiftKey, limits }));
  }

  switch (key.key) {
    // tab/shift+tab step right/left like the arrows and stop at the sheet's
    // edges, as excel does for one cell; they always collapse (excel commits it).
    case 'Tab':
      return moveAction(moveFocus(state, key.shiftKey ? 'left' : 'right', { extend: false, limits }));

    // enter navigates (down/up); it never opens the editor — F2 does. this is
    // the excel split: type to overwrite, F2 to edit in place, enter to commit.
    case 'Enter':
      return moveAction(moveFocus(state, key.shiftKey ? 'up' : 'down', { extend: false, limits }));

    case 'Home':
      return moveAction(
        mod(key) ? selectionAt({ row: 0, col: 0 }) : selectionAt({ row: activeCell(state).row, col: 0 })
      );

    case 'End':
      return moveAction(
        mod(key)
          ? selectionAt({ row: limits.usedRows - 1, col: limits.usedCols - 1 })
          : selectionAt({ row: activeCell(state).row, col: limits.usedCols - 1 })
      );

    case 'PageDown':
      return moveAction(
        moveFocus(state, 'down', { extend: key.shiftKey, step: limits.rowsPerPage, limits })
      );

    case 'PageUp':
      return moveAction(
        moveFocus(state, 'up', { extend: key.shiftKey, step: limits.rowsPerPage, limits })
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
