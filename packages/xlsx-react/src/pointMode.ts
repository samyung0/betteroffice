/**
 * Point mode for formula drafts, as in Excel and Google Sheets: while a draft
 * starting with `=` waits for an operand at its caret, a click, a drag or (in a
 * typed edit) an arrow key writes a cell or range reference there instead of
 * committing the draft. Pointing again replaces that reference; typing ends it.
 */

import { cellAddress } from '@betteroffice/xlsx';
import type { CellAddr } from '@betteroffice/xlsx';

/** A reference point mode wrote: its span in the draft `text` it produced, and its cells. */
export interface PointedRef {
  text: string;
  start: number;
  end: number;
  anchor: CellAddr;
  focus: CellAddr;
}

// what may come right before an operand, and right after one.
const OPERAND_LEADS = new Set(['=', '(', ',', '+', '-', '*', '/', '^', '&', '<', '>', ':']);
const OPERAND_ENDS = new Set([')', ',', '+', '-', '*', '/', '^', '&', '<', '>', '=', ':', '%']);

/** Whether an operand typed at `caret` would fit the formula draft `text`. */
export function referenceSlot(text: string, caret: number): boolean {
  if (!text.startsWith('=') || caret < 1 || caret > text.length) return false;
  let quotes = 0;
  for (let i = 0; i < caret; i++) if (text[i] === '"') quotes++;
  if (quotes % 2 === 1) return false;
  let before = caret - 1;
  while (before > 0 && text[before] === ' ') before--;
  let after = caret;
  while (after < text.length && text[after] === ' ') after++;
  return OPERAND_LEADS.has(text[before]) && (after === text.length || OPERAND_ENDS.has(text[after]));
}

/** `B3`, or `B3:C5` for a range, whichever corner the pointer started from. */
export function referenceText(anchor: CellAddr, focus: CellAddr): string {
  const top = Math.min(anchor.row, focus.row);
  const left = Math.min(anchor.col, focus.col);
  const bottom = Math.max(anchor.row, focus.row);
  const right = Math.max(anchor.col, focus.col);
  const first = cellAddress(top, left);
  return top === bottom && left === right ? first : `${first}:${cellAddress(bottom, right)}`;
}

/** The pointed reference while the draft is as it left it, caret after it; else null. */
export function livePointedRef(
  pointed: PointedRef | null,
  text: string,
  caret: number
): PointedRef | null {
  return pointed && pointed.text === text && caret === pointed.end ? pointed : null;
}

/** `text` with a reference to `anchor:focus` written over `start..end`. */
export function writeReference(
  text: string,
  start: number,
  end: number,
  anchor: CellAddr,
  focus: CellAddr
): PointedRef {
  const reference = referenceText(anchor, focus);
  return {
    text: text.slice(0, start) + reference + text.slice(end),
    start,
    end: start + reference.length,
    anchor,
    focus,
  };
}
