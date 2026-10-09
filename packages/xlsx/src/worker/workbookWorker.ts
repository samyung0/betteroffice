/**
 * The workbook's own thread: the XLSX engine and its Yrs document, the only
 * copy of the workbook. It answers {@link WorkbookWorkerRequest}s in the order
 * they were posted and posts every update the workbook emits before the reply
 * of the request that caused it.
 */

import type { CellEdit, RangeStylePatch, WorkbookHandle } from '../wasm/loader';
import {
  initWasm,
  isPngExportAvailable,
  isProposalsAvailable,
  openWorkbook,
  StaleProposalError,
} from '../wasm/loader';
import type { GridMeta } from '../display-list/types';
import { cellAddress, normalizeRange } from '../selection/index';
import {
  SHEET_GONE,
  WORKBOOK_METHODS,
  type WorkbookCallRequest,
  type WorkbookFrame,
  type WorkbookMethod,
  type WorkbookOpened,
  type WorkbookView,
  type WorkbookWorkerMessage,
  type WorkbookWorkerRequest,
} from './protocol';

interface WorkerScope {
  onmessage: ((event: MessageEvent<WorkbookWorkerRequest>) => void) | null;
  postMessage(message: WorkbookWorkerMessage, transfer?: Transferable[]): void;
}

const scope = self as unknown as WorkerScope;
const methods = new Set<string>(WORKBOOK_METHODS);
// As the editor caps them: the overlay draws merged ranges in view.
const MAX_VISIBLE_MERGED_RANGES = 1024;
// The engine's `rangeCells` cap (MAX_RANGE_CELLS in workbook.rs). A drawn
// window may hold up to 250,000 cells, so it is read in blocks under it.
const MAX_RANGE_CELLS = 100_000;
// Calls that read or write the sheet in their first argument.
const SHEET_FIRST = new Set<WorkbookMethod>([
  'cell',
  'cellPosition',
  'rangeCells',
  'editCell',
  'editCells',
  'moveChart',
  'patchRangeStyle',
  'setNumberFormat',
  'selectionFormatting',
  'captureFormat',
  'applyFormat',
  'mergedRanges',
  'printDisplayList',
]);

let workbook: WorkbookHandle | null = null;
let epoch = 0;
let seq = 0;
let queue = Promise.resolve();

scope.onmessage = (event) => {
  const request = event.data;
  queue = queue.then(() => handle(request));
};

async function handle(request: WorkbookWorkerRequest): Promise<void> {
  seq = request.id;
  try {
    const value = await answer(request);
    reply({ type: 'reply', id: request.id, ok: true, value }, transfers(value));
  } catch (error) {
    reply({
      type: 'reply',
      id: request.id,
      ok: false,
      error: error instanceof Error ? error.message : String(error),
      ...(error instanceof StaleProposalError ? { stale: error.cells } : {}),
    });
  }
}

async function answer(request: WorkbookWorkerRequest): Promise<unknown> {
  if (request.type === 'open') {
    await initWasm();
    const opened = openWorkbook(request.bytes, {
      collaborative: request.collaborative,
      clientId: request.clientId,
    });
    if (request.initialUpdate) {
      try {
        opened.applyUpdate(request.initialUpdate);
      } catch (error) {
        opened.dispose();
        throw error;
      }
    }
    workbook = opened;
    opened.onUpdate((update, origin) => {
      epoch += 1;
      reply({ type: 'update', update, origin }, [update.buffer]);
    });
    return {
      clientId: opened.clientId,
      proposals: isProposalsAvailable(),
      png: isPngExportAvailable(),
    } satisfies WorkbookOpened;
  }
  if (request.type === 'dispose') {
    workbook?.dispose();
    workbook = null;
    return null;
  }
  const open = workbook;
  if (!open) throw new Error('workbook is not open');
  if (request.type === 'frame') return frame(open, request.view);
  if (request.type === 'toggle') {
    const sheet = activeSheet(open, request.sheetId);
    const patch: RangeStylePatch = {};
    patch[request.property] = !open.selectionFormatting(sheet, request.range)[request.property];
    return open.patchRangeStyle(sheet, request.range, patch);
  }
  return call(open, request);
}

function call(open: WorkbookHandle, request: WorkbookCallRequest): unknown {
  if (!methods.has(request.method)) throw new Error(`unknown method ${request.method}`);
  let args = request.args;
  if (request.sheetId !== undefined) {
    if (request.method === 'setActiveSheet') {
      const index = open.sheetInfo().sheetIds.indexOf(request.sheetId);
      if (index < 0) throw new Error(SHEET_GONE);
      args = [index];
    } else {
      const sheet = activeSheet(open, request.sheetId);
      if (SHEET_FIRST.has(request.method)) args = [sheet, ...args.slice(1)];
      else if (request.method === 'applyOps')
        args = [(args[0] as Record<string, unknown>[]).map((op) => withSheet(op, sheet))];
    }
  }
  const method = open[request.method] as (...values: unknown[]) => unknown;
  return method.apply(open, args);
}

/** The active sheet's index while it is still `sheetId`. */
function activeSheet(open: WorkbookHandle, sheetId: string): number {
  const info = open.sheetInfo();
  if (info.sheetIds[info.activeSheet] !== sheetId) throw new Error(SHEET_GONE);
  return info.activeSheet;
}

function withSheet(op: Record<string, unknown>, sheet: number): Record<string, unknown> {
  return typeof op.sheet === 'number' ? { ...op, sheet } : op;
}

function frame(open: WorkbookHandle, requested: WorkbookView): WorkbookFrame {
  const sheetInfo = open.sheetInfo();
  const sheet = sheetInfo.activeSheet;
  const view = requested.initialScroll
    ? {
        ...requested,
        viewport: {
          ...requested.viewport,
          x: sheetInfo.initialScrollX,
          y: sheetInfo.initialScrollY,
        },
      }
    : requested;
  const displayList = open.displayList(view.viewport);
  const grid = displayList.grid;
  const selection = view.selection;
  const range = selection ? normalizeRange(selection) : null;
  const rangeA1 = range
    ? `${cellAddress(range.top, range.left)}:${cellAddress(range.bottom, range.right)}`
    : null;
  const quietly = <T,>(read: () => T, fallback: T): T => {
    try {
      return read();
    } catch {
      return fallback;
    }
  };
  return {
    epoch,
    seq,
    view,
    sheetInfo,
    displayList,
    cells: grid ? windowCells(open, sheet, grid) : {},
    focus: selection
      ? quietly(() => open.cell(sheet, selection.focus.row, selection.focus.col), null)
      : null,
    formatting: rangeA1 ? quietly(() => open.selectionFormatting(sheet, rangeA1), {}) : {},
    selectionMerged: rangeA1 ? quietly(() => open.mergedRanges(sheet, rangeA1), []) : [],
    visibleMerged: grid ? visibleMerged(open, sheet, grid, quietly) : [],
    history: quietly(() => open.historyState(), {
      canUndo: false,
      canRedo: false,
      undoDepth: 0,
      redoDepth: 0,
    }),
    proposals: quietly(() => open.listProposals(), []),
  };
}

/** The merged ranges over the drawn window's span, as the editor reads them. */
function visibleMerged(
  open: WorkbookHandle,
  sheet: number,
  grid: GridMeta,
  quietly: <T>(read: () => T, fallback: T) => T
) {
  const rows = grid.rowOffsets.length - 1;
  const cols = grid.colOffsets.length - 1;
  if (rows <= 0 || cols <= 0) return [];
  const span = `${cellAddress(grid.startRow, grid.startCol)}:${cellAddress(
    grid.startRow + rows - 1,
    grid.startCol + cols - 1
  )}`;
  return quietly(() => open.mergedRanges(sheet, span), []).slice(0, MAX_VISIBLE_MERGED_RANGES);
}

/**
 * The editable text of every non-empty drawn cell: one range read per run of
 * drawn tracks, split into row blocks under the engine's range cap. Empty
 * cells stay out, so the frame grows with the content drawn, not the window.
 */
function windowCells(open: WorkbookHandle, sheet: number, grid: GridMeta) {
  const cells: Record<string, CellEdit> = {};
  const rows = runs(tracks(grid.startRow, grid.rowIndices, grid.rowOffsets.length - 1));
  const cols = runs(tracks(grid.startCol, grid.colIndices, grid.colOffsets.length - 1));
  for (const [left, right] of cols) {
    // a sheet has at most 16,384 columns, so a block holds at least six rows.
    const step = Math.floor(MAX_RANGE_CELLS / (right - left + 1));
    for (const [first, last] of rows)
      for (let top = first; top <= last; top += step) {
        const bottom = Math.min(last, top + step - 1);
        const block = open.rangeCells(
          sheet,
          `${cellAddress(top, left)}:${cellAddress(bottom, right)}`
        );
        block.forEach((line, row) =>
          line.forEach((cell, col) => {
            if (cell.input) cells[`${top + row}:${left + col}`] = cell;
          })
        );
      }
  }
  return cells;
}

function tracks(start: number, indices: number[] | undefined, count: number): number[] {
  return indices ?? Array.from({ length: Math.max(0, count) }, (_, index) => start + index);
}

/** Consecutive indices as inclusive [first, last] runs. */
function runs(indices: number[]): Array<[number, number]> {
  const out: Array<[number, number]> = [];
  for (const index of indices) {
    const last = out[out.length - 1];
    if (last && index === last[1] + 1) last[1] = index;
    else out.push([index, index]);
  }
  return out;
}

function transfers(value: unknown): Transferable[] {
  return value instanceof Uint8Array && value.byteOffset === 0 && value.byteLength === value.buffer.byteLength
    ? [value.buffer as ArrayBuffer]
    : [];
}

function reply(message: WorkbookWorkerMessage, transfer: Transferable[] = []): void {
  scope.postMessage(message, transfer);
}
