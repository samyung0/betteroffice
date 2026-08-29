import initWasmModule, { XlsxViewDocument } from './generated/viewer/xlsx_view_wasm.js';
import type { InitInput } from './generated/viewer/xlsx_view_wasm.js';
import type { ChartRegion, DisplayList } from '../display-list/types';
import type { CellPosition, MergedRange, SheetInfo, Viewport } from './loader';

export type WasmInitInput = InitInput | Promise<InitInput>;

export interface WorkbookViewerHandle {
  sheetInfo(): SheetInfo;
  setActiveSheet(index: number): void;
  displayList(viewport: Viewport): DisplayList;
  chartAtPoint(viewport: Viewport, x: number, y: number): ChartRegion | null;
  cellPosition(sheet: number, row: number, col: number): CellPosition;
  cellText(sheet: number, row: number, col: number): string;
  mergedRanges(sheet: number, range: string): MergedRange[];
  dispose(): void;
}

let initialized = false;
let initialization: Promise<void> | undefined;

export function initWasm(
  input: WasmInitInput = new URL('./generated/viewer/xlsx_view_wasm_bg.wasm', import.meta.url)
): Promise<void> {
  if (initialized) return Promise.resolve();
  if (initialization) return initialization;
  initialization = initWasmModule({ module_or_path: input }).then(
    () => {
      initialized = true;
    },
    (error: unknown) => {
      initialization = undefined;
      throw toError(error);
    }
  );
  return initialization;
}

export function isWasmAvailable(): boolean {
  return typeof WebAssembly === 'object';
}

export function openWorkbook(bytes: Uint8Array): WorkbookViewerHandle {
  if (!initialized) throw new Error('xlsx viewer wasm is not initialized; call initWasm() first');
  const document = call(() => XlsxViewDocument.open(bytes));
  let disposed = false;

  const read = <T,>(operation: () => T): T => {
    if (disposed) throw new Error('workbook viewer handle is disposed');
    return call(operation);
  };

  return {
    sheetInfo: () => parseJson(() => read(() => document.sheetInfoJson())),
    setActiveSheet: (index) => read(() => document.setActiveSheet(index)),
    displayList: (viewport) =>
      parseJson(() => read(() => document.displayListJson(JSON.stringify(viewport)))),
    chartAtPoint: (viewport, x, y) =>
      parseJson(() =>
        read(() => document.chartAtPointJson(JSON.stringify({ viewport, x, y })))
      ),
    cellPosition: (sheet, row, col) =>
      parseJson(() =>
        read(() => document.cellPositionJson(JSON.stringify({ sheet, row, col })))
      ),
    cellText: (sheet, row, col) => read(() => document.cellText(sheet, row, col)),
    mergedRanges: (sheet, range) =>
      parseJson<{ ranges: MergedRange[] }>(() =>
        read(() => document.mergedRangesJson(JSON.stringify({ sheet, range })))
      ).ranges,
    dispose(): void {
      if (disposed) return;
      disposed = true;
      call(() => document.free());
    },
  };
}

export function wasmVersion(): string {
  if (!initialized) throw new Error('xlsx viewer wasm is not initialized; call initWasm() first');
  return XlsxViewDocument.version();
}

function parseJson<T>(operation: () => string): T {
  return JSON.parse(operation()) as T;
}

function call<T>(operation: () => T): T {
  try {
    return operation();
  } catch (error) {
    throw toError(error);
  }
}

function toError(error: unknown): Error {
  if (error instanceof Error) return error;
  return new Error(typeof error === 'string' ? error : String(error));
}

