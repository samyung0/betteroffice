import type { CollaborationUpdateOrigin } from '../collaboration/types';
import type { DisplayList } from '../display-list/types';
import type { Selection } from '../selection/types';
import type {
  CellEdit,
  HistoryState,
  MergedRange,
  Proposal,
  SelectionFormatting,
  SheetInfo,
  Viewport,
  WorkbookHandle,
} from '../wasm/loader';

/** The handle methods the worker answers; every one runs in the worker. */
export const WORKBOOK_METHODS = [
  'encodeStateVector',
  'encodeStateAsUpdate',
  'applyUpdate',
  'sheetInfo',
  'calculationStatus',
  'displayList',
  'printDisplayList',
  'chartAtPoint',
  'moveChart',
  'setActiveSheet',
  'editCell',
  'editCells',
  'applyOps',
  'undo',
  'redo',
  'cell',
  'searchText',
  'cellPosition',
  'rangeCells',
  'patchRangeStyle',
  'setNumberFormat',
  'selectionFormatting',
  'captureFormat',
  'applyFormat',
  'mergedRanges',
  'historyState',
  'renderPng',
  'renderRangePng',
  'save',
  'propose',
  'listProposals',
  'acceptProposal',
  'rejectProposal',
] as const satisfies readonly (keyof WorkbookHandle)[];

export type WorkbookMethod = (typeof WORKBOOK_METHODS)[number];

/** Methods that change the workbook (they may emit updates). */
export const MUTATING_METHODS: ReadonlySet<WorkbookMethod> = new Set([
  'applyUpdate',
  'moveChart',
  'setActiveSheet',
  'editCell',
  'editCells',
  'applyOps',
  'undo',
  'redo',
  'patchRangeStyle',
  'setNumberFormat',
  'applyFormat',
  'propose',
  'acceptProposal',
  'rejectProposal',
]);

/** What the editor shows: the scrolled window and the selection on the active sheet. */
export interface WorkbookView {
  viewport: Viewport;
  /** Echoed back; the worker draws in sheet pixels. */
  zoom: number;
  selection: Selection | null;
  /** Draw the active sheet at its saved scroll; the frame's viewport says where. */
  initialScroll?: boolean;
}

/**
 * Everything the editor reads to draw one state of the workbook, built in the
 * worker after every request posted before it.
 */
export interface WorkbookFrame {
  /** Workbook changes applied before this frame (local and remote). */
  epoch: number;
  /** The last request the worker handled before building it. */
  seq: number;
  view: WorkbookView;
  sheetInfo: SheetInfo;
  displayList: DisplayList;
  /**
   * The editable text of every non-empty cell in the drawn window, by
   * `row:col`; a drawn cell left out is empty.
   */
  cells: Record<string, CellEdit>;
  /** The selection's focus cell, drawn or not. */
  focus: CellEdit | null;
  formatting: SelectionFormatting;
  /** Merged ranges inside the selection and inside the drawn window, at most 1,024 each. */
  selectionMerged: MergedRange[];
  visibleMerged: MergedRange[];
  history: HistoryState;
  proposals: Proposal[];
}

export interface WorkbookOpenRequest {
  type: 'open';
  id: number;
  bytes: Uint8Array;
  collaborative: boolean;
  clientId?: number;
  initialUpdate?: Uint8Array;
}

export interface WorkbookCallRequest {
  type: 'call';
  id: number;
  method: WorkbookMethod;
  args: unknown[];
  /**
   * The active sheet the caller acted on. The call runs on the active sheet
   * only while it is still that one (a peer may have removed or moved it);
   * otherwise it fails with {@link SHEET_GONE}.
   */
  sheetId?: string;
}

export interface WorkbookToggleRequest {
  type: 'toggle';
  id: number;
  sheetId: string;
  range: string;
  property: 'bold' | 'italic' | 'strikethrough';
}

export interface WorkbookFrameRequest {
  type: 'frame';
  id: number;
  view: WorkbookView;
}

export type WorkbookWorkerRequest =
  | WorkbookOpenRequest
  | WorkbookCallRequest
  | WorkbookToggleRequest
  | WorkbookFrameRequest;

export interface WorkbookOpened {
  clientId: number;
  proposals: boolean;
  png: boolean;
}

export type WorkbookWorkerMessage =
  | { type: 'reply'; id: number; ok: true; value: unknown }
  | { type: 'reply'; id: number; ok: false; error: string; stale?: string[] }
  | { type: 'update'; update: Uint8Array; origin: CollaborationUpdateOrigin };

/** The error text of a call whose sheet stopped being the active one. */
export const SHEET_GONE = 'the sheet is no longer the active sheet';
