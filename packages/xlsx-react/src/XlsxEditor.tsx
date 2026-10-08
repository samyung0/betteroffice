/**
 * `<XlsxEditor />` — the editor shell. A dpr-aware canvas paints the grid; DOM
 * overlays (selection marquee, active-cell outline, in-cell editor) sit above it,
 * positioned from the same display-list geometry the painter uses. A top bar
 * holds the name box, formula bar, and save/undo/redo; an offscreen `role=grid`
 * mirror serves screen readers. All compute lives in `@betteroffice/xlsx`; this
 * layer is framework glue — keyboard/mouse wiring, focus flow, and DOM chrome.
 */

import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import {
  buildA11yGrid,
  cellAddress,
  cellAtPoint,
  cellRect,
  chartRegionAtPoint,
  extendTo,
  fromTsv,
  hyperlinkAtCell,
  moveFocus,
  normalizeRange,
  openWorkbookWorker,
  paintDisplayList,
  parseHyperlinkLocation,
  rangeRect,
  selectionAt,
  selectionKeyReducer,
  safeExternalHyperlink,
  SheetGoneError,
  StaleProposalError,
  toTsv,
  zoomedViewport,
} from '@betteroffice/xlsx';
import type {
  CellAddr,
  CellBounds,
  CellEdit,
  CellInputEdit,
  CapturedFormat,
  ChartRegion,
  DisplayList,
  DrawCmd,
  Direction,
  EditResult,
  Selection,
  SelectionLimits,
  SheetInfo,
  Viewport,
  WorkbookFrame,
  WorkbookProxy,
} from '@betteroffice/xlsx';
import type {
  AwarenessPeer,
  CollaborationAwareness,
  CollaborationReplica,
} from '@betteroffice/xlsx/collaboration';
import type { Translations } from '@betteroffice/xlsx-i18n';
import { LocaleProvider, useTranslation } from './i18n';
import { EditorToolbar } from './components/EditorToolbar';
import type {
  FormattingAction,
  HorizontalAlignment,
  MergeAction,
  NumberFormat,
  SelectionFormatting,
  TextWrapping,
  VerticalAlignment,
} from './components/Toolbar';
import {
  ToolbarButton,
  ToolbarGroup,
  chromeFont,
  toolbarColors,
} from './components/ui/ToolbarPrimitives';
import { IconSetContext, type IconSet, ToolbarIcon } from './components/ui/ToolbarIcon';
import {
  ALL_STYLE_PROPERTIES,
  type FreezeAmount,
  frozenCount,
  freezePaneOp,
  newSheetName,
  xlsxCommandEdits,
  type XlsxCommand,
  type XlsxCommandState,
} from './commands';
import {
  expandRangeToMergedCells,
  PresenceStrip,
  RemoteSelections,
} from './presence/Presence';
import { ProposalsPanel } from './proposals/ProposalsPanel';

/**
 * The imperative surface handed to {@link XlsxEditorProps.onReady}: the open
 * workbook handle plus a `refreshProposals` to re-read the pending list after an
 * external caller (e.g. a demo agent) stages proposals on the same handle.
 */
export interface XlsxEditorApi {
  clearSelection: () => void;
  focus: () => void;
  /** The workbook, in its own worker. */
  handle: WorkbookProxy;
  /** Commits pending input, or rejects with the engine's error. */
  flush: () => Promise<void>;
  refreshProposals: () => void;
  save: () => Promise<Uint8Array>;
  /** Scrolls the focus cell into view. */
  selectCells: (sheet: number, selection: Selection) => Promise<boolean>;
  /** Runs a menu command on the current selection, as its toolbar button would. */
  run: (command: XlsxCommand) => void;
  /** The sheet metadata of the grid on screen. */
  sheetInfo: () => SheetInfo | null;
  /** The part of the active sheet on screen, in sheet pixels (for an image of it). */
  visibleViewport: () => Viewport | null;
}

export interface XlsxEditorCollaborationOptions {
  /** Peer-unique Yrs client ID. Generated securely when omitted. */
  clientId?: number;
  /** Shared Yrs state applied before the replica is exposed. */
  initialUpdate?: Uint8Array;
  /** Receive the editor-owned collaboration replica. */
  onReplica?: (replica: CollaborationReplica | null) => void;
  /** Presence-capable provider connected to the editor-owned replica. */
  provider?: CollaborationAwareness | null;
}

/**
 * Props for {@link XlsxEditor}.
 */
export interface XlsxEditorProps {
  /** Raw .xlsx bytes to open. When omitted the shell paints a demo frame. */
  file?: Uint8Array;
  /** Download name for the save button; falls back to `workbook.xlsx`. */
  fileName?: string;
  /** Receive saved bytes instead of triggering a browser download. */
  onSave?: (bytes: Uint8Array) => void;
  /** Called after a user edit changes the workbook. */
  onChange?: () => void;
  /** Reports whether uncommitted cell, formula or chart-nudge input exists. */
  onPendingChange?: (pending: boolean) => void;
  /** Open a network-ready Yrs replica and repaint when peer updates arrive. */
  collaboration?: XlsxEditorCollaborationOptions;
  i18n?: Translations;
  /**
   * Called when a workbook opens, with a handle to stage agent proposals and a
   * way to refresh the panel afterward. Enables demo/host agents without
   * exposing the wasm object through the render tree. A returned cleanup runs
   * before the workbook is replaced or disposed.
   */
  onReady?: (api: XlsxEditorApi) => void | (() => void);
  /** Called once per opened workbook, right after its first grid is painted. */
  onFirstPaint?: () => void;
  className?: string;
  /** Blocks user edits; navigation and selection remain available. */
  readOnly?: boolean;
  /**
   * Flat chrome for hosts with their own toolbars: the formatting toolbar as one
   * 40px row (save, PNG export and print in its menu) over a 40px formula bar.
   */
  singleRowToolbar?: boolean;
  /** Called with what a host menu shows enabled and checked, whenever it changes. */
  onCommandStateChange?: (state: XlsxCommandState) => void;
  showCustomNumberFormat?: boolean;
  /** Replaces every toolbar icon; without it the built-in drawings stay. */
  icons?: IconSet;
  showSearchMenus?: boolean;
  showProposals?: boolean;
  showFontPicker?: boolean;
  showFontSizePicker?: boolean;
  showZoomControl?: boolean;
  /** The zoom it opens at (1 = 100%); the toolbar and `run('zoom:…')` change it. */
  initialZoom?: number;
}

/**
 * the open in-cell editor: which cell it targets, its current draft text and
 * the sheet it was opened on (it lands there or nowhere).
 */
interface EditState {
  row: number;
  col: number;
  value: string;
  sheetId: string;
}

/** a committed cell input drawn over its cell until a frame shows it. */
interface OptimisticEdit {
  seq: number;
  sheetId: string;
  row: number;
  col: number;
  text: string;
}

const NO_HISTORY = { canUndo: false, canRedo: false };
const NUMBER_INPUT = /^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?%?$/i;

// the in-cell editor while its cell is scrolled away: in the window (so typing
// never scrolls the grid to it) but invisible and click-through.
const OFFSCREEN_EDITOR: React.CSSProperties = {
  position: 'absolute',
  left: 0,
  top: 0,
  width: 1,
  height: 1,
  padding: 0,
  border: 0,
  opacity: 0,
  pointerEvents: 'none',
};

/** The frame on screen and the state it was painted from. */
interface Painted {
  frame: DisplayList;
  zoom: number;
  viewport: Viewport;
  handle: WorkbookProxy | null;
  sheet: number;
  /** The engine frame's `seq`: it shows every request posted before it. */
  seq: number;
  canvas: HTMLCanvasElement;
}

function sameCell(a: CellAddr, b: CellAddr): boolean {
  return a.row === b.row && a.col === b.col;
}

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

// the painted frame is the view the scroll container shows now.
function paintedIsLive(painted: Painted, scroll: HTMLElement, zoom: number): boolean {
  return (
    painted.zoom === zoom &&
    painted.viewport.x === scroll.scrollLeft / zoom &&
    painted.viewport.y === scroll.scrollTop / zoom &&
    painted.viewport.width === scroll.clientWidth / zoom &&
    painted.viewport.height === scroll.clientHeight / zoom
  );
}

/**
 * Dispatched (bubbling) from the grid once a frame from the workbook's worker
 * is painted, with the frame's `epoch` and `seq`. Hosts time input to frame
 * with it.
 */
export const XLSX_FRAME_PRESENTED_EVENT = 'xlsx-frame-presented';

// keys that type nothing: a scrolled-away edit stays where it is for them.
const NON_TYPING_KEYS = new Set(['Shift', 'Control', 'Alt', 'Meta', 'CapsLock', 'AltGraph', 'Fn']);

const COL_W = 96;
const ROW_H = 24;
const BRAND = '#217346';
const DEFAULT_XLSX_TOOLBAR_HEIGHT = 87;
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const CHART_NUDGE_PX = 1;
const CHART_NUDGE_MULTIPLIER = 10;
// how long a run of arrow presses may stay local before it lands as one edit.
const CHART_NUDGE_SETTLE_MS = 250;
const CHART_NUDGE_KEYS: Record<string, [number, number] | undefined> = {
  ArrowLeft: [-1, 0],
  ArrowRight: [1, 0],
  ArrowUp: [0, -1],
  ArrowDown: [0, 1],
};
// chrome shortcuts that stay live while a chart is selected; every other key
// stops at the chart rather than reaching the cells behind it.
const CHART_GLOBAL_KEYS = new Set(['z', 'y', 's']);

// a placeholder grid frame so the shell paints something real when no file is
// open. real files render through the wasm display list instead.
function buildDemoDisplayList(width: number, height: number, cellText: string): DisplayList {
  const commands: DrawCmd[] = [
    { op: 'fillRect', x: 0, y: 0, w: width, h: height, color: '#ffffff' },
  ];
  const cols = Math.ceil(width / COL_W);
  const rows = Math.ceil(height / ROW_H);
  for (let c = 0; c <= cols; c++) {
    commands.push({
      op: 'line',
      x1: c * COL_W,
      y1: 0,
      x2: c * COL_W,
      y2: height,
      width: 1,
      color: '#e0e0e0',
    });
  }
  for (let r = 0; r <= rows; r++) {
    commands.push({
      op: 'line',
      x1: 0,
      y1: r * ROW_H,
      x2: width,
      y2: r * ROW_H,
      width: 1,
      color: '#e0e0e0',
    });
  }
  commands.push({ op: 'fillRect', x: 0, y: 0, w: width, h: ROW_H, color: '#f3f3f3' });
  commands.push({ op: 'fillRect', x: 0, y: 0, w: COL_W, h: height, color: '#f3f3f3' });
  commands.push({
    op: 'text',
    x: COL_W + 8,
    y: ROW_H + 18,
    text: cellText,
    fontSize: 14,
    color: '#202020',
    clip: { x: COL_W, y: ROW_H, w: COL_W * 3, h: ROW_H },
    align: 'left',
  });
  return { width, height, commands };
}

// the scroll offset on one axis that shows [start, end) whole, moving as little
// as possible: scroll coordinates (cellPosition) and the extent outside the
// frozen pane. A track wider than that shows its start.
function revealOffset(scroll: number, start: number, end: number, body: number): number {
  if (start < scroll || end - start > body) return start;
  return end > scroll + body ? end - body : scroll;
}

// median of the gaps between consecutive offsets, or a fallback when the window
// has no tracks. the median ignores outliers like a single very wide column, so
// the extent-to-count estimate below is not skewed by one atypical track.
function medianTrack(offsets: number[] | undefined, fallback: number): number {
  if (!offsets || offsets.length < 2) return fallback;
  const gaps: number[] = [];
  for (let i = 1; i < offsets.length; i++) gaps.push(offsets[i] - offsets[i - 1]);
  gaps.sort((a, b) => a - b);
  const mid = gaps[gaps.length >> 1];
  return mid > 0 ? mid : fallback;
}

// derive nav bounds from the scrollable extent: rows/cols estimated from the
// content size over a representative (median) track size, rowsPerPage from the
// viewport. a slack of one keeps the row/col just past the used edge reachable.
function deriveLimits(
  dl: DisplayList | null,
  info: SheetInfo,
  viewportHeight: number
): SelectionLimits {
  const rowH = medianTrack(dl?.grid?.rowOffsets, ROW_H);
  const colW = medianTrack(dl?.grid?.colOffsets, COL_W);
  const rows = Math.max(1, Math.round(info.contentHeight / rowH)) + 1;
  const cols = Math.max(1, Math.round(info.contentWidth / colW)) + 1;
  const rowsPerPage = Math.max(1, Math.floor(viewportHeight / rowH));
  return { rows, cols, rowsPerPage };
}

// trigger a browser download of a byte blob under the given name and mime type.
function downloadBytes(bytes: Uint8Array, name: string, mime: string): void {
  const blob = new Blob([new Uint8Array(bytes)], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

// the png download name derived from the workbook name: swap .xlsx for .png.
function pngName(fileName: string | undefined): string {
  return `${(fileName ?? 'workbook.xlsx').replace(/\.xlsx$/i, '')}.png`;
}

function scaledRect(rect: { x: number; y: number; w: number; h: number }, zoom: number) {
  return {
    x: rect.x * zoom,
    y: rect.y * zoom,
    w: rect.w * zoom,
    h: rect.h * zoom,
  };
}

const visuallyHidden: React.CSSProperties = {
  position: 'absolute',
  width: 1,
  height: 1,
  margin: -1,
  padding: 0,
  border: 0,
  overflow: 'hidden',
  clip: 'rect(0 0 0 0)',
  whiteSpace: 'nowrap',
};

const xlsxToolbarStyles: Record<string, React.CSSProperties> = {
  shell: {
    flex: '0 0 auto',
    minHeight: DEFAULT_XLSX_TOOLBAR_HEIGHT,
    padding: '4px 0 5px',
    borderBottom: '1px solid var(--xlsx-divider, #e2e8f0)',
    background: 'var(--xlsx-chrome-bg, #ffffff)',
    color: 'var(--xlsx-text, #0f172a)',
    boxSizing: 'border-box',
  },
  rail: {
    display: 'flex',
    alignItems: 'center',
    minHeight: 32,
    margin: '0 8px',
    padding: '2px 8px',
    borderRadius: 4,
    background: 'var(--xlsx-chrome-bg, #ffffff)',
    boxSizing: 'border-box',
    overflowX: 'auto',
    overflowY: 'hidden',
  },
  group: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: 1,
    padding: '0 6px',
    borderRight: '1px solid var(--xlsx-divider, rgba(226, 232, 240, 0.9))',
    flex: '0 0 auto',
  },
  formulaGroup: {
    display: 'flex',
    alignItems: 'center',
    gap: 4,
    flex: '1 1 320px',
    minWidth: 240,
    padding: '0 6px',
    borderRight: '1px solid var(--xlsx-divider, rgba(226, 232, 240, 0.9))',
  },
  nameBox: {
    appearance: 'none',
    width: 64,
    height: 28,
    flex: '0 0 auto',
    boxSizing: 'border-box',
    border: '1px solid var(--xlsx-border, #e2e8f0)',
    borderRadius: 6,
    background: 'var(--xlsx-input-bg, #f8fafc)',
    color: 'var(--xlsx-text, #0f172a)',
    font: '600 12px ui-monospace, SFMono-Regular, Menlo, monospace',
    textAlign: 'center',
    outlineColor: 'var(--xlsx-focus, #2563eb)',
  },
  formulaMark: {
    display: 'grid',
    placeItems: 'center',
    width: 20,
    height: 28,
    flex: '0 0 auto',
    color: 'var(--xlsx-text-muted, #64748b)',
    font: 'italic 700 12px Georgia, serif',
    userSelect: 'none',
  },
  formulaInput: {
    appearance: 'none',
    flex: '1 1 260px',
    minWidth: 140,
    height: 28,
    boxSizing: 'border-box',
    border: '1px solid var(--xlsx-border, #e2e8f0)',
    borderRadius: 6,
    padding: '0 8px',
    background: 'var(--xlsx-input-bg, #ffffff)',
    color: 'var(--xlsx-text, #0f172a)',
    font: `13px ${chromeFont}`,
    outlineColor: 'var(--xlsx-focus, #2563eb)',
  },
  flatShell: {
    flex: '0 0 auto',
    borderBottom: '1px solid var(--xlsx-divider, #e2e8f0)',
    background: 'var(--xlsx-chrome-bg, #ffffff)',
    color: 'var(--xlsx-text, #0f172a)',
  },
  flatFormulaGroup: {
    display: 'flex',
    alignItems: 'center',
    gap: 4,
    flex: 1,
    minWidth: 0,
  },
  flatRail: {
    display: 'flex',
    alignItems: 'center',
    height: 39,
    padding: '0 8px',
    boxSizing: 'border-box',
    overflowX: 'auto',
    overflowY: 'hidden',
  },
  proposals: {
    marginLeft: 'auto',
    paddingLeft: 6,
    flex: '0 0 auto',
  },
  count: {
    display: 'inline-grid',
    placeItems: 'center',
    minWidth: 15,
    height: 15,
    marginLeft: -3,
    padding: '0 4px',
    borderRadius: 8,
    background: 'var(--xlsx-text, #0f172a)',
    color: 'var(--xlsx-chrome-bg, #ffffff)',
    fontSize: 10,
    fontWeight: 700,
    lineHeight: 1,
    boxSizing: 'border-box',
  },
};

/**
 * The xlsx editor React component.
 */
export function XlsxEditor({
  i18n,
  icons,
  ...props
}: XlsxEditorProps) {
  return (
    <LocaleProvider i18n={i18n}>
      <IconSetContext.Provider value={icons ?? null}>
        <XlsxEditorContent {...props} />
      </IconSetContext.Provider>
    </LocaleProvider>
  );
}

function XlsxEditorContent({
  file,
  fileName,
  onSave,
  onChange,
  collaboration,
  onReady,
  onFirstPaint,
  onPendingChange,
  className,
  readOnly = false,
  singleRowToolbar = false,
  onCommandStateChange,
  showCustomNumberFormat = true,
  showSearchMenus = true,
  showProposals = true,
  showFontPicker = true,
  showFontSizePicker = true,
  showZoomControl = true,
  initialZoom = 1,
}: Omit<XlsxEditorProps, 'i18n' | 'icons'>) {
  const { t } = useTranslation();
  const collaborationEnabled = collaboration !== undefined;
  const collaborationClientId = collaboration?.clientId;
  const collaborationInitialUpdate = collaboration?.initialUpdate;
  const collaborationOnReplica = collaboration?.onReplica;
  const collaborationProvider = collaboration?.provider;
  const toolbarRef = useRef<HTMLDivElement>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const handleRef = useRef<WorkbookProxy | null>(null);
  const frameRef = useRef<DisplayList | null>(null);
  // the exact frame on screen, stored with the zoom it was painted at. scroll
  // repaints are rAF-coalesced and a mutation republishes the frame, so hit
  // testing reads this snapshot rather than the live scroll offset or the
  // current model — either would answer for pixels that are not on screen.
  const paintedRef = useRef<Painted | null>(null);
  const rafRef = useRef<number | null>(null);
  const editorInputRef = useRef<HTMLInputElement>(null);
  const formulaInputRef = useRef<HTMLInputElement>(null);
  const draggingRef = useRef(false);
  const clickStartRef = useRef<CellAddr | null>(null);
  // an in-flight chart drag: which chart, and where the pointer went down.
  const chartDragRef = useRef<{ id: string; clientX: number; clientY: number } | null>(null);
  // an arrow-key burst not yet landed: the chart and the logical px accumulated.
  const nudgeRef = useRef<{ id: string; dx: number; dy: number } | null>(null);
  const nudgeTimerRef = useRef<number | null>(null);
  const suppressBlurRef = useRef(false);
  const flushNudgeRef = useRef<() => void>(() => {});
  // commits the open draft: null without a workbook, else whether it landed.
  const settlePendingEditsRef = useRef<() => Promise<boolean> | null>(() => null);
  // the open cell or formula-bar draft and the sheet it was typed on: the
  // commit lands on that sheet while it is still the active one, else nowhere.
  const pendingDraftRef = useRef<(EditState & { sheet: number }) | null>(null);
  // the sheet a formula-bar draft started on.
  const formulaSheetIdRef = useRef('');
  // latest onReady, read (not depended on) by the open effect so a changing
  // callback identity never reopens the workbook.
  const onReadyRef = useRef(onReady);
  onReadyRef.current = onReady;
  const onFirstPaintRef = useRef(onFirstPaint);
  onFirstPaintRef.current = onFirstPaint;
  const firstPaintPendingRef = useRef(false);
  const onChangeRef = useRef(onChange);
  const runCommandRef = useRef<(command: XlsxCommand) => void>(() => {});
  const onCommandStateChangeRef = useRef(onCommandStateChange);
  onCommandStateChangeRef.current = onCommandStateChange;
  onChangeRef.current = onChange;
  const readOnlyRef = useRef(readOnly);
  readOnlyRef.current = readOnly;
  // one frame request in the worker at a time; a view change meanwhile asks
  // again once it is answered, so frames never queue up behind a scroll.
  const frameRequestRef = useRef({ inFlight: false, again: false });
  // the next frame opens the active sheet at its saved scroll.
  const initialScrollRef = useRef(false);
  // the last change posted, and the requests posted when a peer's change
  // last arrived: a frame older than either may be stale.
  const lastMutationRef = useRef(0);
  const remoteSeenRef = useRef(0);
  const revealSeqRef = useRef(0);

  // the last frame the worker drew, and what it says about the workbook.
  const [shown, setShown] = useState<WorkbookFrame | null>(null);
  const shownRef = useRef(shown);
  shownRef.current = shown;
  // a sheet switch posted and not drawn yet: the tab the user picked, and the
  // request that switched it.
  const [pendingSheet, setPendingSheet] = useState<{ index: number; seq: number } | null>(
    null
  );
  const sheetInfo = shown?.sheetInfo ?? null;
  const [error, setError] = useState<string | null>(null);
  const [renderError, setRenderError] = useState<string | null>(null);
  // the placeholder grid painted while no file is open.
  const [demo, setDemo] = useState<DisplayList | null>(null);
  const frame = file ? (shown?.displayList ?? null) : demo;
  const [selection, setSelection] = useState<Selection | null>(null);
  const selectionRef = useRef(selection);
  selectionRef.current = selection;
  const flushEditorRef = useRef<() => Promise<void>>(async () => {});
  const [editing, setEditing] = useState<EditState | null>(null);
  const [formulaDraft, setFormulaDraft] = useState<string | null>(null);
  const [toolbarHeight, setToolbarHeight] = useState(DEFAULT_XLSX_TOOLBAR_HEIGHT);
  const [zoom, setZoom] = useState(initialZoom);
  const zoomRef = useRef(zoom);
  zoomRef.current = zoom;
  // the sheet point at the grid's top-left when the zoom changed: it stays
  // there, as in Google Sheets, once the scroll area has the new size.
  const zoomAnchorRef = useRef<{ x: number; y: number } | null>(null);
  const changeZoom = useCallback((next: number) => {
    if (next === zoomRef.current) return;
    const scroll = scrollRef.current;
    if (scroll) zoomAnchorRef.current = zoomedViewport(scroll, zoomRef.current);
    setZoom(next);
  }, []);
  // keyboard moves and typing scroll the focus cell into view, after the
  // commit that moved it.
  const revealPendingRef = useRef(false);
  const [revealRequest, setRevealRequest] = useState(0);
  const requestReveal = useCallback(() => {
    revealPendingRef.current = true;
    setRevealRequest((n) => n + 1);
  }, []);
  // the scroll area's extent in sheet pixels when a revealed cell lies past
  // the used range, for the sheet it was revealed on.
  const [reach, setReach] = useState<{ sheet: number; width: number; height: number } | null>(
    null
  );
  // a revealed cell's position, kept while the scroll area grows to reach it.
  const revealCellRef = useRef<{ cell: CellBounds; row: number; col: number } | null>(null);
  const [dragging, setDragging] = useState(false);
  // the selected chart, and the live pointer offset while it is dragged.
  // `movable` rides along so the arrow keys never depend on a frame lookup.
  const [selectedChart, setSelectedChart] = useState<{ id: string; movable: boolean } | null>(
    null
  );
  const [chartDragOffset, setChartDragOffset] = useState<{ x: number; y: number } | null>(null);
  // logical-px preview of an arrow burst that has not landed yet.
  const [nudgeOffset, setNudgeOffset] = useState<{ x: number; y: number } | null>(null);
  const [borderStyleChoice, setBorderStyleChoice] = useState<SelectionFormatting['borderStyle']>();
  const [borderColorChoice, setBorderColorChoice] = useState<string>();
  const [capturedFormat, setCapturedFormat] = useState<CapturedFormat | null>(null);
  const paintSourceRef = useRef<string | null>(null);
  const [proposalsPanelOpen, setProposalsPanelOpen] = useState(false);
  const [collaborationReplica, setCollaborationReplica] =
    useState<CollaborationReplica | null>(null);
  const [awarenessPeers, setAwarenessPeers] = useState<readonly AwarenessPeer[]>([]);
  // a1 lists keyed by proposal id: cells that drifted since a proposal was
  // staged, surfaced when accepting it throws a StaleProposalError.
  const [staleFor, setStaleFor] = useState<Record<string, string[]>>({});
  // what the open workbook's engine was built with.
  const [features, setFeatures] = useState({ png: false, proposals: false });
  // committed cell inputs the worker has not drawn yet.
  const [optimistic, setOptimistic] = useState<OptimisticEdit[]>([]);
  // changes posted to the worker and not answered yet.
  const [inFlight, setInFlight] = useState(0);

  const activeSheet = pendingSheet?.index ?? sheetInfo?.activeSheet ?? 0;
  const activeSheetId = sheetInfo?.sheetIds[activeSheet] ?? '';
  // the frame on screen draws the sheet the user is on.
  const shownHere = shown && shown.sheetInfo.activeSheet === activeSheet ? shown : null;
  const selectionFormatting: SelectionFormatting = shownHere?.formatting ?? {};
  const historyState = shown?.history ?? NO_HISTORY;
  const mergedRanges = shownHere?.selectionMerged ?? [];
  const visibleMergedRanges = shownHere?.visibleMerged ?? [];
  const proposals = shown?.proposals ?? [];
  // overlays follow the frame on screen, so they stay on its pixels.
  const paintedZoom = file && shown ? shown.view.zoom : zoom;
  pendingDraftRef.current = editing
    ? { sheet: activeSheet, ...editing }
    : selection && formulaDraft !== null
      ? {
          sheet: activeSheet,
          sheetId: formulaSheetIdRef.current,
          ...selection.focus,
          value: formulaDraft,
        }
      : null;

  // the editable text of a cell, when the frame on screen is current: it was
  // built after every change posted and every peer's change seen, on this sheet.
  const knownInput = (row: number, col: number): string | undefined => {
    if (!shownHere || shownHere.seq < Math.max(lastMutationRef.current, remoteSeenRef.current))
      return undefined;
    const cell = shownHere.cells[`${row}:${col}`];
    if (cell) return cell.input;
    const focus = shownHere.view.selection?.focus;
    return focus && focus.row === row && focus.col === col ? shownHere.focus?.input : undefined;
  };

  // the name box and formula bar's cell: its address at once, its text from
  // the frame (or the input just committed to it) once the frame has it.
  const focusedCell: CellEdit | null = useMemo(() => {
    if (!selection || !sheetInfo) return null;
    const { row, col } = selection.focus;
    const a1 = cellAddress(row, col);
    const pending = optimistic.find(
      (entry) => entry.sheetId === activeSheetId && entry.row === row && entry.col === col
    );
    if (pending) return { a1, input: pending.text, isFormula: pending.text.startsWith('=') };
    const cell =
      shownHere?.cells[`${row}:${col}`] ??
      (shownHere?.view.selection && sameCell(shownHere.view.selection.focus, selection.focus)
        ? shownHere.focus
        : null);
    return cell ?? { a1, input: '', isFormula: false };
  }, [selection, sheetInfo, optimistic, activeSheetId, shownHere]);

  const reportError = useCallback((e: unknown) => {
    if (e instanceof SheetGoneError) return;
    setError(messageOf(e));
  }, []);

  // ask the worker for the frame of the current view. One request at a time:
  // a change meanwhile asks again once it is answered, at the next animation
  // frame.
  const requestFrame = useCallback(() => {
    const proxy = handleRef.current;
    const scroll = scrollRef.current;
    if (!proxy || !scroll || scroll.clientWidth === 0 || scroll.clientHeight === 0) return;
    const request = frameRequestRef.current;
    if (request.inFlight) {
      request.again = true;
      return;
    }
    request.inFlight = true;
    request.again = false;
    const initialScroll = initialScrollRef.current;
    initialScrollRef.current = false;
    const viewZoom = zoomRef.current;
    proxy
      .frame({
        viewport: zoomedViewport(scroll, viewZoom),
        zoom: viewZoom,
        selection: selectionRef.current,
        ...(initialScroll ? { initialScroll } : {}),
      })
      .then(
        (next) => {
          if (handleRef.current === proxy) setShown(next);
        },
        (e: unknown) => {
          if (handleRef.current !== proxy) return;
          frameRef.current = null;
          paintedRef.current = null;
          setRenderError(messageOf(e));
        }
      )
      .finally(() => {
        request.inFlight = false;
        if (!request.again || handleRef.current !== proxy) return;
        // what changed meanwhile is drawn at the next animation frame, with
        // the scroll (a reveal's included) it has by then.
        requestAnimationFrame(() => {
          if (request.again && handleRef.current === proxy) requestFrame();
        });
      });
  }, []);

  // a change posted to the worker: counted until it is answered, and the
  // frame showing it asked for.
  const track = useCallback(
    <T,>(proxy: WorkbookProxy, done: Promise<T>): Promise<T> => {
      lastMutationRef.current = proxy.posted;
      setInFlight((count) => count + 1);
      requestFrame();
      const settled = () => setInFlight((count) => count - 1);
      done.then(settled, settled);
      return done;
    },
    [requestFrame]
  );

  // a change landed: the host hears of it when it changed the workbook.
  const applyResult = useCallback((result: EditResult | undefined) => {
    if (result?.applied) onChangeRef.current?.();
  }, []);

  // a change on the active sheet, refused if a peer removed or moved it first.
  const onSheet = useCallback<WorkbookProxy['onSheet']>(
    (sheetId, method, ...args) => {
      const proxy = handleRef.current;
      if (!proxy) return Promise.reject(new Error('Workbook is still loading'));
      return track(proxy, proxy.onSheet(sheetId, method, ...args));
    },
    [track]
  );

  // a cell input committed: posted, and drawn over its cell until the
  // worker's frame shows it. A draft whose sheet a peer removed is dropped.
  const commitCell = useCallback(
    (sheetId: string, sheet: number, row: number, col: number, value: string) => {
      if (!handleRef.current) return Promise.resolve(false);
      const done = onSheet(sheetId, 'editCell', sheet, row, col, value);
      const seq = lastMutationRef.current;
      setOptimistic((list) => [
        ...list.filter(
          (entry) => entry.sheetId !== sheetId || entry.row !== row || entry.col !== col
        ),
        { seq, sheetId, row, col, text: value },
      ]);
      return done.then(
        (result) => {
          applyResult(result);
          return true;
        },
        (e: unknown) => {
          if (e instanceof SheetGoneError) return true;
          reportError(e);
          return false;
        }
      );
    },
    [onSheet, applyResult, reportError]
  );

  const clearSelection = useCallback(() => {
    if (!settlePendingEditsRef.current()) return;
    setSelection(null);
    setSelectedChart(null);
    setEditing(null);
    setFormulaDraft(null);
    setCapturedFormat(null);
    paintSourceRef.current = null;
  }, []);

  const selectCells = useCallback(
    async (sheet: number, nextSelection: Selection): Promise<boolean> => {
      const proxy = handleRef.current;
      if (!proxy || !Number.isInteger(sheet)) return false;
      try {
        const info = await proxy.sheetInfo();
        if (sheet < 0 || sheet >= info.sheetNames.length) return false;
        await proxy.cellPosition(sheet, nextSelection.anchor.row, nextSelection.anchor.col);
        const position = await proxy.cellPosition(
          sheet,
          nextSelection.focus.row,
          nextSelection.focus.col
        );
        if (handleRef.current !== proxy || !settlePendingEditsRef.current()) return false;
        const switched = track(proxy, proxy.onSheet(info.sheetIds[sheet], 'setActiveSheet', sheet));
        setPendingSheet({ index: sheet, seq: lastMutationRef.current });
        setSelection({
          anchor: { ...nextSelection.anchor },
          focus: { ...nextSelection.focus },
        });
        setSelectedChart(null);
        setEditing(null);
        setFormulaDraft(null);
        setCapturedFormat(null);
        paintSourceRef.current = null;
        requestAnimationFrame(() => {
          const scroll = scrollRef.current;
          if (!scroll) return;
          scroll.scrollLeft = position.x * zoomRef.current;
          scroll.scrollTop = position.y * zoomRef.current;
        });
        setError(null);
        await switched;
        return true;
      } catch {
        return false;
      }
    },
    [track]
  );

  useEffect(() => {
    if (!readOnly) return;
    setEditing(null);
    setFormulaDraft(null);
    setCapturedFormat(null);
    paintSourceRef.current = null;
    chartDragRef.current = null;
    setChartDragOffset(null);
    nudgeRef.current = null;
    setNudgeOffset(null);
    if (nudgeTimerRef.current != null) clearTimeout(nudgeTimerRef.current);
    nudgeTimerRef.current = null;
  }, [readOnly]);

  useEffect(() => {
    const toolbar = toolbarRef.current;
    if (!toolbar) return;
    const updateHeight = () => setToolbarHeight(toolbar.offsetHeight);
    updateHeight();
    const observer = new ResizeObserver(updateHeight);
    observer.observe(toolbar);
    return () => observer.disconnect();
  }, [readOnly]);

  // the pending proposals ride on every frame: a new one re-reads them and
  // repaints their ghosts after an external caller staged some.
  const refreshProposals = requestFrame;

  // open the workbook in its worker when the file changes; dispose it on
  // change/unmount and reset all editing state so a dropped file starts clean.
  useEffect(() => {
    setEditing(null);
    setFormulaDraft(null);
    setSelectedChart(null);
    setReach(null);
    // a burst belongs to the document it was typed on: its timer would fire
    // against whatever workbook `handleRef` holds by then.
    nudgeRef.current = null;
    setNudgeOffset(null);
    if (nudgeTimerRef.current != null) {
      clearTimeout(nudgeTimerRef.current);
      nudgeTimerRef.current = null;
    }
    setStaleFor({});
    setProposalsPanelOpen(false);
    setCollaborationReplica(null);
    setBorderStyleChoice(undefined);
    setBorderColorChoice(undefined);
    setCapturedFormat(null);
    setRenderError(null);
    setShown(null);
    setPendingSheet(null);
    setOptimistic([]);
    setInFlight(0);
    setFeatures({ png: false, proposals: false });
    paintSourceRef.current = null;
    paintedRef.current = null;
    frameRef.current = null;
    handleRef.current = null;
    setSelection(null);
    if (!file) {
      setError(null);
      return;
    }
    let proxy: WorkbookProxy | null = null;
    let unsubscribeUpdates = () => {};
    let cleanupReady = () => {};
    let disposed = false;
    const runReadyCleanup = () => {
      const cleanup = cleanupReady;
      cleanupReady = () => {};
      try {
        cleanup();
      } catch {}
    };
    void openWorkbookWorker(file, {
      collaborative: collaborationEnabled,
      clientId: collaborationClientId,
      initialUpdate: collaborationInitialUpdate,
    }).then(
      (opened) => {
        if (disposed) {
          opened.dispose();
          return;
        }
        proxy = opened;
        handleRef.current = opened;
        unsubscribeUpdates = opened.onUpdate((_update, origin) => {
          if (disposed || origin !== 'remote') return;
          // a peer's change: the next frame shows it (and drops a draft
          // whose sheet it removed).
          remoteSeenRef.current = opened.posted;
          setStaleFor({});
          setError(null);
          requestFrame();
        });
        initialScrollRef.current = true;
        firstPaintPendingRef.current = true;
        setFeatures({ png: opened.pngExportAvailable, proposals: opened.proposalsAvailable });
        setSelection(selectionAt({ row: 0, col: 0 }));
        setCollaborationReplica(opened);
        setError(null);
        const cleanup = onReadyRef.current?.({
          clearSelection,
          flush: () => flushEditorRef.current(),
          handle: opened,
          refreshProposals,
          focus: () => scrollRef.current?.focus(),
          save: async () => {
            const settled = settlePendingEditsRef.current();
            if (!settled || !(await settled)) {
              throw new Error('Could not commit pending workbook edits');
            }
            return opened.save();
          },
          selectCells,
          run: (command) => runCommandRef.current(command),
          sheetInfo: () => shownRef.current?.sheetInfo ?? null,
          visibleViewport: () => {
            const scroll = scrollRef.current;
            if (!scroll || !scroll.clientWidth || !scroll.clientHeight) return null;
            return zoomedViewport(scroll, zoomRef.current);
          },
        });
        if (typeof cleanup === 'function') cleanupReady = cleanup;
      },
      (e: unknown) => {
        if (disposed) return;
        handleRef.current = null;
        setSelection(null);
        setError(messageOf(e));
      }
    );
    return () => {
      disposed = true;
      runReadyCleanup();
      unsubscribeUpdates();
      proxy?.dispose();
      handleRef.current = null;
    };
  }, [
    file,
    collaborationEnabled,
    collaborationClientId,
    collaborationInitialUpdate,
    clearSelection,
    refreshProposals,
    requestFrame,
    selectCells,
  ]);

  useEffect(() => {
    if (!collaborationOnReplica || !collaborationReplica) return;
    collaborationOnReplica(collaborationReplica);
    return () => collaborationOnReplica(null);
  }, [collaborationOnReplica, collaborationReplica]);

  useEffect(() => {
    setAwarenessPeers([]);
    if (!collaborationProvider) return;
    return collaborationProvider.onAwareness((peers) => setAwarenessPeers([...peers]));
  }, [collaborationProvider]);

  useEffect(() => {
    return () => collaborationProvider?.setCursor(null);
  }, [collaborationProvider]);

  useEffect(() => {
    if (!collaborationProvider) return;
    const sheet = sheetInfo?.sheetIds[activeSheet];
    if (!selection || !sheet) {
      collaborationProvider.setCursor(null);
      return;
    }
    collaborationProvider.setCursor({
      sheet,
      anchor: { ...selection.anchor },
      head: { ...selection.focus },
    });
  }, [collaborationProvider, selection, sheetInfo, activeSheet]);

  // the placeholder grid, painted at once while no file is open.
  const paintDemo = useCallback(() => {
    const scroll = scrollRef.current;
    const canvas = canvasRef.current;
    const ctx = canvas?.getContext('2d');
    if (!scroll || !canvas || !ctx) return;
    const dpr = window.devicePixelRatio || 1;
    const w = scroll.clientWidth;
    const h = scroll.clientHeight;
    if (w === 0 || h === 0) return;
    canvas.width = Math.round(w * dpr);
    canvas.height = Math.round(h * dpr);
    canvas.style.width = `${w}px`;
    canvas.style.height = `${h}px`;
    const viewport = zoomedViewport(scroll, zoomRef.current);
    const dl = buildDemoDisplayList(viewport.width, viewport.height, t('editor.demoCellText'));
    paintDisplayList(ctx, dl, dpr * zoomRef.current);
    frameRef.current = dl;
    paintedRef.current = {
      frame: dl,
      zoom: zoomRef.current,
      viewport,
      handle: null,
      sheet: 0,
      seq: 0,
      canvas,
    };
    setRenderError(null);
    setDemo(dl);
  }, [t]);

  // a frame from the worker: painted in the same commit as the overlays it
  // places, so the selection, the open editor and the pixels never part.
  useLayoutEffect(() => {
    const scroll = scrollRef.current;
    const canvas = canvasRef.current;
    if (!shown || !scroll || !canvas) return;
    const { view } = shown;
    if (view.initialScroll) {
      scroll.scrollLeft = view.viewport.x * view.zoom;
      scroll.scrollTop = view.viewport.y * view.zoom;
    }
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const dpr = window.devicePixelRatio || 1;
    const w = view.viewport.width * view.zoom;
    const h = view.viewport.height * view.zoom;
    canvas.width = Math.round(w * dpr);
    canvas.height = Math.round(h * dpr);
    canvas.style.width = `${w}px`;
    canvas.style.height = `${h}px`;
    paintDisplayList(ctx, shown.displayList, dpr * view.zoom);
    frameRef.current = shown.displayList;
    paintedRef.current = {
      frame: shown.displayList,
      zoom: view.zoom,
      viewport: view.viewport,
      handle: handleRef.current,
      sheet: shown.sheetInfo.activeSheet,
      seq: shown.seq,
      canvas,
    };
    setRenderError(null);
    setOptimistic((list) =>
      list.some((entry) => entry.seq < shown.seq)
        ? list.filter((entry) => entry.seq > shown.seq)
        : list
    );
    const switching = pendingSheet !== null && shown.seq < pendingSheet.seq;
    if (pendingSheet && !switching) setPendingSheet(null);
    // a peer removed or moved away the sheet a draft is open on: drop the
    // draft rather than land it on whichever sheet is active now.
    const draft = pendingDraftRef.current;
    const sheetId = shown.sheetInfo.sheetIds[shown.sheetInfo.activeSheet];
    if (!switching && draft && draft.sheetId !== sheetId) {
      pendingDraftRef.current = null;
      // nothing keeps the focus, so the keys still being typed do nothing
      // until a click instead of editing the sheet now active.
      suppressBlurRef.current = true;
      editorInputRef.current?.blur();
      suppressBlurRef.current = false;
      formulaInputRef.current?.blur();
      setEditing(null);
      setFormulaDraft(null);
    }
    scroll.dispatchEvent(
      new CustomEvent(XLSX_FRAME_PRESENTED_EVENT, {
        bubbles: true,
        detail: { epoch: shown.epoch, seq: shown.seq },
      })
    );
    if (firstPaintPendingRef.current) {
      firstPaintPendingRef.current = false;
      onFirstPaintRef.current?.();
    }
  }, [shown]);

  // a zoom change: the anchor goes back under the top-left corner, and the
  // frame at the new zoom is asked for (the demo grid is drawn at once).
  useLayoutEffect(() => {
    const scroll = scrollRef.current;
    const anchor = zoomAnchorRef.current;
    zoomAnchorRef.current = null;
    if (!scroll || !anchor) return;
    scroll.scrollLeft = anchor.x * zoom;
    scroll.scrollTop = anchor.y * zoom;
    if (handleRef.current) requestFrame();
    else if (!file) paintDemo();
  }, [zoom, requestFrame, paintDemo, file]);

  // the selection rides on every frame (the focus cell's text, the range's
  // formatting): a new one asks for the frame that describes it.
  useEffect(() => {
    requestFrame();
  }, [selection, requestFrame]);

  // repaint on scroll/resize (rAF-coalesced): the worker draws the window the
  // scroll container shows; with no file, the placeholder grid.
  useEffect(() => {
    const scroll = scrollRef.current;
    if (!scroll) return;
    const schedulePaint = () => {
      if (rafRef.current != null) return;
      rafRef.current = requestAnimationFrame(() => {
        rafRef.current = null;
        if (!handleRef.current) {
          if (!file) paintDemo();
          return;
        }
        // a frame already drew this exact view.
        const painted = paintedRef.current;
        if (painted?.handle && paintedIsLive(painted, scroll, zoomRef.current)) return;
        requestFrame();
      });
    };
    if (!file) paintDemo();
    scroll.addEventListener('scroll', schedulePaint, { passive: true });
    const observer = new ResizeObserver(schedulePaint);
    observer.observe(scroll);
    return () => {
      scroll.removeEventListener('scroll', schedulePaint);
      observer.disconnect();
      if (rafRef.current != null) {
        cancelAnimationFrame(rafRef.current);
        rafRef.current = null;
      }
    };
  }, [file, paintDemo, requestFrame]);

  // clear a stuck drag if the mouse is released outside the grid.
  useEffect(() => {
    const stop = () => {
      draggingRef.current = false;
      setDragging(false);
    };
    window.addEventListener('mouseup', stop);
    return () => window.removeEventListener('mouseup', stop);
  }, []);

  // rebuilt from the live frame so the offscreen mirror never lags a mutation;
  // the visible window is small, so a rebuild per paint frame is cheap enough.
  const a11yGrid = useMemo(() => {
    if (!frame || !sheetInfo) return null;
    return buildA11yGrid(frame, selection, sheetInfo.sheetNames[activeSheet] ?? '', {
      gridLabel: t('a11y.gridLabel'),
      rowHeaderLabel: t('a11y.rowHeaderLabel'),
      columnHeaderLabel: t('a11y.columnHeaderLabel'),
      cellLabel: t('a11y.cellLabel'),
      cellLabelSelected: t('a11y.cellLabelSelected'),
      emptyCellLabel: t('a11y.emptyCellLabel'),
      emptyCellLabelSelected: t('a11y.emptyCellLabelSelected'),
    });
  }, [frame, selection, sheetInfo, activeSheet, t]);

  // preventScroll everywhere: the sticky overlay host sits below the full-height
  // canvas in flow, so a plain focus() scrolls the grid to bring it into view.
  const focusContainer = useCallback(() => {
    scrollRef.current?.focus({ preventScroll: true });
  }, []);

  // focus the in-cell editor when it opens, without scrolling the grid. it
  // stays mounted while its cell is off screen, so focus never leaves it.
  useEffect(() => {
    if (editing) editorInputRef.current?.focus({ preventScroll: true });
  }, [editing]);

  settlePendingEditsRef.current = () => {
    const handle = handleRef.current;
    if (!handle) return null;
    flushNudgeRef.current();
    chartDragRef.current = null;
    setChartDragOffset(null);
    setDragging(false);
    const draft = pendingDraftRef.current;
    if (!draft || readOnlyRef.current) return Promise.resolve(true);
    // the live input stays current during IME composition, when state lags.
    const value = editing ? (editorInputRef.current?.value ?? draft.value) : draft.value;
    pendingDraftRef.current = null;
    suppressBlurRef.current = true;
    editorInputRef.current?.blur();
    suppressBlurRef.current = false;
    setEditing(null);
    setFormulaDraft(null);
    return commitCell(draft.sheetId, draft.sheet, draft.row, draft.col, value);
  };

  const selectedRangeA1 = useCallback((target: Selection): string => {
    const range = normalizeRange(target);
    return `${cellAddress(range.top, range.left)}:${cellAddress(range.bottom, range.right)}`;
  }, []);

  const limits = useCallback((): SelectionLimits => {
    return deriveLimits(
      frameRef.current,
      sheetInfo!,
      (scrollRef.current?.clientHeight ?? 0) / zoom
    );
  }, [sheetInfo, zoom]);

  // map a viewport-local pointer event to a sheet cell via the frame geometry.
  const pointToCell = useCallback(
    (clientX: number, clientY: number): CellAddr | null => {
      const canvas = canvasRef.current;
      const grid = frameRef.current?.grid;
      if (!canvas || !grid) return null;
      const rect = canvas.getBoundingClientRect();
      return cellAtPoint(grid, (clientX - rect.left) / zoom, (clientY - rect.top) / zoom);
    },
    [zoom]
  );

  // which chart a pointer event lands on. containment runs over the regions the
  // painted frame published — engine geometry, engine clipping, engine paint
  // order — so no geometry is rebuilt here and the answer cannot drift from the
  // pixels by a scroll frame or a mutation the canvas has not drawn yet.
  const pointToChart = useCallback((clientX: number, clientY: number): ChartRegion | null => {
    const canvas = canvasRef.current;
    const painted = paintedRef.current;
    if (!canvas || !painted) return null;
    const rect = canvas.getBoundingClientRect();
    return chartRegionAtPoint(
      painted.frame.charts,
      (clientX - rect.left) / painted.zoom,
      (clientY - rect.top) / painted.zoom
    );
  }, []);

  const selectedChartRegion = useMemo(
    () =>
      selectedChart
        ? (frame?.charts?.find((chart) => chart.id === selectedChart.id) ?? null)
        : null,
    [frame, selectedChart]
  );

  // a selected chart that scrolled out of the painted frame has no outline to
  // show, so drop it rather than leave an invisible selection swallowing keys.
  useEffect(() => {
    if (!selectedChart || !frame) return;
    if (frame.charts?.some((chart) => chart.id === selectedChart.id)) return;
    // land what the burst already earned before the selection goes away, and
    // drop any armed drag with it.
    flushNudgeRef.current();
    chartDragRef.current = null;
    setChartDragOffset(null);
    setSelectedChart(null);
  }, [frame, selectedChart]);

  // slide a chart through the engine's edit path, so the new anchor is
  // undoable and reaches the drawing part on save.
  const moveChartBy = useCallback(
    (id: string, dx: number, dy: number) => {
      if (!handleRef.current || readOnly || (dx === 0 && dy === 0)) return;
      onSheet(activeSheetId, 'moveChart', activeSheet, id, dx, dy).then(applyResult, reportError);
    },
    [activeSheet, activeSheetId, applyResult, onSheet, readOnly, reportError]
  );

  // land a run of arrow nudges as one edit. key repeat fires as fast as the os
  // pleases and every move is a whole-workbook semantic sync, so presses only
  // accumulate a local delta and preview it; the burst lands once it settles.
  const flushNudge = useCallback(() => {
    if (nudgeTimerRef.current != null) {
      clearTimeout(nudgeTimerRef.current);
      nudgeTimerRef.current = null;
    }
    const pending = nudgeRef.current;
    nudgeRef.current = null;
    setNudgeOffset(null);
    if (!pending) return;
    moveChartBy(pending.id, pending.dx, pending.dy);
  }, [moveChartBy]);

  flushNudgeRef.current = flushNudge;

  const nudgeChart = useCallback((id: string, dx: number, dy: number) => {
    const pending = nudgeRef.current;
    if (pending && pending.id !== id) flushNudgeRef.current();
    const base = nudgeRef.current ?? { id, dx: 0, dy: 0 };
    const next = { id, dx: base.dx + dx, dy: base.dy + dy };
    nudgeRef.current = next;
    onPendingChange?.(true);
    setNudgeOffset({ x: next.dx, y: next.dy });
    if (nudgeTimerRef.current != null) clearTimeout(nudgeTimerRef.current);
    nudgeTimerRef.current = setTimeout(() => {
      nudgeTimerRef.current = null;
      // a burst that returns to where it started is not an edit.
      const settled = nudgeRef.current;
      nudgeRef.current = null;
      setNudgeOffset(null);
      if (settled && (settled.dx !== 0 || settled.dy !== 0)) {
        moveChartByRef.current(settled.id, settled.dx, settled.dy);
      }
    }, CHART_NUDGE_SETTLE_MS) as unknown as number;
  }, [onPendingChange]);

  const moveChartByRef = useRef(moveChartBy);
  moveChartByRef.current = moveChartBy;

  // commit a chart drag on release, as one edit for the whole gesture. a drag
  // the window loses (focus leaves mid-gesture) is dropped, not committed
  // later against whatever the pointer has since moved over.
  useEffect(() => {
    const commit = (event: MouseEvent) => {
      const drag = chartDragRef.current;
      // any release ends the gesture; only a primary one lands it, so a
      // non-primary release cannot leave the drag armed for a later mouseup.
      chartDragRef.current = null;
      setChartDragOffset(null);
      if (!drag || event.button !== 0) return;
      moveChartBy(
        drag.id,
        (event.clientX - drag.clientX) / zoom,
        (event.clientY - drag.clientY) / zoom
      );
    };
    const cancel = () => {
      chartDragRef.current = null;
      setChartDragOffset(null);
      flushNudgeRef.current();
    };
    window.addEventListener('mouseup', commit);
    window.addEventListener('blur', cancel);
    return () => {
      window.removeEventListener('mouseup', commit);
      window.removeEventListener('blur', cancel);
    };
  }, [moveChartBy, zoom]);

  // on unmount this only stops the timer: cleanups run in hook order, so the
  // open effect above has already disposed the workbook and a flush here would
  // find no handle to move a chart on. an unfinished burst is dropped.
  useEffect(
    () => () => {
      if (nudgeTimerRef.current != null) clearTimeout(nudgeTimerRef.current);
      nudgeTimerRef.current = null;
      nudgeRef.current = null;
    },
    []
  );

  const openEditor = useCallback(
    (seed?: string) => {
      const handle = handleRef.current;
      if (!handle || !selection || readOnly) return;
      const { row, col } = selection.focus;
      const sheetId = activeSheetId;
      const known = seed ?? knownInput(row, col);
      if (known !== undefined) {
        setEditing({ row, col, value: known, sheetId });
        requestReveal();
        return;
      }
      // the frame on screen is not current: the worker reads the cell after
      // every change posted before it. A key typed meanwhile opens its own edit.
      void handle
        .onSheet(sheetId, 'cell', activeSheet, row, col)
        .then(
          (cell) => cell.input,
          () => ''
        )
        .then((value) => {
          const focus = selectionRef.current?.focus;
          if (handleRef.current !== handle || readOnlyRef.current) return;
          if (!focus || !sameCell(focus, { row, col })) return;
          setEditing((current) => current ?? { row, col, value, sheetId });
          requestReveal();
        });
    },
    [selection, activeSheet, activeSheetId, readOnly, requestReveal, knownInput]
  );

  // commit the open editor, optionally stepping the selection like excel.
  // `refocus: false` from the input's own blur: focus is already going
  // elsewhere, and moving it back would cancel that (Chromium).
  const commitEditor = useCallback(
    (move?: Direction, refocus = true) => {
      // the open draft (null once a peer's change dropped it).
      const draft = pendingDraftRef.current;
      if (!handleRef.current || !editing || !draft || readOnly) return;
      if (refocus) suppressBlurRef.current = true;
      const { row, col, value, sheetId } = editing;
      void commitCell(sheetId, draft.sheet, row, col, value);
      setEditing(null);
      const base = selectionAt({ row, col });
      setSelection(move ? moveFocus(base, move, { limits: limits() }) : base);
      if (move) requestReveal();
      if (refocus) focusContainer();
    },
    [editing, commitCell, limits, focusContainer, readOnly, requestReveal]
  );

  const cancelEditor = useCallback(() => {
    suppressBlurRef.current = true;
    setEditing(null);
    focusContainer();
  }, [focusContainer]);

  const clearCells = useCallback(() => {
    if (!handleRef.current || !selection || readOnly) return;
    const r = normalizeRange(selection);
    const edits: CellInputEdit[] = [];
    for (let row = r.top; row <= r.bottom; row++) {
      for (let col = r.left; col <= r.right; col++) edits.push({ row, col, input: '' });
    }
    onSheet(activeSheetId, 'editCells', activeSheet, edits).then(applyResult, reportError);
  }, [selection, activeSheet, activeSheetId, applyResult, onSheet, readOnly, reportError]);

  const copySelection = useCallback(async () => {
    const handle = handleRef.current;
    if (!handle || !selection) return;
    try {
      const cells = await handle.onSheet(
        activeSheetId,
        'rangeCells',
        activeSheet,
        selectedRangeA1(selection)
      );
      const tsv = toTsv(
        cells.map((row) => row.map((c) => ({ input: c.input, isFormula: c.isFormula })))
      );
      await navigator.clipboard.writeText(tsv);
    } catch {
      // clipboard denied or read failed — nothing to paste, leave state as-is.
    }
  }, [selection, activeSheet, activeSheetId, selectedRangeA1]);

  const cutSelection = useCallback(async () => {
    await copySelection();
    clearCells();
  }, [copySelection, clearCells]);

  const pasteSelection = useCallback(async () => {
    const handle = handleRef.current;
    if (!handle || !selection || readOnly) return;
    let text: string;
    try {
      text = await navigator.clipboard.readText();
    } catch {
      return;
    }
    if (readOnlyRef.current || handleRef.current !== handle) return;
    const grid = fromTsv(text);
    if (grid.length === 0) return;
    const r = normalizeRange(selection);
    const edits: CellInputEdit[] = [];
    let width = 1;
    grid.forEach((rowArr, dr) => {
      width = Math.max(width, rowArr.length);
      rowArr.forEach((input, dc) => edits.push({ row: r.top + dr, col: r.left + dc, input }));
    });
    onSheet(activeSheetId, 'editCells', activeSheet, edits).then(applyResult, reportError);
    setSelection({
      anchor: { row: r.top, col: r.left },
      focus: { row: r.top + grid.length - 1, col: r.left + width - 1 },
    });
  }, [selection, activeSheet, activeSheetId, applyResult, onSheet, readOnly, reportError]);


  const formatSelection = useCallback(
    (action: FormattingAction) => {
      const handle = handleRef.current;
      if (!handle || !selection || readOnly) return;
      const range = selectedRangeA1(selection);
      if (action === 'paintFormat') {
        if (capturedFormat) {
          setCapturedFormat(null);
          paintSourceRef.current = null;
          return;
        }
        const normalized = normalizeRange(selection);
        const source = `${activeSheet}:${normalized.top}:${normalized.left}:${normalized.bottom}:${normalized.right}`;
        handle.onSheet(activeSheetId, 'captureFormat', activeSheet, range).then((format) => {
          setCapturedFormat(format);
          paintSourceRef.current = source;
        }, reportError);
        return;
      }
      const sheetId = activeSheetId;
      const sheet = activeSheet;
      let result: Promise<EditResult>;
      if (action === 'currency') {
        result = onSheet(sheetId, 'setNumberFormat', sheet, range, 'currency');
      } else if (action === 'percent') {
        result = onSheet(sheetId, 'setNumberFormat', sheet, range, 'percent');
      } else if (action === 'increaseDecimal') {
        result = onSheet(sheetId, 'setNumberFormat', sheet, range, 'increaseDecimal');
      } else if (action === 'decreaseDecimal') {
        result = onSheet(sheetId, 'setNumberFormat', sheet, range, 'decreaseDecimal');
      } else if (action === 'bold' || action === 'italic' || action === 'strikethrough') {
        // flipped from the range's state when the worker reaches it.
        result = track(handle, handle.toggle(sheetId, range, action));
      } else if (action.type === 'numberFormat') {
        result =
          action.value === 'custom'
            ? onSheet(sheetId, 'setNumberFormat', sheet, range, {
                type: 'custom',
                pattern: selectionFormatting.numberFormatPattern ?? '0.00',
              })
            : onSheet(sheetId, 'setNumberFormat', sheet, range, action.value);
      } else if (action.type === 'fontFamily') {
        result = onSheet(sheetId, 'patchRangeStyle', sheet, range, { fontFamily: action.value });
      } else if (action.type === 'fontSize') {
        result = onSheet(sheetId, 'patchRangeStyle', sheet, range, { fontSize: action.value });
      } else if (action.type === 'textColor') {
        result = onSheet(sheetId, 'patchRangeStyle', sheet, range, { textColor: action.value });
      } else if (action.type === 'fillColor') {
        result = onSheet(sheetId, 'patchRangeStyle', sheet, range, { fillColor: action.value });
      } else if (action.type === 'clearColor') {
        result = onSheet(sheetId, 'patchRangeStyle', sheet, range, { clear: [action.value] });
      } else if (action.type === 'borderPreset') {
        result = onSheet(sheetId, 'patchRangeStyle', sheet, range, {
          border: {
            preset: action.value,
            style: borderStyleChoice ?? selectionFormatting.borderStyle ?? 'solid',
            color: borderColorChoice ?? selectionFormatting.borderColor ?? '#000000',
          },
        });
      } else if (action.type === 'borderStyle') {
        setBorderStyleChoice(action.value);
        result = onSheet(sheetId, 'patchRangeStyle', sheet, range, {
          border: { style: action.value },
        });
      } else if (action.type === 'borderColor') {
        setBorderColorChoice(action.value);
        result = onSheet(sheetId, 'patchRangeStyle', sheet, range, {
          border: { color: action.value },
        });
      } else if (action.type === 'horizontalAlignment') {
        result = onSheet(sheetId, 'patchRangeStyle', sheet, range, {
          horizontalAlignment: action.value,
        });
      } else if (action.type === 'verticalAlignment') {
        result = onSheet(sheetId, 'patchRangeStyle', sheet, range, {
          verticalAlignment: action.value,
        });
      } else {
        result = onSheet(sheetId, 'patchRangeStyle', sheet, range, {
          textWrapping: action.value,
        });
      }
      result.then(applyResult, reportError);
      focusContainer();
    },
    [
      selection,
      selectedRangeA1,
      capturedFormat,
      activeSheet,
      activeSheetId,
      selectionFormatting,
      borderStyleChoice,
      borderColorChoice,
      applyResult,
      focusContainer,
      onSheet,
      readOnly,
      reportError,
      track,
    ]
  );

  useEffect(() => {
    if (!handleRef.current || !selection || !capturedFormat || dragging || readOnly) return;
    const normalized = normalizeRange(selection);
    const key = `${activeSheet}:${normalized.top}:${normalized.left}:${normalized.bottom}:${normalized.right}`;
    if (key === paintSourceRef.current) return;
    onSheet(
      activeSheetId,
      'applyFormat',
      activeSheet,
      selectedRangeA1(selection),
      capturedFormat
    ).then(applyResult, reportError);
    setCapturedFormat(null);
    paintSourceRef.current = null;
    focusContainer();
  }, [
    selection,
    capturedFormat,
    dragging,
    activeSheet,
    activeSheetId,
    selectedRangeA1,
    applyResult,
    focusContainer,
    onSheet,
    readOnly,
    reportError,
  ]);

  const undo = useCallback(() => {
    const handle = handleRef.current;
    if (!handle || readOnly) return;
    flushNudgeRef.current();
    track(handle, handle.undo()).then(applyResult, reportError);
  }, [applyResult, readOnly, reportError, track]);

  const redo = useCallback(() => {
    const handle = handleRef.current;
    if (!handle || readOnly) return;
    flushNudgeRef.current();
    track(handle, handle.redo()).then(applyResult, reportError);
  }, [applyResult, readOnly, reportError, track]);

  const print = useCallback(() => window.print(), []);

  const searchMenus = useCallback(() => undefined, []);

  const mergeSelection = useCallback(
    (action: MergeAction) => {
      const handle = handleRef.current;
      if (!handle || !selection || readOnly) return;
      const range = normalizeRange(selection);
      const cell = (row: number, col: number) => ({ row, col });
      const mergeRange = (top: number, left: number, bottom: number, right: number) => ({
        start: cell(top, left),
        end: cell(bottom, right),
      });
      const sheetId = activeSheetId;
      if (action === 'unmerge') {
        // the merges the selection holds when the worker reaches it.
        handle
          .onSheet(sheetId, 'mergedRanges', activeSheet, selectedRangeA1(selection))
          .then((merged) => {
            if (merged.length === 0) return;
            return onSheet(
              sheetId,
              'applyOps',
              merged.map((range) => ({
                type: 'unmergeCells',
                sheet: activeSheet,
                range: mergeRange(range.start.row, range.start.col, range.end.row, range.end.col),
              }))
            ).then(applyResult);
          })
          .then(undefined, reportError);
        focusContainer();
        return;
      }
      const ops: unknown[] = [];
      if (action === 'all') {
        ops.push({
          type: 'mergeCells',
          sheet: activeSheet,
          range: mergeRange(range.top, range.left, range.bottom, range.right),
        });
      } else if (action === 'horizontal') {
        for (let row = range.top; row <= range.bottom; row++) {
          ops.push({
            type: 'mergeCells',
            sheet: activeSheet,
            range: mergeRange(row, range.left, row, range.right),
          });
        }
      } else {
        for (let col = range.left; col <= range.right; col++) {
          ops.push({
            type: 'mergeCells',
            sheet: activeSheet,
            range: mergeRange(range.top, col, range.bottom, col),
          });
        }
      }
      onSheet(sheetId, 'applyOps', ops).then(applyResult, reportError);
      focusContainer();
    },
    [
      selection,
      activeSheet,
      activeSheetId,
      selectedRangeA1,
      applyResult,
      focusContainer,
      onSheet,
      readOnly,
      reportError,
    ]
  );

  // accept a proposal (optionally forcing past drift): apply it, drop any stale
  // warning, refresh the list, repaint, and return focus to the grid.
  const acceptProposal = useCallback(
    (id: string, force?: boolean) => {
      const handle = handleRef.current;
      if (!handle || readOnly) return;
      track(handle, handle.acceptProposal(id, { force })).then(
        (result) => {
          applyResult(result);
          setStaleFor(({ [id]: _dropped, ...rest }) => rest);
          focusContainer();
        },
        (e: unknown) => {
          if (e instanceof StaleProposalError) {
            setStaleFor((m) => ({ ...m, [id]: e.cells }));
            refreshProposals();
          } else reportError(e);
        }
      );
    },
    [applyResult, refreshProposals, focusContainer, readOnly, reportError, track]
  );

  // reject a proposal: drop it and its warning; the next frame draws the
  // canvas without its ghost and its border chrome.
  const rejectProposal = useCallback(
    (id: string) => {
      const handle = handleRef.current;
      if (!handle || readOnly) return;
      track(handle, handle.rejectProposal(id)).then(
        () => setStaleFor(({ [id]: _dropped, ...rest }) => rest),
        reportError
      );
    },
    [readOnly, reportError, track]
  );

  const save = useCallback(async () => {
    const handle = handleRef.current;
    if (!handle) return;
    const settled = settlePendingEditsRef.current();
    if (!settled || !(await settled)) return;
    try {
      const bytes = await handle.save();
      if (onSave) onSave(bytes);
      else downloadBytes(bytes, fileName ?? 'workbook.xlsx', XLSX_MIME);
    } catch (e) {
      setError(messageOf(e));
    }
  }, [onSave, fileName]);

  // render the current scroll window to png via the raster backend and download
  // it — the same display list the canvas paints, rasterized in the core.
  const exportPng = useCallback(async () => {
    const handle = handleRef.current;
    const scroll = scrollRef.current;
    if (!handle || !scroll) return;
    try {
      const png = await handle.renderPng(zoomedViewport(scroll, zoom));
      downloadBytes(png, pngName(fileName), 'image/png');
    } catch (e) {
      setError(messageOf(e));
    }
  }, [fileName, zoom]);

  // commit the formula bar draft to the focused cell.
  const commitFormula = useCallback(
    (move?: Direction) => {
      const draft = pendingDraftRef.current;
      if (!handleRef.current || !selection || formulaDraft == null || !draft || readOnly) return;
      const { row, col } = selection.focus;
      void commitCell(draft.sheetId, draft.sheet, row, col, formulaDraft);
      setFormulaDraft(null);
      if (move) {
        setSelection((prev) => (prev ? moveFocus(prev, move, { limits: limits() }) : prev));
        requestReveal();
      }
    },
    [selection, formulaDraft, commitCell, limits, readOnly, requestReveal]
  );

  flushEditorRef.current = async () => {
    if (!handleRef.current) throw new Error('Workbook is still loading');
    const settled = settlePendingEditsRef.current();
    if (!settled || !(await settled)) {
      throw new Error('Could not commit pending workbook edits');
    }
  };
  // pending: input typed and not committed, or committed and not yet in the
  // workbook (its change still on the way to the worker).
  useEffect(() => {
    onPendingChange?.(
      editing !== null || formulaDraft !== null || nudgeOffset !== null || inFlight > 0
    );
  }, [editing, formulaDraft, nudgeOffset, inFlight, onPendingChange]);

  // grid-level keyboard: chrome shortcuts first, then the pure selection reducer.
  const onKeyDown = useCallback(
    (e: React.KeyboardEvent) => {
      const handle = handleRef.current;
      if (!handle || !selection || !sheetInfo || editing) return;
      const mod = e.metaKey || e.ctrlKey;
      const lower = e.key.toLowerCase();

      // a selected chart owns the keyboard: arrows nudge it, escape drops it,
      // and nothing else reaches the cells hidden behind it — the grid overlay
      // is suppressed while it is selected, so a delete or a keystroke there
      // would edit a target the user cannot see. undo/redo/save stay global.
      if (selectedChart && !(mod && CHART_GLOBAL_KEYS.has(lower))) {
        if (e.key === 'Escape') {
          // escape cancels the whole gesture, pointer or keyboard: an armed
          // drag must not land a move the user just abandoned.
          chartDragRef.current = null;
          setChartDragOffset(null);
          nudgeRef.current = null;
          setNudgeOffset(null);
          if (nudgeTimerRef.current != null) {
            clearTimeout(nudgeTimerRef.current);
            nudgeTimerRef.current = null;
          }
          setSelectedChart(null);
          e.preventDefault();
          return;
        }
        const nudge = mod ? undefined : CHART_NUDGE_KEYS[e.key];
        if (nudge) {
          const step = CHART_NUDGE_PX * (e.shiftKey ? CHART_NUDGE_MULTIPLIER : 1);
          if (!readOnly && selectedChart.movable) {
            nudgeChart(selectedChart.id, nudge[0] * step, nudge[1] * step);
          }
          e.preventDefault();
        }
        return;
      }

      if (mod) {
        if (lower === 'c') {
          void copySelection();
          e.preventDefault();
          return;
        }
        if (lower === 'v') {
          void pasteSelection();
          e.preventDefault();
          return;
        }
        if (lower === 'x') {
          void cutSelection();
          e.preventDefault();
          return;
        }
        if (lower === 'z') {
          e.shiftKey ? redo() : undo();
          e.preventDefault();
          return;
        }
        if (lower === 'y') {
          redo();
          e.preventDefault();
          return;
        }
        if (lower === 's') {
          void save();
          e.preventDefault();
          return;
        }
      }

      const action = selectionKeyReducer(
        selection,
        {
          key: e.key,
          shiftKey: e.shiftKey,
          metaKey: e.metaKey,
          ctrlKey: e.ctrlKey,
          altKey: e.altKey,
        },
        limits()
      );
      switch (action.type) {
        case 'move':
          setSelection(action.selection);
          // select all keeps the view, as the menu's Select all does.
          if (!(mod && lower === 'a')) requestReveal();
          e.preventDefault();
          break;
        case 'startEdit':
          openEditor(action.initialInput);
          e.preventDefault();
          break;
        case 'clear':
          clearCells();
          e.preventDefault();
          break;
        case 'none':
          break;
      }
    },
    [
      selection,
      sheetInfo,
      editing,
      limits,
      copySelection,
      pasteSelection,
      cutSelection,
      undo,
      redo,
      save,
      openEditor,
      clearCells,
      selectedChart,
      nudgeChart,
      readOnly,
      requestReveal,
    ]
  );

  const onMouseDown = useCallback(
    (e: React.MouseEvent) => {
      // the editor input is a dom overlay above the canvas, so a press inside
      // it is the editor's own — it places a caret, and must not commit, reach
      // a chart painted under it, or move the grid.
      const dismissesEditor = editing != null;
      if (dismissesEditor && editorInputRef.current?.contains(e.target as Node)) return;
      // a release the window never saw would otherwise leave the last gesture
      // armed and commit its accumulated delta on some later, unrelated mouseup.
      chartDragRef.current = null;
      // a pointer gesture ends any arrow burst, so the two never interleave.
      flushNudgeRef.current();

      // a chart takes the press as a whole object: it finishes any open edit,
      // then selects, leaving the grid selection where the edit left it. the
      // commit is stated here rather than left to the blur `focusContainer`
      // triggers, so both branches finish the edit for the same visible reason.
      const chart = pointToChart(e.clientX, e.clientY);
      if (chart) {
        if (dismissesEditor) commitEditor();
        setSelectedChart({ id: chart.id, movable: !readOnly && chart.movable });
        // only a primary press starts a drag: a right-press opens a context
        // menu whose release this window never sees, and an armed drag would
        // then land on whatever the next unrelated click released over.
        if (!readOnly && chart.movable && e.button === 0) {
          chartDragRef.current = { id: chart.id, clientX: e.clientX, clientY: e.clientY };
        }
        setChartDragOffset(null);
        // no click start: a press on a chart is not a cell click, so it must
        // not follow a hyperlink in whatever cell sits behind it.
        clickStartRef.current = null;
        focusContainer();
        e.preventDefault();
        return;
      }

      setSelectedChart(null);
      setChartDragOffset(null);
      // an async reopen leaves the previous frame painted, so a chart here can
      // outlive `selection` for a moment. that is fine: the chart branch has
      // already returned above, and selecting a chart the user can still see is
      // what should happen. only the cell path needs a selection to extend.
      if (!selection) return;
      const addr = pointToCell(e.clientX, e.clientY);
      if (!addr) return;
      if (dismissesEditor) commitEditor();
      if (e.shiftKey) setSelection((prev) => (prev ? extendTo(prev, addr, limits()) : prev));
      else setSelection(selectionAt(addr));
      // no click start: a press that dismissed the editor must not also follow a
      // hyperlink in the cell it selects.
      clickStartRef.current = dismissesEditor ? null : addr;
      draggingRef.current = true;
      setDragging(true);
      focusContainer();
    },
    [editing, selection, pointToCell, pointToChart, limits, focusContainer, commitEditor, readOnly]
  );

  const onMouseMove = useCallback(
    (e: React.MouseEvent<HTMLDivElement>) => {
      const chartDrag = chartDragRef.current;
      if (chartDrag) {
        // the button came back up somewhere this window never heard about, so
        // the gesture is over; do not keep accumulating it.
        if (e.buttons === 0) {
          chartDragRef.current = null;
          setChartDragOffset(null);
          return;
        }
        e.currentTarget.style.cursor = 'move';
        // a tooltip from whatever cell was hovered before must not ride along.
        e.currentTarget.title = '';
        setChartDragOffset({
          x: e.clientX - chartDrag.clientX,
          y: e.clientY - chartDrag.clientY,
        });
        return;
      }
      const addr = pointToCell(e.clientX, e.clientY);
      if (!addr) {
        if (!draggingRef.current) {
          e.currentTarget.style.cursor = 'default';
          e.currentTarget.title = '';
        }
        return;
      }
      if (!draggingRef.current) {
        const chart = pointToChart(e.clientX, e.clientY);
        if (chart) {
          // a pinned chart still selects on click, so it must not read as a cell.
          e.currentTarget.style.cursor = !readOnly && chart.movable ? 'move' : 'pointer';
          e.currentTarget.title = '';
          return;
        }
        const hyperlink = frameRef.current
          ? hyperlinkAtCell(frameRef.current, addr.row, addr.col)
          : null;
        e.currentTarget.style.cursor = hyperlink ? 'pointer' : 'default';
        e.currentTarget.title = hyperlink?.tooltip ?? '';
        return;
      }
      setSelection((prev) => (prev ? extendTo(prev, addr, limits()) : prev));
    },
    [pointToCell, pointToChart, limits, readOnly]
  );

  const activateHyperlink = useCallback(
    (addr: CellAddr): boolean => {
      const handle = handleRef.current;
      const currentFrame = frameRef.current;
      if (!handle || !currentFrame || !sheetInfo) return false;
      const hyperlink = hyperlinkAtCell(currentFrame, addr.row, addr.col);
      if (!hyperlink) return false;
      const href = safeExternalHyperlink(hyperlink);
      if (href) {
        window.open(href, '_blank', 'noopener,noreferrer');
        return true;
      }
      if (!hyperlink.location) return true;
      const currentName = sheetInfo.sheetNames[activeSheet] ?? '';
      const destination = parseHyperlinkLocation(hyperlink.location, currentName);
      if (!destination) return true;
      const targetSheet = sheetInfo.sheetNames.findIndex(
        (name) => name.toLowerCase() === destination.sheetName.toLowerCase()
      );
      if (targetSheet < 0) return true;
      track(handle, handle.onSheet(sheetInfo.sheetIds[targetSheet], 'setActiveSheet', targetSheet))
        .then(() => handle.cellPosition(targetSheet, destination.row, destination.col))
        .then((position) => {
          const scroll = scrollRef.current;
          if (!scroll) return;
          scroll.scrollLeft = position.x * zoomRef.current;
          scroll.scrollTop = position.y * zoomRef.current;
        }, reportError);
      setPendingSheet({ index: targetSheet, seq: lastMutationRef.current });
      setSelection(selectionAt({ row: destination.row, col: destination.col }));
      setEditing(null);
      setFormulaDraft(null);
      setCapturedFormat(null);
      paintSourceRef.current = null;
      setError(null);
      return true;
    },
    [activeSheet, sheetInfo, reportError, track]
  );


  const onClick = useCallback(
    (e: React.MouseEvent) => {
      if (editing) {
        clickStartRef.current = null;
        return;
      }
      const addr = pointToCell(e.clientX, e.clientY);
      const start = clickStartRef.current;
      clickStartRef.current = null;
      if (!addr || !start || addr.row !== start.row || addr.col !== start.col) return;
      activateHyperlink(addr);
    },
    [activateHyperlink, editing, pointToCell]
  );

  const onDoubleClick = useCallback(
    (e: React.MouseEvent) => {
      if (editing) return;
      if (pointToChart(e.clientX, e.clientY)) return;
      const addr = pointToCell(e.clientX, e.clientY);
      if (
        addr &&
        frameRef.current &&
        hyperlinkAtCell(frameRef.current, addr.row, addr.col)
      ) {
        return;
      }
      if (selection) openEditor();
    },
    [editing, openEditor, pointToCell, pointToChart, selection]
  );

  const onMouseLeave = useCallback((e: React.MouseEvent<HTMLDivElement>) => {
    e.currentTarget.style.cursor = 'default';
    e.currentTarget.title = '';
  }, []);

  // overlays sit on the frame on screen, and only while it draws this sheet.
  const grid = !file || shownHere ? frame?.grid : undefined;

  const renderedSelection = selection
    ? expandRangeToMergedCells(normalizeRange(selection), visibleMergedRanges)
    : null;
  const renderedFocus = selection
    ? expandRangeToMergedCells(
        {
          top: selection.focus.row,
          left: selection.focus.col,
          bottom: selection.focus.row,
          right: selection.focus.col,
        },
        visibleMergedRanges
      )
    : null;
  const selRect = grid && renderedSelection ? rangeRect(grid, renderedSelection) : null;
  const focusRect = grid && renderedFocus ? rangeRect(grid, renderedFocus) : null;
  const editRect = grid && editing ? cellRect(grid, editing.row, editing.col) : null;
  const scaledSelectionRect = selRect ? scaledRect(selRect, paintedZoom) : null;
  const scaledFocusRect = focusRect ? scaledRect(focusRect, paintedZoom) : null;
  const scaledEditRect = editRect ? scaledRect(editRect, paintedZoom) : null;
  // committed inputs on this sheet the worker has not drawn yet, over their cells.
  const optimisticCells = grid
    ? optimistic.flatMap((entry) => {
        if (entry.sheetId !== activeSheetId) return [];
        const rect = cellRect(grid, entry.row, entry.col);
        return rect ? [{ ...entry, rect: scaledRect(rect, paintedZoom) }] : [];
      })
    : [];

  // the selected chart's outline, placed from the engine-published region and
  // offset by the live drag so the box tracks the pointer before it commits.
  const chartOutlineRect = selectedChartRegion
    ? scaledRect(selectedChartRegion.rect, paintedZoom)
    : null;

  // the scroll area: the used range, or as far as a revealed cell needed.
  const ownReach = reach?.sheet === activeSheet ? reach : null;
  const reachWidth = Math.max(sheetInfo?.contentWidth ?? 0, ownReach?.width ?? 0);
  const reachHeight = Math.max(sheetInfo?.contentHeight ?? 0, ownReach?.height ?? 0);
  const reachRef = useRef({ width: reachWidth, height: reachHeight });
  reachRef.current = { width: reachWidth, height: reachHeight };
  const spacerWidth = sheetInfo ? reachWidth * zoom : undefined;
  const spacerHeight = sheetInfo ? reachHeight * zoom : undefined;

  // whether the frame on screen shows a cell whole. Only a frame that is the
  // live view answers: after a scroll not painted yet, the engine does. A track
  // clamped against a frozen pane starts exactly at its edge, so that start is
  // not proof the cell is whole.
  const shownWhole = (row: number, col: number): boolean => {
    const scroll = scrollRef.current;
    const painted = paintedRef.current;
    const grid = painted?.frame.grid;
    if (
      !scroll ||
      !painted ||
      !grid ||
      !sheetInfo ||
      painted.sheet !== activeSheet ||
      painted.seq < lastMutationRef.current ||
      !paintedIsLive(painted, scroll, zoom)
    )
      return false;
    const rect = cellRect(grid, row, col);
    if (!rect) return false;
    // on one axis: the cell fits between the pane and the window's edge, or
    // spans all of it (a cell wider than the view shows as much as it can).
    // Scrolled, a start exactly at the pane edge may be a clamped track.
    const whole = (
      start: number,
      size: number,
      pane: number | undefined,
      extent: number,
      scrolled: number
    ) =>
      pane !== undefined &&
      (((start > pane || (start === pane && (pane === 0 || scrolled === 0))) &&
        start + size <= extent) ||
        (start <= pane && start + size >= extent));
    const { frozenCols, frozenRows } = sheetInfo;
    return (
      (col < frozenCols ||
        whole(
          rect.x,
          rect.w,
          frozenCols ? grid.colOffsets[frozenCols] : 0,
          scroll.clientWidth / zoom,
          painted.viewport.x
        )) &&
      (row < frozenRows ||
        whole(
          rect.y,
          rect.h,
          frozenRows ? grid.rowOffsets[frozenRows] : 0,
          scroll.clientHeight / zoom,
          painted.viewport.y
        ))
    );
  };

  // scroll so the focus cell at `cell` shows whole: the frozen panes stay put.
  const revealTo = (cell: CellBounds, row: number, col: number) => {
    const scroll = scrollRef.current;
    const grid = frameRef.current?.grid;
    const info = shownRef.current?.sheetInfo;
    if (!scroll || !grid || !info) return;
    const zoom = zoomRef.current;
    const { width: reachWidth, height: reachHeight } = reachRef.current;
    const width = scroll.clientWidth / zoom;
    const height = scroll.clientHeight / zoom;
    const left = scroll.scrollLeft / zoom;
    const top = scroll.scrollTop / zoom;
    // the frozen panes' extent (it does not move with the scroll); undefined
    // while they fill the window.
    const paneWidth = info.frozenCols ? grid.colOffsets[info.frozenCols] : 0;
    const paneHeight = info.frozenRows ? grid.rowOffsets[info.frozenRows] : 0;
    const axis = (
      scrolled: number,
      pinned: boolean,
      start: number,
      size: number,
      extent: number,
      pane: number | undefined
    ) =>
      pinned || pane === undefined || extent <= pane
        ? scrolled
        : revealOffset(scrolled, start, start + size, extent - pane);
    const x = axis(left, col < info.frozenCols, cell.x, cell.width, width, paneWidth);
    const y = axis(top, row < info.frozenRows, cell.y, cell.height, height, paneHeight);
    if (x === left && y === top) return;
    // whole pixels, rounded toward the cell so no sliver of it stays hidden: a
    // start alignment rounds down, an end alignment up.
    const snap = (to: number, start: number) =>
      to === start ? Math.floor(to * zoom) : Math.ceil(to * zoom);
    const scrollLeft = x === left ? scroll.scrollLeft : snap(x, cell.x);
    const scrollTop = y === top ? scroll.scrollTop : snap(y, cell.y);
    // a cell past the used range: grow the scroll area first, then come back.
    const needWidth = (scrollLeft + scroll.clientWidth) / zoom;
    const needHeight = (scrollTop + scroll.clientHeight) / zoom;
    if (needWidth > reachWidth || needHeight > reachHeight) {
      revealCellRef.current = { cell, row, col };
      setReach({
        sheet: info.activeSheet,
        width: Math.max(reachWidth, needWidth),
        height: Math.max(reachHeight, needHeight),
      });
      return;
    }
    scroll.scrollLeft = scrollLeft;
    scroll.scrollTop = scrollTop;
    // the worker draws the new window; an editor opened by this keystroke
    // stays mounted off screen until that frame places it on its cell.
    requestFrame();
  };

  useLayoutEffect(() => {
    // a revealed cell past the used range, once the scroll area reaches it.
    const waiting = revealCellRef.current;
    revealCellRef.current = null;
    if (waiting) return revealTo(waiting.cell, waiting.row, waiting.col);
    const handle = handleRef.current;
    if (!revealPendingRef.current || !handle || !frameRef.current?.grid || !selection || !sheetInfo)
      return;
    revealPendingRef.current = false;
    const { row, col } = selection.focus;
    if (shownWhole(row, col)) return;
    // the worker places the cell after every change posted before; a later
    // reveal supersedes this one.
    const request = ++revealSeqRef.current;
    handle.onSheet(activeSheetId, 'cellPosition', activeSheet, row, col).then(
      (cell) => {
        if (request === revealSeqRef.current && handleRef.current === handle)
          revealTo(cell, row, col);
      },
      () => {}
    );
  }, [revealRequest, reach]);

  const formulaValue = formulaDraft ?? focusedCell?.input ?? '';
  const normalizedSelection = selection ? normalizeRange(selection) : null;
  const selectionRows = normalizedSelection
    ? normalizedSelection.bottom - normalizedSelection.top + 1
    : 1;
  const selectionColumns = normalizedSelection
    ? normalizedSelection.right - normalizedSelection.left + 1
    : 1;

  // switch sheets: retarget the core, reset scroll + selection; the next
  // frame draws the sheet at its saved scroll. A tab names its sheet by id (a
  // peer may move it meanwhile); a sheet just added is switched to by index.
  const switchSheet = (index: number, sheetId?: string) => {
    const handle = handleRef.current;
    if (!handle) return;
    // the burst belongs to the sheet it was typed on.
    if (!settlePendingEditsRef.current()) return;
    initialScrollRef.current = true;
    track(
      handle,
      sheetId === undefined
        ? handle.setActiveSheet(index)
        : handle.onSheet(sheetId, 'setActiveSheet', index)
    ).then(undefined, (e: unknown) => {
      setPendingSheet(null);
      reportError(e);
    });
    setPendingSheet({ index, seq: lastMutationRef.current });
    setSelection(selectionAt({ row: 0, col: 0 }));
    setSelectedChart(null);
    setEditing(null);
    setFormulaDraft(null);
    setCapturedFormat(null);
    paintSourceRef.current = null;
    // keys go to the new sheet's grid, not the tab that was clicked.
    focusContainer();
  };

  // ops on the selection that no toolbar button carries, as one undo step.
  const applyStructure = (ops: unknown[]) => {
    if (!handleRef.current || readOnly || !settlePendingEditsRef.current()) return false;
    onSheet(activeSheetId, 'applyOps', ops).then(applyResult, reportError);
    return true;
  };


  const freeze = (rows: number, columns: number) =>
    applyStructure([freezePaneOp(activeSheet, rows, columns)]);

  runCommandRef.current = (command: XlsxCommand) => {
    if (readOnly && xlsxCommandEdits(command)) return;
    const range = selection ? normalizeRange(selection) : null;
    const rows = range ? range.bottom - range.top + 1 : 0;
    const columns = range ? range.right - range.left + 1 : 0;
    const sheet = activeSheet;
    if (command === 'undo') return undo();
    if (command === 'redo') return redo();
    if (command === 'deleteValues') return clearCells();
    if (command === 'bold' || command === 'italic' || command === 'strikethrough')
      return formatSelection(command);
    if (command === 'insertSheet') {
      if (!sheetInfo) return;
      const index = activeSheet + 1;
      if (applyStructure([{ type: 'addSheet', index, name: newSheetName(sheetInfo.sheetNames) }]))
        switchSheet(index);
      return;
    }
    if (command === 'selectAll') {
      if (!sheetInfo) return;
      const { rows: lastRows, cols: lastCols } = limits();
      setSelection({
        anchor: { row: 0, col: 0 },
        focus: { row: lastRows - 1, col: lastCols - 1 },
      });
      // The grid takes the keys next, Ctrl/Cmd+C first.
      focusContainer();
      return;
    }
    const [kind, value] = command.split(':') as [string, string | undefined];
    if (kind === 'zoom') return changeZoom(Number(value) / 100);
    if (kind === 'merge') return mergeSelection(value as MergeAction);
    if (kind === 'numberFormat')
      return formatSelection({ type: 'numberFormat', value: value as NumberFormat });
    if (kind === 'align')
      return formatSelection({
        type: 'horizontalAlignment',
        value: value as HorizontalAlignment,
      });
    if (kind === 'valign')
      return formatSelection({ type: 'verticalAlignment', value: value as VerticalAlignment });
    if (kind === 'wrap')
      return formatSelection({ type: 'textWrapping', value: value as TextWrapping });
    if (!range || !sheetInfo) return;
    const amount = (value === 'current' ? value : Number(value)) as FreezeAmount;
    if (kind === 'freezeRows')
      return void freeze(frozenCount(amount, selection!.focus.row), sheetInfo.frozenCols);
    if (kind === 'freezeColumns')
      return void freeze(sheetInfo.frozenRows, frozenCount(amount, selection!.focus.col));
    const cells = {
      start: { row: range.top, col: range.left },
      end: { row: range.bottom, col: range.right },
    };
    const ops: Record<string, unknown[]> = {
      insertRowAbove: [{ type: 'insertRows', sheet, at: range.top, count: rows }],
      insertRowBelow: [{ type: 'insertRows', sheet, at: range.bottom + 1, count: rows }],
      insertColumnLeft: [{ type: 'insertCols', sheet, at: range.left, count: columns }],
      insertColumnRight: [{ type: 'insertCols', sheet, at: range.right + 1, count: columns }],
      deleteRows: [{ type: 'deleteRows', sheet, at: range.top, count: rows }],
      deleteColumns: [{ type: 'deleteCols', sheet, at: range.left, count: columns }],
      clearFormatting: [
        {
          type: 'patchRangeStyle',
          sheet,
          range: cells,
          patch: { clear: [...ALL_STYLE_PROPERTIES] },
        },
        { type: 'setRangeNumberFormat', sheet, range: cells, format: { type: 'automatic' } },
      ],
    };
    if (ops[command]) applyStructure(ops[command]);
  };

  const focusRow = selection?.focus.row ?? -1;
  const focusColumn = selection?.focus.col ?? -1;
  const frozenRows = sheetInfo?.frozenRows ?? 0;
  const frozenColumns = sheetInfo?.frozenCols ?? 0;
  const canUnmerge = mergedRanges.length > 0;
  const commandState = useMemo<XlsxCommandState>(
    () => ({
      canUndo: historyState.canUndo,
      canRedo: historyState.canRedo,
      selection:
        focusRow < 0
          ? null
          : { rows: selectionRows, columns: selectionColumns, focusRow, focusColumn },
      frozenRows,
      frozenColumns,
      formatting: selectionFormatting,
      canMerge: selectionRows > 1 || selectionColumns > 1,
      canUnmerge,
      zoom,
    }),
    [
      historyState,
      focusRow,
      focusColumn,
      selectionRows,
      selectionColumns,
      frozenRows,
      frozenColumns,
      selectionFormatting,
      canUnmerge,
      zoom,
    ]
  );
  useEffect(() => {
    onCommandStateChangeRef.current?.(commandState);
  }, [commandState]);

  return (
    <div
      className={className}
      role="application"
      aria-label={t('editor.appLabel')}
      style={{
        position: 'relative',
        display: 'flex',
        flexDirection: 'column',
        width: '100%',
        height: '100%',
        minWidth: 0,
        color: 'var(--xlsx-text, #202124)',
        background: 'var(--xlsx-chrome-bg, #ffffff)',
        fontFamily: chromeFont,
      }}
    >
      {/* Read-only keeps the toolbar and formula bar, disabled, so the grid stays put. */}
      <div
          ref={toolbarRef}
          data-testid="xlsx-toolbar"
          style={singleRowToolbar ? xlsxToolbarStyles.flatShell : xlsxToolbarStyles.shell}
        >
        <EditorToolbar
          currentFormatting={{
            ...selectionFormatting,
            borderStyle: borderStyleChoice ?? selectionFormatting.borderStyle,
            borderColor: borderColorChoice ?? selectionFormatting.borderColor,
            paintFormat: capturedFormat !== null,
          }}
          selectionShape={{
            rows: selectionRows,
            columns: selectionColumns,
            canUnmerge: mergedRanges.length > 0,
          }}
          onSearchMenus={searchMenus}
          onUndo={undo}
          onRedo={redo}
          canUndo={historyState.canUndo}
          canRedo={historyState.canRedo}
          onPrint={print}
          zoom={zoom}
          onZoomChange={changeZoom}
          onFormat={formatSelection}
          onMerge={mergeSelection}
          singleRow={singleRowToolbar}
          showSearchMenus={showSearchMenus}
          showCustomNumberFormat={showCustomNumberFormat}
          showFontPicker={showFontPicker}
          showFontSizePicker={showFontSizePicker}
          showZoomControl={showZoomControl}
          disabled={readOnly}
        >
          <EditorToolbar.Toolbar />
          <div
            style={singleRowToolbar ? xlsxToolbarStyles.flatRail : xlsxToolbarStyles.rail}
            role="group"
            aria-label={t('toolbar.formulaBarLabel')}
          >
            {!singleRowToolbar && (
              <ToolbarGroup
                style={{ ...xlsxToolbarStyles.group, paddingLeft: 0 }}
                label={t('toolbar.fileActionsLabel')}
              >
                <ToolbarButton
                  testId="xlsx-save"
                  onClick={save}
                  disabled={!sheetInfo || readOnly}
                  title={t('toolbar.save')}
                >
                  <ToolbarIcon name="save" size={18} />
                </ToolbarButton>
                <ToolbarButton
                  testId="xlsx-export-png"
                  onClick={exportPng}
                  disabled={!sheetInfo || !features.png}
                  title={t('toolbar.exportPng')}
                >
                  <ToolbarIcon name="image" size={18} />
                </ToolbarButton>
              </ToolbarGroup>
            )}
            <div
              style={
                singleRowToolbar
                  ? xlsxToolbarStyles.flatFormulaGroup
                  : xlsxToolbarStyles.formulaGroup
              }
              role="group"
              aria-label={t('toolbar.formulaBarLabel')}
            >
              <input
                data-testid="xlsx-name-box"
                readOnly
                value={focusedCell?.a1 ?? ''}
                placeholder={t('toolbar.nameBoxPlaceholder')}
                aria-label={t('toolbar.nameBoxPlaceholder')}
                style={xlsxToolbarStyles.nameBox}
              />
              <span style={xlsxToolbarStyles.formulaMark} aria-hidden="true">
                fx
              </span>
              <input
                ref={formulaInputRef}
                data-testid="xlsx-formula-input"
                value={formulaValue}
                placeholder={t('toolbar.formulaPlaceholder')}
                aria-label={t('toolbar.formulaPlaceholder')}
                disabled={!sheetInfo}
                // Read-only, not disabled: its text still selects and copies.
                readOnly={readOnly}
                onChange={(e) => {
                  onPendingChange?.(true);
                  if (formulaDraft === null) formulaSheetIdRef.current = activeSheetId;
                  setFormulaDraft(e.target.value);
                }}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') {
                    commitFormula(e.shiftKey ? 'up' : 'down');
                    focusContainer();
                    e.preventDefault();
                  } else if (e.key === 'Escape') {
                    setFormulaDraft(null);
                    focusContainer();
                    e.preventDefault();
                  }
                }}
                // an app or tab switch keeps the draft, as it keeps a cell edit.
                onBlur={() => {
                  if (document.hasFocus()) commitFormula();
                }}
                style={xlsxToolbarStyles.formulaInput}
              />
            </div>
            <PresenceStrip
              peers={awarenessPeers}
              sheetIds={sheetInfo?.sheetIds ?? []}
              sheetNames={sheetInfo?.sheetNames ?? []}
              activeSheet={activeSheet}
            />
            {features.proposals && showProposals && (
              <div style={xlsxToolbarStyles.proposals}>
                <ToolbarButton
                  testId="xlsx-proposals-button"
                  onClick={() => setProposalsPanelOpen((open) => !open)}
                  disabled={!sheetInfo || readOnly}
                  active={proposalsPanelOpen}
                  ariaExpanded={proposalsPanelOpen}
                  title={t('proposals.panelLabel')}
                  style={{ width: proposals.length > 0 ? 42 : 28 }}
                >
                  <ToolbarIcon name="proposals" size={18} />
                  {proposals.length > 0 && (
                    <span data-testid="xlsx-proposals-count" style={xlsxToolbarStyles.count}>
                      {proposals.length}
                    </span>
                  )}
                </ToolbarButton>
              </div>
            )}
          </div>
        </EditorToolbar>
        </div>
      {!readOnly && features.proposals && showProposals && proposalsPanelOpen && (
        <ProposalsPanel
          proposals={proposals}
          staleFor={staleFor}
          onAccept={acceptProposal}
          onReject={rejectProposal}
          style={{
            top: toolbarHeight + 4,
            right: 8,
            width: 'min(320px, calc(100% - 16px))',
            maxHeight: `min(420px, calc(100% - ${toolbarHeight + 12}px))`,
          }}
        />
      )}

      <div
        ref={scrollRef}
        className="xlsx-editor__scroll"
        data-testid="xlsx-scroll"
        tabIndex={0}
        onKeyDown={onKeyDown}
        onMouseDown={onMouseDown}
        onMouseMove={onMouseMove}
        onMouseLeave={onMouseLeave}
        onClick={onClick}
        onDoubleClick={onDoubleClick}
        style={{ position: 'relative', flex: 1, overflow: 'auto', minHeight: 0, outline: 'none' }}
      >
        <div
          style={{
            position: 'absolute',
            top: 0,
            left: 0,
            width: spacerWidth ?? '100%',
            height: spacerHeight ?? '100%',
          }}
        />
        {/* one sticky layer pins the canvas and overlays to the viewport top-left
            so overlay children share the canvas's coordinate space — a separate
            sticky sibling would sit below the full-height canvas in flow and
            scroll-jump when a child (the in-cell editor) is focused. */}
        <div style={{ position: 'sticky', top: 0, left: 0, width: 0, height: 0 }}>
          <canvas
            ref={canvasRef}
            style={{ display: 'block', position: 'absolute', top: 0, left: 0 }}
          />

          <div
            data-testid="xlsx-overlay-host"
            style={{
              position: 'absolute',
              top: 0,
              left: 0,
              width: 0,
              height: 0,
              pointerEvents: 'none',
            }}
          >
            {scaledSelectionRect && !selectedChart && (
              <div
                data-testid="xlsx-selection"
                style={{
                  position: 'absolute',
                  left: scaledSelectionRect.x,
                  top: scaledSelectionRect.y,
                  width: scaledSelectionRect.w,
                  height: scaledSelectionRect.h,
                  boxSizing: 'border-box',
                  border: `1px solid ${BRAND}`,
                  background: 'rgba(33, 115, 70, 0.12)',
                }}
              />
            )}
            {chartOutlineRect && (
              <div
                data-testid="xlsx-chart-selection"
                data-chart-id={selectedChart?.id}
                aria-hidden
                style={{
                  position: 'absolute',
                  left:
                    chartOutlineRect.x +
                    (chartDragOffset?.x ?? 0) +
                    (nudgeOffset?.x ?? 0) * paintedZoom,
                  top:
                    chartOutlineRect.y +
                    (chartDragOffset?.y ?? 0) +
                    (nudgeOffset?.y ?? 0) * paintedZoom,
                  width: chartOutlineRect.w,
                  height: chartOutlineRect.h,
                  boxSizing: 'border-box',
                  border: `2px solid ${BRAND}`,
                  boxShadow: '0 1px 6px rgba(0, 0, 0, 0.25)',
                  background: chartDragOffset ? 'rgba(33, 115, 70, 0.08)' : 'transparent',
                }}
              />
            )}
            {scaledFocusRect && !selectedChart && !editing && (
              <div
                style={{
                  position: 'absolute',
                  left: scaledFocusRect.x,
                  top: scaledFocusRect.y,
                  width: scaledFocusRect.w,
                  height: scaledFocusRect.h,
                  boxSizing: 'border-box',
                  border: `2px solid ${BRAND}`,
                }}
              />
            )}
            <RemoteSelections
              peers={awarenessPeers}
              grid={grid}
              sheetIds={sheetInfo?.sheetIds ?? []}
              activeSheet={activeSheet}
              zoom={paintedZoom}
              mergedRanges={visibleMergedRanges}
            />
            {optimisticCells.map((entry) => (
              <div
                key={`${entry.row}:${entry.col}`}
                data-testid="xlsx-pending-cell"
                style={{
                  position: 'absolute',
                  left: entry.rect.x,
                  top: entry.rect.y,
                  width: entry.rect.w,
                  height: entry.rect.h,
                  boxSizing: 'border-box',
                  padding: '0 3px',
                  overflow: 'hidden',
                  whiteSpace: 'nowrap',
                  lineHeight: `${entry.rect.h}px`,
                  textAlign: NUMBER_INPUT.test(entry.text) ? 'right' : 'left',
                  fontSize: 13 * paintedZoom,
                  fontFamily: 'system-ui, sans-serif',
                  color: '#202124',
                  background: '#ffffff',
                }}
              >
                {entry.text}
              </div>
            ))}
            {/* stays mounted and focused while its cell is scrolled away, so
                keys and composition keep landing in the edit; the next key
                scrolls the cell back, as in Excel and Sheets. */}
            {!readOnly && editing && (
              <input
                ref={editorInputRef}
                data-testid="xlsx-cell-editor"
                value={editing.value}
                onChange={(e) => {
                  onPendingChange?.(true);
                  if (!shownWhole(editing.row, editing.col)) requestReveal();
                  setEditing((prev) => (prev ? { ...prev, value: e.target.value } : prev));
                }}
                onKeyDown={(e) => {
                  e.stopPropagation();
                  // a key that types (or composes) brings the cell back whole.
                  if (
                    !NON_TYPING_KEYS.has(e.key) &&
                    !e.ctrlKey &&
                    !e.metaKey &&
                    !shownWhole(editing.row, editing.col)
                  )
                    requestReveal();
                  if (e.key === 'Enter') {
                    commitEditor(e.shiftKey ? 'up' : 'down');
                    e.preventDefault();
                  } else if (e.key === 'Tab') {
                    commitEditor(e.shiftKey ? 'left' : 'right');
                    e.preventDefault();
                  } else if (e.key === 'Escape') {
                    cancelEditor();
                    e.preventDefault();
                  }
                }}
                onBlur={(e) => {
                  if (suppressBlurRef.current) {
                    suppressBlurRef.current = false;
                    return;
                  }
                  // the window or tab lost focus: the edit stays open and takes
                  // the next key on return, as in Excel and Sheets.
                  if (!document.hasFocus()) return;
                  commitEditor(undefined, false);
                  // a press on nothing focusable: once focus has settled on the
                  // page body, the grid takes the keys again (a focus move a blur
                  // handler redirects can be dropped).
                  if (!e.relatedTarget)
                    requestAnimationFrame(() => {
                      if (document.hasFocus() && document.activeElement === document.body)
                        focusContainer();
                    });
                }}
                style={
                  scaledEditRect
                    ? {
                        position: 'absolute',
                        left: scaledEditRect.x,
                        top: scaledEditRect.y,
                        width: scaledEditRect.w,
                        height: scaledEditRect.h,
                        boxSizing: 'border-box',
                        border: `2px solid ${BRAND}`,
                        padding: '0 3px',
                        font: `${13 * paintedZoom}px system-ui, sans-serif`,
                        background: '#ffffff',
                        pointerEvents: 'auto',
                        outline: 'none',
                      }
                    : OFFSCREEN_EDITOR
                }
              />
            )}
          </div>
        </div>
      </div>

      {a11yGrid && (
        <>
          <div style={visuallyHidden} role="grid" aria-label={a11yGrid.label}>
            <div role="row">
              <span role="columnheader" />
              {a11yGrid.columnHeaders.map((h) => (
                <span key={h.col} role="columnheader">
                  {h.label}
                </span>
              ))}
            </div>
            {a11yGrid.rows.map((r) => (
              <div key={r.row} role="row">
                <span role="rowheader">{r.header}</span>
                {r.cells.map((c) => (
                  <span key={c.col} role="gridcell" aria-selected={c.selected}>
                    {c.label}
                  </span>
                ))}
              </div>
            ))}
          </div>
          {a11yGrid.charts.map((chart, index) => (
            <div key={`${index}:${chart.label}`} style={visuallyHidden} role="img" aria-label={chart.label} />
          ))}
        </>
      )}

      {renderError && (
        <div
          data-testid="xlsx-render-error"
          role="alert"
          style={{
            position: 'absolute',
            inset: 0,
            display: 'grid',
            placeItems: 'center',
            padding: 16,
            textAlign: 'center',
            color: 'var(--xlsx-error, #b00020)',
            background: 'var(--xlsx-chrome-bg, #ffffff)',
          }}
        >
          {renderError}
        </div>
      )}

      {error && (
        <div
          data-testid="xlsx-error"
          role="alert"
          style={{
            position: 'absolute',
            inset: 0,
            display: 'grid',
            placeItems: 'center',
            padding: 16,
            textAlign: 'center',
            color: 'var(--xlsx-error, #b00020)',
          }}
        >
          {t('editor.openError')}: {error}
        </div>
      )}

      {sheetInfo && sheetInfo.sheetNames.length > 0 && (
        <div
          data-testid="xlsx-sheet-tabs"
          role="tablist"
          aria-label={t('editor.sheetTabsLabel')}
          style={{
            display: 'flex',
            flex: '0 0 auto',
            gap: 2,
            padding: 'var(--xlsx-tabs-padding, 4px 6px)',
            borderTop: '1px solid var(--xlsx-divider, #e0e0e0)',
            background: 'var(--xlsx-tabs-bg, #fafafa)',
            overflowX: 'auto',
          }}
        >
          {sheetInfo.sheetNames.map((name, i) => {
            const active = i === activeSheet;
            return (
              <button
                key={i}
                role="tab"
                aria-selected={active}
                onClick={() => switchSheet(i, sheetInfo.sheetIds[i])}
                style={{
                  flex: '0 0 auto',
                  border: 'none',
                  padding: 'var(--xlsx-tab-padding, 4px 12px)',
                  fontSize: 'var(--xlsx-tab-font-size, inherit)',
                  whiteSpace: 'nowrap',
                  cursor: 'pointer',
                  borderBottom: active
                    ? `2px solid var(--xlsx-accent, ${BRAND})`
                    : '2px solid transparent',
                  fontWeight: active ? 600 : 400,
                  color: active ? toolbarColors.text : toolbarColors.muted,
                  background: active ? 'var(--xlsx-active-tab-bg, #ffffff)' : 'transparent',
                }}
              >
                {name}
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}
