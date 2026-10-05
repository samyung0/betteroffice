import type { StyleProperty } from '@betteroffice/xlsx';
import type {
  HorizontalAlignment,
  MergeAction,
  NumberFormat,
  SelectionFormatting,
  TextWrapping,
  VerticalAlignment,
} from './components/Toolbar';

/** Frozen rows or columns: none, one, two, or up to the selection's focus. */
export type FreezeAmount = 0 | 1 | 2 | 'current';

const FREEZE: readonly FreezeAmount[] = [0, 1, 2, 'current'];
const NUMBER_FORMATS = [
  'automatic',
  'plainText',
  'number',
  'percent',
  'scientific',
  'currency',
  'date',
  'time',
] as const satisfies readonly Exclude<NumberFormat, 'custom'>[];
const ALIGN = ['left', 'center', 'right'] as const satisfies readonly HorizontalAlignment[];
const VALIGN = ['top', 'middle', 'bottom'] as const satisfies readonly VerticalAlignment[];
const WRAP = ['overflow', 'wrap', 'clip'] as const satisfies readonly TextWrapping[];
const MERGE = [
  'all',
  'horizontal',
  'vertical',
  'unmerge',
] as const satisfies readonly MergeAction[];
/** Zoom percentages, as the toolbar's zoom list. */
export const ZOOM_PERCENTS = [50, 75, 90, 100, 125, 150, 200] as const;

/** Every command a host menu can run by id (`XlsxEditorApi.run`). */
export const XLSX_COMMANDS = [
  'undo',
  'redo',
  'selectAll',
  'deleteValues',
  'deleteRows',
  'deleteColumns',
  'insertRowAbove',
  'insertRowBelow',
  'insertColumnLeft',
  'insertColumnRight',
  'insertSheet',
  ...FREEZE.map((amount) => `freezeRows:${amount}` as const),
  ...FREEZE.map((amount) => `freezeColumns:${amount}` as const),
  ...ZOOM_PERCENTS.map((percent) => `zoom:${percent}` as const),
  ...NUMBER_FORMATS.map((format) => `numberFormat:${format}` as const),
  'bold',
  'italic',
  'strikethrough',
  ...ALIGN.map((value) => `align:${value}` as const),
  ...VALIGN.map((value) => `valign:${value}` as const),
  ...WRAP.map((value) => `wrap:${value}` as const),
  ...MERGE.map((value) => `merge:${value}` as const),
  'clearFormatting',
] as const;

export type XlsxCommand = (typeof XLSX_COMMANDS)[number];

export function isXlsxCommand(value: unknown): value is XlsxCommand {
  return (XLSX_COMMANDS as readonly unknown[]).includes(value);
}

type CommandName<C> = C extends `${infer Name}:${string}` ? Name : C;

/** Whether a command changes the workbook (freezing panes does): a read-only editor runs only those that don't. */
const EDITS: Record<CommandName<XlsxCommand>, boolean> = {
  undo: true,
  redo: true,
  selectAll: false,
  deleteValues: true,
  deleteRows: true,
  deleteColumns: true,
  insertRowAbove: true,
  insertRowBelow: true,
  insertColumnLeft: true,
  insertColumnRight: true,
  insertSheet: true,
  freezeRows: true,
  freezeColumns: true,
  zoom: false,
  numberFormat: true,
  bold: true,
  italic: true,
  strikethrough: true,
  align: true,
  valign: true,
  wrap: true,
  merge: true,
  clearFormatting: true,
};

export function xlsxCommandEdits(command: XlsxCommand): boolean {
  return EDITS[command.split(':')[0] as CommandName<XlsxCommand>];
}

/** What a host menu needs to enable, check and label its items. */
export interface XlsxCommandState {
  canUndo: boolean;
  canRedo: boolean;
  /** The selection's size and focus cell, zero-based. */
  selection: { rows: number; columns: number; focusRow: number; focusColumn: number } | null;
  frozenRows: number;
  frozenColumns: number;
  formatting: SelectionFormatting;
  canMerge: boolean;
  canUnmerge: boolean;
  zoom: number;
}

/** The freeze pane op for `rows` x `columns` frozen; none when both are zero. */
export function freezePaneOp(sheet: number, rows: number, columns: number) {
  return {
    type: 'setFreezePane',
    sheet,
    pane:
      rows === 0 && columns === 0
        ? null
        : { rows, cols: columns, top_left: { row: rows, col: columns } },
  };
}

/** How many rows or columns an amount freezes, given the focus index. */
export function frozenCount(amount: FreezeAmount, focus: number): number {
  return amount === 'current' ? focus + 1 : amount;
}

/** The first `Sheet<n>` name no sheet uses yet. */
export function newSheetName(names: readonly string[]): string {
  const taken = new Set(names.map((name) => name.toLowerCase()));
  for (let n = names.length + 1; ; n++) if (!taken.has(`sheet${n}`)) return `Sheet${n}`;
}

/** Every style property, so clearing formatting resets them all. */
export const ALL_STYLE_PROPERTIES = [
  'bold',
  'italic',
  'strikethrough',
  'fontFamily',
  'fontSize',
  'textColor',
  'fillColor',
  'borders',
  'horizontalAlignment',
  'verticalAlignment',
  'textWrapping',
] as const satisfies readonly StyleProperty[];
