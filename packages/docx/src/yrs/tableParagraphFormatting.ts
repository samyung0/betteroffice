import type {
  ParagraphFormatting,
  Table,
  TableCellFormatting,
  TableFormatting,
  TableLook,
  TableRowFormatting,
} from '../types/document';
import type { StyleResolver } from '../styles';
import type { Style } from '../types/styles';
import { mergeParagraphFormatting } from '../utils/paragraphFormattingMerge';
import type { YrsSession } from './index';

/** What the table-style paragraph formatting of a cell reads from its table. */
interface TableShape {
  formatting?: Pick<TableFormatting, 'look' | 'styleRowBandSize' | 'styleColBandSize'>;
  columnWidths?: number[];
  rows: Array<{
    formatting?: Pick<TableRowFormatting, 'gridBefore' | 'gridAfter'>;
    cells: Array<{ formatting?: Pick<TableCellFormatting, 'gridSpan'> }>;
  }>;
}

export function tableCellParagraphFormatting(
  table: TableShape,
  style: Style | undefined,
  rowIndex: number,
  startColumn: number,
  endColumn: number,
  columns: number
): ParagraphFormatting | undefined {
  let result = style?.pPr;
  if (!style?.tblStylePr?.some((part) => part.pPr)) return result;
  const look = table.formatting?.look ?? style.tblPr?.look;
  const flag = (key: Exclude<keyof TableLook, 'value'>, mask: number) =>
    look?.[key] ?? (Number.parseInt(look?.value ?? (look ? '0' : '04A0'), 16) & mask) !== 0;
  const firstRow = flag('firstRow', 0x20);
  const lastRow = flag('lastRow', 0x40);
  const firstColumn = flag('firstColumn', 0x80);
  const lastColumn = flag('lastColumn', 0x100);
  const gridBefore = table.rows[rowIndex].formatting?.gridBefore ?? 0;
  startColumn += gridBefore;
  endColumn += gridBefore;
  const atFirstRow = !!firstRow && rowIndex === 0;
  const atLastRow = !!lastRow && rowIndex === table.rows.length - 1;
  const atFirstColumn = !!firstColumn && startColumn === 0;
  const atLastColumn = !!lastColumn && endColumn === columns;
  const rowBandSize = table.formatting?.styleRowBandSize ?? style.tblPr?.styleRowBandSize ?? 1;
  const columnBandSize = table.formatting?.styleColBandSize ?? style.tblPr?.styleColBandSize ?? 1;
  const regions: string[] = [];
  if (!flag('noHBand', 0x200) && !atFirstRow && !atLastRow && rowBandSize > 0) {
    regions.push(
      Math.floor((rowIndex - Number(!!firstRow)) / rowBandSize) % 2 === 0
        ? 'band1Horz'
        : 'band2Horz'
    );
  }
  if (!flag('noVBand', 0x400) && !atFirstColumn && !atLastColumn && columnBandSize > 0) {
    regions.push(
      Math.floor((startColumn - Number(!!firstColumn)) / columnBandSize) % 2 === 0
        ? 'band1Vert'
        : 'band2Vert'
    );
  }
  for (const [region, active] of [
    ['firstCol', atFirstColumn],
    ['lastCol', atLastColumn],
    ['firstRow', atFirstRow],
    ['lastRow', atLastRow],
    ['nwCell', atFirstRow && atFirstColumn],
    ['neCell', atFirstRow && atLastColumn],
    ['swCell', atLastRow && atFirstColumn],
    ['seCell', atLastRow && atLastColumn],
  ] as const) {
    if (active) regions.push(region);
  }
  for (const region of regions) {
    result = mergeParagraphFormatting(
      result,
      style.tblStylePr.find((part) => part.type === region)?.pPr
    );
  }
  return result;
}

export function tableColumnCount(table: TableShape): number {
  return Math.max(
    table.columnWidths?.length ?? 0,
    ...table.rows.map((row) =>
      row.cells.reduce(
        (sum, cell) => sum + (cell.formatting?.gridSpan ?? 1),
        (row.formatting?.gridBefore ?? 0) + (row.formatting?.gridAfter ?? 0)
      )
    )
  );
}

type Attrs = Record<string, unknown>;

/** A table embed's payload in a story, as far as its cells' paragraph formatting needs it. */
interface TablePayloadShape {
  tblPr?: Attrs;
  grid?: unknown;
  rows?: Array<{ trPr?: Attrs; cells?: Array<{ tcPr?: Attrs; story?: string }> }>;
}

const CELL_STORY = /:t\d+:r\d+c\d+$/;
const CONTROL_STORY = /:sdt\d+$/;

/**
 * The table-style paragraph formatting each cell story of a table embed gives
 * its paragraphs, from the table's current rows, position and style, as the
 * seed folds it in under a paragraph's style.
 */
export function tablePayloadCellFormatting(
  payload: TablePayloadShape,
  styles: StyleResolver | null
): Map<string, ParagraphFormatting> {
  const cells = new Map<string, ParagraphFormatting>();
  if (!styles) return cells;
  const tblPr = payload.tblPr ?? {};
  const original = (tblPr._originalFormatting ?? {}) as TableFormatting;
  const defaultStyle = styles.getDefaultTableStyle();
  const styleId = (typeof tblPr.styleId === 'string' ? tblPr.styleId : undefined) ?? defaultStyle?.styleId;
  const style = (styleId ? styles.getStyle(styleId) : undefined) ?? defaultStyle;
  const rows = Array.isArray(payload.rows) ? payload.rows : [];
  const table: TableShape = {
    formatting: { ...original, look: (tblPr.look as TableLook | undefined) ?? original.look },
    columnWidths: Array.isArray(payload.grid) ? (payload.grid as number[]) : undefined,
    rows: rows.map((row) => ({
      formatting: row.trPr?._originalFormatting as TableRowFormatting | undefined,
      cells: (row.cells ?? []).map((cell) => ({ formatting: { gridSpan: Number(cell.tcPr?.colspan) || 1 } })),
    })),
  };
  const columns = tableColumnCount(table);
  const occupied: boolean[][] = [];
  rows.forEach((row, rowIndex) => {
    let column = 0;
    for (const cell of row.cells ?? []) {
      while (occupied[rowIndex]?.[column]) column += 1;
      const colspan = Number(cell.tcPr?.colspan) || 1;
      const rowspan = Number(cell.tcPr?.rowspan) || 1;
      for (let r = rowIndex; r < rowIndex + rowspan; r += 1) {
        occupied[r] ??= [];
        for (let c = column; c < column + colspan; c += 1) occupied[r][c] = true;
      }
      const formatting = tableCellParagraphFormatting(table, style, rowIndex, column, column + colspan, columns);
      if (cell.story && formatting) cells.set(cell.story, formatting);
      column += colspan;
    }
  });
  return cells;
}

/** The cell story whose paragraphs `story` shares the table formatting of: itself, or a control's cell. */
export function enclosingCellStory(story: string): string | undefined {
  let id = story;
  while (CONTROL_STORY.test(id)) id = id.replace(CONTROL_STORY, '');
  return CELL_STORY.test(id) ? id : undefined;
}

/**
 * The table-style paragraph formatting a paragraph in `story` gets from the
 * table cell holding it, read from the session's current table, so cells
 * added or moved in the session (by a peer too) get their own.
 */
export function cellParagraphFormatting(
  session: YrsSession,
  styles: StyleResolver | null,
  story: string
): ParagraphFormatting | undefined {
  const cell = enclosingCellStory(story);
  if (!cell || !styles) return undefined;
  const parent = cell.replace(CELL_STORY, '');
  let payloads = tableReads?.tables.get(parent);
  if (!payloads) {
    payloads = session.storyTables(parent) as TablePayloadShape[];
    tableReads?.tables.set(parent, payloads);
  }
  for (const payload of payloads) {
    if (!payload.rows?.some((row) => row.cells?.some((entry) => entry.story === cell))) continue;
    let cells = tableReads?.cells.get(payload);
    if (!cells) {
      cells = tablePayloadCellFormatting(payload, styles);
      tableReads?.cells.set(payload, cells);
    }
    return cells.get(cell);
  }
  return undefined;
}

/** Table reads `withTableReads` keeps for edits that change no table. */
let tableReads:
  | { tables: Map<string, TablePayloadShape[]>; cells: Map<TablePayloadShape, Map<string, ParagraphFormatting>> }
  | undefined;

/**
 * Runs `edit`, which changes paragraphs but no table, reading each story's
 * tables and each table's cell formatting once.
 */
export function withTableReads<T>(edit: () => T): T {
  const outer = tableReads;
  tableReads ??= { tables: new Map(), cells: new Map() };
  try {
    return edit();
  } finally {
    tableReads = outer;
  }
}

export type RowSpanInfo = { rowSpan: number; skip: boolean };

/** Each cell's rowspan by `row-gridColumn`; a covered vMerge continuation is `skip`, as the seed folds it. */
export function calculateRowSpans(table: Pick<Table, 'rows'>): Map<string, RowSpanInfo> {
  const result = new Map<string, RowSpanInfo>();
  const active = new Map<number, number>();
  table.rows.forEach((row, rowIndex) => {
    let column = 0;
    const cells = row.cells.map((cell) => {
      const current = column;
      column += cell.formatting?.gridSpan ?? 1;
      return { column: current, vMerge: cell.formatting?.vMerge, key: `${rowIndex}-${current}` };
    });
    const empty =
      cells.length > 0 &&
      cells.every((cell) => cell.vMerge === 'continue' && active.has(cell.column));
    if (empty) {
      for (const cell of cells) {
        active.delete(cell.column);
        result.set(cell.key, { rowSpan: 1, skip: false });
      }
      return;
    }
    for (const cell of cells) {
      if (cell.vMerge === 'restart') {
        active.set(cell.column, rowIndex);
        result.set(cell.key, { rowSpan: 1, skip: false });
      } else if (cell.vMerge === 'continue') {
        const start = active.get(cell.column);
        if (start === undefined) result.set(cell.key, { rowSpan: 1, skip: false });
        else {
          const owner = result.get(`${start}-${cell.column}`);
          if (owner) owner.rowSpan += 1;
          result.set(cell.key, { rowSpan: 1, skip: true });
        }
      } else {
        active.delete(cell.column);
        result.set(cell.key, { rowSpan: 1, skip: false });
      }
    }
  });
  return result;
}
