/**
 * Grid and chart pointer behaviour against the real wasm core: happy-dom has no
 * layout or canvas, so the viewport size and the 2d context are stubbed and
 * click points come from the same display-list geometry the editor hit-tests.
 *
 * Both suites live here on purpose. happy-dom registers into one shared process
 * global, so a second file with its own register/unregister pair would tear the
 * dom down under whichever suite bun happens to run second.
 */

import { GlobalRegistrator } from '@happy-dom/global-registrator';
import {
  afterAll,
  afterEach,
  beforeAll,
  describe,
  expect,
  it,
  setDefaultTimeout,
  spyOn,
} from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import {
  cellRect,
  initWasm,
  MAX_SCROLL_EXTENT_PX,
  openWorkbook,
  selectionAt,
  toTsv,
} from '@betteroffice/xlsx';
import type { CellAddr, ChartRegion, GridMeta, WorkbookProxy } from '@betteroffice/xlsx';
import { XLSX_MAX_COLUMNS, XLSX_MAX_ROWS } from '@betteroffice/xlsx/collaboration';
import {
  freezePaneOp,
  XLSX_COMMANDS,
  xlsxCommandEdits,
  type XlsxCommand,
  type XlsxCommandState,
} from './commands';
import { useState } from 'react';
import { zhCN } from '@betteroffice/xlsx-i18n';
import { XlsxEditor, type XlsxEditorApi, type XlsxEditorProps } from './XlsxEditor';

const WASM = resolve(import.meta.dir, '../../xlsx/src/wasm/generated/xlsx_wasm_bg.wasm');
const FIXTURE = resolve(import.meta.dir, '../../xlsx/test-fixtures/sample.xlsx');
const CHART_FIXTURE = resolve(import.meta.dir, '../../xlsx/test-fixtures/charts.xlsx');
// a few used cells, so most of the window lies past the used range.
const SMALL_FIXTURE = resolve(import.meta.dir, '../../xlsx/test-fixtures/defined-names.xlsx');
const UNDRAWABLE_FIXTURE = resolve(
  import.meta.dir,
  '../../xlsx/test-fixtures/unsupported-charts.xlsx'
);
const VIEWPORT = { width: 800, height: 600 };
const LINK_TARGET = 'https://example.com/report';
const LINK_CELL: CellAddr = { row: 5, col: 4 };

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const { act, cleanup, configure, fireEvent, render, waitFor } = await import(
  '@testing-library/react'
);
// the workbook opens in its own worker: Bun starts it and its WASM per editor.
// a timed-out wait reports its message, not the whole happy-dom tree.
configure({
  asyncUtilTimeout: 10_000,
  getElementError: (message) => new Error(message ?? 'element not found'),
});
setDefaultTimeout(30_000);

function stubContext(): CanvasRenderingContext2D {
  const noop = () => {};
  return {
    save: noop,
    restore: noop,
    setTransform: noop,
    clearRect: noop,
    beginPath: noop,
    rect: noop,
    clip: noop,
    setLineDash: noop,
    moveTo: noop,
    lineTo: noop,
    quadraticCurveTo: noop,
    bezierCurveTo: noop,
    closePath: noop,
    fill: noop,
    stroke: noop,
    fillRect: noop,
    fillText: noop,
    measureText: () => ({ width: 0 }),
  } as unknown as CanvasRenderingContext2D;
}

// the stubs live on shared prototypes, so they are installed and restored
// around this file rather than leaking into every other happy-dom suite.
const LAYOUT = [
  ['clientWidth', VIEWPORT.width],
  ['clientHeight', VIEWPORT.height],
] as const;
const originalGetContext = HTMLCanvasElement.prototype.getContext;
const originalOpen = window.open;
const originalLayout = LAYOUT.map(([property]) => {
  return [property, Object.getOwnPropertyDescriptor(HTMLElement.prototype, property)] as const;
});

// no committed fixture carries a hyperlink, so the engine installs one: a
// structural op through the ops escape hatch, saved back out as workbook bytes.
function withHyperlink(bytes: Uint8Array): Uint8Array {
  const handle = openWorkbook(bytes);
  try {
    handle.applyOps([
      {
        type: 'setHyperlinks',
        sheet: 0,
        hyperlinks: [
          { range: { start: LINK_CELL, end: LINK_CELL }, external_target: LINK_TARGET },
        ],
      },
    ]);
    return handle.save();
  } finally {
    handle.dispose();
  }
}

// a used range past the window on both axes, so the keyboard can leave it.
const FAR_CELL: CellAddr = { row: 40, col: 20 };
function withFarCell(bytes: Uint8Array, frozen = 0): Uint8Array {
  const handle = openWorkbook(bytes);
  try {
    handle.editCell(0, FAR_CELL.row, FAR_CELL.col, 'far');
    if (frozen) handle.applyOps([freezePaneOp(0, frozen, frozen)]);
    return handle.save();
  } finally {
    handle.dispose();
  }
}

// the small fixture's sheet holding just `cells` (A1 address -> number).
async function withCells(bytes: Uint8Array, cells: Record<string, number>): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes);
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  const rows = new Map<number, string[]>();
  for (const [address, value] of Object.entries(cells)) {
    const row = Number(/\d+$/.exec(address)![0]);
    rows.set(row, [...(rows.get(row) ?? []), `<c r="${address}"><v>${value}</v></c>`]);
  }
  const data = [...rows]
    .sort(([a], [b]) => a - b)
    .map(([row, cellXml]) => `<row r="${row}">${cellXml.join('')}</row>`)
    .join('');
  zip.file(
    'xl/worksheets/sheet1.xml',
    sheet.replace(/<sheetData>.*<\/sheetData>/, `<sheetData>${data}</sheetData>`)
  );
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array' }));
}

// column B wider than the window.
function withWideColumn(bytes: Uint8Array): Uint8Array {
  const handle = openWorkbook(bytes);
  try {
    handle.applyOps([{ type: 'setColWidth', sheet: 0, col: 1, width: 200 }]);
    return handle.save();
  } finally {
    handle.dispose();
  }
}

// column F styled with a 36pt font: a value typed there grows its row.
async function withTallFontColumn(bytes: Uint8Array): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes);
  let styles = await zip.file('xl/styles.xml')!.async('string');
  styles = styles
    .replace(/<fonts count="(\d+)">/, (_, n) => `<fonts count="${Number(n) + 1}">`)
    .replace('</fonts>', '<font><sz val="36"/><name val="Calibri"/></font></fonts>');
  const fonts = Number(/<fonts count="(\d+)">/.exec(styles)![1]);
  styles = styles
    .replace(/<cellXfs count="(\d+)">/, (_, n) => `<cellXfs count="${Number(n) + 1}">`)
    .replace(
      '</cellXfs>',
      `<xf numFmtId="0" fontId="${fonts - 1}" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>`
    );
  const xfs = Number(/<cellXfs count="(\d+)">/.exec(styles)![1]);
  zip.file('xl/styles.xml', styles);
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  zip.file(
    'xl/worksheets/sheet1.xml',
    sheet.replace('</cols>', `<col min="6" max="6" width="9" style="${xfs - 1}"/></cols>`)
  );
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array' }));
}

// cell positions the editor asks the worker for, and frames, from here on.
function countEngine(proxy: WorkbookProxy) {
  const counts = { positions: 0, paints: 0 };
  const onSheet = proxy.onSheet;
  const frame = proxy.frame;
  proxy.onSheet = ((sheetId, method, ...args) => {
    if (method === 'cellPosition') counts.positions += 1;
    return onSheet(sheetId, method, ...args);
  }) as WorkbookProxy['onSheet'];
  proxy.frame = (view) => {
    counts.paints += 1;
    return frame(view);
  };
  return counts;
}

// the workbook of the editor mounted last.
let currentWorkbook: WorkbookProxy | null = null;

// every request posted so far answered and the frames it asked for painted:
// a round trip to the worker, the commits it triggers, again until quiet.
// a clipboard `write` that records the text of each ClipboardItem it takes.
function writeInto(written: string[]) {
  return async ([item]: ClipboardItem[]) => {
    written.push(await (await item.getType('text/plain')).text());
  };
}

function idleWith(proxy: () => WorkbookProxy | null | undefined) {
  return act(async () => {
    for (let round = 0; round < 4; round += 1) {
      await proxy()?.historyState();
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  });
}

interface Fixture {
  bytes: Uint8Array;
  grid: GridMeta;
  charts: ChartRegion[];
}

function fixtureFrom(bytes: Uint8Array): Fixture {
  const probe = openWorkbook(bytes);
  try {
    const frame = probe.displayList({ x: 0, y: 0, ...VIEWPORT });
    return { bytes, grid: frame.grid as GridMeta, charts: frame.charts ?? [] };
  } finally {
    probe.dispose();
  }
}

let plain: Fixture;
let small: Fixture;
// a used range starting away from A1: C5:D8 holding 1 to 8, row by row.
let offset: Fixture;
// A1 and Z100000 only: a used range of 2.6M cells holding two.
let sparse: Fixture;
// J90000:J90010 only.
let farColumn: Fixture;
// A1 and AZ100000: 5.2M cells from A1, past the copy ceiling.
let pastCeiling: Fixture;
// eight values spread over A1:Y10000, 250,000 cells from A1.
const SPREAD: Record<string, number> = {
  A1: 1, Y1: 2, C2: 3, M4000: 4, N4001: 5, B7999: 6, X8001: 7, Y10000: 8,
};
let spread: Fixture;
let linked: Fixture;
let charted: Fixture;
let undrawable: Fixture;
let wide: Fixture;
let wideFrozen: Fixture;
// frozen panes larger than the window on both axes.
let paneFilled: Fixture;
let wideColumn: Fixture;
let tallFont: Fixture;
let opened: string[] = [];

beforeAll(async () => {
  HTMLCanvasElement.prototype.getContext = (() =>
    stubContext()) as unknown as HTMLCanvasElement['getContext'];
  for (const [property, value] of LAYOUT) {
    Object.defineProperty(HTMLElement.prototype, property, {
      configurable: true,
      get: () => value,
    });
  }
  window.open = ((url?: string | URL) => {
    opened.push(String(url));
    return null;
  }) as typeof window.open;
  await initWasm(new Uint8Array(readFileSync(WASM)));
  const source = new Uint8Array(readFileSync(FIXTURE));
  plain = fixtureFrom(source);
  small = fixtureFrom(new Uint8Array(readFileSync(SMALL_FIXTURE)));
  const smallBytes = new Uint8Array(readFileSync(SMALL_FIXTURE));
  offset = fixtureFrom(
    await withCells(smallBytes, { C5: 1, D5: 2, C6: 3, D6: 4, C7: 5, D7: 6, C8: 7, D8: 8 })
  );
  sparse = fixtureFrom(await withCells(smallBytes, { A1: 1, Z100000: 2 }));
  pastCeiling = fixtureFrom(await withCells(smallBytes, { A1: 1, AZ100000: 2 }));
  spread = fixtureFrom(await withCells(smallBytes, SPREAD));
  farColumn = fixtureFrom(
    await withCells(
      smallBytes,
      Object.fromEntries(Array.from({ length: 11 }, (_, index) => [`J${90_000 + index}`, index + 1]))
    )
  );
  linked = fixtureFrom(withHyperlink(source));
  charted = fixtureFrom(new Uint8Array(readFileSync(CHART_FIXTURE)));
  undrawable = fixtureFrom(new Uint8Array(readFileSync(UNDRAWABLE_FIXTURE)));
  wide = fixtureFrom(withFarCell(source));
  wideFrozen = fixtureFrom(withFarCell(source, 2));
  paneFilled = fixtureFrom(withFarCell(source, 40));
  wideColumn = fixtureFrom(withWideColumn(source));
  tallFont = fixtureFrom(await withTallFontColumn(source));
});

afterAll(async () => {
  HTMLCanvasElement.prototype.getContext = originalGetContext;
  window.open = originalOpen;
  for (const [property, descriptor] of originalLayout) {
    if (descriptor) Object.defineProperty(HTMLElement.prototype, property, descriptor);
  }
  // last: bun shares one process across test files, and happy-dom's fetch
  // rejects the file: urls other suites initialise their wasm from.
  await GlobalRegistrator.unregister();
});

afterEach(() => {
  cleanup();
  opened = [];
  currentWorkbook = null;
});

function pointAt(fixture: Fixture, addr: CellAddr): { clientX: number; clientY: number } {
  const rect = cellRect(fixture.grid, addr.row, addr.col);
  if (!rect) throw new Error(`cell ${addr.row},${addr.col} is outside the painted window`);
  return { clientX: rect.x + rect.w / 2, clientY: rect.y + rect.h / 2 };
}

async function mountEditor(
  fixture: Fixture = plain,
  onSave?: (bytes: Uint8Array) => void,
  onChange?: () => void,
  props: Partial<XlsxEditorProps> = {}
) {
  const ready: { handle: WorkbookProxy | null; api?: XlsxEditorApi; painted: number } = {
    handle: null,
    painted: 0,
  };
  const onFirstPaint = () => {
    ready.painted += 1;
  };
  const file = fixture.bytes.slice();
  const element = (extra: Partial<XlsxEditorProps>) => (
    <XlsxEditor
      {...props}
      {...extra}
      file={file}
      onChange={onChange}
      onSave={onSave}
      onFirstPaint={onFirstPaint}
      onReady={(api) => {
        ready.handle = api.handle;
        ready.api = api;
        currentWorkbook = api.handle;
      }}
    />
  );
  const view = render(element({}));
  const nameBox = () => view.getByTestId('xlsx-name-box') as HTMLInputElement;
  await waitFor(() => expect(ready.painted).toBe(1));
  const idle = () => idleWith(() => ready.handle);
  await idle();
  const surface = view.getByTestId('xlsx-scroll');
  const editor = () => view.queryByTestId('xlsx-cell-editor') as HTMLInputElement | null;
  const press = (addr: CellAddr) => {
    fireEvent.mouseDown(surface, pointAt(fixture, addr));
    fireEvent.mouseUp(surface, pointAt(fixture, addr));
    fireEvent.click(surface, pointAt(fixture, addr));
  };
  return {
    surface,
    nameBox,
    editor,
    idle,
    workbook: () => ready.handle!,
    api: () => ready.api!,
    /** The same workbook with other props (no reopen). */
    setProps: (extra: Partial<XlsxEditorProps>) =>
      act(async () => {
        view.rerender(element(extra));
      }),
    /** The toolbar controls still taking input. */
    liveControls: () =>
      Array.from(
        view
          .getByTestId('xlsx-toolbar')
          .querySelectorAll('button:not(:disabled), input:not(:disabled):not([readonly])'),
        (node) => node.getAttribute('data-testid')
      ),
    /** A cell's editable text on sheet 0, once every change posted landed. */
    input: async (row: number, col: number) => {
      await idle();
      return (await ready.handle!.cell(0, row, col)).input;
    },
    flush: async () => {
      await act(async () => {
        await ready.api!.flush();
      });
      await idle();
    },
    run: async (command: XlsxCommand) => {
      await act(async () => ready.api!.run(command));
      await idle();
    },
    shown: () => ready.api!.visibleViewport()!,
    typeFormula: (value: string) =>
      fireEvent.change(view.getByTestId('xlsx-formula-input'), { target: { value } }),
    reopenWith: (next: Fixture) =>
      act(async () => {
        view.rerender(
          <XlsxEditor
            {...props}
            file={next.bytes.slice()}
            onChange={onChange}
            onSave={onSave}
            onFirstPaint={onFirstPaint}
            onReady={(api) => {
              ready.handle = api.handle;
              ready.api = api;
              currentWorkbook = api.handle;
            }}
          />
        );
      }),
    click: async (addr: CellAddr) => {
      press(addr);
      await idle();
    },
    doubleClick: async (addr: CellAddr) => {
      press(addr);
      press(addr);
      fireEvent.doubleClick(surface, pointAt(fixture, addr));
      await idle();
    },
    pressInEditor: (addr: CellAddr) => fireEvent.mouseDown(editor()!, pointAt(fixture, addr)),
    doubleClickInEditor: (addr: CellAddr) => {
      const target = editor()!;
      fireEvent.mouseDown(target, pointAt(fixture, addr));
      fireEvent.mouseUp(target, pointAt(fixture, addr));
      fireEvent.click(target, pointAt(fixture, addr));
      fireEvent.doubleClick(target, pointAt(fixture, addr));
    },
    type: (value: string) => fireEvent.change(editor()!, { target: { value } }),
    formula: () => (view.getByTestId('xlsx-formula-input') as HTMLInputElement).value,
    formulaInput: () => view.getByTestId('xlsx-formula-input') as HTMLInputElement,
    canUndo: () => !(view.getByTestId('xlsx-undo') as HTMLButtonElement).disabled,
    outline: () => view.queryByTestId('xlsx-chart-selection'),
    selectionBox: () => view.queryByTestId('xlsx-selection'),
    error: () => view.queryByTestId('xlsx-error'),
    outlineAt: () => {
      const box = view.getByTestId('xlsx-chart-selection');
      return {
        x: Math.round(parseFloat(box.style.left)),
        y: Math.round(parseFloat(box.style.top)),
      };
    },
  };
}

// the centre of a chart's visible region, which is where the editor paints it.
function chartCenter(region: ChartRegion): { clientX: number; clientY: number } {
  return {
    clientX: region.clip.x + region.clip.w / 2,
    clientY: region.clip.y + region.clip.h / 2,
  };
}

function chartRounded(region: ChartRegion): { x: number; y: number } {
  return { x: Math.round(region.rect.x), y: Math.round(region.rect.y) };
}

// past the window an arrow burst stays local in, so it has landed or been
// discarded by the time the assertion runs.
function settle(): Promise<void> {
  return act(async () => {
    await new Promise((resolve) => setTimeout(resolve, 400));
  });
}

async function selectChart(
  view: Awaited<ReturnType<typeof mountEditor>>,
  chart: ChartRegion
): Promise<void> {
  fireEvent.mouseDown(view.surface, chartCenter(chart));
  await waitFor(() => view.outline()!);
  fireEvent.mouseUp(window, chartCenter(chart));
}

describe('XlsxEditor grid pointer handling', () => {
  it('reports semantic workbook changes without treating selection as an edit', async () => {
    let changes = 0;
    const view = await mountEditor(plain, undefined, () => {
      changes += 1;
    });

    await view.click({ row: 2, col: 0 });
    expect(changes).toBe(0);

    await view.doubleClick({ row: 2, col: 0 });
    view.type('Edited item');
    await view.click({ row: 3, col: 1 });
    expect(changes).toBe(1);
  });

  it('flushes open cell and formula drafts before a host checkpoint without blur', async () => {
    const view = await mountEditor();
    await view.doubleClick({ row: 2, col: 0 });
    view.type('Unblurred cell');
    expect((await view.input(2, 0))).toBe('Line item 1');
    await view.flush();
    expect((await view.input(2, 0))).toBe('Unblurred cell');
    view.typeFormula('Unblurred formula');
    expect((await view.input(2, 0))).toBe('Unblurred cell');
    await view.flush();
    const reopened = openWorkbook(await view.workbook().save());
    try {
      expect(reopened.cell(0, 2, 0).input).toBe('Unblurred formula');
    } finally {
      reopened.dispose();
    }
  });

  it('commits the open editor and moves the selection when another cell is clicked', async () => {
    const view = await mountEditor();

    await view.doubleClick({ row: 2, col: 0 });
    expect(view.editor()?.value).toBe('Line item 1');

    view.type('Edited item');
    await view.click({ row: 3, col: 1 });

    expect(view.editor() === null).toBe(true);
    expect(view.nameBox().value).toBe('B4');
    expect((await view.input(2, 0))).toBe('Edited item');
  });

  it('commits and reopens on the target when another cell is double-clicked', async () => {
    const view = await mountEditor();

    await view.doubleClick({ row: 2, col: 0 });
    view.type('Edited item');
    await view.doubleClick({ row: 3, col: 1 });

    expect(view.editor()?.value).toBe('200');
    expect(view.nameBox().value).toBe('B4');
    expect((await view.input(2, 0))).toBe('Edited item');
  });

  it('leaves a formula cell unchanged when the pointer moves on', async () => {
    const view = await mountEditor();

    await view.doubleClick({ row: 2, col: 3 });
    expect(view.editor()?.value).toBe('=B3+C3');

    await view.click({ row: 6, col: 0 });

    expect(view.editor() === null).toBe(true);
    expect(view.nameBox().value).toBe('A7');
    expect((await view.input(2, 3))).toBe('=B3+C3');

    await view.doubleClick({ row: 3, col: 3 });
    expect(view.editor()?.value).toBe('=B4+C4');

    await view.doubleClick({ row: 7, col: 0 });

    expect(view.editor()?.value).toBe('Line item 6');
    expect(view.nameBox().value).toBe('A8');
    expect((await view.input(3, 3))).toBe('=B4+C4');
  });

  it('keeps a press inside the open editor from committing or moving on', async () => {
    const view = await mountEditor();

    await view.doubleClick({ row: 2, col: 0 });
    view.type('Edited item');
    view.pressInEditor({ row: 6, col: 0 });

    expect(view.editor()?.value).toBe('Edited item');
    expect(view.nameBox().value).toBe('A3');
    expect((await view.input(2, 0))).toBe('Line item 1');
  });

  it('keeps a double-click inside the open editor from reopening it', async () => {
    const view = await mountEditor();

    await view.doubleClick({ row: 2, col: 0 });
    view.type('Edited item');
    view.doubleClickInEditor({ row: 2, col: 0 });

    expect(view.editor()?.value).toBe('Edited item');
  });

  it('dismisses the editor without following a hyperlink in the clicked cell', async () => {
    const view = await mountEditor(linked);

    await view.doubleClick({ row: 2, col: 0 });
    view.type('Edited item');
    await view.click(LINK_CELL);

    expect(view.editor() === null).toBe(true);
    expect(view.nameBox().value).toBe('E6');
    expect(opened).toEqual([]);

    await view.click(LINK_CELL);

    expect(opened).toEqual([LINK_TARGET]);
  });
});

describe('XlsxEditor keyboard', () => {
  // dozens of keys, each with a full paint: generous for a loaded runner.
  const MANY_KEYS_MS = 30_000;
  // each key, then the worker's answers to what it asked for.
  const press = async (target: Element, key: string, init: KeyboardEventInit = {}) => {
    await act(async () => {
      fireEvent.keyDown(target, { key, ...init });
    });
    await idleWith(() => currentWorkbook);
  };
  // the worker calls the editor makes (a read or an edit list over a whole
  // column is recorded but kept from the worker) and what reaches the clipboard.
  const watch = (view: Awaited<ReturnType<typeof mountEditor>>) => {
    const proxy = view.workbook();
    const forward = proxy.onSheet.bind(proxy);
    const sent: Array<{ method: string; args: unknown[] }> = [];
    const spy = spyOn(proxy, 'onSheet').mockImplementation(((sheetId, method, ...args) => {
      sent.push({ method, args });
      if (method === 'rangeCells' && /\d{7}$/.test(String(args[1]))) return Promise.resolve([]);
      if (method === 'editCells' && (args[1] as unknown[]).length > 10_000)
        return Promise.resolve(undefined);
      return forward(sheetId, method, ...args);
    }) as WorkbookProxy['onSheet']);
    const written: string[] = [];
    const clipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: { write: writeInto(written) },
    });
    return {
      sent: (method: string) => sent.filter((call) => call.method === method).map((call) => call.args),
      written,
      restore: () => {
        spy.mockRestore();
        if (clipboard) Object.defineProperty(navigator, 'clipboard', clipboard);
        else Reflect.deleteProperty(navigator, 'clipboard');
      },
    };
  };
  const scrollTo = async (
    view: Awaited<ReturnType<typeof mountEditor>>,
    left: number,
    top: number
  ) => {
    await act(async () => {
      view.surface.scrollLeft = left;
      view.surface.scrollTop = top;
      fireEvent.scroll(view.surface);
      await new Promise((resolve) => requestAnimationFrame(resolve));
    });
    await view.idle();
  };
  const cellAt = (a1: string) => {
    const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(a1)!;
    const col = [...letters].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
    return { row: Number(digits) - 1, col };
  };
  // the open editor's box over its cell.
  const editorBox = (input: HTMLInputElement) => {
    const left = parseFloat(input.style.left);
    const top = parseFloat(input.style.top);
    return {
      left,
      top,
      right: left + parseFloat(input.style.width),
      bottom: top + parseFloat(input.style.height),
    };
  };

  it('scrolls a cell the arrow keys reach into view and types into it', async () => {
    const view = await mountEditor(wide);
    await view.click({ row: 3, col: 1 });
    for (let step = 0; step < 15; step += 1) await press(view.surface, 'ArrowRight');
    for (let step = 0; step < 30; step += 1) await press(view.surface, 'ArrowDown');
    expect(view.nameBox().value).toBe('Q34');
    expect(view.surface.scrollLeft).toBeGreaterThan(0);
    expect(view.surface.scrollTop).toBeGreaterThan(0);

    await press(view.surface, '7');
    const input = view.editor();
    expect(input?.value).toBe('7');
    const box = editorBox(input!);
    expect(box.left).toBeGreaterThanOrEqual(0);
    expect(box.top).toBeGreaterThanOrEqual(0);
    expect(box.right).toBeLessThanOrEqual(VIEWPORT.width);
    expect(box.bottom).toBeLessThanOrEqual(VIEWPORT.height);
    expect(document.activeElement === input).toBe(true);

    await press(input!, 'Enter');
    expect((await view.input(33, 16))).toBe('7');
    expect(view.nameBox().value).toBe('Q35');
  }, MANY_KEYS_MS);

  // Excel moves over the whole sheet (A1:XFD1048576) and Google Sheets over the
  // sheet's grid (1,000 rows by 26 columns new); sample.xlsx uses A1:D62.
  it('moves the selection past the used range', async () => {
    const view = await mountEditor(plain);
    await view.click({ row: 0, col: 0 });
    for (let step = 0; step < 70; step += 1) await press(view.surface, 'ArrowDown');
    expect(view.nameBox().value).toBe('A71');
    for (let step = 0; step < 12; step += 1) await press(view.surface, 'ArrowRight');
    expect(view.nameBox().value).toBe('M71');
  }, MANY_KEYS_MS);

  it('selects the cells a drag or a Shift+click covers past the used range', async () => {
    const view = await mountEditor(small);
    const from = { row: 9, col: 1 };
    const to = { row: 14, col: 3 };
    await act(async () => {
      fireEvent.mouseDown(view.surface, pointAt(small, from));
      fireEvent.mouseMove(view.surface, { ...pointAt(small, to), buttons: 1 });
      fireEvent.mouseUp(window, pointAt(small, to));
    });
    expect(view.nameBox().value).toBe('B10');
    const pressed = cellRect(small.grid, from.row, from.col)!;
    const under = cellRect(small.grid, to.row, to.col)!;
    const box = view.selectionBox()!.style;
    expect([box.left, box.top, box.width, box.height].map((value) => Math.round(parseFloat(value)))).toEqual(
      [pressed.x, pressed.y, under.x + under.w - pressed.x, under.y + under.h - pressed.y].map(Math.round)
    );

    await act(async () => {
      fireEvent.mouseDown(view.surface, { ...pointAt(small, { row: 20, col: 4 }), shiftKey: true });
      fireEvent.mouseUp(window, pointAt(small, { row: 20, col: 4 }));
    });
    expect(view.nameBox().value).toBe('B10');
  }, MANY_KEYS_MS);

  // sample.xlsx's Budget sheet uses A1:D62.
  it('lands Ctrl+End on the last used cell', async () => {
    const view = await mountEditor(plain);
    await view.click({ row: 0, col: 0 });
    await press(view.surface, 'End', { ctrlKey: true });
    expect(view.nameBox().value).toBe('D62');
  }, MANY_KEYS_MS);

  it('copies and clears only the used part of a whole-column selection', async () => {
    const view = await mountEditor(plain);
    const calls = watch(view);
    try {
      expect(await view.input(1, 1)).not.toBe('');
      const column = { anchor: { row: 0, col: 1 }, focus: { row: XLSX_MAX_ROWS - 1, col: 1 } };
      await act(async () => {
        expect(await view.api().selectCells(0, column)).toBe(true);
      });
      await view.idle();
      await press(view.surface, 'c', { ctrlKey: true });
      await view.idle();
      expect(calls.sent('rangeCells').map((args) => args[1])).toEqual(['B1:B62']);
      expect(calls.written).toHaveLength(1);
      expect(calls.written[0].split(/\r?\n/)).toHaveLength(62);

      await press(view.surface, 'Delete');
      expect(calls.sent('clearContents').map((args) => args[1])).toEqual(['B1:B62']);
      expect(await view.input(1, 1)).toBe('');
    } finally {
      calls.restore();
    }
  }, MANY_KEYS_MS);

  it('copies the whole sheet from A1 with its leading blanks, and clears only the used cells', async () => {
    const view = await mountEditor(offset);
    const calls = watch(view);
    const copy = async (selection: { anchor: CellAddr; focus: CellAddr }) => {
      await act(async () => {
        expect(await view.api().selectCells(0, selection)).toBe(true);
      });
      await view.idle();
      await press(view.surface, 'c', { ctrlKey: true });
      await view.idle();
    };
    const sheet = { anchor: { row: 0, col: 0 }, focus: { row: XLSX_MAX_ROWS - 1, col: XLSX_MAX_COLUMNS - 1 } };
    try {
      await copy(sheet);
      expect(calls.sent('rangeCells').map((args) => args[1])).toEqual(['A1:D8']);
      expect(calls.written[calls.written.length - 1].split(/\r?\n/)).toEqual([
        '\t\t\t',
        '\t\t\t',
        '\t\t\t',
        '\t\t\t',
        '\t\t1\t2',
        '\t\t3\t4',
        '\t\t5\t6',
        '\t\t7\t8',
      ]);

      // wholly past the used range's end: nothing is asked of the worker.
      await copy({ anchor: { row: 8, col: 0 }, focus: { row: 20, col: 5 } });
      expect(calls.sent('rangeCells')).toHaveLength(1);
      expect(calls.written[calls.written.length - 1]).toBe('');

      await copy(sheet);
      await press(view.surface, 'Delete');
      expect(calls.sent('clearContents').map((args) => args[1])).toEqual(['C5:D8']);
      for (const [row, col] of [[4, 2], [7, 3]]) expect(await view.input(row, col)).toBe('');
    } finally {
      calls.restore();
    }
  }, MANY_KEYS_MS);

  // A1 and Z100000: select all covers A1:Z100000, the engine clears two cells.
  it('clears a sparse used range through the engine at the cost of its stored cells', async () => {
    const view = await mountEditor(sparse);
    const calls = watch(view);
    try {
      await view.click({ row: 0, col: 0 });
      await press(view.surface, 'a', { ctrlKey: true });
      await press(view.surface, 'Delete');
      expect(calls.sent('clearContents').map((args) => args[1])).toEqual(['A1:Z100000']);
      expect(calls.sent('editCells')).toEqual([]);
      expect(await view.input(0, 0)).toBe('');
      expect(await view.input(99_999, 25)).toBe('');
    } finally {
      calls.restore();
    }
  }, MANY_KEYS_MS);

  // J90000:J90010: select all is A1:J90010, 900,100 cells read in ten pieces.
  it('cuts a far column through a whole-sheet selection read in pieces', async () => {
    const view = await mountEditor(farColumn);
    const calls = watch(view);
    try {
      await view.click({ row: 0, col: 0 });
      await press(view.surface, 'a', { ctrlKey: true });
      await press(view.surface, 'x', { ctrlKey: true });
      // the pieces are read one after another, then the clear goes out.
      await waitFor(() => expect(calls.sent('clearContents')).toHaveLength(1), { timeout: 20_000 });
      const ranges = calls.sent('rangeCells').map((args) => args[1]);
      expect(ranges).toHaveLength(10);
      expect([ranges[0], ranges[9]]).toEqual(['A1:J10000', 'A90001:J90010']);
      const rows = calls.written[0].split(/\r?\n/);
      expect(rows).toHaveLength(90_010);
      expect([rows[89_998], rows[89_999], rows[90_009]]).toEqual([
        '\t\t\t\t\t\t\t\t\t',
        '\t\t\t\t\t\t\t\t\t1',
        '\t\t\t\t\t\t\t\t\t11',
      ]);
      expect(calls.sent('clearContents').map((args) => args[1])).toEqual(['J90000:J90010']);
      expect(await view.input(89_999, 9)).toBe('');
    } finally {
      calls.restore();
    }
  }, MANY_KEYS_MS);

  // A1 and AZ100000: 5.2M cells, past XLSX_COPY_MAX_CELLS.
  it('cuts nothing past the copy ceiling, with a notice', async () => {
    const notices: string[] = [];
    const view = await mountEditor(pastCeiling, undefined, undefined, {
      onNotice: (notice) => notices.push(notice),
    });
    const calls = watch(view);
    try {
      await view.click({ row: 0, col: 0 });
      await press(view.surface, 'a', { ctrlKey: true });
      await press(view.surface, 'x', { ctrlKey: true });
      await view.idle();
      expect(notices).toEqual(['copy-too-large']);
      expect(calls.sent('rangeCells')).toEqual([]);
      expect(calls.written).toEqual([]);
      expect(calls.sent('clearContents')).toEqual([]);
      expect(await view.input(0, 0)).toBe('1');
      expect(await view.input(99_999, 51)).toBe('2');
    } finally {
      calls.restore();
    }
  }, MANY_KEYS_MS);

  // 25 columns x 10,000 rows: three reads of at most 100,000 cells.
  it('copies a 250,000-cell selection in pieces exactly as the cells read', async () => {
    const view = await mountEditor(spread);
    const calls = watch(view);
    try {
      await view.click({ row: 0, col: 0 });
      await press(view.surface, 'a', { ctrlKey: true });
      await press(view.surface, 'c', { ctrlKey: true });
      await waitFor(() => expect(calls.written).toHaveLength(1), { timeout: 20_000 });
      expect(calls.sent('rangeCells').map((args) => args[1])).toEqual([
        'A1:Y4000',
        'A4001:Y8000',
        'A8001:Y10000',
      ]);
      // the oracle: the sheet's eight values on a blank 10,000 x 25 grid.
      const grid = Array.from({ length: 10_000 }, () => Array.from({ length: 25 }, () => ''));
      for (const [address, value] of Object.entries(SPREAD)) {
        const [, letter, row] = /^([A-Z])(\d+)$/.exec(address)!;
        grid[Number(row) - 1][letter.charCodeAt(0) - 65] = String(value);
      }
      expect(calls.written).toEqual([
        toTsv(grid.map((row) => row.map((input) => ({ input, isFormula: false })))),
      ]);
    } finally {
      calls.restore();
    }
  }, MANY_KEYS_MS);

  it('refuses a paste past the engine update cap with a notice and changes nothing', async () => {
    const notices: string[] = [];
    const view = await mountEditor(plain, undefined, undefined, {
      onNotice: (notice) => notices.push(notice),
    });
    const before = await view.input(1, 1);
    // 90,000 formulas may take 896 bytes each: over 64 MiB.
    const text = Array.from({ length: 300 }, () => Array(300).fill('=A1').join('\t')).join('\r\n');
    const clipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: { readText: async () => text, writeText: async () => {} },
    });
    try {
      await view.click({ row: 0, col: 0 });
      await press(view.surface, 'v', { ctrlKey: true });
      await view.idle();
      expect(notices).toEqual(['edit-too-large']);
      expect(await view.input(1, 1)).toBe(before);
      expect(await view.input(299, 299)).toBe('');
    } finally {
      if (clipboard) Object.defineProperty(navigator, 'clipboard', clipboard);
      else Reflect.deleteProperty(navigator, 'clipboard');
    }
  }, MANY_KEYS_MS);

  it('refuses edits while a large update is sending, keeping selection and copy', async () => {
    const view = await mountEditor(plain, undefined, undefined, { sending: true });
    const calls = watch(view);
    try {
      await view.click({ row: 1, col: 1 });
      const before = await view.input(1, 1);
      await press(view.surface, '7');
      expect(view.editor()).toBeNull();
      await press(view.surface, 'Delete');
      await press(view.surface, 'v', { ctrlKey: true });
      await press(view.surface, 'z', { ctrlKey: true });
      // the toolbar keeps only what edits nothing, as in View.
      expect(view.liveControls()).toEqual([
        'xlsx-search-menus',
        'xlsx-zoom',
        null,
        'xlsx-toolbar-more',
        'xlsx-export-png',
        'xlsx-name-box',
      ]);
      expect(view.formulaInput().readOnly).toBe(true);
      // selection and copy still work.
      await press(view.surface, 'ArrowDown');
      expect(view.nameBox().value).toBe('B3');
      await press(view.surface, 'c', { ctrlKey: true });
      await view.idle();
      expect(calls.written).toHaveLength(1);
      expect(calls.sent('clearContents')).toEqual([]);
      expect(calls.sent('editCells')).toEqual([]);
      expect(calls.sent('undo')).toEqual([]);
      expect(await view.input(1, 1)).toBe(before);

      await view.setProps({ sending: false });
      await press(view.surface, '7');
      await press(view.editor()!, 'Enter');
      expect(await view.input(2, 1)).toBe('7');
    } finally {
      calls.restore();
    }
  }, MANY_KEYS_MS);

  it('shows the loading cursor on the grid and formula bar while sending', async () => {
    const view = await mountEditor(plain, undefined, undefined, { sending: true });
    expect(view.surface.getAttribute('aria-busy')).toBe('true');
    expect(view.surface.style.cursor).toBe('progress');
    expect(view.formulaInput().style.cursor).toBe('progress');
    fireEvent.mouseMove(view.surface, pointAt(plain, { row: 2, col: 2 }));
    expect(view.surface.style.cursor).toBe('progress');
    await view.setProps({ sending: false });
    expect(view.surface.getAttribute('aria-busy')).toBeNull();
    expect(view.surface.style.cursor).toBe('default');
    expect(view.formulaInput().style.cursor).toBe('');
  }, MANY_KEYS_MS);

  // the worker calls made, the first read held until released, and what
  // reaches the clipboard.
  const holdFirstRead = (view: Awaited<ReturnType<typeof mountEditor>>) => {
    const proxy = view.workbook();
    const forward = proxy.onSheet.bind(proxy);
    let open!: () => void;
    const held = new Promise<void>((resolve) => {
      open = resolve;
    });
    const sent: string[] = [];
    const spy = spyOn(proxy, 'onSheet').mockImplementation(((sheetId, method, ...args) => {
      sent.push(method);
      if (method === 'rangeCells' && sent.filter((name) => name === 'rangeCells').length === 1)
        return held.then(() => forward(sheetId, method, ...args));
      return forward(sheetId, method, ...args);
    }) as WorkbookProxy['onSheet']);
    const written: string[] = [];
    const clipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: { write: writeInto(written) },
    });
    return {
      sent,
      written,
      release: () =>
        act(async () => {
          open();
        }),
      restore: () => {
        spy.mockRestore();
        if (clipboard) Object.defineProperty(navigator, 'clipboard', clipboard);
        else Reflect.deleteProperty(navigator, 'clipboard');
      },
    };
  };

  it('refuses edits while a cut reads its selection, then clears what it copied', async () => {
    const view = await mountEditor(farColumn);
    const read = holdFirstRead(view);
    try {
      await view.click({ row: 0, col: 0 });
      await press(view.surface, 'a', { ctrlKey: true });
      await press(view.surface, 'x', { ctrlKey: true });
      expect(view.surface.getAttribute('aria-busy')).toBe('true');
      await press(view.surface, 'Delete');
      await press(view.surface, '7');
      expect(view.editor()).toBeNull();
      expect(read.sent.filter((name) => name === 'clearContents' || name === 'editCells')).toEqual([]);
      await read.release();
      await waitFor(() => expect(view.surface.getAttribute('aria-busy')).toBeNull(), { timeout: 20_000 });
      await view.idle();
      expect(read.sent.filter((name) => name === 'clearContents')).toHaveLength(1);
      expect(await view.input(89_999, 9)).toBe('');
    } finally {
      read.restore();
    }
  }, MANY_KEYS_MS);

  it('ends a cut as a copy when sending starts during its read', async () => {
    const view = await mountEditor(farColumn);
    const read = holdFirstRead(view);
    try {
      await view.click({ row: 0, col: 0 });
      await press(view.surface, 'a', { ctrlKey: true });
      await press(view.surface, 'x', { ctrlKey: true });
      await view.setProps({ sending: true });
      await read.release();
      await waitFor(() => expect(read.written).toHaveLength(1), { timeout: 20_000 });
      await view.idle();
      expect(read.sent.filter((name) => name === 'clearContents')).toEqual([]);
      expect(await view.input(89_999, 9)).toBe('1');
    } finally {
      read.restore();
    }
  }, MANY_KEYS_MS);

  // the latest copy or cut owns the clipboard: an older read writes nothing.
  it('keeps the clipboard of a cut made while an earlier copy still read', async () => {
    const view = await mountEditor(plain);
    const read = holdFirstRead(view);
    try {
      await view.click({ row: 0, col: 0 });
      await press(view.surface, 'a', { ctrlKey: true });
      await press(view.surface, 'c', { ctrlKey: true });
      await view.click({ row: 1, col: 1 });
      await press(view.surface, 'x', { ctrlKey: true });
      expect([read.written, await view.input(1, 1)]).toEqual([['Q1'], '']);
      await read.release();
      await view.idle();
      expect(read.written).toEqual(['Q1']);
    } finally {
      read.restore();
    }
  }, MANY_KEYS_MS);

  // the browser's permission lasts a few seconds after the key press: the
  // write starts there, filled once the read answers.
  it('starts the clipboard write at the key press, before the cells are read', async () => {
    const view = await mountEditor(plain);
    const read = holdFirstRead(view);
    let started = 0;
    const written: string[] = [];
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        write: (items: ClipboardItem[]) => {
          started += 1;
          return writeInto(written)(items);
        },
      },
    });
    try {
      await view.click({ row: 1, col: 1 });
      await press(view.surface, 'c', { ctrlKey: true });
      expect([started, written]).toEqual([1, []]);
      await read.release();
      await view.idle();
      expect(written).toEqual(['Q1']);
    } finally {
      read.restore();
    }
  }, MANY_KEYS_MS);

  it('shows the copy-failed notice when the browser refuses the write, a cut clearing nothing', async () => {
    const notices: string[] = [];
    const view = await mountEditor(plain, undefined, undefined, {
      onNotice: (notice) => notices.push(notice),
    });
    const calls = watch(view);
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        write: async () => {
          throw new Error('Document is not focused.');
        },
      },
    });
    try {
      await view.click({ row: 1, col: 1 });
      await press(view.surface, 'c', { ctrlKey: true });
      await press(view.surface, 'x', { ctrlKey: true });
      await view.idle();
      expect([notices, calls.sent('clearContents'), await view.input(1, 1)]).toEqual([
        ['copy-failed', 'copy-failed'],
        [],
        'Q1',
      ]);
    } finally {
      calls.restore();
    }
  }, MANY_KEYS_MS);

  it('writes the text once read where the browser has no ClipboardItem', async () => {
    const view = await mountEditor(plain);
    const item = Object.getOwnPropertyDescriptor(globalThis, 'ClipboardItem');
    const clipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
    const written: string[] = [];
    Reflect.deleteProperty(globalThis, 'ClipboardItem');
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: { writeText: async (text: string) => void written.push(text) },
    });
    try {
      await view.click({ row: 1, col: 1 });
      await press(view.surface, 'c', { ctrlKey: true });
      await view.idle();
      expect(written).toEqual(['Q1']);
    } finally {
      if (item) Object.defineProperty(globalThis, 'ClipboardItem', item);
      if (clipboard) Object.defineProperty(navigator, 'clipboard', clipboard);
      else Reflect.deleteProperty(navigator, 'clipboard');
    }
  }, MANY_KEYS_MS);

  // J90000:J90010: the cut's first of ten pieces is held, then J90010 is copied.
  it('clears nothing when a copy made during a cut read supersedes it', async () => {
    const view = await mountEditor(farColumn);
    const read = holdFirstRead(view);
    try {
      await view.click({ row: 0, col: 0 });
      await press(view.surface, 'a', { ctrlKey: true });
      await press(view.surface, 'x', { ctrlKey: true });
      await press(view.surface, 'End', { ctrlKey: true });
      await press(view.surface, 'c', { ctrlKey: true });
      expect(read.written).toEqual(['11']);
      await read.release();
      await waitFor(() => expect(view.surface.getAttribute('aria-busy')).toBeNull(), { timeout: 20_000 });
      await view.idle();
      expect([
        read.sent.filter((name) => name === 'rangeCells').length,
        read.sent.filter((name) => name === 'clearContents'),
        read.written,
      ]).toEqual([2, [], ['11']]);
      expect(await view.input(89_999, 9)).toBe('1');
    } finally {
      read.restore();
    }
  }, MANY_KEYS_MS);

  // each edit path with an edit ready to land: the changes before the pause
  // leave an undo and a redo, and the paint brush is armed on B2.
  for (const [mode, pause] of [
    ['read-only', { readOnly: true }],
    ['sending', { sending: true }],
  ] as const) {
    it(`refuses every edit while ${mode}, and lands none once it ends`, async () => {
      let changes = 0;
      const view = await mountEditor(plain, undefined, () => {
        changes += 1;
      });
      const calls = watch(view);
      Object.defineProperty(navigator, 'clipboard', {
        configurable: true,
        value: {
          readText: async () => 'pasted',
          write: writeInto(calls.written),
        },
      });
      const paint = () => document.querySelector('[title="Paint format"]') as HTMLButtonElement;
      try {
        await view.click({ row: 1, col: 1 });
        const below = await view.input(2, 1);
        await press(view.surface, '1');
        await press(view.editor()!, 'Enter');
        await press(view.surface, '2');
        await press(view.editor()!, 'Enter');
        await press(view.surface, 'z', { ctrlKey: true });
        await view.click({ row: 1, col: 1 });
        await act(async () => {
          fireEvent.click(paint());
        });
        await view.idle();
        expect(paint().getAttribute('aria-pressed')).toBe('true');
        expect(changes).toBe(3);

        await view.setProps(pause);
        // the brush would land on the next selection.
        await view.click({ row: 4, col: 3 });
        for (const key of ['7', 'F2', 'Delete', 'Backspace']) await press(view.surface, key);
        await view.doubleClick({ row: 4, col: 3 });
        expect(view.editor()).toBeNull();
        await press(view.surface, 'v', { ctrlKey: true });
        await press(view.surface, 'x', { ctrlKey: true });
        await press(view.surface, 'z', { ctrlKey: true });
        await press(view.surface, 'y', { ctrlKey: true });
        await press(view.surface, 'z', { ctrlKey: true, shiftKey: true });
        for (const command of XLSX_COMMANDS.filter(xlsxCommandEdits)) await view.run(command);
        expect(view.liveControls()).toEqual([
          'xlsx-search-menus',
          'xlsx-zoom',
          null,
          'xlsx-toolbar-more',
          'xlsx-export-png',
          'xlsx-name-box',
        ]);
        expect(view.formulaInput().readOnly).toBe(true);
        await view.idle();
        // a cut copies only.
        expect(calls.written).toHaveLength(1);
        expect(changes).toBe(3);

        await view.setProps({ readOnly: false, sending: false });
        await settle();
        await view.idle();
        expect(changes).toBe(3);
        expect(calls.sent('clearContents')).toEqual([]);
        expect(calls.sent('editCells')).toEqual([]);
        expect(calls.sent('applyFormat')).toEqual([]);
        expect(calls.sent('applyOps')).toEqual([]);
        // neither the undo nor the redo landed.
        expect(await view.input(1, 1)).toBe('1');
        expect(await view.input(2, 1)).toBe(below);
        await press(view.surface, '7');
        await press(view.editor()!, 'Enter');
        expect(await view.input(4, 3)).toBe('7');
        expect(changes).toBe(4);
      } finally {
        calls.restore();
      }
    }, MANY_KEYS_MS);

    it(`refuses chart drags and nudges while ${mode}`, async () => {
      const [chart] = charted.charts;
      let changes = 0;
      const view = await mountEditor(charted, undefined, () => {
        changes += 1;
      });
      await selectChart(view, chart);
      await view.setProps(pause);
      for (let step = 0; step < 3; step++) await press(view.surface, 'ArrowRight');
      expect(view.outlineAt().x).toBe(Math.round(chart.rect.x));
      const start = chartCenter(chart);
      fireEvent.mouseDown(view.surface, start);
      fireEvent.mouseMove(view.surface, {
        clientX: start.clientX + 40,
        clientY: start.clientY + 24,
        buttons: 1,
      });
      await act(async () => {
        fireEvent.mouseUp(window, { clientX: start.clientX + 40, clientY: start.clientY + 24 });
      });
      await settle();
      await view.idle();
      expect(view.outlineAt().x).toBe(Math.round(chart.rect.x));
      expect(changes).toBe(0);
    }, MANY_KEYS_MS);
  }

  it('keeps an open cell edit read-only while sending, then commits it', async () => {
    const view = await mountEditor(plain);
    const calls = watch(view);
    try {
      await view.doubleClick({ row: 2, col: 2 });
      view.type('draft');
      await view.setProps({ sending: true });
      expect(view.editor()!.readOnly).toBe(true);
      await press(view.editor()!, 'Enter');
      expect(view.editor()!.value).toBe('draft');
      expect(calls.sent('editCell')).toEqual([]);
      await view.setProps({ sending: false });
      await press(view.editor()!, 'Enter');
      expect(await view.input(2, 2)).toBe('draft');
    } finally {
      calls.restore();
    }
  }, MANY_KEYS_MS);

  it('refuses a sheet switch, the name box, clearSelection and selectCells while sending with an edit open', async () => {
    const view = await mountEditor(plain);
    const tabs = () => [...view.surface.ownerDocument.querySelectorAll('[role="tab"]')];
    await view.doubleClick({ row: 2, col: 2 });
    view.type('draft');
    await view.setProps({ sending: true });
    await act(async () => {
      fireEvent.click(tabs()[1]);
    });
    await view.idle();
    fireEvent.change(view.nameBox(), { target: { value: 'D5' } });
    await press(view.nameBox(), 'Enter');
    await act(async () => view.api().clearSelection());
    let selecting!: Promise<boolean>;
    await act(async () => {
      selecting = view.api().selectCells(1, selectionAt({ row: 4, col: 0 }));
      await selecting;
    });
    expect(await selecting).toBe(false);
    expect([tabs()[0].getAttribute('aria-selected'), view.editor()?.value]).toEqual(['true', 'draft']);
    await view.setProps({ sending: false });
    await press(view.editor()!, 'Enter');
    expect(await view.input(2, 2)).toBe('draft');
  }, MANY_KEYS_MS);

  it('keeps an open cell or formula-bar draft on its cell while sending when another cell is pressed', async () => {
    const view = await mountEditor(plain);
    const d5 = await view.input(4, 3);
    await view.doubleClick({ row: 2, col: 2 });
    view.type('draft');
    await view.setProps({ sending: true });
    await view.click({ row: 4, col: 3 });
    expect([view.editor()?.value, view.nameBox().value]).toEqual(['draft', 'C3']);
    await view.setProps({ sending: false });
    await press(view.editor()!, 'Enter');
    expect(await view.input(2, 2)).toBe('draft');

    await view.click({ row: 10, col: 1 });
    await act(async () => view.formulaInput().focus());
    view.typeFormula('=B3*2');
    await view.setProps({ sending: true });
    await view.click({ row: 4, col: 3 });
    expect([view.formula(), view.nameBox().value]).toEqual(['=B3*2', 'B11']);
    await view.setProps({ sending: false });
    await press(view.formulaInput(), 'Enter');
    expect([await view.input(10, 1), await view.input(4, 3)]).toEqual(['=B3*2', d5]);
  }, MANY_KEYS_MS);

  // text the engine never refuses: only the busy press refusal keeps it on B2.
  it('commits a formula-bar text draft typed before sending to its own cell after a press', async () => {
    const view = await mountEditor(plain);
    const d5 = await view.input(4, 3);
    await view.click({ row: 1, col: 1 });
    await act(async () => view.formulaInput().focus());
    view.typeFormula('hello');
    await view.setProps({ sending: true });
    await view.click({ row: 4, col: 3 });
    expect([view.nameBox().value, view.formula()]).toEqual(['B2', 'hello']);
    await view.setProps({ sending: false });
    await act(async () => view.formulaInput().focus());
    await press(view.formulaInput(), 'Enter');
    expect([await view.input(1, 1), await view.input(4, 3)]).toEqual(['hello', d5]);
  }, MANY_KEYS_MS);

  it('reveals a far cell on another sheet the name box goes to', async () => {
    const view = await mountEditor(plain);
    await view.click({ row: 0, col: 0 });
    const box = view.nameBox();
    fireEvent.change(box, { target: { value: "'Summary'!A5000" } });
    await press(box, 'Enter');
    await view.idle();
    expect((await view.workbook().sheetInfo()).activeSheet).toBe(1);
    expect(view.nameBox().value).toBe('A5000');
    const cell = await view.workbook().cellPosition(1, 4999, 0);
    const shown = view.shown();
    expect(cell.y).toBeGreaterThanOrEqual(shown.y);
    expect(cell.y).toBeLessThan(shown.y + shown.height);
    // the scroll area grew to the cell, so a browser can scroll there.
    const spacer = view.surface.firstElementChild as HTMLElement;
    expect(parseFloat(spacer.style.height)).toBeGreaterThan(cell.y);
  }, MANY_KEYS_MS);

  it('selects a cell past the scroll limit without growing the scroll area beyond it', async () => {
    const view = await mountEditor(plain);
    await view.click({ row: 0, col: 0 });
    const box = view.nameBox();
    fireEvent.change(box, { target: { value: 'XFD1048576' } });
    await press(box, 'Enter');
    await view.idle();
    expect(view.nameBox().value).toBe('XFD1048576');
    const spacer = view.surface.firstElementChild as HTMLElement;
    expect(parseFloat(spacer.style.height)).toBe(MAX_SCROLL_EXTENT_PX);
    expect(parseFloat(spacer.style.width)).toBeLessThanOrEqual(MAX_SCROLL_EXTENT_PX);
    // the next key still lands in the selected cell.
    await press(view.surface, '7');
    await press(view.editor()!, 'Enter');
    expect(await view.input(1_048_575, 16_383)).toBe('7');
  }, MANY_KEYS_MS);

  it('selects the cell an address typed into the name box names', async () => {
    const view = await mountEditor(plain);
    await view.click({ row: 0, col: 0 });
    const box = view.nameBox();
    expect(box.readOnly).toBe(false);
    fireEvent.change(box, { target: { value: 'c200' } });
    await press(box, 'Enter');
    expect(view.nameBox().value).toBe('C200');
    await press(view.surface, '7');
    await press(view.editor()!, 'Enter');
    expect(await view.input(199, 2)).toBe('7');

    // another sheet's address opens that sheet; a range goes to its first cell.
    fireEvent.change(view.nameBox(), { target: { value: "'Summary'!b3:c4" } });
    await press(view.nameBox(), 'Enter');
    await view.idle();
    expect((await view.workbook().sheetInfo()).activeSheet).toBe(1);
    expect(view.nameBox().value).toBe('B3');
    // text that names no cell stays in the box.
    fireEvent.change(view.nameBox(), { target: { value: 'nowhere' } });
    await press(view.nameBox(), 'Enter');
    expect(view.nameBox().value).toBe('nowhere');
  }, MANY_KEYS_MS);

  it('keeps a revealed cell clear of the frozen panes', async () => {
    const view = await mountEditor(wideFrozen);
    const pane = cellRect(wideFrozen.grid, 2, 2)!;
    await view.click({ row: 3, col: 1 });
    for (let step = 0; step < 15; step += 1) await press(view.surface, 'ArrowRight');
    for (let step = 0; step < 30; step += 1) await press(view.surface, 'ArrowDown');
    await press(view.surface, '7');
    const box = editorBox(view.editor()!);
    expect(box.left).toBeGreaterThanOrEqual(pane.x);
    expect(box.top).toBeGreaterThanOrEqual(pane.y);
    expect(box.right).toBeLessThanOrEqual(VIEWPORT.width);
    expect(box.bottom).toBeLessThanOrEqual(VIEWPORT.height);

    // back into the frozen columns: the rows stay where they are.
    const top = view.surface.scrollTop;
    await press(view.editor()!, 'Escape');
    for (let step = 0; step < 15; step += 1) await press(view.surface, 'ArrowLeft');
    expect(view.nameBox().value).toBe('B34');
    expect(view.surface.scrollTop).toBe(top);
  }, MANY_KEYS_MS);

  it('brings a cell scrolled out of view back when typing into it', async () => {
    const view = await mountEditor(wide);
    await view.click({ row: 3, col: 1 });
    await scrollTo(view, 1500, 600);

    await press(view.surface, '7');
    expect(view.editor()?.value).toBe('7');
    const box = editorBox(view.editor()!);
    // at the window's top-left, give or take the engine's f32 rounding.
    expect(box.left).toBeCloseTo(0, 1);
    expect(box.top).toBeCloseTo(0, 1);
    await press(view.editor()!, 'Enter');
    expect((await view.input(3, 1))).toBe('7');
  });

  it('leaves the scroll alone for a move that stays on screen', async () => {
    const view = await mountEditor(wide);
    await view.click({ row: 3, col: 1 });
    await press(view.surface, 'ArrowRight');
    await press(view.surface, 'ArrowDown');
    expect(view.nameBox().value).toBe('C5');
    expect(view.surface.scrollLeft).toBe(0);
    expect(view.surface.scrollTop).toBe(0);
  });

  it('keeps the view on select all, as the menu does', async () => {
    const view = await mountEditor(wide);
    await view.click({ row: 3, col: 1 });
    await press(view.surface, 'a', { ctrlKey: true });
    expect(view.selectionBox() !== null).toBe(true);
    expect(view.surface.scrollLeft).toBe(0);
    expect(view.surface.scrollTop).toBe(0);
  });

  it('keeps typing into an open edit whose cell scrolled away, and scrolls back', async () => {
    const view = await mountEditor(wide);
    await view.click({ row: 3, col: 1 });
    await press(view.surface, '7');
    const input = view.editor()!;
    await scrollTo(view, 1500, 600);

    // still mounted and focused, out of sight.
    expect(view.editor() === input).toBe(true);
    expect(document.activeElement === input).toBe(true);
    expect(input.style.opacity).toBe('0');

    await press(input, '5');
    fireEvent.change(input, { target: { value: '75' } });
    expect(view.editor() === input).toBe(true);
    expect(input.style.opacity).toBe('');
    const box = editorBox(input);
    expect(box.left).toBeCloseTo(0, 1);
    expect(box.top).toBeCloseTo(0, 1);
    await press(input, 'Enter');
    expect((await view.input(3, 1))).toBe('75');
  });

  it('keeps an IME composition in the edit while its cell scrolls back', async () => {
    const view = await mountEditor(wide);
    await view.click({ row: 3, col: 1 });
    await press(view.surface, 'F2');
    const input = view.editor()!;
    await scrollTo(view, 1500, 600);

    fireEvent.compositionStart(input);
    await press(input, 'Process', { isComposing: true, keyCode: 229 });
    expect(view.editor() === input).toBe(true);
    expect(document.activeElement === input).toBe(true);
    expect(input.style.opacity).toBe('');
    fireEvent.change(input, { target: { value: '日本' } });
    fireEvent.compositionEnd(input, { data: '日本' });
    await press(input, 'Enter');
    expect((await view.input(3, 1))).toBe('日本');
  });

  it('types into a cell the frozen panes leave no room to show', async () => {
    const view = await mountEditor(paneFilled);
    await view.click({ row: 3, col: 1 });
    await press(view.surface, 'End', { ctrlKey: true });
    const { row, col } = cellAt(view.nameBox().value);
    expect(row).toBeGreaterThan(40);

    await press(view.surface, '7');
    const input = view.editor()!;
    expect(document.activeElement === input).toBe(true);
    await press(input, '8');
    fireEvent.change(input, { target: { value: '78' } });
    await press(input, 'Enter');
    expect((await view.input(row, col))).toBe('78');
  });

  it('leaves a scrolled-away edit where it is for keys that type nothing', async () => {
    const view = await mountEditor(wide);
    await view.click({ row: 3, col: 1 });
    await press(view.surface, '7');
    await scrollTo(view, 1500, 600);
    const input = view.editor()!;
    for (const key of ['Meta', 'Control', 'Shift', 'Alt']) await press(input, key);
    await press(input, 'c', { metaKey: true });
    expect([view.surface.scrollLeft, view.surface.scrollTop]).toEqual([1500, 600]);
    expect(input.style.opacity).toBe('0');
    expect(document.activeElement === input).toBe(true);
  });

  it('draws an Enter that commits and scrolls in at most two frames, the last one live', async () => {
    const view = await mountEditor(wide);
    // the last row the window shows whole: Enter moves below the edge.
    const offsets = wide.grid.rowOffsets;
    const last = offsets.filter((bottom) => bottom <= VIEWPORT.height).length - 2;
    await view.click({ row: last, col: 1 });
    await press(view.surface, '7');
    const counts = countEngine(view.workbook());
    await press(view.editor()!, 'Enter');
    // what the browser does next: the scroll event, then a frame.
    await act(async () => {
      fireEvent.scroll(view.surface);
      await new Promise((resolve) => requestAnimationFrame(resolve));
    });
    await view.idle();
    expect(view.surface.scrollTop).toBeGreaterThan(0);
    expect(await view.input(last, 1)).toBe('7');
    // the commit's frame, then the scrolled window's.
    expect(counts.paints).toBeLessThanOrEqual(2);
    expect(view.shown().y).toBe(view.surface.scrollTop);
    expect(view.selectionBox() !== null).toBe(true);
  });

  it('judges a key against the live view, not a frame its scroll has not repainted', async () => {
    const view = await mountEditor(wide);
    await view.click({ row: 3, col: 1 });
    // scrolled, and the key lands before the scroll's frame is painted.
    await act(async () => {
      view.surface.scrollLeft = 1500;
    });
    await press(view.surface, 'ArrowRight');
    expect(view.nameBox().value).toBe('C4');
    expect(view.surface.scrollLeft).toBe(Math.floor((await view.workbook().cellPosition(0, 3, 2)).x));
  });

  it('brings a partly hidden edit fully into view when typing into it', async () => {
    const view = await mountEditor(wide);
    await view.click({ row: 3, col: 1 });
    await press(view.surface, '7');
    const start = (await view.workbook().cellPosition(0, 3, 1)).x;
    await scrollTo(view, Math.ceil(start) + 10, 0);
    expect(view.editor()!.style.opacity).toBe('');
    await press(view.editor()!, '5');
    expect(view.surface.scrollLeft).toBe(Math.floor(start));
  });

  it('shows a cell wider than the view as it is, typing or moving down', async () => {
    const view = await mountEditor(wideColumn);
    await view.click({ row: 3, col: 1 });
    await press(view.surface, '7');
    const input = view.editor()!;
    // inside the wide cell, to the right of its start.
    await scrollTo(view, view.surface.scrollLeft + 500, 0);
    const counts = countEngine(view.workbook());
    const left = view.surface.scrollLeft;
    for (const key of ['a', 'b', 'c']) {
      await press(input, key);
      fireEvent.change(input, { target: { value: `${input.value}${key}` } });
    }
    expect(view.surface.scrollLeft).toBe(left);
    await press(input, 'Escape');
    for (let step = 0; step < 3; step += 1) await press(view.surface, 'ArrowDown');
    expect(view.nameBox().value).toBe('B7');
    expect(view.surface.scrollLeft).toBe(left);
    // no position asked of the worker: the frame on screen answered.
    expect(counts.positions).toBe(0);
  });

  it('judges the cell an Enter lands on against the row its commit grew', async () => {
    const view = await mountEditor(tallFont);
    // Enter lands on the last row the window shows whole.
    const target = tallFont.grid.rowOffsets.filter((bottom) => bottom <= VIEWPORT.height).length - 2;
    await view.click({ row: target - 1, col: 5 });
    await press(view.surface, '7');
    await press(view.editor()!, 'Enter');
    expect((await view.input(target - 1, 5))).toBe('7');
    const cell = await view.workbook().cellPosition(0, target, 5);
    expect(cell.y + cell.height).toBeLessThanOrEqual(view.surface.scrollTop + VIEWPORT.height);
    expect(view.surface.scrollTop).toBeGreaterThan(0);
  });

  it('asks the engine nothing for the first cell under frozen panes at scroll 0', async () => {
    const view = await mountEditor(wideFrozen);
    const counts = countEngine(view.workbook());
    await view.click({ row: 2, col: 2 });
    await press(view.surface, '7');
    const input = view.editor()!;
    for (const key of ['a', 'b', 'c']) {
      await press(input, key);
      fireEvent.change(input, { target: { value: `${input.value}${key}` } });
    }
    expect([view.surface.scrollLeft, view.surface.scrollTop]).toEqual([0, 0]);
    expect(counts.positions).toBe(0);
  });

  it('commits an edit a press on nothing focusable ends, and gives the grid the keys', async () => {
    const view = await mountEditor(wide);
    await view.click({ row: 3, col: 1 });
    await press(view.surface, '7');
    await act(async () => {
      fireEvent.blur(view.editor()!, { relatedTarget: null });
    });
    expect(view.editor() === null).toBe(true);
    expect((await view.input(3, 1))).toBe('7');
    // the grid takes the keys once focus has settled, a frame later.
    await act(async () => {
      await new Promise((resolve) => requestAnimationFrame(resolve));
    });
    // identity checks stay cheap to print when they fail.
    expect(document.activeElement?.getAttribute('data-testid')).toBe('xlsx-scroll');
  });

  it('keeps a formula-bar draft across a window or tab switch, commits it on a blur', async () => {
    const view = await mountEditor(wide);
    await view.click({ row: 3, col: 1 });
    const formula = view.formulaInput();
    const before = (await view.input(3, 1));
    act(() => formula.focus());
    view.typeFormula('42');
    const away = spyOn(document, 'hasFocus').mockReturnValue(false);
    try {
      await act(async () => {
        fireEvent.blur(formula);
      });
    } finally {
      away.mockRestore();
    }
    expect((await view.input(3, 1))).toBe(before);
    expect(formula.value).toBe('42');
    await act(async () => {
      fireEvent.blur(formula);
    });
    expect((await view.input(3, 1))).toBe('42');
  });

  it('keeps the edit open across a window or tab switch', async () => {
    const view = await mountEditor(wide);
    await view.click({ row: 3, col: 1 });
    await press(view.surface, '7');
    const input = view.editor()!;
    const before = (await view.input(3, 1));
    const away = spyOn(document, 'hasFocus').mockReturnValue(false);
    try {
      await act(async () => {
        fireEvent.blur(input, { relatedTarget: null });
      });
    } finally {
      away.mockRestore();
    }
    expect(view.editor() === input).toBe(true);
    expect((await view.input(3, 1))).toBe(before);
    // back: the next key types into the edit.
    fireEvent.change(input, { target: { value: '78' } });
    await press(input, 'Enter');
    expect((await view.input(3, 1))).toBe('78');
  });

  it('gives the grid the keys after a sheet tab is clicked', async () => {
    const view = await mountEditor(wide);
    const tab = view.surface.ownerDocument.querySelectorAll<HTMLButtonElement>('[role="tab"]')[1];
    await act(async () => {
      tab.focus();
      fireEvent.click(tab);
    });
    expect(tab.getAttribute('aria-selected')).toBe('true');
    expect(document.activeElement?.getAttribute('data-testid')).toBe('xlsx-scroll');
  });

  it('keeps the top-left cell where it is when the zoom changes, as Google Sheets does', async () => {
    const view = await mountEditor(wide);
    await scrollTo(view, 480, 240);
    for (const command of ['zoom:200', 'zoom:50', 'zoom:90', 'zoom:125'] as const) {
      await view.run(command);
      expect(view.shown().x).toBeCloseTo(480, 6);
      expect(view.shown().y).toBeCloseTo(240, 6);
    }
  });

  it('follows the pointer and the keyboard at 200%, clear of the frozen panes', async () => {
    const view = await mountEditor(wideFrozen);
    await view.run('zoom:200');
    // a point twice as far from the corner is the same cell at 200%.
    const b4 = pointAt(wideFrozen, { row: 3, col: 1 });
    const zoomed = { clientX: b4.clientX * 2, clientY: b4.clientY * 2 };
    fireEvent.mouseDown(view.surface, zoomed);
    fireEvent.mouseUp(view.surface, zoomed);
    fireEvent.click(view.surface, zoomed);
    expect(view.nameBox().value).toBe('B4');

    for (let step = 0; step < 10; step += 1) await press(view.surface, 'ArrowRight');
    for (let step = 0; step < 20; step += 1) await press(view.surface, 'ArrowDown');
    expect(view.nameBox().value).toBe('L24');
    await press(view.surface, '7');
    const pane = cellRect(wideFrozen.grid, 2, 2)!;
    const box = editorBox(view.editor()!);
    expect(box.left).toBeGreaterThanOrEqual(pane.x * 2);
    expect(box.top).toBeGreaterThanOrEqual(pane.y * 2);
    expect(box.right).toBeLessThanOrEqual(VIEWPORT.width);
    expect(box.bottom).toBeLessThanOrEqual(VIEWPORT.height);
    await press(view.editor()!, 'Enter');
    expect((await view.input(23, 11))).toBe('7');
  }, MANY_KEYS_MS);

  it('leaves the formula bar focused when an edit it committed scrolls back', async () => {
    const view = await mountEditor(wide);
    await view.click({ row: 3, col: 1 });
    await press(view.surface, '7');
    const input = view.editor()!;
    await scrollTo(view, 1500, 600);

    const formula = view.formulaInput();
    act(() => formula.focus());
    expect(view.editor() === null).toBe(true);
    expect((await view.input(3, 1))).toBe('7');
    await scrollTo(view, 0, 0);
    expect(document.activeElement === formula).toBe(true);
    expect(input.isConnected).toBe(false);
  });
});

describe('XlsxEditor chart objects', () => {
  it('selects a chart instead of the cells behind it, and deselects off it', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);

    fireEvent.mouseDown(view.surface, chartCenter(chart));
    const outline = await waitFor(() => view.outline()!);
    expect(outline.getAttribute('data-chart-id')).toBe(chart.id);
    // the press must not reach the cells under the chart at all: the grid
    // selection stays exactly where it was, not merely hidden.
    expect(view.nameBox().value).toBe('A1');
    expect(view.selectionBox() === null).toBe(true);
    fireEvent.mouseUp(window, chartCenter(chart));

    fireEvent.mouseDown(view.surface, { clientX: 2, clientY: 2 });
    await waitFor(() => expect(view.outline() === null).toBe(true));
    expect(view.selectionBox() !== null).toBe(true);
    expect(view.nameBox().value).toBe('A1');
  });

  // the renderer paints a chart it cannot draw as a neutral box rather than
  // failing the frame. it is still an object on the sheet, so it must select
  // and move like any other.
  it('selects and moves a chart the renderer could not draw', async () => {
    const chart = undrawable.charts.find((candidate) => candidate.placeholder);
    expect(chart).toBeDefined();
    expect(chart!.movable).toBe(true);
    const view = await mountEditor(undrawable);
    await selectChart(view, chart!);

    expect(view.outline()!.getAttribute('data-chart-id')).toBe(chart!.id);
    await act(async () => {
      fireEvent.keyDown(view.surface, { key: 'ArrowRight', shiftKey: true });
    });
    await waitFor(() => expect(view.canUndo()).toBe(true), { timeout: 2000 });
    expect(view.error() === null).toBe(true);
    await waitFor(() => expect(view.outlineAt().x).toBe(Math.round(chart!.rect.x + 10)));
  });

  it('keeps cell-editing keys off the cells hidden behind a selected chart', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);
    const before = view.formula();

    fireEvent.mouseDown(view.surface, chartCenter(chart));
    await waitFor(() => view.outline()!);
    fireEvent.mouseUp(window, chartCenter(chart));

    for (const key of ['Delete', 'Backspace', 'x', 'Enter']) {
      await act(async () => {
        fireEvent.keyDown(view.surface, { key });
      });
    }

    expect(view.editor() === null).toBe(true);
    expect(view.outline() !== null).toBe(true);
    expect(view.formula()).toBe(before);
    expect(view.canUndo()).toBe(false);
  });

  it('commits an open edit before taking the press, then selects the chart', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);

    await view.doubleClick({ row: 1, col: 1 });
    view.type('Edited item');
    fireEvent.mouseDown(view.surface, chartCenter(chart));

    const outline = await waitFor(() => view.outline()!);
    expect(outline.getAttribute('data-chart-id')).toBe(chart.id);
    expect(view.editor() === null).toBe(true);
    expect((await view.input(1, 1))).toBe('Edited item');
    fireEvent.mouseUp(window, chartCenter(chart));
  });

  it('commits an edit whose input has scrolled out of the window', async () => {
    const SCROLLED_BY = 300;
    const probe = openWorkbook(charted.bytes.slice());
    let target: ChartRegion;
    try {
      target = (probe.displayList({ x: 0, y: SCROLLED_BY, ...VIEWPORT }).charts ?? [])[0];
    } finally {
      probe.dispose();
    }
    const view = await mountEditor(charted);

    await view.doubleClick({ row: 1, col: 1 });
    view.type('Edited item');

    await act(async () => {
      view.surface.scrollTop = SCROLLED_BY;
      fireEvent.scroll(view.surface);
      await new Promise((resolve) => requestAnimationFrame(resolve));
    });
    await view.idle();
    // the input stays mounted, out of sight, while its cell is away; the press
    // commits it without a blur.
    expect(view.editor()?.style.opacity).toBe('0');

    fireEvent.mouseDown(view.surface, chartCenter(target));
    await waitFor(() => view.outline()!);
    fireEvent.mouseUp(window, chartCenter(target));

    expect((await view.input(1, 1))).toBe('Edited item');
  });

  it('leaves a press inside the open editor to the editor, not a chart behind it', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);

    await view.doubleClick({ row: 1, col: 1 });
    view.type('Edited item');
    // the press lands on the input, at coordinates the chart also covers: the
    // editor is a dom overlay, so without its guard the hit test would reach
    // the chart painted underneath.
    fireEvent.mouseDown(view.editor()!, chartCenter(chart));

    expect(view.editor()?.value).toBe('Edited item');
    expect(view.outline() === null).toBe(true);
  });

  it('drags a selected chart and repins it through the engine', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);
    const start = chartCenter(chart);

    fireEvent.mouseDown(view.surface, start);
    await waitFor(() => view.outline()!);
    fireEvent.mouseMove(view.surface, {
      clientX: start.clientX + 40,
      clientY: start.clientY + 24,
      buttons: 1,
    });
    expect(view.outline()!.style.left).toBe(`${chart.rect.x + 40}px`);

    await act(async () => {
      fireEvent.mouseUp(window, { clientX: start.clientX + 40, clientY: start.clientY + 24 });
    });

    await waitFor(() => {
      const at = view.outlineAt();
      expect(at.x).toBe(Math.round(chart.rect.x + 40));
      expect(at.y).toBe(Math.round(chart.rect.y + 24));
    });
  });

  it('lands a run of arrow nudges as one undoable edit', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);
    await selectChart(view, chart);

    // key repeat: five presses preview locally and touch nothing.
    for (let press = 0; press < 5; press++) {
      await act(async () => {
        fireEvent.keyDown(view.surface, { key: 'ArrowRight' });
      });
    }
    expect(view.outlineAt().x).toBe(Math.round(chart.rect.x + 5));
    expect(view.canUndo()).toBe(false);

    await waitFor(() => expect(view.canUndo()).toBe(true), { timeout: 2000 });
    expect(view.outlineAt().x).toBe(Math.round(chart.rect.x + 5));

    // one burst is one undo step, not five.
    await act(async () => {
      fireEvent.keyDown(view.surface, { key: 'z', ctrlKey: true });
    });
    await waitFor(() => expect(view.outlineAt().x).toBe(Math.round(chart.rect.x)));
    expect(view.canUndo()).toBe(false);
  });

  it('drops a pending burst when the file is swapped under it', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);
    await selectChart(view, chart);

    await act(async () => {
      fireEvent.keyDown(view.surface, { key: 'ArrowRight', shiftKey: true });
    });
    expect(view.canUndo()).toBe(false);

    // the editor stays mounted across a file swap, so a burst still in flight
    // would fire against the workbook that replaced the one it was typed on.
    await view.reopenWith(charted);
    await waitFor(() => expect(view.nameBox().value).toBe('A1'));
    await settle();

    expect(view.error() === null).toBe(true);
    expect(view.canUndo()).toBe(false);
    const still = ((await view.workbook().displayList({ x: 0, y: 0, ...VIEWPORT })).charts ?? []).find(
      (candidate) => candidate.id === chart.id
    );
    expect(Math.round(still!.rect.x)).toBe(chartRounded(chart).x);
  });

  it('discards a burst that returns to where it started', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);
    await selectChart(view, chart);

    await act(async () => {
      fireEvent.keyDown(view.surface, { key: 'ArrowRight', shiftKey: true });
      fireEvent.keyDown(view.surface, { key: 'ArrowLeft', shiftKey: true });
    });
    expect(view.outlineAt().x).toBe(Math.round(chart.rect.x));

    await settle();
    expect(view.canUndo()).toBe(false);
    expect(view.outlineAt().x).toBe(Math.round(chart.rect.x));
  });

  it('lands a pending burst before a save reads the workbook', async () => {
    const [chart] = charted.charts;
    const saved: Uint8Array[] = [];
    const view = await mountEditor(charted, (bytes) => saved.push(bytes));
    await selectChart(view, chart);

    await act(async () => {
      fireEvent.keyDown(view.surface, { key: 'ArrowRight', shiftKey: true });
    });
    expect(view.canUndo()).toBe(false);

    await act(async () => {
      fireEvent.keyDown(view.surface, { key: 's', ctrlKey: true });
    });
    await view.idle();

    expect(view.canUndo()).toBe(true);
    expect(saved).toHaveLength(1);
    const reopened = openWorkbook(saved[0]);
    try {
      const moved = (reopened.displayList({ x: 0, y: 0, ...VIEWPORT }).charts ?? []).find(
        (candidate) => candidate.id === chart.id
      );
      expect(Math.round(moved!.rect.x)).toBe(Math.round(chart.rect.x + 10));
    } finally {
      reopened.dispose();
    }
  });

  it('drops a selection and its pending burst on escape', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);
    await selectChart(view, chart);

    await act(async () => {
      fireEvent.keyDown(view.surface, { key: 'ArrowRight', shiftKey: true });
    });
    fireEvent.keyDown(view.surface, { key: 'Escape' });

    await waitFor(() => expect(view.outline() === null).toBe(true));
    await settle();
    expect(view.canUndo()).toBe(false);
  });

  it('cancels an armed drag on escape instead of landing it on release', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);
    const start = chartCenter(chart);

    fireEvent.mouseDown(view.surface, start);
    await waitFor(() => view.outline()!);
    fireEvent.mouseMove(view.surface, {
      clientX: start.clientX + 40,
      clientY: start.clientY + 24,
      buttons: 1,
    });
    fireEvent.keyDown(view.surface, { key: 'Escape' });
    await waitFor(() => expect(view.outline() === null).toBe(true));

    await act(async () => {
      fireEvent.mouseUp(window, { clientX: start.clientX + 40, clientY: start.clientY + 24 });
    });

    expect(view.canUndo()).toBe(false);
  });

  it('never arms a drag from a non-primary press', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);
    const start = chartCenter(chart);

    fireEvent.mouseDown(view.surface, { ...start, button: 2 });
    const outline = await waitFor(() => view.outline()!);
    expect(outline.getAttribute('data-chart-id')).toBe(chart.id);
    // the context menu swallows the matching release, so the next unrelated
    // primary release anywhere must not land a move.
    await act(async () => {
      fireEvent.mouseUp(window, { clientX: start.clientX + 200, clientY: start.clientY + 150 });
    });

    expect(view.canUndo()).toBe(false);
    expect(view.outlineAt()).toEqual(chartRounded(chart));
  });

  it('disarms a drag whose release was lost when the next press arrives', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);
    const start = chartCenter(chart);

    fireEvent.mouseDown(view.surface, start);
    await waitFor(() => view.outline()!);
    // no mouseup: the release went somewhere this window never saw.
    fireEvent.mouseDown(view.surface, { clientX: 2, clientY: 2 });
    await act(async () => {
      fireEvent.mouseUp(window, { clientX: start.clientX + 300, clientY: start.clientY + 300 });
    });

    expect(view.canUndo()).toBe(false);
  });

  it('hit-tests the frame it painted, not a scroll offset it has not drawn yet', async () => {
    const SCROLLED_BY = 400;
    const target = charted.charts[0];
    const point = chartCenter(target);
    const probe = openWorkbook(charted.bytes.slice());
    try {
      // the fixture must genuinely disagree at this point across the two
      // viewports, or this proves nothing.
      const stale = probe.chartAtPoint(
        { x: 0, y: SCROLLED_BY, ...VIEWPORT },
        point.clientX,
        point.clientY
      );
      expect(stale?.id).not.toBe(target.id);
    } finally {
      probe.dispose();
    }

    const view = await mountEditor(charted);
    // freeze the repaint: scrolling now advances the scroll offset while the
    // canvas still shows the frame painted at the old one.
    const scheduled = window.requestAnimationFrame;
    window.requestAnimationFrame = (() => 0) as typeof window.requestAnimationFrame;
    try {
      view.surface.scrollTop = SCROLLED_BY;
      fireEvent.scroll(view.surface);
      fireEvent.mouseDown(view.surface, point);
      const outline = await waitFor(() => view.outline()!);
      expect(outline.getAttribute('data-chart-id')).toBe(target.id);
      fireEvent.mouseUp(window, point);
    } finally {
      window.requestAnimationFrame = scheduled;
    }
  });

  it('selects a chart pinned to the sheet but never drags it', async () => {
    const pinned = charted.charts.find((chart) => !chart.movable);
    expect(pinned).toBeDefined();
    const view = await mountEditor(charted);
    const start = chartCenter(pinned!);

    fireEvent.mouseDown(view.surface, start);
    const outline = await waitFor(() => view.outline()!);
    expect(outline.getAttribute('data-chart-id')).toBe(pinned!.id);

    await act(async () => {
      fireEvent.mouseMove(view.surface, {
        clientX: start.clientX + 40,
        clientY: start.clientY + 24,
        buttons: 1,
      });
      fireEvent.mouseUp(window, { clientX: start.clientX + 40, clientY: start.clientY + 24 });
    });

    expect(view.outlineAt().x).toBe(chartRounded(pinned!).x);

    await act(async () => {
      fireEvent.keyDown(view.surface, { key: 'ArrowRight', shiftKey: true });
    });
    await settle();
    expect(view.nameBox().value).toBe('A1');
    expect(view.outlineAt().x).toBe(chartRounded(pinned!).x);
    // the engine refuses to repin an absolute anchor, so a nudge that reached
    // it would surface as an error overlay rather than doing nothing.
    expect(view.error() === null).toBe(true);
    expect(view.canUndo()).toBe(false);

    fireEvent.mouseMove(view.surface, start);
    expect(view.surface.style.cursor).toBe('pointer');
  });

  it('abandons a drag whose release the window never saw', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);
    const start = chartCenter(chart);

    fireEvent.mouseDown(view.surface, start);
    await waitFor(() => view.outline()!);
    // the button comes back up off-window: the next move reports none held.
    fireEvent.mouseMove(view.surface, {
      clientX: start.clientX + 60,
      clientY: start.clientY + 60,
      buttons: 0,
    });

    await act(async () => {
      fireEvent.mouseUp(window, { clientX: start.clientX + 500, clientY: start.clientY + 500 });
    });

    expect(view.outlineAt().x).toBe(chartRounded(chart).x);
    expect(view.canUndo()).toBe(false);
  });

  it('drops a selection that scrolls out of the painted frame', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);

    fireEvent.mouseDown(view.surface, chartCenter(chart));
    await waitFor(() => view.outline()!);
    fireEvent.mouseUp(window, chartCenter(chart));

    await act(async () => {
      view.surface.scrollTop = 4000;
      fireEvent.scroll(view.surface);
    });

    // no invisible selection left swallowing the keyboard.
    await waitFor(() => expect(view.outline() === null).toBe(true));
    await waitFor(() => expect(view.selectionBox() !== null).toBe(true));
  });
});

describe('XlsxEditor with a peer', () => {
  for (const source of ['cell', 'formula bar'] as const)
    it(`drops a half-typed ${source} edit whose sheet the peer removes, and takes no more keys`, async () => {
      const peer = openWorkbook(plain.bytes.slice(), { collaborative: true, clientId: 3101 });
      let mine: WorkbookProxy | undefined;
      const view = render(
        <XlsxEditor
          file={plain.bytes.slice()}
          collaboration={{ clientId: 3102 }}
          onReady={(api) => {
            mine = api.handle;
          }}
        />
      );
      try {
        await waitFor(() => expect(mine).toBeDefined());
        const editor = mine!;
        const idle = () => idleWith(() => editor);
        editor.onUpdate((update, origin) => {
          if (origin === 'local') peer.applyUpdate(update);
        });
        // the peer's edits reach the editor as a room would deliver them.
        const fromPeer = async (change: () => void) => {
          await act(async () => {
            const before = await editor.encodeStateVector();
            change();
            await editor.applyUpdate(peer.encodeStateAsUpdate(before));
          });
          await idle();
        };
        await fromPeer(() => {
          peer.applyOps([{ type: 'addSheet', index: 3, name: 'Added' }]);
        });
        const added = view.getAllByRole('tab')[3];
        expect(added.textContent).toBe('Added');
        await act(async () => {
          fireEvent.click(added);
        });
        await idle();
        const surface = view.getByTestId('xlsx-scroll');
        const formula = view.getByTestId('xlsx-formula-input') as HTMLInputElement;
        if (source === 'cell') {
          await act(async () => {
            fireEvent.keyDown(surface, { key: 'd' });
          });
          const input = view.getByTestId('xlsx-cell-editor') as HTMLInputElement;
          expect(document.activeElement === input).toBe(true);
          fireEvent.change(input, { target: { value: 'draft' } });
        } else {
          act(() => formula.focus());
          fireEvent.change(formula, { target: { value: 'draft' } });
        }

        // the peer undoes adding the sheet the draft is open on.
        await fromPeer(() => {
          peer.undo();
        });
        expect(view.getAllByRole('tab')).toHaveLength(3);
        expect(view.queryByTestId('xlsx-cell-editor') === null).toBe(true);
        expect(formula.value).not.toBe('draft');
        // the keys still being typed go nowhere until a click: neither the
        // grid nor the formula bar has the focus to take them.
        // (identity checks stay cheap to print when they fail.)
        expect(document.activeElement?.getAttribute('data-testid') ?? document.activeElement?.tagName).toBe('BODY');
        await act(async () => {
          for (const key of ['x', 'Enter']) fireEvent.keyDown(document.activeElement!, { key });
        });
        await idle();
        expect(view.queryByTestId('xlsx-cell-editor') === null).toBe(true);
        for (let sheet = 0; sheet < (await editor.sheetInfo()).sheetNames.length; sheet += 1)
          expect((await editor.cell(sheet, 0, 0)).input).not.toBe('draft');
        for (let sheet = 0; sheet < peer.sheetInfo().sheetNames.length; sheet += 1)
          expect(peer.cell(sheet, 0, 0).input).not.toBe('draft');
        // one click on the grid gives it the keys again.
        fireEvent.mouseDown(surface, pointAt(plain, { row: 1, col: 1 }));
        fireEvent.mouseUp(surface, pointAt(plain, { row: 1, col: 1 }));
        await act(async () => {
          fireEvent.keyDown(document.activeElement!, { key: 'd' });
        });
        expect(view.queryByTestId('xlsx-cell-editor') === null).toBe(false);
      } finally {
        cleanup();
        peer.dispose();
      }
    });

  it('commits a draft to its own sheet when a peer inserts one before it', async () => {
    const peer = openWorkbook(plain.bytes.slice(), { collaborative: true, clientId: 3201 });
    let api: XlsxEditorApi | undefined;
    const view = render(
      <XlsxEditor
        file={plain.bytes.slice()}
        collaboration={{ clientId: 3202 }}
        onReady={(ready) => {
          api = ready;
        }}
      />
    );
    try {
      await waitFor(() => expect(api).toBeDefined());
      const editor = api!.handle;
      await idleWith(() => editor);
      editor.onUpdate((update, origin) => {
        if (origin === 'local') peer.applyUpdate(update);
      });
      const summary = view.getAllByRole('tab')[1];
      expect(summary.textContent).toBe('Summary');
      await act(async () => {
        fireEvent.click(summary);
      });
      await idleWith(() => editor);
      await act(async () => {
        fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'd' });
      });
      const input = view.getByTestId('xlsx-cell-editor') as HTMLInputElement;
      fireEvent.change(input, { target: { value: 'draft' } });

      // the peer's new first sheet reaches the worker, and a host flush
      // follows before any frame shows the shifted indexes.
      const before = await editor.encodeStateVector();
      peer.applyOps([{ type: 'addSheet', index: 0, name: 'First' }]);
      await act(async () => {
        void editor.applyUpdate(peer.encodeStateAsUpdate(before));
        await api!.flush();
      });
      const names = (await editor.sheetInfo()).sheetNames;
      const landed = [];
      for (const [sheet, name] of names.entries())
        if ((await editor.cell(sheet, 0, 0)).input === 'draft') landed.push(name);
      expect(landed).toEqual(['Summary']);
    } finally {
      cleanup();
      peer.dispose();
    }
  });

  it('keeps every key typed while a peer pastes a large range, the edit on its cell', async () => {
    const peer = openWorkbook(plain.bytes.slice(), { collaborative: true, clientId: 3301 });
    let api: XlsxEditorApi | undefined;
    const view = render(
      <XlsxEditor
        file={plain.bytes.slice()}
        collaboration={{ clientId: 3302 }}
        onReady={(ready) => {
          api = ready;
        }}
      />
    );
    try {
      await waitFor(() => expect(api).toBeDefined());
      const editor = api!.handle;
      const idle = () => idleWith(() => editor);
      await idle();
      // the peer pastes 4,000 cells, several times, as the user types.
      const pastes: Uint8Array[] = [];
      peer.onUpdate((update) => pastes.push(update));
      for (let paste = 0; paste < 3; paste += 1) {
        const edits = [];
        for (let row = 20; row < 420; row += 1)
          for (let col = 0; col < 10; col += 1)
            edits.push({ row, col, input: String(paste * 1000 + row + col) });
        peer.editCells(0, edits);
      }
      const surface = view.getByTestId('xlsx-scroll');
      const target = { row: 3, col: 1 };
      fireEvent.mouseDown(surface, pointAt(plain, target));
      fireEvent.mouseUp(surface, pointAt(plain, target));
      await idle();
      const typed: string[] = [];
      for (let entry = 0; entry < 3; entry += 1) {
        // each paste reaches the editor between two keys of an entry.
        void editor.applyUpdate(pastes[entry]);
        const text = `typed ${entry}`;
        await act(async () => {
          fireEvent.keyDown(surface, { key: text[0] });
        });
        const input = view.getByTestId('xlsx-cell-editor') as HTMLInputElement;
        for (let length = 2; length <= text.length; length += 1) {
          fireEvent.change(input, { target: { value: text.slice(0, length) } });
          // the open edit stays on its cell while the peer's frames land.
          expect(input.isConnected).toBe(true);
          expect(document.activeElement === input).toBe(true);
        }
        expect(input.value).toBe(text);
        await act(async () => {
          fireEvent.keyDown(input, { key: 'Enter' });
        });
        typed.push(text);
      }
      await idle();
      for (const [entry, text] of typed.entries())
        expect((await editor.cell(0, target.row + entry, target.col)).input).toBe(text);
      expect((await editor.cell(0, 419, 9)).input).toBe(String(2000 + 419 + 9));
      expect(view.getByTestId('xlsx-name-box')).toHaveProperty('value', 'B7');
    } finally {
      cleanup();
      peer.dispose();
    }
  }, 30_000);
});

describe('XlsxEditor host integration', () => {
  it('reports its command state a bounded number of times while the worker opens', async () => {
    // a host that keeps the command state, as Capy's does: a fresh value per
    // render would re-render the host, and so the editor, without end.
    let reports = 0;
    let painted = false;
    const file = plain.bytes.slice();
    function Host() {
      const [, setState] = useState<XlsxCommandState | null>(null);
      return (
        <XlsxEditor
          file={file}
          onCommandStateChange={(state) => {
            reports += 1;
            // past the bound the host stops re-rendering, so a loop ends.
            if (reports < 10) setState(state);
          }}
          onFirstPaint={() => {
            painted = true;
          }}
        />
      );
    }
    render(<Host />);
    // the loop shows before the worker answers: count until the first paint.
    await waitFor(() => expect(painted || reports >= 10).toBe(true));
    expect(reports).toBeLessThan(10);
  });

  it('reports the first painted grid once per opened workbook', async () => {
    let painted = 0;
    const onFirstPaint = () => {
      painted += 1;
    };
    const view = render(<XlsxEditor file={plain.bytes.slice()} onFirstPaint={onFirstPaint} />);
    await waitFor(() => expect(painted).toBe(1));
    await act(async () => {
      fireEvent.scroll(view.getByTestId('xlsx-scroll'));
    });
    expect(painted).toBe(1);
    view.rerender(<XlsxEditor file={wide.bytes.slice()} onFirstPaint={onFirstPaint} />);
    await waitFor(() => expect(painted).toBe(2));
  });

  it('keeps viewing mode navigable without exposing user mutations', async () => {
    let api: XlsxEditorApi | undefined;
    let changes = 0;
    const view = render(
      <XlsxEditor
        file={plain.bytes.slice()}
        onChange={() => changes++}
        onReady={(ready) => {
          api = ready;
        }}
        readOnly
      />
    );
    await waitFor(() => expect(api).toBeDefined());
    const idle = () => idleWith(() => api?.handle);
    await idle();
    const surface = view.getByTestId('xlsx-scroll');
    const target = { row: 2, col: 0 };
    const before = (await api!.handle.cell(0, target.row, target.col)).input;

    fireEvent.doubleClick(surface, pointAt(plain, target));
    fireEvent.keyDown(surface, { key: 'x' });
    fireEvent.keyDown(surface, { key: 'Delete' });
    await idle();

    // The toolbar and formula bar stay, disabled, so the grid keeps its place;
    // the zoom box and its list button and the name box stay usable, as zoom
    // and going to a cell edit nothing.
    const toolbar = view.getByTestId('xlsx-toolbar');
    const live = toolbar.querySelectorAll('button:not(:disabled), input:not(:disabled):not([readonly])');
    expect(Array.from(live, (node) => node.getAttribute('data-testid'))).toEqual([
      'xlsx-search-menus',
      'xlsx-zoom',
      null,
      'xlsx-toolbar-more',
      'xlsx-export-png',
      'xlsx-name-box',
    ]);
    const formula = view.getByTestId('xlsx-formula-input') as HTMLInputElement;
    expect([formula.disabled, formula.readOnly]).toEqual([false, true]);
    expect(view.queryByTestId('xlsx-cell-editor') === null).toBe(true);
    expect((await api!.handle.cell(0, target.row, target.col)).input).toBe(before);
    expect(changes).toBe(0);

    await act(async () => {
      expect(await api!.selectCells(0, selectionAt({ row: 3, col: 1 }))).toBe(true);
    });
    await waitFor(() => {
      const selected = view.getByRole('gridcell', { selected: true });
      expect(selected.textContent).toContain('B4');
    });
    await act(async () => api!.clearSelection());
    await waitFor(() =>
      expect(view.queryAllByRole('gridcell', { selected: true })).toHaveLength(0)
    );
    expect(await api!.selectCells(99, selectionAt({ row: 0, col: 0 }))).toBe(false);
  });

  it('notifies on applied edits and saves through the host API', async () => {
    let api: XlsxEditorApi | undefined;
    let changes = 0;
    const view = render(
      <XlsxEditor
        file={plain.bytes.slice()}
        onChange={() => changes++}
        onReady={(ready) => {
          api = ready;
        }}
      />
    );
    await waitFor(() => expect(api).toBeDefined());
    const idle = () => idleWith(() => api?.handle);
    await idle();
    const surface = view.getByTestId('xlsx-scroll');
    const target = { row: 2, col: 0 };

    await act(async () => {
      expect(await api!.selectCells(0, selectionAt(target))).toBe(true);
    });
    await idle();
    fireEvent.doubleClick(surface, pointAt(plain, target));
    const editor = await waitFor(() => view.getByTestId('xlsx-cell-editor'));
    fireEvent.change(editor, { target: { value: 'Host edit' } });
    fireEvent.keyDown(editor, { key: 'Enter' });
    await idle();
    expect(changes).toBe(1);

    await act(async () => {
      await api!.selectCells(0, selectionAt({ row: 3, col: 1 }));
      api!.clearSelection();
    });
    let saved!: Uint8Array;
    await act(async () => {
      saved = await api!.save();
    });
    expect(changes).toBe(1);

    const reopened = openWorkbook(saved);
    try {
      expect(reopened.cell(0, target.row, target.col).input).toBe('Host edit');
    } finally {
      reopened.dispose();
    }
  });

  it('repaints a change the host makes through its handle, and edits from it', async () => {
    const view = await mountEditor();
    const target = { row: 2, col: 0 };
    await view.click(target);
    const before = view.formula();
    await act(async () => {
      await view.workbook().editCell(0, target.row, target.col, 'From the host');
    });
    await view.idle();
    expect(before).not.toBe('From the host');
    expect(view.formula()).toBe('From the host');
    // F2 opens on the host's value, so Enter does not write the old one back.
    await act(async () => {
      fireEvent.keyDown(view.surface, { key: 'F2' });
    });
    await view.idle();
    expect(view.editor()!.value).toBe('From the host');
  });

  it('does not finish an asynchronous paste after entering viewing mode', async () => {
    const file = plain.bytes.slice();
    let api: XlsxEditorApi | undefined;
    let resolveClipboard!: (text: string) => void;
    const clipboardText = new Promise<string>((resolve) => {
      resolveClipboard = resolve;
    });
    const originalClipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: { readText: () => clipboardText },
    });

    try {
      const onReady = (ready: XlsxEditorApi) => {
        api = ready;
      };
      const view = render(<XlsxEditor file={file} onReady={onReady} />);
      await waitFor(() => expect(api).toBeDefined());
      const idle = () => idleWith(() => api?.handle);
      await idle();
      const target = { row: 2, col: 0 };
      const before = (await api!.handle.cell(0, target.row, target.col)).input;
      await act(async () => {
        await api!.selectCells(0, selectionAt(target));
      });
      await idle();

      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'v', ctrlKey: true });
      view.rerender(<XlsxEditor file={file} onReady={onReady} readOnly />);
      await act(async () => resolveClipboard('late paste'));
      await idle();

      expect((await api!.handle.cell(0, target.row, target.col)).input).toBe(before);
    } finally {
      if (originalClipboard) {
        Object.defineProperty(navigator, 'clipboard', originalClipboard);
      } else {
        Reflect.deleteProperty(navigator, 'clipboard');
      }
    }
  });
});

describe('XlsxEditor proposal review', () => {
  it('reviews a staged proposal through accept, undo, reject, stale warning, and force apply', async () => {
    let api: XlsxEditorApi | undefined;
    const view = render(
      <XlsxEditor file={plain.bytes.slice()} onReady={(ready) => { api = ready; }} />
    );
    await waitFor(() => expect(api).toBeDefined());
    const workbook = api!.handle;
    const idle = () => idleWith(() => workbook);
    await idle();
    const input = async () => (await workbook.cell(0, 6, 4)).input;
    const before = await input();
    const stage = async (value: string) => {
      await act(async () => {
        await workbook.propose('Audit agent', 'Review this change', [
          { sheet: 0, row: 6, col: 4, input: value },
        ]);
        api!.refreshProposals();
      });
      await idle();
    };

    await stage('12');
    expect(await input()).toBe(before);
    fireEvent.click(view.getByTestId('xlsx-proposals-button'));
    expect(view.getByTestId('xlsx-proposal').textContent).toContain('Audit agent');
    fireEvent.click(view.getByTestId('xlsx-proposal-accept'));
    await idle();
    expect(await input()).toBe('12');
    expect(await workbook.listProposals()).toHaveLength(0);
    fireEvent.click(view.getByTestId('xlsx-undo'));
    await idle();
    expect(await input()).toBe(before);

    await stage('24');
    fireEvent.click(view.getByTestId('xlsx-proposal-reject'));
    await idle();
    expect(await workbook.listProposals()).toHaveLength(0);
    expect(await input()).toBe(before);

    await stage('42');
    await act(async () => {
      await workbook.editCell(0, 6, 4, '99');
      api!.refreshProposals();
    });
    await idle();
    fireEvent.click(view.getByTestId('xlsx-proposal-accept'));
    await waitFor(() =>
      expect(view.getByTestId('xlsx-proposal-stale').textContent).toContain('E7')
    );
    expect(await input()).toBe('99');
    fireEvent.click(view.getByTestId('xlsx-proposal-force'));
    await idle();
    expect(await input()).toBe('42');
    expect(await workbook.listProposals()).toHaveLength(0);

    await act(async () => {
      await workbook.editCell(0, 1, 6, '10');
    });
    await stage('=G2*2');
    expect(view.getByTestId('xlsx-proposal-cell-new').textContent).toBe('20');
    await act(async () => {
      await workbook.editCell(0, 1, 6, '99');
      api!.refreshProposals();
    });
    await idle();
    fireEvent.click(view.getByTestId('xlsx-proposal-accept'));
    await waitFor(() =>
      expect(view.getByTestId('xlsx-proposal-cell-new').textContent).toBe('198')
    );
    expect(await input()).toBe('42');
    fireEvent.click(view.getByTestId('xlsx-proposal-accept'));
    await idle();
    expect(await input()).toBe('=G2*2');
    expect(await workbook.listProposals()).toHaveLength(0);
  });
});

describe('XlsxEditor pending host edits', () => {
  it('settles a pending chart move before selecting another sheet', async () => {
    const source = openWorkbook(charted.bytes);
    source.applyOps([{ type: 'addSheet', index: 1, name: 'Extra' }]);
    const file = source.save();
    source.dispose();
    let api: XlsxEditorApi | undefined;
    const view = render(
      <XlsxEditor
        file={file}
        onReady={(ready) => {
          api = ready;
        }}
      />
    );
    await waitFor(() => expect(api).toBeDefined());
    const idle = () => idleWith(() => api?.handle);
    await idle();
    const surface = view.getByTestId('xlsx-scroll');
    const chart = (await api!.handle.displayList({ x: 0, y: 0, ...VIEWPORT })).charts![0];
    fireEvent.mouseDown(surface, chartCenter(chart));
    fireEvent.mouseUp(window, chartCenter(chart));
    await act(async () => {
      fireEvent.keyDown(surface, { key: 'ArrowRight' });
    });
    expect(
      Math.round(parseFloat(view.getByTestId('xlsx-chart-selection').style.left))
    ).toBe(Math.round(chart.rect.x + 1));
    await act(async () => {
      expect(await api!.selectCells(1, selectionAt({ row: 0, col: 0 }))).toBe(true);
    });
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 350));
    });
    await act(async () => {
      await api!.selectCells(0, selectionAt({ row: 0, col: 0 }));
    });
    const after = (await api!.handle.displayList({ x: 0, y: 0, ...VIEWPORT })).charts!.find(
      (c) => c.id === chart.id
    )!;
    expect(after.rect.x).toBe(chart.rect.x + 1);
  });

  it('commits the current cell draft before selecting another cell', async () => {
    let api: XlsxEditorApi | undefined;
    const view = render(
      <XlsxEditor
        file={plain.bytes.slice()}
        onReady={(ready) => {
          api = ready;
        }}
      />
    );
    await waitFor(() => expect(api).toBeDefined());
    const idle = () => idleWith(() => api?.handle);
    await idle();
    const target = { row: 2, col: 0 };
    await act(async () => {
      await api!.selectCells(0, selectionAt(target));
    });
    await idle();
    fireEvent.doubleClick(view.getByTestId('xlsx-scroll'), pointAt(plain, target));
    const editor = await waitFor(() => view.getByTestId('xlsx-cell-editor'));
    fireEvent.change(editor, { target: { value: 'Draft that must survive' } });
    await act(async () => {
      await api!.selectCells(0, selectionAt({ row: 3, col: 1 }));
    });
    expect((await api!.handle.cell(0, target.row, target.col)).input).toBe(
      'Draft that must survive'
    );
  });
  for (const source of ['cell', 'formula'] as const) {
    for (const action of ['save', 'clear', 'select'] as const) {
      it(`commits a ${source} draft before host ${action}`, async () => {
        let api: XlsxEditorApi | undefined;
        let changes = 0;
        const view = render(
          <XlsxEditor
            file={plain.bytes.slice()}
            onChange={() => {
              changes += 1;
            }}
            onReady={(ready) => {
              api = ready;
            }}
          />
        );
        await waitFor(() => expect(api).toBeDefined());
        const idle = () => idleWith(() => api?.handle);
        await idle();
        const target = { row: 2, col: 0 };
        await act(async () => {
          await api!.selectCells(0, selectionAt(target));
        });
        await idle();
        if (source === 'cell') {
          fireEvent.doubleClick(view.getByTestId('xlsx-scroll'), pointAt(plain, target));
        }
        const input = view.getByTestId(
          source === 'cell' ? 'xlsx-cell-editor' : 'xlsx-formula-input'
        );
        fireEvent.change(input, { target: { value: 'Saved draft' } });
        expect(await api!.selectCells(-1, selectionAt(target))).toBe(false);
        expect((await api!.handle.cell(0, target.row, target.col)).input).toBe('Line item 1');
        let saved: Uint8Array | undefined;
        await act(async () => {
          if (action === 'save') saved = await api!.save();
          else if (action === 'clear') api!.clearSelection();
          else await api!.selectCells(0, selectionAt({ row: 3, col: 1 }));
        });
        await idle();
        expect((await api!.handle.cell(0, target.row, target.col)).input).toBe('Saved draft');
        expect(changes).toBe(1);
        if (saved) {
          const reopened = openWorkbook(saved);
          try {
            expect(reopened.cell(0, target.row, target.col).input).toBe('Saved draft');
          } finally {
            reopened.dispose();
          }
        }
      });
    }
  }
  it('commits the next cell edit on blur after a host save', async () => {
    let api: XlsxEditorApi | undefined;
    const view = render(
      <XlsxEditor
        file={plain.bytes.slice()}
        onReady={(ready) => {
          api = ready;
        }}
      />
    );
    await waitFor(() => expect(api).toBeDefined());
    const idle = () => idleWith(() => api?.handle);
    await idle();
    const target = { row: 2, col: 0 };
    await act(async () => {
      await api!.selectCells(0, selectionAt(target));
    });
    await idle();
    fireEvent.change(view.getByTestId('xlsx-formula-input'), {
      target: { value: 'First draft' },
    });
    await act(async () => {
      await api!.save();
    });
    await idle();
    fireEvent.doubleClick(view.getByTestId('xlsx-scroll'), pointAt(plain, target));
    const editor = await waitFor(() => view.getByTestId('xlsx-cell-editor'));
    fireEvent.change(editor, { target: { value: 'Second draft' } });
    fireEvent.blur(editor);
    expect((await api!.handle.cell(0, target.row, target.col)).input).toBe('Second draft');
  });
});

describe('XlsxEditor menu commands', () => {
  // Budget: A1 title, A2 "Item", A3..A6 "Line item 1..4"; B2..D2 Q1, Q2, Total.
  async function mountCommands(props: XlsxEditorProps = {}) {
    let api: XlsxEditorApi | undefined;
    let state: XlsxCommandState | undefined;
    const view = render(
      <XlsxEditor
        file={plain.bytes.slice()}
        onCommandStateChange={(next) => {
          state = next;
        }}
        onReady={(ready) => {
          api = ready;
        }}
        {...props}
      />
    );
    await waitFor(() => expect(api).toBeDefined());
    const idle = () => idleWith(() => api?.handle);
    await idle();
    const select = async (anchor: CellAddr, focus: CellAddr = anchor) => {
      await act(async () => {
        await api!.selectCells((await api!.handle.sheetInfo()).activeSheet, { anchor, focus });
      });
      await idle();
    };
    const run = async (command: XlsxCommand) => {
      await act(async () => {
        api!.run(command);
      });
      await idle();
    };
    const column = (col: number, rows = 8) =>
      Promise.all(
        Array.from({ length: rows }, async (_, row) => (await api!.handle.cell(0, row, col)).input)
      );
    const row = (at: number, cols = 4) =>
      Promise.all(
        Array.from({ length: cols }, async (_, col) => (await api!.handle.cell(0, at, col)).input)
      );
    // Types into the toolbar's zoom box and presses Enter.
    const typeZoom = async (value: string) => {
      await act(async () => {
        const box = view.getByTestId('xlsx-zoom');
        fireEvent.focus(box);
        fireEvent.input(box, { target: { value } });
        fireEvent.keyDown(box, { key: 'Enter' });
      });
      await idle();
    };
    return { api: () => api!, state: () => state!, select, run, column, row, typeZoom, view };
  }

  it('inserts as many rows as are selected above the selection, as one undo step', async () => {
    const view = await mountCommands();
    const before = await view.column(0);
    await view.select({ row: 2, col: 0 }, { row: 3, col: 1 });
    await view.run('insertRowAbove');
    expect(await view.column(0, 10)).toEqual([...before.slice(0, 2), '', '', ...before.slice(2)]);
    await view.run('undo');
    expect(await view.column(0)).toEqual(before);
  });

  it('inserts rows below the selection', async () => {
    const view = await mountCommands();
    const before = await view.column(0);
    await view.select({ row: 2, col: 0 });
    await view.run('insertRowBelow');
    expect(await view.column(0, 9)).toEqual([...before.slice(0, 3), '', ...before.slice(3)]);
  });

  it('inserts columns left and right of the selection', async () => {
    const view = await mountCommands();
    const header = await view.row(1);
    await view.select({ row: 1, col: 1 });
    await view.run('insertColumnLeft');
    expect(await view.row(1, 5)).toEqual([header[0], '', ...header.slice(1)]);
    await view.select({ row: 1, col: 2 });
    await view.run('insertColumnRight');
    expect(await view.row(1, 6)).toEqual([header[0], '', header[1], '', ...header.slice(2)]);
  });

  it('deletes the selected rows and columns', async () => {
    const view = await mountCommands();
    const before = await view.column(0);
    await view.select({ row: 2, col: 0 }, { row: 3, col: 0 });
    await view.run('deleteRows');
    expect(await view.column(0, 4)).toEqual([...before.slice(0, 2), ...before.slice(4, 6)]);
    const header = await view.row(1);
    await view.select({ row: 1, col: 1 });
    await view.run('deleteColumns');
    expect(await view.row(1, 3)).toEqual([header[0], ...header.slice(2)]);
  });

  it('deletes the values of the selected cells', async () => {
    const view = await mountCommands();
    await view.select({ row: 2, col: 0 }, { row: 2, col: 2 });
    await view.run('deleteValues');
    expect(await view.row(2, 3)).toEqual(['', '', '']);
    expect(await view.row(3, 1)).toEqual(['Line item 2']);
  });

  it('freezes rows and columns, by count or up to the selection, independently', async () => {
    const view = await mountCommands();
    const frozen = async () => {
      const info = await view.api().handle.sheetInfo();
      return [info.frozenRows, info.frozenCols];
    };
    await view.select({ row: 3, col: 2 });
    await view.run('freezeRows:1');
    expect(await frozen()).toEqual([1, 0]);
    await view.run('freezeColumns:current');
    expect(await frozen()).toEqual([1, 3]);
    await view.run('freezeRows:current');
    expect(await frozen()).toEqual([4, 3]);
    expect([view.state().frozenRows, view.state().frozenColumns]).toEqual([4, 3]);
    await view.run('freezeRows:0');
    await view.run('freezeColumns:2');
    expect(await frozen()).toEqual([0, 2]);
    await view.run('freezeColumns:0');
    expect(await frozen()).toEqual([0, 0]);
  });

  it('adds a sheet after the active one and opens it', async () => {
    const view = await mountCommands();
    await view.run('insertSheet');
    const info = await view.api().handle.sheetInfo();
    expect(info.sheetNames).toEqual(['Budget', 'Sheet4', 'Summary', 'Styled']);
    expect(info.activeSheet).toBe(1);
    expect(view.view.getAllByRole('tab')[1].getAttribute('aria-selected')).toBe('true');
  });

  it('clears every style and the number format of the selection', async () => {
    const view = await mountCommands();
    await view.select({ row: 2, col: 1 }, { row: 3, col: 2 });
    await view.run('bold');
    await view.run('numberFormat:percent');
    await view.run('align:center');
    let formatting = await view.api().handle.selectionFormatting(0, 'B3:C4');
    expect([formatting.bold, formatting.numberFormat, formatting.horizontalAlignment]).toEqual([
      true,
      'percent',
      'center',
    ]);
    await view.run('clearFormatting');
    formatting = await view.api().handle.selectionFormatting(0, 'B3:C4');
    expect(formatting.bold).toBe(false);
    expect(formatting.numberFormat).toBe('automatic');
    expect(formatting.horizontalAlignment).not.toBe('center');
    expect(view.state().canUndo).toBe(true);
  });

  it('reports the visible part of the sheet in sheet pixels, at any zoom', async () => {
    const view = await mountCommands();
    expect(view.api().visibleViewport()).toEqual({ x: 0, y: 0, ...VIEWPORT });
    await view.run('zoom:200');
    expect(view.api().visibleViewport()).toEqual({
      x: 0,
      y: 0,
      width: VIEWPORT.width / 2,
      height: VIEWPORT.height / 2,
    });
  });

  it('opens at the zoom it is given, and its zoom box takes a typed zoom as Google Sheets does', async () => {
    const view = await mountCommands({ initialZoom: 1.5 });
    const box = view.view.getByTestId('xlsx-zoom') as HTMLInputElement;
    expect([view.state().zoom, box.value]).toEqual([1.5, '150%']);
    // Sheets takes 50–200% in whole percents: 300 is 200%, 20 is 50%.
    for (const [typed, zoom] of [
      ['300', 2],
      ['20', 0.5],
      ['120.6', 1.21],
      ['abc', 1.21],
    ] as const) {
      await view.typeZoom(typed);
      expect([view.state().zoom, box.value], typed).toEqual([zoom, `${zoom * 100}%`]);
    }
  });

  it('keeps the zoom box usable while read-only, as zoom edits nothing', async () => {
    const view = await mountCommands({ readOnly: true });
    expect((view.view.getByTestId('xlsx-zoom') as HTMLInputElement).disabled).toBe(false);
    await view.typeZoom('150');
    expect(view.state().zoom).toBe(1.5);
  });

  it('runs the toolbar commands a menu repeats: merge, wrapping and zoom', async () => {
    const view = await mountCommands();
    await view.select({ row: 2, col: 1 }, { row: 3, col: 2 });
    expect(view.state().canMerge).toBe(true);
    await view.run('merge:all');
    expect(await view.api().handle.mergedRanges(0, 'B3:C4')).toHaveLength(1);
    expect(view.state().canUnmerge).toBe(true);
    await view.run('merge:unmerge');
    expect(await view.api().handle.mergedRanges(0, 'B3:C4')).toHaveLength(0);
    await view.run('wrap:wrap');
    expect((await view.api().handle.selectionFormatting(0, 'B3')).textWrapping).toBe('wrap');
    await view.run('zoom:150');
    expect(view.state().zoom).toBe(1.5);
  });

  it('leaves a merge whose unmerge read answers after sending started', async () => {
    const view = await mountEditor();
    await act(async () => {
      await view.api().selectCells(0, { anchor: { row: 2, col: 1 }, focus: { row: 3, col: 2 } });
    });
    await view.run('merge:all');
    const proxy = view.workbook();
    const forward = proxy.onSheet.bind(proxy);
    let open!: () => void;
    const held = new Promise<void>((resolve) => {
      open = resolve;
    });
    const spy = spyOn(proxy, 'onSheet').mockImplementation(((sheetId, method, ...args) =>
      method === 'mergedRanges'
        ? held.then(() => forward(sheetId, method, ...args))
        : forward(sheetId, method, ...args)) as WorkbookProxy['onSheet']);
    try {
      await view.run('merge:unmerge');
      await view.setProps({ sending: true });
      await act(async () => open());
      await view.idle();
    } finally {
      spy.mockRestore();
    }
    expect(await view.workbook().mergedRanges(0, 'B3:C4')).toHaveLength(1);
  });
});

describe('XlsxEditor active cell', () => {
  // a press, a pointer move of `dy` CSS px and the release: a jittery click or a short drag.
  const pressAndMove = async (
    view: Awaited<ReturnType<typeof mountEditor>>,
    at: { clientX: number; clientY: number },
    dy: number
  ) => {
    await act(async () => {
      fireEvent.mouseDown(view.surface, at);
      fireEvent.mouseMove(view.surface, { ...at, clientY: at.clientY + dy, buttons: 1 });
      fireEvent.mouseUp(window, { ...at, clientY: at.clientY + dy });
    });
    await view.idle();
  };
  // Budget cut to A1:D5: row 8 lies past the used range.
  let trimmed: Fixture;
  beforeAll(() => {
    const handle = openWorkbook(plain.bytes.slice());
    try {
      handle.applyOps([{ type: 'deleteRows', sheet: 0, at: 5, count: 58 }]);
      trimmed = fixtureFrom(handle.save());
    } finally {
      handle.dispose();
    }
  });

  it('keeps the pressed cell active when a press past the used range moves a little', async () => {
    const view = await mountEditor(trimmed);
    await pressAndMove(view, pointAt(trimmed, { row: 7, col: 0 }), 2);
    expect(view.nameBox().value).toBe('A8');
  });

  // Excel and Sheets keep the pressed cell active while a drag grows the range.
  it('keeps the pressed cell active when the pointer slips onto the cell above', async () => {
    const view = await mountEditor();
    const cell = cellRect(plain.grid, 5, 1)!;
    await pressAndMove(view, { clientX: cell.x + cell.w / 2, clientY: cell.y + 1 }, -2);
    expect(view.nameBox().value).toBe('B6');
    await act(async () => {
      fireEvent.keyDown(view.surface, { key: 'x' });
    });
    await act(async () => {
      fireEvent.keyDown(view.editor()!, { key: 'Enter' });
    });
    expect([await view.input(5, 1), await view.input(4, 1)]).toEqual(['x', '300']);
  });

  it('moves only the far end with Shift+arrows, and steps from the active cell after', async () => {
    const view = await mountEditor();
    await view.click({ row: 2, col: 1 });
    for (const key of ['ArrowDown', 'ArrowDown', 'ArrowRight'])
      await act(async () => {
        fireEvent.keyDown(view.surface, { key, shiftKey: true });
      });
    await view.idle();
    expect(view.nameBox().value).toBe('B3');
    await act(async () => {
      fireEvent.keyDown(view.surface, { key: 'ArrowDown' });
    });
    await view.idle();
    expect(view.nameBox().value).toBe('B4');
  });

  // Excel and Sheets keep the active cell where it was.
  it('keeps the active cell where it was on select all, so typing goes there', async () => {
    const view = await mountEditor();
    await view.click({ row: 4, col: 2 });
    await act(async () => {
      fireEvent.keyDown(view.surface, { key: 'a', ctrlKey: true });
    });
    await view.idle();
    expect(view.nameBox().value).toBe('C5');
    await act(async () => {
      fireEvent.keyDown(view.surface, { key: 'x' });
    });
    await act(async () => {
      fireEvent.keyDown(view.editor()!, { key: 'Enter' });
    });
    expect([await view.input(4, 2), await view.input(0, 0)]).toEqual(['x', 'Quarterly Budget Report']);
  });

  // Excel: Ctrl+A at C5, then Shift+click E10 selects C5:E10.
  it('extends a Shift+click after select all from the active cell', async () => {
    const view = await mountEditor();
    await view.click({ row: 4, col: 2 });
    await act(async () => {
      fireEvent.keyDown(view.surface, { key: 'a', ctrlKey: true });
    });
    await view.idle();
    await act(async () => {
      fireEvent.mouseDown(view.surface, { ...pointAt(plain, { row: 9, col: 4 }), shiftKey: true });
      fireEvent.mouseUp(window, pointAt(plain, { row: 9, col: 4 }));
    });
    await view.idle();
    const from = cellRect(plain.grid, 4, 2)!;
    const to = cellRect(plain.grid, 9, 4)!;
    const box = view.selectionBox()!.style;
    expect([
      view.nameBox().value,
      ...[box.left, box.top, box.width, box.height].map((value) => Math.round(parseFloat(value))),
    ]).toEqual(['C5', ...[from.x, from.y, to.x + to.w - from.x, to.y + to.h - from.y].map(Math.round)]);
  });

  // Budget's title is merged over A1:D1.
  it('selects a merged cell by its top-left cell, so typing replaces what it shows', async () => {
    const view = await mountEditor();
    await view.click({ row: 0, col: 2 });
    expect([view.nameBox().value, view.formula()]).toEqual(['A1', 'Quarterly Budget Report']);
    await act(async () => {
      fireEvent.keyDown(view.surface, { key: 'x' });
    });
    await act(async () => {
      fireEvent.keyDown(view.editor()!, { key: 'Enter' });
    });
    expect([await view.input(0, 0), await view.input(0, 2)]).toEqual(['x', '']);
  });
});

describe('XlsxEditor formula entry', () => {
  const press = async (target: Element, key: string, init: KeyboardEventInit = {}) => {
    await act(async () => {
      fireEvent.keyDown(target, { key, ...init });
    });
    await idleWith(() => currentWorkbook);
  };

  // Excel and Sheets close the parentheses a typed formula leaves open.
  it('closes the parentheses a formula typed in the cell or the formula bar leaves open', async () => {
    const view = await mountEditor();
    await view.click({ row: 10, col: 1 });
    await press(view.surface, '=');
    view.type('=SUM(B3:B5');
    await press(view.editor()!, 'Enter');
    expect(await view.input(10, 1)).toBe('=SUM(B3:B5)');
    view.typeFormula('=ROUND(SUM(B3:B5');
    await press(view.formulaInput(), 'Enter');
    expect(await view.input(11, 1)).toBe('=ROUND(SUM(B3:B5))');
    // a call whose function takes no arguments is complete once closed.
    await press(view.surface, '=');
    view.type('=TODAY(');
    await press(view.editor()!, 'Enter');
    expect(await view.input(12, 1)).toBe('=TODAY()');
    // an unfinished argument list is repaired on entry by closing it.
    await press(view.surface, '=');
    view.type('=SUM(B3,');
    await press(view.editor()!, 'Enter');
    expect(await view.input(13, 1)).toBe('=SUM(B3,)');
  });

  const refusal = (view: Awaited<ReturnType<typeof mountEditor>>) =>
    view.surface.ownerDocument.querySelector('[data-testid="xlsx-edit-refusal"]')?.textContent ?? null;

  it('refuses a formula that does not parse with a line under the formula bar, the edit never closing', async () => {
    const view = await mountEditor();
    await view.click({ row: 10, col: 1 });
    await press(view.surface, '=');
    view.type('=B3*/2');
    const input = view.editor();
    await press(input!, 'Enter');
    // the same input, never unmounted, still holds the draft on its cell.
    expect(view.editor() === input).toBe(true);
    expect([view.editor()?.value, view.nameBox().value]).toEqual(['=B3*/2', 'B11']);
    expect(refusal(view)).toBe("There's a problem with this formula. Check its parentheses, quotes and operators.");
    expect(view.error() === null).toBe(true);
    expect(await view.input(10, 1)).toBe('900');
    view.type('=B3*2');
    expect(refusal(view)).toBe(null);
    await press(view.editor()!, 'Enter');
    expect(await view.input(10, 1)).toBe('=B3*2');

    view.typeFormula('=B3+');
    await press(view.formulaInput(), 'Enter');
    expect([view.formula(), view.nameBox().value]).toEqual(['=B3+', 'B12']);
    expect(refusal(view) === null).toBe(false);
    expect(await view.input(11, 1)).toBe('1000');
  });

  it('shows the refusal of text past 32,767 characters on the same line, translated', async () => {
    const view = await mountEditor(plain, undefined, undefined, { i18n: zhCN });
    await view.click({ row: 10, col: 1 });
    await press(view.surface, 'x');
    view.type('x'.repeat(40_000));
    await press(view.editor()!, 'Enter');
    expect([view.editor()?.value.length, view.nameBox().value]).toEqual([40_000, 'B11']);
    expect(refusal(view)).toBe('一个单元格最多可容纳 32,767 个字符。');
    expect(view.error() === null).toBe(true);
    expect(await view.input(10, 1)).toBe('900');
  });

  it('shows a pasted text past 32,767 characters on the same line', async () => {
    const view = await mountEditor();
    const clipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: { readText: async () => 'x'.repeat(40_000) },
    });
    try {
      await view.click({ row: 10, col: 1 });
      await press(view.surface, 'v', { ctrlKey: true });
      await view.idle();
      expect(refusal(view)).toBe('A cell holds at most 32,767 characters.');
      expect(view.error() === null).toBe(true);
      expect(await view.input(10, 1)).toBe('900');
    } finally {
      if (clipboard) Object.defineProperty(navigator, 'clipboard', clipboard);
      else Reflect.deleteProperty(navigator, 'clipboard');
    }
  });

  // holds every cell edit back until released, as a busy engine would.
  const holdEdits = (view: Awaited<ReturnType<typeof mountEditor>>) => {
    const proxy = view.workbook();
    const forward = proxy.onSheet.bind(proxy);
    let open = () => {};
    const gate = new Promise<void>((resolve) => {
      open = resolve;
    });
    const spy = spyOn(proxy, 'onSheet').mockImplementation(((sheetId, method, ...args) =>
      method === 'editCell'
        ? gate.then(() => forward(sheetId, method, ...args))
        : forward(sheetId, method, ...args)) as WorkbookProxy['onSheet']);
    return {
      release: async () => {
        await act(async () => open());
        await view.idle();
      },
      edits: () => spy.mock.calls.filter(([, method]) => method === 'editCell').length,
      restore: () => spy.mockRestore(),
    };
  };

  // a formula typed into B11 (900; B12 holds 1000) in an edit opened by typing.
  const typeInB11 = async (view: Awaited<ReturnType<typeof mountEditor>>, value: string) => {
    await view.click({ row: 10, col: 1 });
    await press(view.surface, '=');
    view.type(value);
  };

  it("waits for the engine's answer before a formula's edit closes", async () => {
    const view = await mountEditor();
    const held = holdEdits(view);
    try {
      await typeInB11(view, '=B3*2');
      await press(view.editor()!, 'Enter');
      expect([view.editor()?.value, view.editor()?.readOnly, view.nameBox().value]).toEqual([
        '=B3*2',
        true,
        'B11',
      ]);
      await held.release();
      expect([view.editor() === null, view.nameBox().value]).toEqual([true, 'B12']);
      expect(await view.input(10, 1)).toBe('=B3*2');
    } finally {
      held.restore();
    }
  });

  // after Enter the user is not typing into that formula (Epo, 2026-10-19).
  it("drops the keys typed while a formula's commit waits for the engine", async () => {
    const view = await mountEditor();
    const held = holdEdits(view);
    try {
      await typeInB11(view, '=B3*2');
      await press(view.editor()!, 'Enter');
      await press(view.editor()!, '5');
      await press(view.editor()!, 'Enter');
      expect([view.editor()?.readOnly, view.editor()?.value]).toEqual([true, '=B3*2']);
      await held.release();
      expect([view.editor() === null, view.nameBox().value, held.edits()]).toEqual([true, 'B12', 1]);
    } finally {
      held.restore();
    }
    expect([await view.input(10, 1), await view.input(11, 1)]).toEqual(['=B3*2', '1000']);
  });

  it('selects a cell pressed while a formula commit waits once the engine took it', async () => {
    const view = await mountEditor();
    const held = holdEdits(view);
    try {
      await typeInB11(view, '=B3*2');
      await view.click({ row: 4, col: 3 });
      expect([view.nameBox().value, view.editor()?.readOnly]).toEqual(['B11', true]);
      await held.release();
      expect([view.nameBox().value, view.editor() === null]).toEqual(['D5', true]);
    } finally {
      held.restore();
    }
    expect(await view.input(10, 1)).toBe('=B3*2');
  });

  it('selects a cell pressed after Enter while the commit waits once it landed', async () => {
    const view = await mountEditor();
    const held = holdEdits(view);
    try {
      await typeInB11(view, '=B3*2');
      await press(view.editor()!, 'Enter');
      await view.click({ row: 4, col: 3 });
      expect(view.nameBox().value).toBe('B11');
      await held.release();
      expect([view.nameBox().value, view.editor() === null]).toEqual(['D5', true]);
    } finally {
      held.restore();
    }
  });

  it('switches to a sheet tab pressed while a formula commit waits once the engine took it', async () => {
    const view = await mountEditor();
    const held = holdEdits(view);
    const selected = () =>
      [...view.surface.ownerDocument.querySelectorAll('[role="tab"]')].map(
        (tab) => tab.getAttribute('aria-selected') === 'true'
      );
    try {
      await typeInB11(view, '=B3*2');
      await press(view.editor()!, 'Enter');
      await act(async () => {
        fireEvent.click(view.surface.ownerDocument.querySelectorAll('[role="tab"]')[1]);
      });
      await view.idle();
      expect(selected().slice(0, 2)).toEqual([true, false]);
      await held.release();
      expect(selected().slice(0, 2)).toEqual([false, true]);
    } finally {
      held.restore();
    }
    expect(await view.input(10, 1)).toBe('=B3*2');
  });

  it('keeps a refused cell draft on its cell when a cell was pressed during the wait', async () => {
    const view = await mountEditor();
    const held = holdEdits(view);
    try {
      await typeInB11(view, '=B3*/2');
      await press(view.editor()!, 'Enter');
      await view.click({ row: 4, col: 3 });
      await held.release();
    } finally {
      held.restore();
    }
    expect([view.editor()?.value, view.nameBox().value, refusal(view) === null]).toEqual([
      '=B3*/2',
      'B11',
      false,
    ]);
  });

  it('stays on the sheet when a draft refused after a sheet tab press, committing it once', async () => {
    const view = await mountEditor();
    const held = holdEdits(view);
    const selected = () =>
      [...view.surface.ownerDocument.querySelectorAll('[role="tab"]')].map(
        (tab) => tab.getAttribute('aria-selected') === 'true'
      );
    try {
      await typeInB11(view, '=B3*/2');
      await press(view.editor()!, 'Enter');
      await act(async () => {
        fireEvent.click(view.surface.ownerDocument.querySelectorAll('[role="tab"]')[1]);
      });
      await held.release();
      expect([selected().slice(0, 2), view.editor()?.value, held.edits()]).toEqual([
        [true, false],
        '=B3*/2',
        1,
      ]);
    } finally {
      held.restore();
    }
  });

  it('inserts a row asked for while a formula commit waits once the engine took it', async () => {
    const view = await mountEditor();
    const held = holdEdits(view);
    try {
      await typeInB11(view, '=B3*2');
      await view.run('insertRowAbove');
      expect((await view.workbook().cell(0, 10, 1)).input).toBe('900');
      await held.release();
    } finally {
      held.restore();
    }
    expect([await view.input(10, 1), await view.input(11, 1)]).toEqual(['', '=B3*2']);
  });

  it('refuses a row insert and a host clear that waited on a commit once sending started', async () => {
    const view = await mountEditor();
    const held = holdEdits(view);
    try {
      await typeInB11(view, '=B3*2');
      await view.run('insertRowAbove');
      await act(async () => view.api().clearSelection());
      await view.setProps({ sending: true });
      await held.release();
    } finally {
      held.restore();
    }
    expect([view.nameBox().value, await view.input(10, 1), await view.input(11, 1)]).toEqual([
      'B11',
      '=B3*2',
      '1000',
    ]);
  });

  // the press commits the formula-bar draft first: a refusal keeps it on its cell.
  it('keeps a refused formula-bar draft on its cell when another cell is pressed', async () => {
    const view = await mountEditor();
    const d5 = await view.input(4, 3);
    await view.click({ row: 10, col: 1 });
    await act(async () => view.formulaInput().focus());
    view.typeFormula('=B3*/2');
    await press(view.formulaInput(), 'Enter');
    await view.click({ row: 4, col: 3 });
    expect([view.formula(), view.nameBox().value, refusal(view) === null]).toEqual([
      '=B3*/2',
      'B11',
      false,
    ]);
    view.typeFormula('=B3*2');
    await press(view.formulaInput(), 'Enter');
    expect([await view.input(10, 1), await view.input(4, 3)]).toEqual(['=B3*2', d5]);
  });

  // a refused draft stays open after its blur (a toolbar action gives the grid
  // the focus): a key in the grid gives the focus back to the bar, moving nothing.
  it('gives the formula bar back the keys while its refused draft is open', async () => {
    const view = await mountEditor();
    await view.click({ row: 10, col: 1 });
    await act(async () => view.formulaInput().focus());
    view.typeFormula('=B3*/2');
    await press(view.formulaInput(), 'Enter');
    await act(async () => view.surface.focus());
    await press(view.surface, 'ArrowDown');
    expect([view.nameBox().value, view.formula(), document.activeElement === view.formulaInput()]).toEqual([
      'B11',
      '=B3*/2',
      true,
    ]);
    view.typeFormula('=B3*2');
    await press(view.formulaInput(), 'Enter');
    expect([await view.input(10, 1), await view.input(11, 1)]).toEqual(['=B3*2', '1000']);
  });

  // J21 lies outside A1:D62, so the menu's Select all makes A1 the active cell.
  it('commits a refused formula-bar draft to its own cell after the menu selected all', async () => {
    const view = await mountEditor();
    await act(async () => {
      await view.api().selectCells(0, selectionAt({ row: 20, col: 9 }));
    });
    await view.idle();
    await act(async () => view.formulaInput().focus());
    view.typeFormula('=B3*/2');
    await press(view.formulaInput(), 'Enter');
    await view.run('selectAll');
    expect([view.nameBox().value, view.formula()]).toEqual(['A1', '=B3*/2']);
    await act(async () => view.formulaInput().focus());
    view.typeFormula('=B3*2');
    await press(view.formulaInput(), 'Enter');
    expect([await view.input(20, 9), await view.input(0, 0)]).toEqual([
      '=B3*2',
      'Quarterly Budget Report',
    ]);
  });

  // Excel's Edit mode after a refusal: the arrows fix the formula, not commit it again.
  it('moves the caret with the arrows after a refusal in an edit opened by typing', async () => {
    const view = await mountEditor();
    await typeInB11(view, '=B3*/2');
    await press(view.editor()!, 'Enter');
    const edits = spyOn(view.workbook(), 'onSheet');
    try {
      await press(view.editor()!, 'ArrowLeft');
      await press(view.editor()!, 'ArrowRight');
      expect([view.editor()?.value, view.nameBox().value]).toEqual(['=B3*/2', 'B11']);
      expect(edits.mock.calls.filter(([, method]) => method === 'editCell').length).toBe(0);
    } finally {
      edits.mockRestore();
    }
  });

  it('leaves the focus where it went while a commit waited, the next edit committing on leave', async () => {
    const view = await mountEditor();
    const held = holdEdits(view);
    try {
      await typeInB11(view, '=B3*2');
      await press(view.editor()!, 'Enter');
      await act(async () => view.nameBox().focus());
      await held.release();
      expect([view.editor() === null, document.activeElement === view.nameBox()]).toEqual([true, true]);
    } finally {
      held.restore();
    }
    await view.click({ row: 12, col: 1 });
    await press(view.surface, 'x');
    await act(async () => view.formulaInput().focus());
    expect(await view.input(12, 1)).toBe('x');
  });

  // the user moved on to the host's page (a chat box) while the engine answered.
  it('leaves the focus outside the editor when a refusal lands after it moved there', async () => {
    const view = await mountEditor();
    const outside = document.createElement('input');
    document.body.append(outside);
    const held = holdEdits(view);
    try {
      await typeInB11(view, '=B3*/2');
      await press(view.editor()!, 'Enter');
      await act(async () => outside.focus());
      await held.release();
      expect([view.editor()?.value, refusal(view) === null, document.activeElement === outside]).toEqual([
        '=B3*/2',
        false,
        true,
      ]);
    } finally {
      held.restore();
      outside.remove();
    }
  });

  it('goes where the name box names once a waiting commit landed, and nowhere after a refusal', async () => {
    const view = await mountEditor();
    const goTo = async (address: string) => {
      await act(async () => view.nameBox().focus());
      fireEvent.change(view.nameBox(), { target: { value: address } });
      await press(view.nameBox(), 'Enter');
    };
    let held = holdEdits(view);
    try {
      await typeInB11(view, '=B3*/2');
      await goTo('D5');
      await held.release();
    } finally {
      held.restore();
    }
    expect([view.editor()?.value, view.nameBox().value, refusal(view) === null]).toEqual([
      '=B3*/2',
      'B11',
      false,
    ]);
    view.type('=B3*2');
    held = holdEdits(view);
    try {
      await goTo('D5');
      await held.release();
    } finally {
      held.restore();
    }
    expect([view.nameBox().value, await view.input(10, 1)]).toEqual(['D5', '=B3*2']);
  });
});

describe('XlsxEditor point mode', () => {
  const press = async (target: Element, key: string, init: KeyboardEventInit = {}) => {
    await act(async () => {
      fireEvent.keyDown(target, { key, ...init });
    });
    await idleWith(() => currentWorkbook);
  };

  it('puts a clicked cell into a formula waiting for a reference instead of committing it', async () => {
    const view = await mountEditor();
    await view.click({ row: 10, col: 1 });
    await press(view.surface, '=');
    expect(view.editor()?.value).toBe('=');
    await view.click({ row: 2, col: 1 });
    expect(view.editor()?.value).toBe('=B3');
    expect(view.nameBox().value).toBe('B11');
    // pointing again replaces the reference; an operator ends it.
    await view.click({ row: 3, col: 1 });
    expect(view.editor()?.value).toBe('=B4');
    view.type('=B4+');
    await view.click({ row: 2, col: 2 });
    expect(view.editor()?.value).toBe('=B4+C3');
    await press(view.editor()!, 'Enter');
    expect(await view.input(10, 1)).toBe('=B4+C3');
  });

  it('commits as before when the formula waits for no reference', async () => {
    const view = await mountEditor();
    await view.click({ row: 10, col: 1 });
    await press(view.surface, '=');
    view.type('=B3');
    await view.click({ row: 2, col: 2 });
    expect(view.editor() === null).toBe(true);
    expect(view.nameBox().value).toBe('C3');
    expect(await view.input(10, 1)).toBe('=B3');
  });

  it('points with the arrow keys in a typed edit, and F2 switches them to the caret', async () => {
    const view = await mountEditor();
    await view.click({ row: 10, col: 1 });
    await press(view.surface, '=');
    await press(view.editor()!, 'ArrowUp');
    expect(view.editor()?.value).toBe('=B10');
    await press(view.editor()!, 'ArrowUp');
    await press(view.editor()!, 'ArrowUp', { shiftKey: true });
    expect(view.editor()?.value).toBe('=B8:B9');
    view.type('=B8:B9+');
    await press(view.editor()!, 'F2');
    await press(view.editor()!, 'ArrowLeft');
    expect(view.editor()?.value).toBe('=B8:B9+');
    view.editor()!.setSelectionRange(7, 7);
    await press(view.editor()!, 'F2');
    await press(view.editor()!, 'ArrowLeft');
    expect(view.editor()?.value).toBe('=B8:B9+A11');
    await press(view.editor()!, 'Enter');
    expect(await view.input(10, 1)).toBe('=B8:B9+A11');
  });

  // Excel's Enter mode: an arrow where no reference fits commits and moves.
  it('commits and moves on an arrow where no reference fits, in an edit opened by typing', async () => {
    const view = await mountEditor();
    await view.click({ row: 10, col: 1 });
    await press(view.surface, 'n');
    view.type('note');
    await press(view.editor()!, 'ArrowDown');
    expect([view.editor() === null, view.nameBox().value]).toEqual([true, 'B12']);
    expect(await view.input(10, 1)).toBe('note');
    await press(view.surface, '=');
    view.type('=B3');
    await press(view.editor()!, 'ArrowRight');
    expect([view.editor() === null, view.nameBox().value]).toEqual([true, 'C12']);
    expect(await view.input(11, 1)).toBe('=B3');
    await press(view.surface, 'm');
    await press(view.editor()!, 'ArrowDown', { shiftKey: true });
    expect([view.nameBox().value, view.selectionBox() === null]).toEqual(['C12', false]);
    expect(await view.input(11, 2)).toBe('m');
    // F2 gives the arrows to the caret: nothing commits.
    await press(view.surface, 'k');
    await press(view.editor()!, 'F2');
    await press(view.editor()!, 'ArrowDown');
    expect(view.editor()?.value).toBe('k');
  });

  // a click in the edit places the caret, as Excel's Edit mode.
  it('moves the caret with the arrows after a click in an edit opened by typing', async () => {
    const view = await mountEditor();
    await view.click({ row: 10, col: 1 });
    await press(view.surface, 'H');
    view.type('Helo world');
    await act(async () => view.pressInEditor({ row: 10, col: 1 }));
    await press(view.editor()!, 'ArrowRight');
    expect([view.editor()?.value, view.nameBox().value]).toEqual(['Helo world', 'B11']);
  });

  // a range that cuts a merged block grows over it, as the selection does.
  it('grows a pointed range over a merged block it cuts', async () => {
    const view = await mountEditor();
    await view.click({ row: 10, col: 1 });
    await press(view.surface, '=');
    const from = pointAt(plain, { row: 2, col: 1 });
    const to = pointAt(plain, { row: 0, col: 2 });
    await act(async () => {
      fireEvent.mouseDown(view.surface, from);
      fireEvent.mouseMove(view.surface, to);
      fireEvent.mouseUp(window, to);
    });
    await view.idle();
    expect(view.editor()?.value).toBe('=A1:D3');
  });

  // re-pointing the same cell changes nothing, so no stale caret may follow.
  it('keeps the formula-bar caret where typing leaves it after the same cell is pointed again', async () => {
    const view = await mountEditor();
    await view.click({ row: 10, col: 1 });
    view.typeFormula('=');
    await view.click({ row: 2, col: 1 });
    await view.click({ row: 2, col: 1 });
    expect(view.formula()).toBe('=B3');
    await act(async () => {
      view.typeFormula('=B3+');
    });
    expect(view.formulaInput().selectionStart).toBe(4);
  });

  it('leaves the arrows to the caret in an edit opened with F2 until F2 is pressed', async () => {
    const view = await mountEditor();
    await view.click({ row: 10, col: 1 });
    await press(view.surface, 'F2');
    view.type('=');
    await press(view.editor()!, 'ArrowUp');
    expect(view.editor()?.value).toBe('=');
    await press(view.editor()!, 'F2');
    await press(view.editor()!, 'ArrowUp');
    expect(view.editor()?.value).toBe('=B10');
  });

  it('drags a range into a formula waiting for a reference', async () => {
    const view = await mountEditor();
    await view.click({ row: 10, col: 1 });
    await press(view.surface, '=');
    view.type('=SUM(');
    const from = pointAt(plain, { row: 2, col: 1 });
    const to = pointAt(plain, { row: 4, col: 1 });
    await act(async () => {
      fireEvent.mouseDown(view.surface, from);
      fireEvent.mouseMove(view.surface, to);
      fireEvent.mouseUp(window, to);
    });
    await view.idle();
    expect(view.editor()?.value).toBe('=SUM(B3:B5');
    expect(view.surface.querySelector('[data-testid="xlsx-pointed-reference"]') === null).toBe(false);
    view.type('=SUM(B3:B5)');
    expect(view.surface.querySelector('[data-testid="xlsx-pointed-reference"]') === null).toBe(true);
    await press(view.editor()!, 'Enter');
    expect(await view.input(10, 1)).toBe('=SUM(B3:B5)');
  });

  it('puts a clicked cell into a formula-bar draft waiting for a reference', async () => {
    const view = await mountEditor();
    await view.click({ row: 10, col: 1 });
    view.typeFormula('=');
    await view.click({ row: 2, col: 1 });
    expect(view.formula()).toBe('=B3');
    expect(view.nameBox().value).toBe('B11');
  });

  // Budget's title is merged over A1:D1.
  it('names a merged cell by its top-left cell, pointed or arrowed into', async () => {
    const view = await mountEditor();
    await view.click({ row: 10, col: 1 });
    await press(view.surface, '=');
    await view.click({ row: 0, col: 2 });
    expect(view.editor()?.value).toBe('=A1');
    await press(view.editor()!, 'Escape');
    await view.click({ row: 1, col: 2 });
    await press(view.surface, '=');
    await press(view.editor()!, 'ArrowUp');
    expect(view.editor()?.value).toBe('=A1');
  });
});

describe('XlsxEditor cell text', () => {
  // B12 in 20pt bold Georgia, dark red, and nothing else in its row.
  let styled: Fixture;
  beforeAll(() => {
    const handle = openWorkbook(plain.bytes.slice());
    try {
      handle.patchRangeStyle(0, 'B12', {
        fontFamily: 'Georgia',
        fontSize: 20,
        bold: true,
        textColor: '#c00000',
      });
      styled = fixtureFrom(handle.save());
    } finally {
      handle.dispose();
    }
  });

  // Excel: the active cell's own font, whatever the rest of the selection uses.
  it("draws the open edit in the active cell's font over a selection with mixed fonts", async () => {
    const view = await mountEditor(styled);
    await view.click({ row: 11, col: 1 });
    await act(async () => {
      fireEvent.mouseDown(view.surface, { ...pointAt(styled, { row: 10, col: 1 }), shiftKey: true });
      fireEvent.mouseUp(window, pointAt(styled, { row: 10, col: 1 }));
    });
    await view.idle();
    expect(view.nameBox().value).toBe('B12');
    await act(async () => {
      fireEvent.keyDown(view.surface, { key: 'x' });
    });
    await view.idle();
    const input = view.editor()!;
    expect([input.style.fontSize, input.style.fontWeight, input.style.color]).toEqual([
      '26.666667px',
      'bold',
      '#c00000',
    ]);
  });

  it('draws the open edit and the text waiting for its frame as the canvas draws the cell', async () => {
    const view = await mountEditor(styled);
    await view.click({ row: 11, col: 1 });
    await act(async () => {
      fireEvent.keyDown(view.surface, { key: 'x' });
    });
    await view.idle();
    const input = view.editor()!;
    // 20pt is 26.67px at 100%, as the canvas paints it.
    expect([
      input.style.fontSize,
      input.style.fontWeight,
      input.style.fontFamily,
      input.style.color,
    ]).toEqual(['26.666667px', 'bold', 'Georgia, sans-serif', '#c00000']);
    // the committed text shows over its cell until the worker's frame draws it.
    act(() => {
      fireEvent.keyDown(input, { key: 'Enter' });
    });
    const pending = view.surface.querySelector<HTMLElement>('[data-testid="xlsx-pending-cell"]');
    expect(pending && [pending.style.fontSize, pending.style.color]).toEqual([
      '26.666667px',
      '#c00000',
    ]);
    await view.idle();
  });
});
