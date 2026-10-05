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
import { afterAll, afterEach, beforeAll, describe, expect, it, spyOn } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { cellRect, initWasm, openWorkbook, selectionAt } from '@betteroffice/xlsx';
import type { CellAddr, ChartRegion, GridMeta, WorkbookHandle } from '@betteroffice/xlsx';
import { freezePaneOp, type XlsxCommand, type XlsxCommandState } from './commands';
import { XlsxEditor, type XlsxEditorApi } from './XlsxEditor';

const WASM = resolve(import.meta.dir, '../../xlsx/src/wasm/generated/xlsx_wasm_bg.wasm');
const FIXTURE = resolve(import.meta.dir, '../../xlsx/test-fixtures/sample.xlsx');
const CHART_FIXTURE = resolve(import.meta.dir, '../../xlsx/test-fixtures/charts.xlsx');
const UNDRAWABLE_FIXTURE = resolve(
  import.meta.dir,
  '../../xlsx/test-fixtures/unsupported-charts.xlsx'
);
const VIEWPORT = { width: 800, height: 600 };
const LINK_TARGET = 'https://example.com/report';
const LINK_CELL: CellAddr = { row: 5, col: 4 };

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');

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

// engine geometry calls and paints from here on.
function countEngine(handle: WorkbookHandle) {
  const counts = { positions: 0, paints: 0 };
  const cellPosition = handle.cellPosition.bind(handle);
  const displayList = handle.displayList.bind(handle);
  handle.cellPosition = (sheet, row, col) => {
    counts.positions += 1;
    return cellPosition(sheet, row, col);
  };
  handle.displayList = (viewport) => {
    counts.paints += 1;
    return displayList(viewport);
  };
  return counts;
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
});

function pointAt(fixture: Fixture, addr: CellAddr): { clientX: number; clientY: number } {
  const rect = cellRect(fixture.grid, addr.row, addr.col);
  if (!rect) throw new Error(`cell ${addr.row},${addr.col} is outside the painted window`);
  return { clientX: rect.x + rect.w / 2, clientY: rect.y + rect.h / 2 };
}

async function mountEditor(
  fixture: Fixture = plain,
  onSave?: (bytes: Uint8Array) => void,
  onChange?: () => void
) {
  const ready: { handle: WorkbookHandle | null; flush?: XlsxEditorApi['flush'] } = { handle: null };
  const view = render(
    <XlsxEditor
      file={fixture.bytes.slice()}
      onChange={onChange}
      onSave={onSave}
      onReady={(api) => {
        ready.handle = api.handle;
        ready.flush = api.flush;
      }}
    />
  );
  const nameBox = () => view.getByTestId('xlsx-name-box') as HTMLInputElement;
  await waitFor(() => expect(nameBox().value).toBe('A1'));
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
    workbook: () => ready.handle!,
    flush: () => act(() => ready.flush!()),
    typeFormula: (value: string) =>
      fireEvent.change(view.getByTestId('xlsx-formula-input'), { target: { value } }),
    reopenWith: (next: Fixture) =>
      act(async () => {
        view.rerender(
          <XlsxEditor
            file={next.bytes.slice()}
            onChange={onChange}
            onSave={onSave}
            onReady={(api) => {
              ready.handle = api.handle;
              ready.flush = api.flush;
            }}
          />
        );
      }),
    click: press,
    doubleClick: (addr: CellAddr) => {
      press(addr);
      press(addr);
      fireEvent.doubleClick(surface, pointAt(fixture, addr));
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

    view.click({ row: 2, col: 0 });
    expect(changes).toBe(0);

    view.doubleClick({ row: 2, col: 0 });
    view.type('Edited item');
    view.click({ row: 3, col: 1 });
    expect(changes).toBe(1);
  });

  it('flushes open cell and formula drafts before a host checkpoint without blur', async () => {
    const view = await mountEditor();
    view.doubleClick({ row: 2, col: 0 });
    view.type('Unblurred cell');
    expect(view.workbook().cell(0, 2, 0).input).toBe('Line item 1');
    view.flush();
    expect(view.workbook().cell(0, 2, 0).input).toBe('Unblurred cell');
    view.typeFormula('Unblurred formula');
    expect(view.workbook().cell(0, 2, 0).input).toBe('Unblurred cell');
    view.flush();
    const reopened = openWorkbook(view.workbook().save());
    try {
      expect(reopened.cell(0, 2, 0).input).toBe('Unblurred formula');
    } finally {
      reopened.dispose();
    }
  });

  it('commits the open editor and moves the selection when another cell is clicked', async () => {
    const view = await mountEditor();

    view.doubleClick({ row: 2, col: 0 });
    expect(view.editor()?.value).toBe('Line item 1');

    view.type('Edited item');
    view.click({ row: 3, col: 1 });

    expect(view.editor()).toBeNull();
    expect(view.nameBox().value).toBe('B4');
    expect(view.workbook().cell(0, 2, 0).input).toBe('Edited item');
  });

  it('commits and reopens on the target when another cell is double-clicked', async () => {
    const view = await mountEditor();

    view.doubleClick({ row: 2, col: 0 });
    view.type('Edited item');
    view.doubleClick({ row: 3, col: 1 });

    expect(view.editor()?.value).toBe('200');
    expect(view.nameBox().value).toBe('B4');
    expect(view.workbook().cell(0, 2, 0).input).toBe('Edited item');
  });

  it('leaves a formula cell unchanged when the pointer moves on', async () => {
    const view = await mountEditor();

    view.doubleClick({ row: 2, col: 3 });
    expect(view.editor()?.value).toBe('=B3+C3');

    view.click({ row: 6, col: 0 });

    expect(view.editor()).toBeNull();
    expect(view.nameBox().value).toBe('A7');
    expect(view.workbook().cell(0, 2, 3).input).toBe('=B3+C3');

    view.doubleClick({ row: 3, col: 3 });
    expect(view.editor()?.value).toBe('=B4+C4');

    view.doubleClick({ row: 7, col: 0 });

    expect(view.editor()?.value).toBe('Line item 6');
    expect(view.nameBox().value).toBe('A8');
    expect(view.workbook().cell(0, 3, 3).input).toBe('=B4+C4');
  });

  it('keeps a press inside the open editor from committing or moving on', async () => {
    const view = await mountEditor();

    view.doubleClick({ row: 2, col: 0 });
    view.type('Edited item');
    view.pressInEditor({ row: 6, col: 0 });

    expect(view.editor()?.value).toBe('Edited item');
    expect(view.nameBox().value).toBe('A3');
    expect(view.workbook().cell(0, 2, 0).input).toBe('Line item 1');
  });

  it('keeps a double-click inside the open editor from reopening it', async () => {
    const view = await mountEditor();

    view.doubleClick({ row: 2, col: 0 });
    view.type('Edited item');
    view.doubleClickInEditor({ row: 2, col: 0 });

    expect(view.editor()?.value).toBe('Edited item');
  });

  it('dismisses the editor without following a hyperlink in the clicked cell', async () => {
    const view = await mountEditor(linked);

    view.doubleClick({ row: 2, col: 0 });
    view.type('Edited item');
    view.click(LINK_CELL);

    expect(view.editor()).toBeNull();
    expect(view.nameBox().value).toBe('E6');
    expect(opened).toEqual([]);

    view.click(LINK_CELL);

    expect(opened).toEqual([LINK_TARGET]);
  });
});

describe('XlsxEditor keyboard', () => {
  // dozens of keys, each with a full paint: generous for a loaded runner.
  const MANY_KEYS_MS = 30_000;
  const press = (target: Element, key: string, init: KeyboardEventInit = {}) =>
    act(async () => {
      fireEvent.keyDown(target, { key, ...init });
    });
  const scrollTo = (view: Awaited<ReturnType<typeof mountEditor>>, left: number, top: number) =>
    act(async () => {
      view.surface.scrollLeft = left;
      view.surface.scrollTop = top;
      fireEvent.scroll(view.surface);
    });
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
    view.click({ row: 3, col: 1 });
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
    expect(document.activeElement).toBe(input);

    await press(input!, 'Enter');
    expect(view.workbook().cell(0, 33, 16).input).toBe('7');
    expect(view.nameBox().value).toBe('Q35');
  }, MANY_KEYS_MS);

  it('keeps a revealed cell clear of the frozen panes', async () => {
    const view = await mountEditor(wideFrozen);
    const pane = cellRect(wideFrozen.grid, 2, 2)!;
    view.click({ row: 3, col: 1 });
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
    view.click({ row: 3, col: 1 });
    await scrollTo(view, 1500, 600);

    await press(view.surface, '7');
    expect(view.editor()?.value).toBe('7');
    const box = editorBox(view.editor()!);
    // at the window's top-left, give or take the engine's f32 rounding.
    expect(box.left).toBeCloseTo(0, 1);
    expect(box.top).toBeCloseTo(0, 1);
    await press(view.editor()!, 'Enter');
    expect(view.workbook().cell(0, 3, 1).input).toBe('7');
  });

  it('leaves the scroll alone for a move that stays on screen', async () => {
    const view = await mountEditor(wide);
    view.click({ row: 3, col: 1 });
    await press(view.surface, 'ArrowRight');
    await press(view.surface, 'ArrowDown');
    expect(view.nameBox().value).toBe('C5');
    expect(view.surface.scrollLeft).toBe(0);
    expect(view.surface.scrollTop).toBe(0);
  });

  it('keeps the view on select all, as the menu does', async () => {
    const view = await mountEditor(wide);
    view.click({ row: 3, col: 1 });
    await press(view.surface, 'a', { ctrlKey: true });
    expect(view.selectionBox()).not.toBeNull();
    expect(view.surface.scrollLeft).toBe(0);
    expect(view.surface.scrollTop).toBe(0);
  });

  it('keeps typing into an open edit whose cell scrolled away, and scrolls back', async () => {
    const view = await mountEditor(wide);
    view.click({ row: 3, col: 1 });
    await press(view.surface, '7');
    const input = view.editor()!;
    await scrollTo(view, 1500, 600);

    // still mounted and focused, out of sight.
    expect(view.editor()).toBe(input);
    expect(document.activeElement).toBe(input);
    expect(input.style.opacity).toBe('0');

    await press(input, '5');
    fireEvent.change(input, { target: { value: '75' } });
    expect(view.editor()).toBe(input);
    expect(input.style.opacity).toBe('');
    const box = editorBox(input);
    expect(box.left).toBeCloseTo(0, 1);
    expect(box.top).toBeCloseTo(0, 1);
    await press(input, 'Enter');
    expect(view.workbook().cell(0, 3, 1).input).toBe('75');
  });

  it('keeps an IME composition in the edit while its cell scrolls back', async () => {
    const view = await mountEditor(wide);
    view.click({ row: 3, col: 1 });
    await press(view.surface, 'F2');
    const input = view.editor()!;
    await scrollTo(view, 1500, 600);

    fireEvent.compositionStart(input);
    await press(input, 'Process', { isComposing: true, keyCode: 229 });
    expect(view.editor()).toBe(input);
    expect(document.activeElement).toBe(input);
    expect(input.style.opacity).toBe('');
    fireEvent.change(input, { target: { value: '日本' } });
    fireEvent.compositionEnd(input, { data: '日本' });
    await press(input, 'Enter');
    expect(view.workbook().cell(0, 3, 1).input).toBe('日本');
  });

  it('types into a cell the frozen panes leave no room to show', async () => {
    const view = await mountEditor(paneFilled);
    view.click({ row: 3, col: 1 });
    await press(view.surface, 'End', { ctrlKey: true });
    const { row, col } = cellAt(view.nameBox().value);
    expect(row).toBeGreaterThan(40);

    await press(view.surface, '7');
    const input = view.editor()!;
    expect(document.activeElement).toBe(input);
    await press(input, '8');
    fireEvent.change(input, { target: { value: '78' } });
    await press(input, 'Enter');
    expect(view.workbook().cell(0, row, col).input).toBe('78');
  });

  it('leaves a scrolled-away edit where it is for keys that type nothing', async () => {
    const view = await mountEditor(wide);
    view.click({ row: 3, col: 1 });
    await press(view.surface, '7');
    await scrollTo(view, 1500, 600);
    const input = view.editor()!;
    for (const key of ['Meta', 'Control', 'Shift', 'Alt']) await press(input, key);
    await press(input, 'c', { metaKey: true });
    expect([view.surface.scrollLeft, view.surface.scrollTop]).toEqual([1500, 600]);
    expect(input.style.opacity).toBe('0');
    expect(document.activeElement).toBe(input);
  });

  it('paints once for an Enter that commits and scrolls', async () => {
    const view = await mountEditor(wide);
    // the last row the window shows whole: Enter moves below the edge.
    const offsets = wide.grid.rowOffsets;
    const last = offsets.filter((bottom) => bottom <= VIEWPORT.height).length - 2;
    view.click({ row: last, col: 1 });
    await press(view.surface, '7');
    const handle = view.workbook();
    const displayList = handle.displayList.bind(handle);
    let paints = 0;
    handle.displayList = (viewport) => {
      paints += 1;
      return displayList(viewport);
    };
    await press(view.editor()!, 'Enter');
    // what the browser does next: the scroll event, then a frame.
    await act(async () => {
      fireEvent.scroll(view.surface);
      await new Promise((resolve) => setTimeout(resolve, 50));
    });
    expect(view.surface.scrollTop).toBeGreaterThan(0);
    expect(view.workbook().cell(0, last, 1).input).toBe('7');
    expect(paints).toBe(1);
  });

  it('judges a key against the live view, not a frame its scroll has not repainted', async () => {
    const view = await mountEditor(wide);
    view.click({ row: 3, col: 1 });
    // scrolled, and the key lands before the scroll's frame is painted.
    await act(async () => {
      view.surface.scrollLeft = 1500;
    });
    await press(view.surface, 'ArrowRight');
    expect(view.nameBox().value).toBe('C4');
    expect(view.surface.scrollLeft).toBe(Math.floor(view.workbook().cellPosition(0, 3, 2).x));
  });

  it('brings a partly hidden edit fully into view when typing into it', async () => {
    const view = await mountEditor(wide);
    view.click({ row: 3, col: 1 });
    await press(view.surface, '7');
    const start = view.workbook().cellPosition(0, 3, 1).x;
    await scrollTo(view, Math.ceil(start) + 10, 0);
    expect(view.editor()!.style.opacity).toBe('');
    await press(view.editor()!, '5');
    expect(view.surface.scrollLeft).toBe(Math.floor(start));
  });

  it('shows a cell wider than the view as it is, typing or moving down', async () => {
    const view = await mountEditor(wideColumn);
    view.click({ row: 3, col: 1 });
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
    expect(counts).toEqual({ positions: 0, paints: 0 });
  });

  it('judges the cell an Enter lands on against the row its commit grew', async () => {
    const view = await mountEditor(tallFont);
    // Enter lands on the last row the window shows whole.
    const target = tallFont.grid.rowOffsets.filter((bottom) => bottom <= VIEWPORT.height).length - 2;
    view.click({ row: target - 1, col: 5 });
    await press(view.surface, '7');
    await press(view.editor()!, 'Enter');
    expect(view.workbook().cell(0, target - 1, 5).input).toBe('7');
    const cell = view.workbook().cellPosition(0, target, 5);
    expect(cell.y + cell.height).toBeLessThanOrEqual(view.surface.scrollTop + VIEWPORT.height);
    expect(view.surface.scrollTop).toBeGreaterThan(0);
  });

  it('asks the engine nothing for the first cell under frozen panes at scroll 0', async () => {
    const view = await mountEditor(wideFrozen);
    const counts = countEngine(view.workbook());
    view.click({ row: 2, col: 2 });
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
    view.click({ row: 3, col: 1 });
    await press(view.surface, '7');
    await act(async () => {
      fireEvent.blur(view.editor()!, { relatedTarget: null });
    });
    expect(view.editor()).toBeNull();
    expect(view.workbook().cell(0, 3, 1).input).toBe('7');
    // identity checks stay cheap to print when they fail.
    expect(document.activeElement?.getAttribute('data-testid')).toBe('xlsx-scroll');
  });

  it('keeps the edit open across a window or tab switch', async () => {
    const view = await mountEditor(wide);
    view.click({ row: 3, col: 1 });
    await press(view.surface, '7');
    const input = view.editor()!;
    const before = view.workbook().cell(0, 3, 1).input;
    const away = spyOn(document, 'hasFocus').mockReturnValue(false);
    try {
      await act(async () => {
        fireEvent.blur(input, { relatedTarget: null });
      });
    } finally {
      away.mockRestore();
    }
    expect(view.editor() === input).toBe(true);
    expect(view.workbook().cell(0, 3, 1).input).toBe(before);
    // back: the next key types into the edit.
    fireEvent.change(input, { target: { value: '78' } });
    await press(input, 'Enter');
    expect(view.workbook().cell(0, 3, 1).input).toBe('78');
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

  it('leaves the formula bar focused when an edit it committed scrolls back', async () => {
    const view = await mountEditor(wide);
    view.click({ row: 3, col: 1 });
    await press(view.surface, '7');
    const input = view.editor()!;
    await scrollTo(view, 1500, 600);

    const formula = view.formulaInput();
    act(() => formula.focus());
    expect(view.editor()).toBeNull();
    expect(view.workbook().cell(0, 3, 1).input).toBe('7');
    await scrollTo(view, 0, 0);
    expect(document.activeElement).toBe(formula);
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
    expect(view.selectionBox()).toBeNull();
    fireEvent.mouseUp(window, chartCenter(chart));

    fireEvent.mouseDown(view.surface, { clientX: 2, clientY: 2 });
    await waitFor(() => expect(view.outline()).toBeNull());
    expect(view.selectionBox()).not.toBeNull();
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
    expect(view.error()).toBeNull();
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

    expect(view.editor()).toBeNull();
    expect(view.outline()).not.toBeNull();
    expect(view.formula()).toBe(before);
    expect(view.canUndo()).toBe(false);
  });

  it('commits an open edit before taking the press, then selects the chart', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);

    view.doubleClick({ row: 1, col: 1 });
    view.type('Edited item');
    fireEvent.mouseDown(view.surface, chartCenter(chart));

    const outline = await waitFor(() => view.outline()!);
    expect(outline.getAttribute('data-chart-id')).toBe(chart.id);
    expect(view.editor()).toBeNull();
    expect(view.workbook().cell(0, 1, 1).input).toBe('Edited item');
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

    view.doubleClick({ row: 1, col: 1 });
    view.type('Edited item');

    await act(async () => {
      view.surface.scrollTop = SCROLLED_BY;
      fireEvent.scroll(view.surface);
    });
    // the input stays mounted, out of sight, while its cell is away; the press
    // commits it without a blur.
    expect(view.editor()?.style.opacity).toBe('0');

    fireEvent.mouseDown(view.surface, chartCenter(target));
    await waitFor(() => view.outline()!);
    fireEvent.mouseUp(window, chartCenter(target));

    expect(view.workbook().cell(0, 1, 1).input).toBe('Edited item');
  });

  it('leaves a press inside the open editor to the editor, not a chart behind it', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);

    view.doubleClick({ row: 1, col: 1 });
    view.type('Edited item');
    // the press lands on the input, at coordinates the chart also covers: the
    // editor is a dom overlay, so without its guard the hit test would reach
    // the chart painted underneath.
    fireEvent.mouseDown(view.editor()!, chartCenter(chart));

    expect(view.editor()?.value).toBe('Edited item');
    expect(view.outline()).toBeNull();
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

    expect(view.error()).toBeNull();
    expect(view.canUndo()).toBe(false);
    const still = (view.workbook().displayList({ x: 0, y: 0, ...VIEWPORT }).charts ?? []).find(
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

    await waitFor(() => expect(view.outline()).toBeNull());
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
    await waitFor(() => expect(view.outline()).toBeNull());

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
    expect(view.error()).toBeNull();
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
    await waitFor(() => expect(view.outline()).toBeNull());
    await waitFor(() => expect(view.selectionBox()).not.toBeNull());
  });
});

describe('XlsxEditor host integration', () => {
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
    const surface = view.getByTestId('xlsx-scroll');
    const target = { row: 2, col: 0 };
    const before = api!.handle.cell(0, target.row, target.col).input;

    fireEvent.doubleClick(surface, pointAt(plain, target));
    fireEvent.keyDown(surface, { key: 'x' });
    fireEvent.keyDown(surface, { key: 'Delete' });

    expect(view.queryByTestId('xlsx-toolbar')).toBeNull();
    expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
    expect(api!.handle.cell(0, target.row, target.col).input).toBe(before);
    expect(changes).toBe(0);

    await act(async () => {
      expect(api!.selectCells(0, selectionAt({ row: 3, col: 1 }))).toBe(true);
    });
    await waitFor(() => {
      const selected = view.getByRole('gridcell', { selected: true });
      expect(selected.textContent).toContain('B4');
    });
    await act(async () => api!.clearSelection());
    await waitFor(() =>
      expect(view.queryAllByRole('gridcell', { selected: true })).toHaveLength(0)
    );
    expect(api!.selectCells(99, selectionAt({ row: 0, col: 0 }))).toBe(false);
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
    const surface = view.getByTestId('xlsx-scroll');
    const target = { row: 2, col: 0 };

    await act(async () => {
      expect(api!.selectCells(0, selectionAt(target))).toBe(true);
    });
    fireEvent.doubleClick(surface, pointAt(plain, target));
    const editor = await waitFor(() => view.getByTestId('xlsx-cell-editor'));
    fireEvent.change(editor, { target: { value: 'Host edit' } });
    fireEvent.keyDown(editor, { key: 'Enter' });
    expect(changes).toBe(1);

    await act(async () => {
      api!.selectCells(0, selectionAt({ row: 3, col: 1 }));
      api!.clearSelection();
    });
    let saved!: Uint8Array;
    await act(async () => {
      saved = api!.save();
    });
    expect(changes).toBe(1);

    const reopened = openWorkbook(saved);
    try {
      expect(reopened.cell(0, target.row, target.col).input).toBe('Host edit');
    } finally {
      reopened.dispose();
    }
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
      const target = { row: 2, col: 0 };
      const before = api!.handle.cell(0, target.row, target.col).input;
      await act(async () => {
        api!.selectCells(0, selectionAt(target));
      });

      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'v', ctrlKey: true });
      view.rerender(<XlsxEditor file={file} onReady={onReady} readOnly />);
      await act(async () => resolveClipboard('late paste'));

      expect(api!.handle.cell(0, target.row, target.col).input).toBe(before);
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
    const before = workbook.cell(0, 6, 4).input;
    const stage = async (input: string) => {
      await act(async () => {
        workbook.propose('Audit agent', 'Review this change', [
          { sheet: 0, row: 6, col: 4, input },
        ]);
        api!.refreshProposals();
      });
    };

    await stage('12');
    expect(workbook.cell(0, 6, 4).input).toBe(before);
    fireEvent.click(view.getByTestId('xlsx-proposals-button'));
    expect(view.getByTestId('xlsx-proposal').textContent).toContain('Audit agent');
    fireEvent.click(view.getByTestId('xlsx-proposal-accept'));
    await waitFor(() => expect(workbook.cell(0, 6, 4).input).toBe('12'));
    expect(workbook.listProposals()).toHaveLength(0);
    fireEvent.click(view.getByTestId('xlsx-undo'));
    await waitFor(() => expect(workbook.cell(0, 6, 4).input).toBe(before));

    await stage('24');
    fireEvent.click(view.getByTestId('xlsx-proposal-reject'));
    await waitFor(() => expect(workbook.listProposals()).toHaveLength(0));
    expect(workbook.cell(0, 6, 4).input).toBe(before);

    await stage('42');
    await act(async () => {
      workbook.editCell(0, 6, 4, '99');
      api!.refreshProposals();
    });
    fireEvent.click(view.getByTestId('xlsx-proposal-accept'));
    await waitFor(() =>
      expect(view.getByTestId('xlsx-proposal-stale').textContent).toContain('E7')
    );
    expect(workbook.cell(0, 6, 4).input).toBe('99');
    fireEvent.click(view.getByTestId('xlsx-proposal-force'));
    await waitFor(() => expect(workbook.cell(0, 6, 4).input).toBe('42'));
    expect(workbook.listProposals()).toHaveLength(0);

    await act(async () => {
      workbook.editCell(0, 1, 6, '10');
    });
    await stage('=G2*2');
    expect(view.getByTestId('xlsx-proposal-cell-new').textContent).toBe('20');
    await act(async () => {
      workbook.editCell(0, 1, 6, '99');
      api!.refreshProposals();
    });
    fireEvent.click(view.getByTestId('xlsx-proposal-accept'));
    await waitFor(() =>
      expect(view.getByTestId('xlsx-proposal-cell-new').textContent).toBe('198')
    );
    expect(workbook.cell(0, 6, 4).input).toBe('42');
    fireEvent.click(view.getByTestId('xlsx-proposal-accept'));
    await waitFor(() => expect(workbook.cell(0, 6, 4).input).toBe('=G2*2'));
    expect(workbook.listProposals()).toHaveLength(0);
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
    const surface = view.getByTestId('xlsx-scroll');
    const chart = api!.handle.displayList({ x: 0, y: 0, ...VIEWPORT }).charts![0];
    fireEvent.mouseDown(surface, chartCenter(chart));
    fireEvent.mouseUp(window, chartCenter(chart));
    await act(async () => {
      fireEvent.keyDown(surface, { key: 'ArrowRight' });
    });
    expect(
      Math.round(parseFloat(view.getByTestId('xlsx-chart-selection').style.left))
    ).toBe(Math.round(chart.rect.x + 1));
    await act(async () => {
      expect(api!.selectCells(1, selectionAt({ row: 0, col: 0 }))).toBe(true);
    });
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 350));
    });
    await act(async () => {
      api!.selectCells(0, selectionAt({ row: 0, col: 0 }));
    });
    const after = api!.handle
      .displayList({ x: 0, y: 0, ...VIEWPORT })
      .charts!.find((c) => c.id === chart.id)!;
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
    const target = { row: 2, col: 0 };
    await act(async () => {
      api!.selectCells(0, selectionAt(target));
    });
    fireEvent.doubleClick(view.getByTestId('xlsx-scroll'), pointAt(plain, target));
    const editor = await waitFor(() => view.getByTestId('xlsx-cell-editor'));
    fireEvent.change(editor, { target: { value: 'Draft that must survive' } });
    await act(async () => {
      api!.selectCells(0, selectionAt({ row: 3, col: 1 }));
    });
    expect(api!.handle.cell(0, target.row, target.col).input).toBe(
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
        const target = { row: 2, col: 0 };
        await act(async () => {
          api!.selectCells(0, selectionAt(target));
        });
        if (source === 'cell') {
          fireEvent.doubleClick(view.getByTestId('xlsx-scroll'), pointAt(plain, target));
        }
        const input = view.getByTestId(
          source === 'cell' ? 'xlsx-cell-editor' : 'xlsx-formula-input'
        );
        fireEvent.change(input, { target: { value: 'Saved draft' } });
        expect(api!.selectCells(-1, selectionAt(target))).toBe(false);
        expect(api!.handle.cell(0, target.row, target.col).input).toBe('Line item 1');
        let saved: Uint8Array | undefined;
        await act(async () => {
          if (action === 'save') saved = api!.save();
          else if (action === 'clear') api!.clearSelection();
          else api!.selectCells(0, selectionAt({ row: 3, col: 1 }));
        });
        expect(api!.handle.cell(0, target.row, target.col).input).toBe('Saved draft');
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
    const target = { row: 2, col: 0 };
    await act(async () => {
      api!.selectCells(0, selectionAt(target));
    });
    fireEvent.change(view.getByTestId('xlsx-formula-input'), {
      target: { value: 'First draft' },
    });
    await act(async () => {
      api!.save();
    });
    fireEvent.doubleClick(view.getByTestId('xlsx-scroll'), pointAt(plain, target));
    const editor = view.getByTestId('xlsx-cell-editor');
    fireEvent.change(editor, { target: { value: 'Second draft' } });
    fireEvent.blur(editor);
    expect(api!.handle.cell(0, target.row, target.col).input).toBe('Second draft');
  });
});

describe('XlsxEditor menu commands', () => {
  // Budget: A1 title, A2 "Item", A3..A6 "Line item 1..4"; B2..D2 Q1, Q2, Total.
  async function mountCommands() {
    let api: XlsxEditorApi | undefined;
    let state: XlsxCommandState | undefined;
    render(
      <XlsxEditor
        file={plain.bytes.slice()}
        onCommandStateChange={(next) => {
          state = next;
        }}
        onReady={(ready) => {
          api = ready;
        }}
      />
    );
    await waitFor(() => expect(api).toBeDefined());
    const select = (anchor: CellAddr, focus: CellAddr = anchor) =>
      act(async () => {
        api!.selectCells(api!.handle.sheetInfo().activeSheet, { anchor, focus });
      });
    const run = (command: XlsxCommand) =>
      act(async () => {
        api!.run(command);
      });
    const column = (col: number, rows = 8) =>
      Array.from({ length: rows }, (_, row) => api!.handle.cell(0, row, col).input);
    const row = (at: number, cols = 4) =>
      Array.from({ length: cols }, (_, col) => api!.handle.cell(0, at, col).input);
    return { api: () => api!, state: () => state!, select, run, column, row };
  }

  it('inserts as many rows as are selected above the selection, as one undo step', async () => {
    const view = await mountCommands();
    const before = view.column(0);
    await view.select({ row: 2, col: 0 }, { row: 3, col: 1 });
    await view.run('insertRowAbove');
    expect(view.column(0, 10)).toEqual([...before.slice(0, 2), '', '', ...before.slice(2)]);
    await view.run('undo');
    expect(view.column(0)).toEqual(before);
  });

  it('inserts rows below the selection', async () => {
    const view = await mountCommands();
    const before = view.column(0);
    await view.select({ row: 2, col: 0 });
    await view.run('insertRowBelow');
    expect(view.column(0, 9)).toEqual([...before.slice(0, 3), '', ...before.slice(3)]);
  });

  it('inserts columns left and right of the selection', async () => {
    const view = await mountCommands();
    const header = view.row(1);
    await view.select({ row: 1, col: 1 });
    await view.run('insertColumnLeft');
    expect(view.row(1, 5)).toEqual([header[0], '', ...header.slice(1)]);
    await view.select({ row: 1, col: 2 });
    await view.run('insertColumnRight');
    expect(view.row(1, 6)).toEqual([header[0], '', header[1], '', ...header.slice(2)]);
  });

  it('deletes the selected rows and columns', async () => {
    const view = await mountCommands();
    const before = view.column(0);
    await view.select({ row: 2, col: 0 }, { row: 3, col: 0 });
    await view.run('deleteRows');
    expect(view.column(0, 4)).toEqual([...before.slice(0, 2), ...before.slice(4, 6)]);
    const header = view.row(1);
    await view.select({ row: 1, col: 1 });
    await view.run('deleteColumns');
    expect(view.row(1, 3)).toEqual([header[0], ...header.slice(2)]);
  });

  it('deletes the values of the selected cells', async () => {
    const view = await mountCommands();
    await view.select({ row: 2, col: 0 }, { row: 2, col: 2 });
    await view.run('deleteValues');
    expect(view.row(2, 3)).toEqual(['', '', '']);
    expect(view.row(3, 1)).toEqual(['Line item 2']);
  });

  it('freezes rows and columns, by count or up to the selection, independently', async () => {
    const view = await mountCommands();
    const frozen = () => {
      const info = view.api().handle.sheetInfo();
      return [info.frozenRows, info.frozenCols];
    };
    await view.select({ row: 3, col: 2 });
    await view.run('freezeRows:1');
    expect(frozen()).toEqual([1, 0]);
    await view.run('freezeColumns:current');
    expect(frozen()).toEqual([1, 3]);
    await view.run('freezeRows:current');
    expect(frozen()).toEqual([4, 3]);
    expect([view.state().frozenRows, view.state().frozenColumns]).toEqual([4, 3]);
    await view.run('freezeRows:0');
    await view.run('freezeColumns:2');
    expect(frozen()).toEqual([0, 2]);
    await view.run('freezeColumns:0');
    expect(frozen()).toEqual([0, 0]);
  });

  it('adds a sheet after the active one and opens it', async () => {
    const view = await mountCommands();
    await view.run('insertSheet');
    const info = view.api().handle.sheetInfo();
    expect(info.sheetNames).toEqual(['Budget', 'Sheet4', 'Summary', 'Styled']);
    expect(info.activeSheet).toBe(1);
  });

  it('clears every style and the number format of the selection', async () => {
    const view = await mountCommands();
    await view.select({ row: 2, col: 1 }, { row: 3, col: 2 });
    await view.run('bold');
    await view.run('numberFormat:percent');
    await view.run('align:center');
    let formatting = view.api().handle.selectionFormatting(0, 'B3:C4');
    expect([formatting.bold, formatting.numberFormat, formatting.horizontalAlignment]).toEqual([
      true,
      'percent',
      'center',
    ]);
    await view.run('clearFormatting');
    formatting = view.api().handle.selectionFormatting(0, 'B3:C4');
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

  it('runs the toolbar commands a menu repeats: merge, wrapping and zoom', async () => {
    const view = await mountCommands();
    await view.select({ row: 2, col: 1 }, { row: 3, col: 2 });
    expect(view.state().canMerge).toBe(true);
    await view.run('merge:all');
    expect(view.api().handle.mergedRanges(0, 'B3:C4')).toHaveLength(1);
    expect(view.state().canUnmerge).toBe(true);
    await view.run('wrap:wrap');
    expect(view.api().handle.selectionFormatting(0, 'B3').textWrapping).toBe('wrap');
    await view.run('zoom:150');
    expect(view.state().zoom).toBe(1.5);
  });
});
