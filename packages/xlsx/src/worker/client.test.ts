import { afterEach, beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { initWasm, openWorkbook } from '../wasm/loader';
import type { WorkbookHandle, WorkbookUpdateOrigin } from '../wasm/loader';
import { StaleProposalError } from '../wasm/staleProposal';
import { openWorkbookWorker, SheetGoneError, type WorkbookProxy } from './client';

const FIXTURE = resolve(import.meta.dir, '../../test-fixtures/sample.xlsx');
const WASM = resolve(import.meta.dir, '../wasm/generated/xlsx_wasm_bg.wasm');
const VIEW = { viewport: { x: 0, y: 0, width: 800, height: 600 }, zoom: 1 };

const sample = () => new Uint8Array(readFileSync(FIXTURE));
let open: Array<{ dispose(): void }> = [];

function local(clientId: number): WorkbookHandle {
  const handle = openWorkbook(sample(), { collaborative: true, clientId });
  open.push(handle);
  return handle;
}

async function worker(clientId: number, initialUpdate?: Uint8Array): Promise<WorkbookProxy> {
  const proxy = await openWorkbookWorker(sample(), {
    collaborative: true,
    clientId,
    initialUpdate,
  });
  open.push(proxy);
  return proxy;
}

function record(replica: {
  onUpdate(listener: (update: Uint8Array, origin: WorkbookUpdateOrigin) => void): () => void;
}) {
  const updates: Array<[WorkbookUpdateOrigin, number[]]> = [];
  replica.onUpdate((update, origin) => updates.push([origin, [...update]]));
  return updates;
}

beforeAll(() => initWasm(new Uint8Array(readFileSync(WASM))));
afterEach(() => {
  for (const handle of open) handle.dispose();
  open = [];
});

describe('workbook worker', () => {
  it('makes the same updates and state as the engine in this thread', async () => {
    const here = local(41);
    const there = await worker(41);
    const hereUpdates = record(here);
    const thereUpdates = record(there);
    const sheet = (await there.sheetInfo()).sheetIds[0];

    here.editCell(0, 2, 0, 'Typed');
    here.editCells(0, [
      { row: 4, col: 1, input: '12' },
      { row: 5, col: 1, input: '=B5*2' },
    ]);
    here.applyOps([{ type: 'insertRows', sheet: 0, at: 1, count: 2 }]);
    here.undo();
    here.patchRangeStyle(0, 'A1:B2', { bold: true });

    await there.onSheet(sheet, 'editCell', 0, 2, 0, 'Typed');
    await there.onSheet(sheet, 'editCells', 0, [
      { row: 4, col: 1, input: '12' },
      { row: 5, col: 1, input: '=B5*2' },
    ]);
    await there.onSheet(sheet, 'applyOps', [{ type: 'insertRows', sheet: 0, at: 1, count: 2 }]);
    await there.undo();
    await there.toggle(sheet, 'A1:B2', 'bold');

    expect(thereUpdates).toEqual(hereUpdates);
    expect(await there.encodeStateAsUpdate()).toEqual(here.encodeStateAsUpdate());
    expect(await there.cell(0, 5, 1)).toEqual(here.cell(0, 5, 1));
    expect(await there.historyState()).toEqual(here.historyState());
  });

  it('applies a peer update and a seeded state before it answers', async () => {
    const peer = local(42);
    const peerUpdates: Uint8Array[] = [];
    peer.onUpdate((update) => peerUpdates.push(update));
    peer.editCell(0, 1, 1, 'from the peer');
    const seeded = await worker(43, peer.encodeStateAsUpdate());
    expect((await seeded.cell(0, 1, 1)).input).toBe('from the peer');

    const updates = record(seeded);
    peer.editCell(0, 1, 2, 'second');
    const result = await seeded.applyUpdate(peerUpdates.at(-1)!);
    expect(result.applied).toBe(true);
    expect(updates.map(([origin]) => origin)).toEqual(['remote']);
    expect((await seeded.cell(0, 1, 2)).input).toBe('second');
  });

  it('builds a frame of the view after the requests posted before it', async () => {
    const proxy = await worker(44);
    const sheet = (await proxy.sheetInfo()).sheetIds[0];
    void proxy.onSheet(sheet, 'editCell', 0, 2, 0, 'Before the frame');
    expect(proxy.mutationsInFlight).toBe(1);
    const frame = await proxy.frame({
      ...VIEW,
      selection: { anchor: { row: 2, col: 0 }, focus: { row: 3, col: 1 } },
    });
    expect(proxy.mutationsInFlight).toBe(0);
    expect(frame.seq).toBe(proxy.posted);
    expect(frame.cells['2:0'].input).toBe('Before the frame');
    expect(frame.focus?.a1).toBe('B4');
    expect(frame.sheetInfo.activeSheet).toBe(0);
    expect(frame.displayList.grid?.startRow).toBe(0);
    expect(frame.history.canUndo).toBe(true);
    expect(frame.formatting).toEqual(openWorkbook(sample()).selectionFormatting(0, 'A3:B4'));
  });

  it("refuses a call on a sheet that stopped being the active one", async () => {
    const peer = local(45);
    const peerUpdates: Uint8Array[] = [];
    peer.onUpdate((update) => peerUpdates.push(update));
    const proxy = await worker(46, peer.encodeStateAsUpdate());
    const [first, second] = (await proxy.sheetInfo()).sheetIds;
    await proxy.onSheet(second, 'setActiveSheet', 0);
    expect((await proxy.sheetInfo()).activeSheet).toBe(1);

    peer.applyOps([{ type: 'removeSheet', index: 1 }]);
    await proxy.applyUpdate(peerUpdates.at(-1)!);
    const refused = (promise: Promise<unknown>) => promise.then(() => null, (error) => error);
    expect(await refused(proxy.onSheet(second, 'editCell', 1, 0, 0, 'lost sheet'))).toBeInstanceOf(
      SheetGoneError
    );
    const info = await proxy.sheetInfo();
    expect(info.sheetIds).not.toContain(second);
    expect(info.sheetIds[info.activeSheet]).not.toBe(second);
    expect(await refused(proxy.onSheet(second, 'setActiveSheet', 0))).toBeInstanceOf(
      SheetGoneError
    );
    await proxy.onSheet(first, 'setActiveSheet', 0);
  });

  it('rethrows a stale proposal as StaleProposalError', async () => {
    const proxy = await worker(47);
    if (!proxy.proposalsAvailable) return;
    const sheet = (await proxy.sheetInfo()).sheetIds[0];
    const proposal = await proxy.propose('agent', null, [
      { sheet: 0, row: 2, col: 1, input: '1' },
    ]);
    await proxy.onSheet(sheet, 'editCell', 0, 2, 1, '2');
    const error = await proxy.acceptProposal(proposal.id).catch((e: unknown) => e);
    expect(error).toBeInstanceOf(StaleProposalError);
    expect((error as StaleProposalError).cells).toEqual(['B3']);
  });

  it('fails every request once disposed', async () => {
    const proxy = await worker(48);
    proxy.dispose();
    const error = await proxy.sheetInfo().then(() => null, (e: unknown) => e);
    expect((error as Error).message).toBe('workbook handle is disposed');
  });
});
