/**
 * The main thread's side of a workbook opened in its own worker: an async
 * proxy over the worker's {@link WorkbookHandle}. Requests run in the worker
 * in the order they are posted; the main thread holds no copy of the workbook.
 */

import type {
  EditResult,
  OpenWorkbookOptions,
  WorkbookHandle,
  WorkbookUpdateListener,
} from '../wasm/loader';
import { StaleProposalError } from '../wasm/staleProposal';
import {
  MUTATING_METHODS,
  SHEET_GONE,
  WORKBOOK_METHODS,
  type WorkbookFrame,
  type WorkbookMethod,
  type WorkbookOpened,
  type WorkbookView,
  type WorkbookWorkerMessage,
  type WorkbookWorkerRequest,
} from './protocol';

type Call<K extends WorkbookMethod> = (
  ...args: Parameters<WorkbookHandle[K]>
) => Promise<ReturnType<WorkbookHandle[K]>>;

export type WorkbookCalls = { [K in WorkbookMethod]: Call<K> };

export interface WorkbookProxy extends WorkbookCalls {
  readonly clientId: number;
  readonly proposalsAvailable: boolean;
  readonly pngExportAvailable: boolean;
  /** Requests posted so far: a frame whose `seq` reaches this has seen them all. */
  readonly posted: number;
  /** Changes posted and not answered yet. */
  readonly mutationsInFlight: number;
  /** The workbook's state for `view`, once every request posted before it ran. */
  frame(view: WorkbookView): Promise<WorkbookFrame>;
  /**
   * `method` on the active sheet while it is still `sheetId` (for
   * `setActiveSheet`, the sheet `sheetId` names), else a {@link SheetGoneError}:
   * a peer's change may run first and remove or move it.
   */
  onSheet<K extends WorkbookMethod>(
    sheetId: string,
    method: K,
    ...args: Parameters<WorkbookHandle[K]>
  ): Promise<ReturnType<WorkbookHandle[K]>>;
  /** Flips bold, italic or strikethrough over `range` from its current state. */
  toggle(
    sheetId: string,
    range: string,
    property: 'bold' | 'italic' | 'strikethrough'
  ): Promise<EditResult>;
  onUpdate(listener: WorkbookUpdateListener): () => void;
  /** Ends the worker; pending requests fail. */
  dispose(): void;
}

/** A call refused because its sheet stopped being the active one. */
export class SheetGoneError extends Error {}

/** The worker as the proxy uses it (a test can pass its own). */
export interface WorkbookWorkerPort {
  onmessage: ((event: MessageEvent<WorkbookWorkerMessage>) => void) | null;
  onerror: ((event: ErrorEvent) => void) | null;
  onmessageerror: ((event: MessageEvent) => void) | null;
  postMessage(message: WorkbookWorkerRequest, transfer?: Transferable[]): void;
  terminate(): void;
}

export interface WorkbookWorkerOptions extends OpenWorkbookOptions {
  /** Shared Yrs state applied before the proxy is returned. */
  initialUpdate?: Uint8Array;
  createWorker?: () => WorkbookWorkerPort;
}

function spawnWorkbookWorker(): WorkbookWorkerPort {
  return new Worker(new URL('./workbookWorker.js', import.meta.url), {
    type: 'module',
    name: 'betteroffice-xlsx-workbook',
  });
}

/** Whether this environment can run the workbook in a worker. */
export function isWorkbookWorkerAvailable(): boolean {
  return typeof Worker !== 'undefined';
}

/**
 * Opens a workbook in a dedicated worker. Rejects with the engine's error if
 * the bytes are not a readable workbook or the initial update is refused.
 */
export async function openWorkbookWorker(
  bytes: Uint8Array,
  options: WorkbookWorkerOptions = {}
): Promise<WorkbookProxy> {
  const worker = (options.createWorker ?? spawnWorkbookWorker)();
  const pending = new Map<
    number,
    { resolve: (value: unknown) => void; reject: (error: Error) => void; mutating: boolean }
  >();
  const listeners = new Set<WorkbookUpdateListener>();
  let nextId = 1;
  let failure: Error | null = null;
  let inFlight = 0;

  const fail = (error: Error) => {
    if (failure) return;
    failure = error;
    worker.terminate();
    for (const request of pending.values()) request.reject(error);
    pending.clear();
    inFlight = 0;
  };

  worker.onmessage = (event) => {
    const message = event.data;
    if (message.type === 'update') {
      for (const listener of [...listeners]) {
        try {
          listener(message.update.slice(), message.origin);
        } catch {}
      }
      return;
    }
    const request = pending.get(message.id);
    if (!request) return;
    pending.delete(message.id);
    if (request.mutating) inFlight -= 1;
    if (message.ok) request.resolve(message.value);
    else if (message.stale) request.reject(new StaleProposalError(message.stale));
    else if (message.error === SHEET_GONE) request.reject(new SheetGoneError(message.error));
    else request.reject(new Error(message.error));
  };
  worker.onerror = (event) => fail(new Error(`workbook worker failed: ${event.message}`));
  worker.onmessageerror = () => fail(new Error('workbook worker sent an unreadable message'));

  const send = <T,>(
    request: DistributiveOmit<WorkbookWorkerRequest, 'id'>,
    transfer: Transferable[] = [],
    mutating = false
  ): Promise<T> => {
    if (failure) return Promise.reject(failure);
    const id = nextId++;
    return new Promise<T>((resolve, reject) => {
      pending.set(id, { resolve: resolve as (value: unknown) => void, reject, mutating });
      if (mutating) inFlight += 1;
      worker.postMessage({ ...request, id } as WorkbookWorkerRequest, transfer);
    });
  };

  const call = (method: WorkbookMethod, args: unknown[], sheetId?: string) => {
    const owned = args.map((arg) => (arg instanceof Uint8Array ? arg.slice() : arg));
    const transfer = owned.flatMap((arg) => (arg instanceof Uint8Array ? [arg.buffer] : []));
    return send(
      { type: 'call', method, args: owned, ...(sheetId === undefined ? {} : { sheetId }) },
      transfer,
      MUTATING_METHODS.has(method)
    );
  };

  const owned = bytes.slice();
  const initial = options.initialUpdate?.slice();
  let opened: WorkbookOpened;
  try {
    opened = await send<WorkbookOpened>(
      {
        type: 'open',
        bytes: owned,
        collaborative: options.collaborative ?? false,
        clientId: options.clientId,
        initialUpdate: initial,
      },
      [owned.buffer, ...(initial ? [initial.buffer] : [])]
    );
  } catch (error) {
    worker.terminate();
    throw error;
  }

  const calls = Object.fromEntries(
    WORKBOOK_METHODS.map((method) => [method, (...args: unknown[]) => call(method, args)])
  ) as unknown as WorkbookCalls;

  return {
    ...calls,
    clientId: opened.clientId,
    proposalsAvailable: opened.proposals,
    pngExportAvailable: opened.png,
    get posted() {
      return nextId - 1;
    },
    get mutationsInFlight() {
      return inFlight;
    },
    frame: (view) => send<WorkbookFrame>({ type: 'frame', view }),
    onSheet: (sheetId, method, ...args) =>
      call(method, args, sheetId) as ReturnType<WorkbookProxy['onSheet']>,
    toggle: (sheetId, range, property) =>
      send<EditResult>({ type: 'toggle', sheetId, range, property }, [], true),
    onUpdate(listener) {
      if (typeof listener !== 'function') throw new TypeError('update listener must be a function');
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
    dispose() {
      listeners.clear();
      fail(new Error('workbook handle is disposed'));
    },
  };
}

type DistributiveOmit<T, K extends PropertyKey> = T extends unknown ? Omit<T, K> : never;
