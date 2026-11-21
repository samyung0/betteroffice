import type {
  YrsEngineApplyProfile,
  YrsResidentCaretSnapshot,
  YrsResidentWorkerSnapshot,
  YrsSelection,
} from './index';
import type { ResidentCaretPaintStyle } from './residentCaret';
import type { ResidentFontError } from './residentProgressiveFonts';

export type ResidentEngineWorkerRequest =
  | {
      id: number;
      type: 'bootstrap';
      snapshot: YrsResidentWorkerSnapshot;
      extras: string;
      expectedFrameEpoch: number;
      /** Q791 prototype: reply with the retained layout JSON (the page lays out nothing). */
      wantLayout?: boolean;
    }
  | {
      id: number;
      type: 'sync';
      snapshot: YrsResidentWorkerSnapshot;
      extras: string;
      expectedFrameEpoch: number;
      paintCaret: boolean;
      wantLayout?: boolean;
    }
  | {
      id: number;
      type: 'buildFrame';
      extras: string;
      expectedFrameEpoch: number;
      paintCaret: boolean;
    }
  | {
      /** Lay out the updates already applied (`applyUpdate`) the way a local
       * edit is laid out, and answer with a FrameDelta of the changed pages. */
      id: number;
      type: 'relayout';
      layoutInput: string;
      extras: string;
      expectedFrameEpoch: number;
      selection: YrsSelection | null;
      layoutRevision: number;
      paintCaret: boolean;
      wantLayout?: boolean;
    }
  | {
      id: number;
      type: 'applyInput';
      text: string;
      selection: YrsSelection;
      expectedFrameEpoch: number;
      profile: boolean;
      paintCaret: boolean;
    }
  | {
      id: number;
      type: 'applyDelete';
      direction: 'backward' | 'forward';
      selection: YrsSelection;
      expectedFrameEpoch: number;
      profile: boolean;
      paintCaret: boolean;
    }
  | {
      id: number;
      type: 'applyUpdate';
      update: Uint8Array;
      selection: YrsSelection | null;
    }
  | {
      id: number;
      type: 'attachCanvases';
      pages: Array<{ pageId: string; canvas: OffscreenCanvas }>;
      activePageIds: string[];
      devicePixelRatio: number;
      zoom: number;
      caretStyle: ResidentCaretPaintStyle;
    }
  | { id: number; type: 'eraseCaret' }
  /** Ask again for the font slices that failed and repaint; no reply. */
  | { id: number; type: 'retryFontSlices' }
  | { id: number; type: 'destroy' };

export type ResidentEngineWorkerRequestWithoutId = ResidentEngineWorkerRequest extends infer Request
  ? Request extends { id: number }
    ? Omit<Request, 'id'>
    : never
  : never;

export type ResidentEngineWorkerResponse =
  | {
      id: number;
      ok: true;
      frame?: ArrayBuffer;
      updates?: ArrayBuffer[];
      engineMs?: number;
      workerTotalMs?: number;
      engineProfile?: YrsEngineApplyProfile;
      caret?: YrsResidentCaretSnapshot;
      selection?: YrsSelection | null;
      /** The presented frame carries the worker-painted caret line. */
      caretPainted?: boolean;
      replayMs?: number;
      replayedPages?: number;
      layoutRevision?: number;
      /** The worker replica's yrs state vector after this operation, so the
       * next sync can ship a diff instead of the whole document state. */
      stateVector?: ArrayBuffer;
      /** Sliced faces whose outlines failed in the worker since its last
       * frame reply, by the page's font id; they draw blank there. */
      fontErrors?: ResidentFontError[];
      /** Painting an edit's frame failed; the edit's frame and updates above
       * still apply. */
      paintError?: string;
      /** Q791 prototype: the retained `{ layout, headersFooters?, notesConverged }`
       * JSON, when the request asked for it and laid out. */
      layoutJson?: string;
    }
  | {
      id: number;
      ok: false;
      error: string;
      residentUnavailable?: boolean;
      /** An edit request that committed before it failed: the edit's updates,
       * which the main replica still takes. */
      updates?: ArrayBuffer[];
      /** A wasm trap poisoned the worker; it refuses every later request. */
      terminal?: boolean;
    };

/** The worker began request `id`; the client's timeout for it runs from here. */
export type ResidentEngineWorkerStarted = { id: number; started: true };

/** Everything the worker posts: answers and start notices. */
export type ResidentEngineWorkerMessage = ResidentEngineWorkerResponse | ResidentEngineWorkerStarted;
