export type OfficeFormat = "docx" | "xlsx" | "pptx";
/** Top-level Yjs roots each engine's state may hold, Capy's contributor map included. */
export declare const OFFICE_DOCUMENT_ROOTS: Readonly<
  Record<OfficeFormat, readonly string[]>
>;
export interface OfficeCheckpoint {
  format: OfficeFormat;
  schemaVersion: 1;
  baseSha256: string;
  state: Uint8Array;
}
export interface OfficeObjectRef {
  format: OfficeFormat;
  kind: "image";
  id: string;
  storyId?: string;
  sheetId?: string;
  slideId?: string;
}
export interface NetEffect {
  id: string;
  kind: "text" | "image" | "visual";
  operation: "add" | "replace" | "remove" | "move";
  label: string;
  before?: string;
  after?: string;
  assetRef?: OfficeObjectRef;
  imageSHA256?: string;
}
export interface OfficeAsset {
  bytes: Uint8Array;
  mimeType: string;
  sha256: string;
}
export interface ExportDeterminism {
  seed: string;
  now: string;
}
export interface OfficeEntry {
  id: string;
  label: string;
  value: string;
  position: string;
}
export type OfficeCommand =
  | {
      type: "replace_text";
      targetId: string;
      expectedText: string;
      text: string;
    }
  | {
      type: "set_cell";
      sheet: string;
      cell: string;
      expectedValue: string;
      value: string;
    };
export interface OfficeTarget {
  id: string;
  path: string[];
  range?: [number, number];
}
export interface OfficeCommandResult {
  state: Uint8Array;
  inverse: OfficeCommand[];
  targets: OfficeTarget[];
}
export type OfficeEditCode =
  | "invalid_input"
  | "stale_target"
  | "unavailable_target"
  | "unsupported_operation";
export declare class OfficeEditError extends Error {
  readonly code: OfficeEditCode;
  constructor(code: OfficeEditCode, message: string);
}
export declare function seedOffice(
  format: OfficeFormat,
  baseBytes: Uint8Array
): Promise<OfficeCheckpoint>;
export declare function compare(
  baseBytes: Uint8Array,
  fromCheckpoint: OfficeCheckpoint,
  toCheckpoint: OfficeCheckpoint
): Promise<NetEffect[]>;
/**
 * Pending XLSX effects read off the checkpoint's overrides; XLSX keeps no
 * stored baseline. With a room (and replicas configured) the call reuses that
 * room's replica when this state holds everything it applied.
 */
export declare function xlsxPendingEffects(
  baseBytes: Uint8Array,
  checkpoint: OfficeCheckpoint,
  room?: string
): Promise<NetEffect[]>;
/**
 * Keeps an XLSX replica per room within a budget of estimated WASM heap
 * bytes; 0 (the initial value) keeps none. A new replica pushes out least
 * recently used ones only once they have been idle for 2 minutes.
 */
export declare function configureOfficeReplicas(budgetBytes: number): void;
/** Drops a room's replica. */
export declare function dropOfficeReplica(room: string): void;
export declare function officeReplicaStats(): {
  replicas: number;
  replicaBytes: number;
  hits: number;
  misses: number;
  evictions: number;
  /** The XLSX engine's linear memory. */
  wasmBytes: number;
};
export declare function resolveAsset(
  baseBytes: Uint8Array,
  checkpoint: OfficeCheckpoint,
  objectRef: OfficeObjectRef
): Promise<OfficeAsset>;
export declare function exportOffice(
  baseBytes: Uint8Array,
  checkpoint: OfficeCheckpoint,
  determinism: ExportDeterminism
): Promise<Uint8Array>;
export declare function inspectOffice(
  baseBytes: Uint8Array,
  checkpoint: OfficeCheckpoint
): Promise<OfficeEntry[]>;
export declare function applyOfficeCommands(
  baseBytes: Uint8Array,
  checkpoint: OfficeCheckpoint,
  commands: OfficeCommand[]
): Promise<OfficeCommandResult>;
export declare function locateOfficeTargets(
  baseBytes: Uint8Array,
  checkpoint: OfficeCheckpoint,
  ids: string[]
): Promise<OfficeTarget[]>;
export declare function runtimeManifest(): Promise<Record<string, string>>;

export interface OfficeBaselineEntry extends OfficeEntry {
  imageSHA256?: string;
  kind: NetEffect["kind"];
  assetRef?: OfficeObjectRef;
}
export declare function officeBaseline(
  baseBytes: Uint8Array,
  checkpoint: OfficeCheckpoint,
): Promise<OfficeBaselineEntry[]>;
export declare function compareBaselines(
  from: OfficeBaselineEntry[],
  to: OfficeBaselineEntry[],
): NetEffect[];

/**
 * The edits saved after the capture landed on seed(export), with their effects
 * against the export. A refusal is an Error whose message starts with
 * "Office rebase:"; the same inputs are always refused again.
 */
export declare function rebaseOffice(
  baseBytes: Uint8Array,
  captured: OfficeCheckpoint,
  latest: OfficeCheckpoint,
  exportedSource: Uint8Array,
): Promise<{
  state: Uint8Array;
  effects: NetEffect[];
}>;
