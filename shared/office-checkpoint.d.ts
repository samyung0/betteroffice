export type OfficeFormat = "docx" | "xlsx" | "pptx";
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
export declare function runtimeManifest(): Promise<Record<string, string>>;
