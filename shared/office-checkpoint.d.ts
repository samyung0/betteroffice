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
export declare function runtimeManifest(): Promise<Record<string, string>>;
