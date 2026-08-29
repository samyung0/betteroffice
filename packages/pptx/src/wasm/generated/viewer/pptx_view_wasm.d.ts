/* tslint:disable */
/* eslint-disable */

export class PptxViewDocument {
    private constructor();
    free(): void;
    [Symbol.dispose](): void;
    mediaBytes(part_path: string): Uint8Array;
    static open(bytes: Uint8Array): PptxViewDocument;
    snapshotJson(): string;
    static version(): string;
}

export class PptxViewRenderer {
    free(): void;
    [Symbol.dispose](): void;
    hitTestJson(x: number, y: number): string;
    layoutSlideJson(document: PptxViewDocument, slide_index: number): string;
    constructor();
    registerFont(family: string, bold: boolean, italic: boolean, bytes: Uint8Array): number;
}

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
    readonly memory: WebAssembly.Memory;
    readonly __wbg_pptxviewdocument_free: (a: number, b: number) => void;
    readonly __wbg_pptxviewrenderer_free: (a: number, b: number) => void;
    readonly pptxviewdocument_mediaBytes: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxviewdocument_open: (a: number, b: number) => [number, number, number];
    readonly pptxviewdocument_snapshotJson: (a: number) => [number, number, number, number];
    readonly pptxviewdocument_version: () => [number, number];
    readonly pptxviewrenderer_hitTestJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxviewrenderer_layoutSlideJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxviewrenderer_new: () => number;
    readonly pptxviewrenderer_registerFont: (a: number, b: number, c: number, d: number, e: number, f: number, g: number) => [number, number, number];
    readonly __wbindgen_externrefs: WebAssembly.Table;
    readonly __wbindgen_malloc: (a: number, b: number) => number;
    readonly __wbindgen_realloc: (a: number, b: number, c: number, d: number) => number;
    readonly __externref_table_dealloc: (a: number) => void;
    readonly __wbindgen_free: (a: number, b: number, c: number) => void;
    readonly __wbindgen_start: () => void;
}

export type SyncInitInput = BufferSource | WebAssembly.Module;

/**
 * Instantiates the given `module`, which can either be bytes or
 * a precompiled `WebAssembly.Module`.
 *
 * @param {{ module: SyncInitInput }} module - Passing `SyncInitInput` directly is deprecated.
 *
 * @returns {InitOutput}
 */
export function initSync(module: { module: SyncInitInput } | SyncInitInput): InitOutput;

/**
 * If `module_or_path` is {RequestInfo} or {URL}, makes a request and
 * for everything else, calls `WebAssembly.instantiate` directly.
 *
 * @param {{ module_or_path: InitInput | Promise<InitInput> }} module_or_path - Passing `InitInput` directly is deprecated.
 *
 * @returns {Promise<InitOutput>}
 */
export default function __wbg_init (module_or_path?: { module_or_path: InitInput | Promise<InitInput> } | InitInput | Promise<InitInput>): Promise<InitOutput>;
