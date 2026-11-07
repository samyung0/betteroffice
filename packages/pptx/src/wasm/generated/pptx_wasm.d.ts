/* tslint:disable */
/* eslint-disable */
export interface DecodedFontFace { font: Uint8Array; payload?: Uint8Array; }


export class PptxDocument {
    private constructor();
    free(): void;
    [Symbol.dispose](): void;
    acceptProposalJson(args: string): string;
    addCommentJson(args: string): string;
    addPictureJson(args: string): string;
    addShapeJson(args: string): string;
    addTextBoxJson(args: string): string;
    addTextBoxProfiledJson(args: string): string;
    addUndoBoundary(): void;
    applyUpdateJson(update: Uint8Array): string;
    bringShapeForwardJson(args: string): string;
    bringShapeToFrontJson(args: string): string;
    canRedo(): boolean;
    canUndo(): boolean;
    changeParagraphLevelJson(args: string): string;
    clearTextFormattingJson(args: string): string;
    clearUpdateObservation(): void;
    commentsJson(): string;
    deleteSlideJson(args: string): string;
    deleteTextJson(args: string): string;
    deleteTextProfiledJson(args: string): string;
    drainUpdateEvent(): Uint8Array;
    encodeDiff(remote_state_vector: Uint8Array): Uint8Array;
    encodeStateAsUpdate(): Uint8Array;
    encodeStateVector(): Uint8Array;
    formatTextJson(args: string): string;
    insertParagraphBreakJson(args: string): string;
    insertSlideJson(args: string): string;
    insertSlideProfiledJson(args: string): string;
    insertTextJson(args: string): string;
    insertTextProfiledJson(args: string): string;
    listProposalsJson(): string;
    mediaBytes(part_path: string): Uint8Array;
    /**
     * An EMF or WMF picture's SVG display copy (#1820), replayed once and cached; `undefined`
     * for a part with none.
     */
    metafileSvg(part_path: string): Uint8Array | undefined;
    moveShapeJson(args: string): string;
    moveShapeProfiledJson(args: string): string;
    moveSlideJson(args: string): string;
    static openCollaborative(bytes: Uint8Array, client_id: number): PptxDocument;
    /**
     * `source` must match the update's exact package fingerprint.
     */
    static openCollaborativeFromUpdate(update: Uint8Array, client_id: number, source: Uint8Array): PptxDocument;
    previewProposalJson(args: string): string;
    proposeJson(args: string): string;
    redoJson(): string;
    rejectProposalJson(args: string): string;
    removeCommentJson(args: string): string;
    removeShapeJson(args: string): string;
    replaceTextJson(args: string): string;
    replyToCommentJson(args: string): string;
    resizeShapeJson(args: string): string;
    /**
     * Serializes the deck back to `.pptx` bytes, edits included.
     */
    saveBytes(): Uint8Array;
    searchTextJson(args: string): string;
    sendShapeBackwardJson(args: string): string;
    sendShapeToBackJson(args: string): string;
    setCommentFlavorJson(args: string): string;
    setCommentPositionJson(args: string): string;
    setCommentStatusJson(args: string): string;
    setParagraphAlignmentJson(args: string): string;
    setParagraphListJson(args: string): string;
    setParagraphSpacingJson(args: string): string;
    setShapeAdjustJson(args: string): string;
    setShapeFillJson(args: string): string;
    setShapeRectJson(args: string): string;
    setShapeStrokeJson(args: string): string;
    setSlideNotesJson(args: string): string;
    setTextAnchorJson(args: string): void;
    setUndoCaptureMode(mode: string): void;
    snapshotJson(): string;
    startUpdateObservation(): void;
    storyJson(args: string): string;
    undoCaptureMode(): string;
    undoJson(): string;
    /**
     * `undoJson` timed at its undo, snapshot and serialize boundaries, as
     * `{"receipt": ..., "profile": {"undoMs", "snapshotMs", "serializeMs"}}`.
     */
    undoProfiledJson(): string;
    static version(): string;
    readonly clientId: number;
}

export class PptxRenderer {
    free(): void;
    [Symbol.dispose](): void;
    fontBytes(font_id: number): Uint8Array;
    hitTestJson(x: number, y: number): string;
    hitTestSlideJson(document: PptxDocument, id: string, x: number, y: number): string;
    layoutProposalDiffSlideJson(document: PptxDocument, id: string, slide_index: number): string;
    layoutProposalSlideJson(document: PptxDocument, id: string, slide_index: number): string;
    /**
     * `layoutSlideJson` with the editor's caret in paragraph `paragraph` of
     * `story_id`, whose list marker then shows even while it is empty.
     */
    layoutSlideAtCaretJson(document: PptxDocument, slide_index: number, story_id: string, paragraph: number): string;
    layoutSlideJson(document: PptxDocument, slide_index: number): string;
    /**
     * `layoutSlideJson` with its stages timed, returned as
     * `{"layout": ..., "profile": {"scopeMs", "layoutMs", "serializeMs"}}`.
     */
    layoutSlideProfiledJson(document: PptxDocument, slide_index: number): string;
    constructor();
    /**
     * Drops an embedded face the page could not load from layout.
     */
    refuseEmbeddedFont(font_id: number): void;
    /**
     * Registers the deck's embedded faces, one per part, as
     * `[{"family" (its alias), "typeface", "bold", "italic", "fontId"}]`.
     */
    registerEmbeddedFontsJson(document: PptxDocument): string;
    registerFallbackFont(family: string, bold: boolean, italic: boolean, bytes: Uint8Array): number;
    registerFont(family: string, bold: boolean, italic: boolean, bytes: Uint8Array): number;
    /**
     * The faces layout can ask for, one slide's (by id) or the whole deck's,
     * each with the script buckets its text needs, as
     * `{"faces":[{"family","bold","italic","scripts"?}]}` in first-use order,
     * without laying anything out (upstream #1934, #2027). A superset of what
     * layout requests: a face a shape or chart budget skips may be listed, so
     * a host can load a deck's faces before its first paint.
     */
    requestedFontFacesJson(document: PptxDocument, slide_id?: string | null): string;
    /**
     * Saves `document` with the normal-autofit scales this renderer refit
     * its edited text to (upstream #1526); `PptxDocument.saveBytes` keeps
     * the stored ones.
     */
    saveBytes(document: PptxDocument): Uint8Array;
    setActiveSlide(document: PptxDocument, id: string, key: string): boolean;
    slideLayoutKey(document: PptxDocument, slide_index: number): string;
    snapshotWithLayoutKeysJson(document: PptxDocument): string;
}

export function compileSlideJson(slide_json: string): string;

/**
 * Decodes one font pack face (codec "zstd" or "woff2t+zstd") with this engine's linked decoder.
 * Returns owned bytes and an optional transformed payload. Non-zstd input passes through.
 * The default output limit is 64 MiB; woff2t decoding is always capped at 64 MiB.
 */
export function decodeFontFace(bytes: Uint8Array, codec: string, dict?: Uint8Array | null, maxOutput?: number | null): DecodedFontFace;

export function decodeTiffPng(data: Uint8Array): Uint8Array;

export function parsePptxJson(data: Uint8Array): string;

export function rendererVersion(): string;

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
    readonly __externref_table_alloc: () => number;
    readonly __externref_table_dealloc: (a: number) => void;
    readonly __wbg_pptxdocument_free: (a: number, b: number) => void;
    readonly __wbg_pptxrenderer_free: (a: number, b: number) => void;
    readonly __wbindgen_exn_store: (a: number) => void;
    readonly __wbindgen_externrefs: WebAssembly.Table;
    readonly __wbindgen_free: (a: number, b: number, c: number) => void;
    readonly __wbindgen_malloc: (a: number, b: number) => number;
    readonly __wbindgen_realloc: (a: number, b: number, c: number, d: number) => number;
    readonly __wbindgen_start: () => void;
    readonly betteroffice_memory_grows: () => number;
    readonly compileSlideJson: (a: number, b: number) => [number, number, number, number];
    readonly decodeFontFace: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number) => [number, number, number];
    readonly decodeTiffPng: (a: number, b: number) => [number, number, number, number];
    readonly memory: WebAssembly.Memory;
    readonly parsePptxJson: (a: number, b: number) => [number, number, number, number];
    readonly pptxdocument_acceptProposalJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_addCommentJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_addPictureJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_addShapeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_addTextBoxJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_addTextBoxProfiledJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_addUndoBoundary: (a: number) => void;
    readonly pptxdocument_applyUpdateJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_bringShapeForwardJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_bringShapeToFrontJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_canRedo: (a: number) => number;
    readonly pptxdocument_canUndo: (a: number) => number;
    readonly pptxdocument_changeParagraphLevelJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_clearTextFormattingJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_clearUpdateObservation: (a: number) => void;
    readonly pptxdocument_clientId: (a: number) => number;
    readonly pptxdocument_commentsJson: (a: number) => [number, number, number, number];
    readonly pptxdocument_deleteSlideJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_deleteTextJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_deleteTextProfiledJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_drainUpdateEvent: (a: number) => [number, number];
    readonly pptxdocument_encodeDiff: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_encodeStateAsUpdate: (a: number) => [number, number];
    readonly pptxdocument_encodeStateVector: (a: number) => [number, number];
    readonly pptxdocument_formatTextJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_insertParagraphBreakJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_insertSlideJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_insertSlideProfiledJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_insertTextJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_insertTextProfiledJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_listProposalsJson: (a: number) => [number, number, number, number];
    readonly pptxdocument_mediaBytes: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_metafileSvg: (a: number, b: number, c: number) => [number, number];
    readonly pptxdocument_moveShapeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_moveShapeProfiledJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_moveSlideJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_openCollaborative: (a: number, b: number, c: number) => [number, number, number];
    readonly pptxdocument_openCollaborativeFromUpdate: (a: number, b: number, c: number, d: number, e: number) => [number, number, number];
    readonly pptxdocument_previewProposalJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_proposeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_redoJson: (a: number) => [number, number, number, number];
    readonly pptxdocument_rejectProposalJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_removeCommentJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_removeShapeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_replaceTextJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_replyToCommentJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_resizeShapeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_saveBytes: (a: number) => [number, number, number, number];
    readonly pptxdocument_searchTextJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_sendShapeBackwardJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_sendShapeToBackJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setCommentFlavorJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setCommentPositionJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setCommentStatusJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setParagraphAlignmentJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setParagraphListJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setParagraphSpacingJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setShapeAdjustJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setShapeFillJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setShapeRectJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setShapeStrokeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setSlideNotesJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setTextAnchorJson: (a: number, b: number, c: number) => [number, number];
    readonly pptxdocument_setUndoCaptureMode: (a: number, b: number, c: number) => [number, number];
    readonly pptxdocument_snapshotJson: (a: number) => [number, number, number, number];
    readonly pptxdocument_startUpdateObservation: (a: number) => [number, number];
    readonly pptxdocument_storyJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_undoCaptureMode: (a: number) => [number, number];
    readonly pptxdocument_undoJson: (a: number) => [number, number, number, number];
    readonly pptxdocument_undoProfiledJson: (a: number) => [number, number, number, number];
    readonly pptxdocument_version: () => [number, number];
    readonly pptxrenderer_fontBytes: (a: number, b: number) => [number, number, number, number];
    readonly pptxrenderer_hitTestJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxrenderer_hitTestSlideJson: (a: number, b: number, c: number, d: number, e: number, f: number) => [number, number, number, number];
    readonly pptxrenderer_layoutProposalDiffSlideJson: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly pptxrenderer_layoutProposalSlideJson: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly pptxrenderer_layoutSlideAtCaretJson: (a: number, b: number, c: number, d: number, e: number, f: number) => [number, number, number, number];
    readonly pptxrenderer_layoutSlideJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxrenderer_layoutSlideProfiledJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxrenderer_new: () => number;
    readonly pptxrenderer_refuseEmbeddedFont: (a: number, b: number) => void;
    readonly pptxrenderer_registerEmbeddedFontsJson: (a: number, b: number) => [number, number, number, number];
    readonly pptxrenderer_registerFallbackFont: (a: number, b: number, c: number, d: number, e: number, f: number, g: number) => [number, number, number];
    readonly pptxrenderer_registerFont: (a: number, b: number, c: number, d: number, e: number, f: number, g: number) => [number, number, number];
    readonly pptxrenderer_requestedFontFacesJson: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly pptxrenderer_saveBytes: (a: number, b: number) => [number, number, number, number];
    readonly pptxrenderer_setActiveSlide: (a: number, b: number, c: number, d: number, e: number, f: number) => number;
    readonly pptxrenderer_slideLayoutKey: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxrenderer_snapshotWithLayoutKeysJson: (a: number, b: number) => [number, number, number, number];
    readonly rendererVersion: () => [number, number];
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
