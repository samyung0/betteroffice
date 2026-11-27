/// <reference lib="webworker" />

import type { YrsResidentCaretRect, YrsResidentWorkerSnapshot } from './index';
import {
  createResidentEngineSession,
  type ResidentEngineSession,
} from './residentEngineSession';
import {
  presentOffscreenPageBackBuffer,
  presentOffscreenPageBackBufferWithCaret,
  rasterizeDisplayPageToBackBuffer,
  type DrawPageOptions,
} from '../layout/render/canvasBackend';
import {
  applyFrameDeltaOwned,
  decodeFrameDelta,
  type RetainedFrame,
} from '../layout/render/frameDelta';
import { GlyphCache } from '../layout/render/glyphCache';
import type {
  ResidentEngineWorkerMessage,
  ResidentEngineWorkerRequest,
  ResidentFirstPagesViewport,
} from './residentEngineWorkerProtocol';
import {
  residentCaretDeviceRect,
  residentCaretSnapshotForFrame,
  type ResidentCaretPaintStyle,
} from './residentCaret';
import {
  FontSliceUnavailableError,
  glyphsByFont,
  ResidentProgressiveFonts,
  sliceFaceOpener,
  type ResidentFontError,
} from './residentProgressiveFonts';
import { decodeFontFace } from '../wasm/fontDecoder';

const scope = self as unknown as DedicatedWorkerGlobalScope;
let session: ResidentEngineSession | null = null;
let unsubscribe: (() => void) | null = null;
let pendingUpdates: Uint8Array[] = [];
let layoutRevision = 0;
// The input the retained layout was laid out from: the region request and the
// replica's document epoch, which moves on with every change it applies or
// makes. A relayout of the same input lays out nothing (a web-font load's
// pass: measuring reads font files in Rust, so the load changes neither).
let laidOut: { input: string; docEpoch: number } | null = null;
// -1 = no fonts applied yet (fresh session); hydrate skips re-registration
// when the snapshot's revision matches what this session already holds.
let fontsRevision = -1;
let operations = Promise.resolve();
let retainedFrame: RetainedFrame | null = null;
let glyphCache: GlyphCache | null = null;
const offscreenCanvases = new Map<string, OffscreenCanvas>();
const offscreenBackBuffers = new Map<string, OffscreenCanvas>();
const pendingOffscreenPageIds = new Set<string>();
let activeOffscreenPageIds = new Set<string>();
let offscreenDpr = 1;
let offscreenZoom = 1;
// Present-time caret painting. `caretPaintRect` is what the current frame
// wants painted, `paintedCaretPageId/Key` what is on screen. Caret-composited
// pages keep their back-buffer raster (`intactBackBuffers`) so the line can be
// dropped by re-presenting without raster or engine work; plain presents
// detach their bitmap and stay zero-copy.
let caretStyle: ResidentCaretPaintStyle = { color: '#000', width: 2 };
let caretPaintRect: YrsResidentCaretRect | null = null;
let paintedCaretPageId: string | null = null;
let paintedCaretKey: string | null = null;
let caretStage: OffscreenCanvas | null = null;
const intactBackBuffers = new Set<string>();
let trap: WebAssembly.RuntimeError | null = null;
// The fonts registered from sliced faces' skeletons, whose outlines this
// worker loads and installs before it paints the glyphs. A face whose
// outlines fail draws blank and fails no page; the next frame reply carries
// the failure, by the page's font id, for the page to report.
let progressiveFonts: ResidentProgressiveFonts | null = null;
let fontErrors: ResidentFontError[] = [];
// The measurement's font chains from the latest frame request, by which list
// markers and watermarks name their sliced faces.
let fontChains: Record<string, number[]> = {};
// The retained layout JSON this request's frame reply carries, when the page
// recorded the pass without laying it out (`wantLayout`).
let replyLayoutJson: string | undefined;
// The replica's source package has media: pages not built yet may hold
// pictures, which only the page can draw (Q815 C).
let hasMedia = false;

scope.onmessage = (event: MessageEvent<ResidentEngineWorkerRequest>) => {
  operations = operations
    .then(() => {
      if (trap) throw trap;
      // Requests queue behind each other here; the client's timeout for this
      // one starts with this notice, not when it was posted.
      reply({ id: event.data.id, started: true });
      return handle(event.data);
    })
    .catch((error) => {
      if (error instanceof WebAssembly.RuntimeError) {
        trap = error;
        reply({
          id: event.data.id,
          ok: false,
          error: `Resident engine worker trapped: ${error.message}`,
          terminal: true,
        });
        return;
      }
      reply({
        id: event.data.id,
        ok: false,
        error: error instanceof Error ? error.message : String(error),
      });
    });
};

async function handle(request: ResidentEngineWorkerRequest): Promise<void> {
  replyLayoutJson = undefined;
  if ('extras' in request) fontChains = chainsOf(request.extras);
  if (request.type === 'destroy') {
    destroySession();
    return;
  }
  if (request.type === 'seed') {
    destroySession();
    session = await createResidentEngineSession(request.seed.workerClientId);
    try {
      session.setMedia(request.seed.media);
      hasMedia = request.seed.media !== '{}';
      session.loadState(request.seed.state);
    } catch (error) {
      // The bootstrap then starts a replica of its own from the full state.
      destroySession();
      throw error;
    }
    const stateVector = exactBuffer(session.encodeStateVector());
    reply({ id: request.id, ok: true, stateVector }, [stateVector]);
    return;
  }
  if (request.type === 'bootstrap') {
    // The worker is a genuine yrs peer. Reusing the main replica's client id
    // makes a fast structural input race overlap one client's clock range and
    // corrupt the update; an id of its own (the host's assigned worker id, or
    // a random one) lets yrs merge queued/local operations safely while the
    // main replica applies worker updates with local origin. A worker that
    // replaces another under the same assigned id hydrates from the main
    // replica, which holds every worker update it applied, and the replaced
    // worker's later replies are dropped, so no clock is taken twice. A
    // seeded replica keeps its state: the snapshot holds what changed since.
    session ??= await createResidentEngineSession(request.snapshot.workerClientId);
    const layoutJson = hydrate(request.snapshot);
    if (request.wantLayout) replyLayoutJson = layoutJson;
    subscribe();
    const started = performance.now();
    // The pages in view first; the page asks for the rest (Q815 C).
    const first = request.firstPages
      ? firstPagesFor(session.layoutPageHeights(), request.firstPages)
      : 0;
    const frame = first
      ? session.buildDisplayListFrameFirst(request.extras, request.expectedFrameEpoch, first)
      : session.buildDisplayListFrame(request.extras, request.expectedFrameEpoch);
    await replyFrame(
      request.id,
      frame,
      performance.now() - started,
      pendingUpdates,
      undefined,
      started
    );
    return;
  }
  if (request.type === 'retryFontSlices') {
    // The host is back online: ask again for the slices that failed, repainting every page.
    progressiveFonts?.retryUnavailable();
    glyphCache = null;
    intactBackBuffers.clear();
    await replayOffscreen(true);
    return;
  }
  if (request.type === 'eraseCaret') {
    caretPaintRect = null;
    if (paintedCaretPageId !== null) await replayOffscreen(false);
    reply({ id: request.id, ok: true });
    return;
  }
  if (!session) throw new Error('Resident engine worker is not initialized');
  if (request.type === 'sync') {
    unsubscribe?.();
    unsubscribe = null;
    const layoutJson = hydrate(request.snapshot);
    if (request.wantLayout) replyLayoutJson = layoutJson;
    subscribe();
    const started = performance.now();
    const frame = session.buildDisplayListFrame(request.extras, request.expectedFrameEpoch);
    await replyFrame(
      request.id,
      frame,
      performance.now() - started,
      pendingUpdates,
      undefined,
      started,
      false,
      request.paintCaret
    );
    return;
  }
  if (request.type === 'buildPages') {
    pendingUpdates = [];
    const started = performance.now();
    const frame = session.buildUnbuiltPagesFrame(
      Uint32Array.from(request.pages),
      request.expectedFrameEpoch
    );
    await replyFrame(
      request.id,
      frame,
      performance.now() - started,
      pendingUpdates,
      undefined,
      started,
      false,
      request.paintCaret
    );
    return;
  }
  if (request.type === 'buildFrame') {
    pendingUpdates = [];
    const started = performance.now();
    const frame = session.buildDisplayListFrame(request.extras, request.expectedFrameEpoch);
    await replyFrame(
      request.id,
      frame,
      performance.now() - started,
      pendingUpdates,
      undefined,
      started,
      false,
      request.paintCaret
    );
    return;
  }
  if (request.type === 'relayout') {
    if (request.selection) session.setSelection(request.selection.anchor, request.selection.head);
    pendingUpdates = [];
    const started = performance.now();
    const unchanged =
      laidOut?.input === request.layoutInput && laidOut.docEpoch === session.docEpoch();
    laidOut = null;
    const frame = unchanged
      ? session.buildDisplayListFrame(request.extras, request.expectedFrameEpoch)
      : session.relayoutFrame(request.layoutInput, request.extras, request.expectedFrameEpoch);
    laidOut = { input: request.layoutInput, docEpoch: session.docEpoch() };
    layoutRevision = request.layoutRevision;
    if (request.wantLayout) replyLayoutJson = session.retainedRegionLayoutJson();
    await replyFrame(
      request.id,
      frame,
      performance.now() - started,
      pendingUpdates,
      undefined,
      started,
      false,
      request.paintCaret
    );
    return;
  }
  if (request.type === 'applyUpdate') {
    session.applyUpdate(request.update);
    if (request.selection) session.setSelection(request.selection.anchor, request.selection.head);
    return;
  }
  if (request.type === 'attachCanvases') {
    const environmentChanged =
      offscreenCanvases.size > 0 &&
      (offscreenDpr !== request.devicePixelRatio || offscreenZoom !== request.zoom);
    // Pages needing pixels: freshly transferred canvases plus retained
    // canvases re-entering the active window (their buffers were zeroed when
    // they left it).
    const forcedPageIds = new Set(request.pages.map((page) => page.pageId));
    for (const pageId of request.activePageIds) {
      if (!activeOffscreenPageIds.has(pageId)) forcedPageIds.add(pageId);
    }
    for (const { pageId, canvas } of request.pages) offscreenCanvases.set(pageId, canvas);
    activeOffscreenPageIds = new Set(request.activePageIds);
    offscreenDpr = request.devicePixelRatio;
    offscreenZoom = request.zoom;
    caretStyle = request.caretStyle;
    for (const [pageId, canvas] of offscreenCanvases) {
      if (!activeOffscreenPageIds.has(pageId)) {
        // out of the page window: release the bitmap but KEEP the canvas —
        // a transferred surface can never be re-transferred, so the element
        // must stay usable for re-entry
        canvas.width = 0;
        canvas.height = 0;
        offscreenBackBuffers.delete(pageId);
        forgetOffscreenPagePixels(pageId);
      }
    }
    // A dpr/zoom change repaints every active page; otherwise only the forced
    // set needs pixels — surviving active pages keep their bitmaps.
    await replayOffscreen(environmentChanged ? true : forcedPageIds);
    reply({ id: request.id, ok: true });
    return;
  }
  session.setSelection(request.selection.anchor, request.selection.head);
  pendingUpdates = [];
  const started = performance.now();
  try {
    const applied =
      request.type === 'applyDelete'
        ? request.profile
          ? session.applyDeleteProfiled(request.direction, request.expectedFrameEpoch)
          : {
              frame: session.applyDelete(request.direction, request.expectedFrameEpoch),
              profile: undefined,
            }
        : request.profile
          ? session.applyInputProfiled(request.text, request.expectedFrameEpoch)
          : {
              frame: session.applyInput(request.text, request.expectedFrameEpoch),
              profile: undefined,
            };
    // The edit laid itself out against the retained request.
    if (laidOut) laidOut.docEpoch = session.docEpoch();
    await replyFrame(
      request.id,
      applied.frame,
      performance.now() - started,
      pendingUpdates,
      applied.profile,
      started,
      true,
      request.paintCaret,
      true
    );
  } catch (error) {
    if (error instanceof WebAssembly.RuntimeError) throw error;
    const message = error instanceof Error ? error.message : String(error);
    // An edit the replica committed before the failure (a caret snapshot that
    // does not match its frame, say) still goes back, or it would be lost.
    const updates = pendingUpdates.map(exactBuffer);
    reply(
      {
        id: request.id,
        ok: false,
        error: message,
        residentUnavailable: message.includes('resident input state is not ready'),
        ...(updates.length > 0 ? { updates } : {}),
      },
      updates
    );
  } finally {
    pendingUpdates = [];
  }
}

function hydrate(snapshot: YrsResidentWorkerSnapshot): string | undefined {
  if (!session) throw new Error('Resident engine worker is not initialized');
  if (snapshot.media !== undefined) {
    session.setMedia(snapshot.media);
    hasMedia = snapshot.media !== '{}';
  }
  session.loadState(snapshot.state);
  if (snapshot.fontsRevision !== fontsRevision) {
    // A mismatched revision always carries the full font set (the client only
    // omits fonts when it knows this session's applied revision matches).
    session.clearFonts();
    const ids = snapshot.fonts.map((font) => font instanceof Uint8Array
      ? session!.registerFont(font)
      : session!.registerSubstituteFont(font.substituteOf, font.family));
    setProgressiveFonts(openProgressiveFonts(snapshot, ids));
    fontsRevision = snapshot.fontsRevision;
  }
  for (const { story, env } of snapshot.renderInputs) session.yrsBlocksForStory(story, env);
  for (const input of snapshot.measureInputs) session.measureParagraphJson(input);
  let layoutJson: string | undefined;
  if (snapshot.layoutWithRegions) {
    // the retained reply omits the tens-of-MB measured arena; it is the
    // page's layout when the page recorded the pass
    layoutJson = session.layoutDocumentWithRegionsRetainedJson(snapshot.layoutInput);
  } else {
    session.layoutDocumentJson(snapshot.layoutInput);
  }
  laidOut = snapshot.layoutWithRegions
    ? { input: snapshot.layoutInput, docEpoch: session.docEpoch() }
    : null;
  if (snapshot.selection) session.setSelection(snapshot.selection.anchor, snapshot.selection.head);
  layoutRevision = snapshot.layoutRevision;
  pendingUpdates = [];
  return layoutJson;
}

/** How many pages from the top fill `viewport`, plus one (as DOCX View's
 * first pages); 0 when that is every page. */
function firstPagesFor(heights: Float64Array, { height, zoom }: ResidentFirstPagesViewport): number {
  let shown = 0;
  for (let top = 0; shown < heights.length && top < height; shown += 1) top += heights[shown]! * zoom;
  return shown + 1 < heights.length ? shown + 1 : 0;
}

function setProgressiveFonts(next: ResidentProgressiveFonts | null): void {
  progressiveFonts?.dispose();
  progressiveFonts = next;
  glyphCache = null;
}

/** The snapshot's sliced faces, by the ids this session registered them under. */
function openProgressiveFonts(
  snapshot: YrsResidentWorkerSnapshot,
  ids: number[]
): ResidentProgressiveFonts | null {
  const owner = session;
  if (!owner || !snapshot.progressiveFonts?.length) return null;
  const fonts = new ResidentProgressiveFonts({
    updateFontOutlines: (fontId, bytes) => owner.updateFontOutlines(fontId, bytes),
  });
  const decode = async (encoded: ArrayBuffer, maxOutput: number) =>
    (await decodeFontFace(encoded, { codec: 'zstd', maxOutput })).font;
  for (const { slot, source, files, failed } of snapshot.progressiveFonts) {
    fonts.add(ids[slot] ?? slot, sliceFaceOpener(source, decode, files, failed), (error) => {
      fontErrors.push(
        error instanceof FontSliceUnavailableError
          ? { slot, message: error.message, unavailable: { file: error.file, reason: error.reason, ...(error.status === undefined ? {} : { status: error.status }) } }
          : { slot, message: error.message }
      );
    });
  }
  snapshot.fonts.forEach((font, index) => {
    if (!(font instanceof Uint8Array)) fonts.addView(ids[index]!, ids[font.substituteOf] ?? font.substituteOf);
  });
  // Outlines a frame stopped waiting for repaint every page once they land.
  fonts.onLateInstall(() => {
    if (progressiveFonts !== fonts) return;
    glyphCache = null;
    intactBackBuffers.clear();
    operations = operations.then(() => replayOffscreen(true)).then(() => {}, () => {});
  });
  return fonts;
}

/**
 * Readies `page`'s sliced faces, waiting for them up to the guard: installs
 * the outlines its glyph runs draw and registers the faces its list markers
 * and watermarks name. Resolves with the painter's options for them.
 */
async function slicedPaint(page: unknown): Promise<Pick<DrawPageOptions, 'outlineOnlyFontIds' | 'fontFamilies'>> {
  const fonts = progressiveFonts;
  if (!fonts) return {};
  const [changed, fontFamilies] = await Promise.all([
    fonts.ensure(glyphsByFont(page)),
    fonts.textFamilies(page, fontChains),
  ]);
  if (changed.size > 0) glyphCache?.invalidateFonts(changed);
  return { outlineOnlyFontIds: fonts.fontIds(), fontFamilies };
}

function chainsOf(extras: string): Record<string, number[]> {
  return extras ? (JSON.parse(extras) as { fontChains?: Record<string, number[]> }).fontChains ?? {} : {};
}

function subscribe(): void {
  if (!session) return;
  unsubscribe = session.onUpdate((update) => pendingUpdates.push(update.slice()));
}

function destroySession(): void {
  unsubscribe?.();
  unsubscribe = null;
  session?.destroy();
  session = null;
  pendingUpdates = [];
  layoutRevision = 0;
  laidOut = null;
  hasMedia = false;
  fontsRevision = -1;
  retainedFrame = null;
  setProgressiveFonts(null);
  fontErrors = [];
  offscreenCanvases.clear();
  offscreenBackBuffers.clear();
  pendingOffscreenPageIds.clear();
  activeOffscreenPageIds.clear();
  caretPaintRect = null;
  paintedCaretPageId = null;
  paintedCaretKey = null;
  caretStage = null;
  intactBackBuffers.clear();
}

function forgetOffscreenPagePixels(pageId: string): void {
  intactBackBuffers.delete(pageId);
  if (paintedCaretPageId === pageId) {
    paintedCaretPageId = null;
    paintedCaretKey = null;
  }
}

async function replyFrame(
  id: number,
  bytes: Uint8Array,
  engineMs: number,
  updates = pendingUpdates,
  engineProfile?: import('./index').YrsEngineApplyProfile,
  requestStarted = performance.now(),
  requireCaret = false,
  paintCaret = false,
  // An edit request: its edit is committed, so a paint failure travels in the
  // reply beside the frame and updates instead of discarding them.
  committed = false
): Promise<void> {
  retainedFrame = applyFrameDeltaOwned(retainedFrame, decodeFrameDelta(bytes));
  for (const pageId of retainedFrame.damagedPageIds) pendingOffscreenPageIds.add(pageId.toString());
  // The decoder's primitive-id arrays are zero-copy views into `bytes`. The
  // FrameDelta buffer is transferred to the main thread below, so retain only
  // these compact identity arrays in worker-owned memory before detaching it.
  retainedFrame = {
    ...retainedFrame,
    pages: retainedFrame.pages.map((page) => ({
      ...page,
      primitiveIds: page.primitiveIds.slice(),
    })),
  };
  const caret = session?.residentCaretSnapshot();
  if (!caret || !residentCaretSnapshotForFrame(caret, retainedFrame)) {
    throw new Error('Resident caret snapshot does not match the produced frame');
  }
  if (requireCaret && !caret.caretRect) {
    throw new Error('Resident input frame omitted collapsed caret geometry');
  }
  const selection = session?.selection() ?? null;
  caretPaintRect = paintCaret ? (caret.caretRect ?? null) : null;
  // Pages no longer in the document release their surfaces entirely (their
  // elements unmounted main-side); off-window pages are only zeroed, so this
  // is the sole place a live document's canvas reference is dropped.
  const livePageIds = new Set(retainedFrame.pages.map((page) => page.pageId.toString()));
  for (const pageId of pendingOffscreenPageIds) {
    if (!livePageIds.has(pageId)) pendingOffscreenPageIds.delete(pageId);
  }
  for (const pageId of offscreenCanvases.keys()) {
    if (!livePageIds.has(pageId)) {
      offscreenCanvases.delete(pageId);
      offscreenBackBuffers.delete(pageId);
      forgetOffscreenPagePixels(pageId);
    }
  }
  const replayStarted = performance.now();
  let replayed = { replayedPages: 0, caretPainted: false };
  let paintError: string | undefined;
  try {
    replayed = await replayOffscreen(false);
  } catch (error) {
    if (!committed || error instanceof WebAssembly.RuntimeError) throw error;
    paintError = error instanceof Error ? error.message : String(error);
  }
  const { replayedPages, caretPainted } = replayed;
  const replayMs = performance.now() - replayStarted;
  const failedFonts = fontErrors;
  fontErrors = [];
  const layoutJson = replyLayoutJson;
  replyLayoutJson = undefined;
  let unbuiltPages = 0;
  for (const page of retainedFrame?.displayList.pages ?? []) if (page.unbuilt) unbuiltPages += 1;
  const frame = exactBuffer(bytes);
  const updateBuffers = updates.map(exactBuffer);
  const stateVector = session ? exactBuffer(session.encodeStateVector()) : undefined;
  reply(
    {
      id,
      ok: true,
      frame,
      updates: updateBuffers,
      engineMs,
      workerTotalMs: performance.now() - requestStarted,
      engineProfile,
      caret,
      selection,
      caretPainted,
      replayMs,
      replayedPages,
      layoutRevision,
      ...(stateVector ? { stateVector } : {}),
      ...(failedFonts.length > 0 ? { fontErrors: failedFonts } : {}),
      ...(paintError === undefined ? {} : { paintError }),
      ...(layoutJson === undefined ? {} : { layoutJson }),
      ...(unbuiltPages === 0 ? {} : { unbuiltPages, media: hasMedia }),
    },
    [frame, ...updateBuffers, ...(stateVector ? [stateVector] : [])]
  );
}

async function replayOffscreen(
  force: boolean | Set<string>
): Promise<{ replayedPages: number; caretPainted: boolean }> {
  const forcedPageIds = force === true ? activeOffscreenPageIds : force;
  if (forcedPageIds) {
    for (const pageId of forcedPageIds) pendingOffscreenPageIds.add(pageId);
  }
  if (!retainedFrame || offscreenCanvases.size === 0) {
    return { replayedPages: 0, caretPainted: false };
  }
  if (!glyphCache && session) {
    glyphCache = new GlyphCache({
      provider: (fontId, glyphId) => session!.outlineGlyphJson(fontId, glyphId),
    });
  }
  const caretTarget =
    caretPaintRect && activeOffscreenPageIds.has(caretPaintRect.pageId) ? caretPaintRect : null;
  const caretDevice = caretTarget
    ? residentCaretDeviceRect(caretTarget, caretStyle, offscreenDpr, offscreenZoom)
    : null;
  const caretKey =
    caretTarget && caretDevice
      ? `${caretTarget.pageId}|${caretDevice.x}|${caretDevice.y}|${caretDevice.width}|${caretDevice.height}|${caretStyle.color}`
      : null;
  const preparations: Array<
    Promise<{
      canvas: OffscreenCanvas;
      buffer: OffscreenCanvas;
      pageId: string;
    }>
  > = [];
  for (let index = 0; index < retainedFrame.pages.length; index += 1) {
    const retainedPage = retainedFrame.pages[index];
    const pageIdString = retainedPage.pageId.toString();
    // Off-window pages hold no pixels; they re-raster through the forced set
    // when they re-enter the window.
    if (!activeOffscreenPageIds.has(pageIdString)) continue;
    const damaged = pendingOffscreenPageIds.has(pageIdString);
    // Beyond damage, a page presents only for caret compositing: the page
    // gaining the painted line and the page losing it.
    const gainsCaret =
      pageIdString === caretTarget?.pageId &&
      !(paintedCaretPageId === pageIdString && paintedCaretKey === caretKey);
    const losesCaret =
      pageIdString === paintedCaretPageId && pageIdString !== caretTarget?.pageId;
    if (!damaged && !gainsCaret && !losesCaret) continue;
    const canvas = offscreenCanvases.get(pageIdString);
    const page = retainedFrame.displayList.pages[index];
    if (!canvas || !page) continue;
    const pageId = pageIdString;
    let buffer = offscreenBackBuffers.get(pageId);
    if (!buffer) {
      buffer = new OffscreenCanvas(1, 1);
      offscreenBackBuffers.set(pageId, buffer);
    }
    const resolvedBuffer = buffer;
    if (!damaged && intactBackBuffers.has(pageId)) {
      preparations.push(Promise.resolve({ canvas, buffer: resolvedBuffer, pageId }));
      continue;
    }
    preparations.push(
      slicedPaint(page)
        .then((sliced) =>
          rasterizeDisplayPageToBackBuffer(
            resolvedBuffer,
            page,
            { glyphCache: glyphCache ?? undefined, ...sliced },
            offscreenDpr,
            offscreenZoom
          )
        )
        .then(() => ({ canvas, buffer: resolvedBuffer, pageId }))
    );
  }
  const prepared = await Promise.all(preparations).catch(async (error) => {
    await Promise.allSettled(preparations);
    throw error;
  });
  let caretPainted =
    caretTarget !== null && paintedCaretPageId === caretTarget.pageId && paintedCaretKey === caretKey;
  // Present only after the entire damaged frame is ready. This loop is
  // synchronous, so the compositor can never observe the renderer's clears.
  for (const { canvas, buffer, pageId } of prepared) {
    if (caretTarget && caretDevice && pageId === caretTarget.pageId) {
      if (!caretStage) caretStage = new OffscreenCanvas(1, 1);
      presentOffscreenPageBackBufferWithCaret(canvas, buffer, caretStage, {
        ...caretDevice,
        color: caretStyle.color,
      });
      intactBackBuffers.add(pageId);
      paintedCaretPageId = pageId;
      paintedCaretKey = caretKey;
      caretPainted = true;
    } else {
      presentOffscreenPageBackBuffer(canvas, buffer);
      intactBackBuffers.delete(pageId);
      if (paintedCaretPageId === pageId) {
        paintedCaretPageId = null;
        paintedCaretKey = null;
      }
    }
    pendingOffscreenPageIds.delete(pageId);
  }
  return { replayedPages: prepared.length, caretPainted };
}

function exactBuffer(bytes: Uint8Array): ArrayBuffer {
  if (bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength) {
    return bytes.buffer as ArrayBuffer;
  }
  return bytes.slice().buffer;
}

function reply(response: ResidentEngineWorkerMessage, transfer: Transferable[] = []): void {
  scope.postMessage(response, transfer);
}
