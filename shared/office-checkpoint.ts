import { createHash, randomInt } from "node:crypto";
import { readFile } from "node:fs/promises";
import { createYrsSession, type YrsSession } from "../packages/docx/src/yrs";
import { yrsToDocument } from "../packages/docx/src/yrs/yrsToDocument";
import { preloadEditWasm } from "../packages/docx/src/wasm/edit";
import { preloadParseWasm } from "../packages/docx/src/wasm/parse";
import { preloadOpcWasm, unzipContainer } from "../packages/docx/src/wasm/opc";
import { writeDocumentWithRust } from "../packages/docx/src/docx/rustSaveFacade";
import { parseRelationshipsXmlWithRust } from "../packages/docx/src/docx/rustParseFacade";
import {
  DOCX_LINEAGE,
  PPTX_LINEAGE,
  RebaseError,
  assertPptxRestorations,
  docxIds,
  docxWrittenParagraphs,
  pptxIds,
  transplant,
} from "./office-rebase";
import initXlsx, {
  XlsxDocument,
  XlsxEffectsReader,
} from "../packages/xlsx/src/wasm/generated/xlsx_wasm.js";
import initPptx, {
  PptxDocument,
} from "../packages/pptx/src/wasm/generated/pptx_wasm.js";
import type {
  DeckSnapshot,
  ShapeSnapshot,
  StorySnapshot,
} from "../packages/pptx/src/types";

export type OfficeFormat = "docx" | "xlsx" | "pptx";
/**
 * Top-level Yjs roots each engine's state may hold, Capy's contributor map
 * included. PPTX and XLSX engines reject any other root.
 */
export const OFFICE_DOCUMENT_ROOTS: Readonly<
  Record<OfficeFormat, readonly string[]>
> = {
  docx: ["stories", "comments", "bookmarks", "__capy_pending_contributors"],
  xlsx: [
    "xlsx",
    "xlsx:cell-formats",
    "xlsx:sheet-order",
    "xlsx:sheets",
    "xlsx:axis-catalog",
    "xlsx:defined-names",
    "__capy_pending_contributors",
  ],
  pptx: [
    "pptx:meta",
    "pptx:slide-order",
    "pptx:slides",
    "pptx:shapes",
    "pptx:stories",
    "pptx:comments",
    "__capy_pending_contributors",
  ],
};
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
/** Yrs location of one edited target: root map name, nested keys, and the text range for stories. */
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
/** The code prefixes the message so it survives worker boundaries that carry strings only. */
export class OfficeEditError extends Error {
  constructor(
    readonly code: OfficeEditCode,
    message: string
  ) {
    super(`${code}: ${message}`);
  }
}
export interface OfficeBaselineEntry {
  imageSHA256?: string;
  id: string;
  kind: NetEffect["kind"];
  label: string;
  value: string;
  position: string;
  assetRef?: OfficeObjectRef;
}
interface Entry extends OfficeBaselineEntry {
  asset?: OfficeAsset;
}
interface Session {
  state(): Uint8Array;
  entries(): Entry[];
  editable(): OfficeEntry[];
  apply(command: OfficeCommand): { id: string; inverse: OfficeCommand };
  locate(id: string): OfficeTarget;
  exportBytes(determinism: ExportDeterminism): Promise<Uint8Array>;
  /** XLSX only: net effects read off the overrides. */
  effects?(): NetEffect[];
  dispose(): void;
}
const initialized = new Map<OfficeFormat | "opc", Promise<void>>();
/** The XLSX engine's linear memory, where replicas live (officeReplicaStats). */
let xlsxMemory: WebAssembly.Memory | undefined;
const assetPaths = {
  docx: "./office-runtime/docx.wasm",
  parse: "./office-runtime/parse.wasm",
  opc: "./office-runtime/opc.wasm",
  xlsx: "./office-runtime/xlsx.wasm",
  pptx: "./office-runtime/pptx.wasm",
};
const hash = (bytes: Uint8Array | string): string =>
  createHash("sha256").update(bytes).digest("hex");
const bytesAt = (path: string): Promise<Buffer> =>
  readFile(new URL(path, import.meta.url));
function initialize(format: OfficeFormat | "opc"): Promise<void> {
  let ready = initialized.get(format);
  if (!ready) {
    ready = (async () => {
      if (format === "docx") {
        await Promise.all([
          bytesAt(assetPaths.docx).then(preloadEditWasm),
          bytesAt(assetPaths.parse).then(preloadParseWasm),
          initialize("opc"),
        ]);
      } else if (format === "opc")
        await preloadOpcWasm(await bytesAt(assetPaths.opc));
      else if (format === "xlsx")
        xlsxMemory = (
          await initXlsx({ module_or_path: await bytesAt(assetPaths.xlsx) })
        ).memory;
      else await initPptx({ module_or_path: await bytesAt(assetPaths.pptx) });
    })();
    initialized.set(format, ready);
    ready.catch(() => initialized.delete(format));
  }
  return ready;
}
function canonical(value: unknown): string {
  return JSON.stringify(value, (_key, item: unknown) => {
    if (item instanceof Map)
      return Object.fromEntries(
        [...item].sort(([a], [b]) => String(a).localeCompare(String(b)))
      );
    if (item && typeof item === "object" && !Array.isArray(item)) {
      return Object.fromEntries(
        Object.entries(item).sort(([a], [b]) => a.localeCompare(b))
      );
    }
    return item;
  });
}
function assetFromDataUrl(src: unknown): OfficeAsset | undefined {
  if (typeof src !== "string") return undefined;
  const match =
    /^data:(image\/[a-zA-Z0-9.+-]+);base64,([A-Za-z0-9+/]*={0,2})$/.exec(src);
  if (!match)
    throw new Error(
      "Image object does not contain supported embedded image bytes"
    );
  const bytes = Uint8Array.from(Buffer.from(match[2], "base64"));
  return { bytes, mimeType: match[1], sha256: hash(bytes) };
}
function visual(
  id: string,
  label: string,
  value: unknown,
  position = ""
): Entry {
  return { id, kind: "visual", label, value: canonical(value), position };
}
function authoredText(value: unknown): string {
  if (Array.isArray(value)) return value.map(authoredText).join("");
  if (!value || typeof value !== "object") return "";
  const node = value as Record<string, unknown>;
  switch (node.type) {
    case "text":
      return typeof node.text === "string" ? node.text : "";
    case "tab":
      return "\t";
    case "break":
      return "\n";
    case "softHyphen":
      return "\u00ad";
    case "noBreakHyphen":
      return "\u2011";
    case "symbol":
      return typeof node.char === "string" ? node.char : "";
    case "paragraph":
      return authoredText(node.content) + "\n";
    case "run":
    case "inlineSdt":
    case "blockSdt":
      return authoredText(node.content);
    case "hyperlink":
      return authoredText(node.structuredChildren ?? node.children);
    case "simpleField":
      return authoredText(node.content);
    case "complexField":
      return authoredText(node.fieldResult);
    case "mathEquation":
      return typeof node.plainText === "string" ? node.plainText : "";
    case "shape":
      return node.shape
        ? authoredText(node.shape)
        : authoredText(
            (node.textBody as Record<string, unknown> | undefined)
              ?.paragraphs ?? node.children
          );
    case "table":
      return authoredText(node.rows);
    case "tableRow":
      return (
        (node.cells as unknown[] | undefined)?.map(authoredText).join("\t") +
        "\n"
      );
    case "tableCell":
      return authoredText(node.content);
    default:
      return "";
  }
}
/** A seeded source image names its package part (`media:<part>`); an inserted one carries its bytes. */
function docxImageAsset(
  src: unknown,
  media: ReadonlyMap<string, { dataUrl?: string }> | undefined
): OfficeAsset | undefined {
  if (typeof src !== "string" || !src.startsWith("media:"))
    return assetFromDataUrl(src);
  const part = src.slice("media:".length);
  const dataUrl = media?.get(part)?.dataUrl;
  if (!dataUrl)
    throw new Error(`DOCX image part ${part} is absent from the source package`);
  return assetFromDataUrl(dataUrl);
}
function docxEntries(
  session: YrsSession,
  stories: ReadonlySet<string>,
  embeds: ReadonlyMap<string, unknown>,
  media: ReadonlyMap<string, { dataUrl?: string }> | undefined
): Entry[] {
  const entries: Entry[] = [];
  for (const storyId of stories) {
    const objects = session.storyObjectIds(storyId);
    let objectIndex = 0;
    let paragraphIndex = 0;
    let text = "";
    let textFormats: Array<{
      offset: number;
      length: number;
      attributes: unknown;
    }> = [];
    let offset = 0;
    for (const segment of session.storySegments(storyId)) {
      if (segment.kind === "text") {
        if (!segment.attributes.del) {
          if (Object.keys(segment.attributes).length)
            textFormats.push({
              offset: text.length,
              length: segment.text.length,
              attributes: segment.attributes,
            });
          text += segment.text;
        }
        offset += segment.text.length;
        continue;
      }
      const stableId = objects[objectIndex++];
      if (!stableId) throw new Error("Missing DOCX object identity");
      if (segment.kind === "pilcrow") {
        const id = `${storyId}:paragraph:${segment.paraId}`;
        entries.push({
          id,
          kind: "text",
          label: `${storyId}, paragraph ${paragraphIndex + 1}`,
          value: text,
          position: `${storyId}:${paragraphIndex}`,
        });
        entries.push(
          visual(
            `${id}:format`,
            `${storyId}, paragraph ${paragraphIndex + 1} formatting`,
            { properties: segment.properties, runs: textFormats }
          )
        );
        text = "";
        textFormats = [];
        paragraphIndex++;
      } else if (segment.embedKind === "image" && !segment.attributes.del) {
        const id = `${storyId}:image:${stableId}`;
        const asset = docxImageAsset(segment.payload.src, media);
        const assetRef: OfficeObjectRef = {
          format: "docx",
          kind: "image",
          id: stableId,
          storyId,
        };
        entries.push({
          id,
          kind: "image",
          label: `${storyId}, image`,
          value: asset?.sha256 ?? canonical(segment.payload),
          position: `${storyId}:${offset}`,
          assetRef,
          asset,
        });
        const { src: _src, ...geometry } = segment.payload;
        entries.push(
          visual(`${id}:format`, `${storyId}, image formatting`, geometry)
        );
      } else if (segment.payload.modelKind === "commentReference") {
        // A comment's reference mark: its comment entry carries the change.
      } else {
        const id = `${storyId}:object:${stableId}`;
        const payload = segment.payload;
        const readable = authoredText(embeds.get(`${storyId}:${offset}`));
        const textual =
          segment.embedKind !== "table" &&
          (readable.length > 0 ||
            ["sdt", "blockSdt", "field", "math", "shape"].includes(
              segment.embedKind
            ));
        if (textual && !segment.attributes.del) {
          entries.push({
            id,
            kind: "text",
            label: `${storyId}, ${segment.embedKind}`,
            value: readable,
            position: `${storyId}:${offset}`,
          });
        }
        entries.push(
          visual(
            textual ? `${id}:format` : id,
            `${storyId}, ${segment.embedKind} formatting`,
            payload,
            `${storyId}:${offset}`
          )
        );
      }
      offset++;
    }
  }
  for (const comment of session.listComments()) {
    entries.push({
      id: `comment:${comment.id}`,
      kind: "text",
      label: comment.parentId ? "Comment reply" : "Comment",
      value: canonical({ author: comment.author, body: comment.body }),
      position: comment.parentId ?? "",
    });
    entries.push(
      visual(`comment:${comment.id}:status`, "Comment status", {
        done: comment.done,
        parentId: comment.parentId,
      })
    );
  }
  return entries;
}
interface XlsxProjection {
  sheets: Array<{
    id: string;
    name: string;
    cells: Array<{
      id: string;
      address: string;
      value: { kind: string; value?: unknown };
      formula: string | null;
      format: unknown;
    }>;
    images: Array<{
      id: string;
      part: string;
      anchor: unknown;
      bytes: number[];
    }>;
    [key: string]: unknown;
  }>;
  definedNames: Array<{
    name: string;
    formula: string;
    local_sheet: number | null;
  }>;
}
function xlsxEntries(doc: XlsxDocument): Entry[] {
  const projection = JSON.parse(
    doc.checkpointProjectionJson()
  ) as XlsxProjection;
  const entries: Entry[] = [];
  projection.sheets.forEach(({ cells, images, id, name, ...layout }, index) => {
    entries.push({
      id,
      kind: "text",
      label: `Sheet ${name}`,
      value: name,
      position: String(index),
    });
    entries.push(visual(`${id}:layout`, `${name} layout`, layout));
    if (Array.isArray(layout.hyperlinks))
      for (const [linkIndex, link] of layout.hyperlinks.entries()) {
        entries.push({
          id: `${id}:link:${linkIndex}`,
          kind: "text",
          label: `${name}, hyperlink`,
          value: canonical(link),
          position: "",
        });
      }
    for (const image of images) {
      const asset = {
        bytes: Uint8Array.from(image.bytes),
        mimeType: imageMimeType(image.part),
        sha256: hash(Uint8Array.from(image.bytes)),
      };
      entries.push({
        id: `${id}:image:${image.id}`,
        kind: "image",
        label: `${name}, image`,
        value: asset.sha256,
        position: canonical(image.anchor),
        assetRef: { format: "xlsx", kind: "image", sheetId: id, id: image.id },
        asset,
      });
    }
    for (const cell of cells) {
      const value =
        cell.formula !== null
          ? `=${cell.formula}`
          : cell.value.kind === "empty"
          ? ""
          : canonical(cell.value);
      if (value)
        entries.push({
          id: cell.id,
          kind: "text",
          label: `${name}!${cell.address}`,
          value,
          position: `${id}:${cell.address}`,
        });
      entries.push(
        visual(
          `${cell.id}:format`,
          `${name}!${cell.address} formatting`,
          cell.format
        )
      );
    }
  });
  for (const item of projection.definedNames)
    entries.push({
      id: `name:${item.local_sheet}:${item.name}`,
      kind: "text",
      label: `Defined name ${item.name}`,
      value: item.formula,
      position: "",
    });
  return entries;
}
function imageMimeType(part: string): string {
  const extension = part.split(".").pop()?.toLowerCase();
  const type = (
    {
      png: "image/png",
      jpg: "image/jpeg",
      jpeg: "image/jpeg",
      gif: "image/gif",
      svg: "image/svg+xml",
      webp: "image/webp",
      emf: "image/emf",
      wmf: "image/wmf",
      tiff: "image/tiff",
      tif: "image/tiff",
      bmp: "image/bmp",
    } as Record<string, string>
  )[extension ?? ""];
  if (!type) throw new Error(`Unsupported image part type: ${part}`);
  return type;
}
function pptxEntries(doc: PptxDocument): Entry[] {
  const deck = JSON.parse(doc.snapshotJson()) as DeckSnapshot;
  const entries: Entry[] = [
    visual("deck:size", "Presentation size", {
      width: deck.widthEmu,
      height: deck.heightEmu,
    }),
  ];
  deck.slides.forEach((slide, slideIndex) => {
    entries.push(
      visual(
        slide.id,
        `Slide ${slideIndex + 1}`,
        { name: slide.name, layout: slide.layoutPartPath },
        String(slideIndex)
      )
    );
    const visit = (shapes: ShapeSnapshot[], parent: string): void =>
      shapes.forEach((shape, shapeIndex) => {
        const label = `Slide ${slideIndex + 1}, ${shape.name}`;
        const { children, textStories, ...geometry } = shape;
        entries.push(
          visual(shape.id, label, geometry, `${parent}:${shapeIndex}`)
        );
        for (const story of textStories)
          story.paragraphs.forEach((paragraph, index) => {
            entries.push({
              id: paragraph.id,
              kind: "text",
              label,
              value: paragraph.runs.map((run) => run.text).join(""),
              position: `${story.id}:${index}`,
            });
            entries.push(
              visual(`${paragraph.id}:format`, `${label} text formatting`, {
                ...paragraph,
                runs: paragraph.runs.map(({ style }) => style),
              })
            );
          });
        if (shape.mediaPartPath) {
          const bytes = doc.mediaBytes(shape.mediaPartPath);
          const mimeType = imageMimeType(shape.mediaPartPath);
          const asset = { bytes, mimeType, sha256: hash(bytes) };
          entries.push({
            id: `${shape.id}:image`,
            kind: "image",
            label,
            value: asset.sha256,
            position: `${slide.id}:${shapeIndex}`,
            assetRef: {
              format: "pptx",
              kind: "image",
              id: shape.id,
              slideId: slide.id,
            },
            asset,
          });
        }
        visit(children, shape.id);
      });
    visit(slide.shapes, slide.id);
  });
  return entries;
}
function checkReplacement(
  current: string,
  command: OfficeCommand,
  format: string
): string {
  if (command.type !== "replace_text")
    throw new OfficeEditError(
      "unsupported_operation",
      `${format} sources support replace_text only`
    );
  if (command.text.includes("\n"))
    throw new OfficeEditError(
      "invalid_input",
      "replacement text must stay within one paragraph"
    );
  if (command.expectedText !== current)
    throw new OfficeEditError(
      "stale_target",
      "the paragraph text differs from expected_text"
    );
  return command.text;
}
/** The span of `before` that `after` changes, in UTF-16 offsets, never splitting a surrogate pair. */
function changedSpan(
  before: string,
  after: string
): { start: number; end: number; text: string } {
  const limit = Math.min(before.length, after.length);
  let prefix = 0;
  while (prefix < limit && before[prefix] === after[prefix]) prefix++;
  const high = before.charCodeAt(prefix - 1);
  if (high >= 0xd800 && high <= 0xdbff) prefix--;
  let suffix = 0;
  while (
    suffix < limit - prefix &&
    before[before.length - 1 - suffix] === after[after.length - 1 - suffix]
  )
    suffix++;
  const low = before.charCodeAt(before.length - suffix);
  if (suffix > 0 && low >= 0xdc00 && low <= 0xdfff) suffix--;
  return {
    start: prefix,
    end: before.length - suffix,
    text: after.slice(prefix, after.length - suffix),
  };
}
interface DocxParagraph {
  paraId: string;
  /** Story offset of the paragraph's replaceable text. */
  start: number;
  /** Paragraph offset of that text: the page and column breaks it opens with lie before it. */
  lead: number;
  length: number;
  text: string;
  plain: boolean;
}
function docxParagraphs(
  session: YrsSession,
  storyId: string
): DocxParagraph[] {
  let segments;
  try {
    segments = session.storySegments(storyId);
  } catch {
    throw new OfficeEditError("unavailable_target", "unknown DOCX story");
  }
  const paragraphs: DocxParagraph[] = [];
  let start = 0;
  let lead = 0;
  let cursor = 0;
  let text = "";
  let plain = true;
  for (const segment of segments) {
    if (segment.kind === "text") {
      if (segment.attributes.del || segment.attributes.ins) plain = false;
      text += segment.text;
      cursor += segment.text.length;
      continue;
    }
    if (segment.kind === "pilcrow") {
      paragraphs.push({
        paraId: segment.paraId,
        start,
        lead,
        length: cursor - start,
        text,
        plain,
      });
      cursor++;
      start = cursor;
      lead = 0;
      text = "";
      plain = true;
      continue;
    }
    // A page or column break the paragraph opens with stays in place,
    // outside its replaceable text.
    if (
      cursor === start &&
      (segment.embedKind === "pageBreak" || segment.embedKind === "columnBreak")
    ) {
      start++;
      lead++;
    } else plain = false;
    cursor++;
  }
  return paragraphs;
}
function docxTarget(id: string): { story: string; paraId: string } {
  const marker = id.lastIndexOf(":paragraph:");
  if (marker <= 0)
    throw new OfficeEditError(
      "unavailable_target",
      "target_id is not a DOCX paragraph"
    );
  return {
    story: id.slice(0, marker),
    paraId: id.slice(marker + ":paragraph:".length),
  };
}
function docxParagraph(session: YrsSession, id: string) {
  const { story, paraId } = docxTarget(id);
  const paragraph = docxParagraphs(session, story).find(
    (item) => item.paraId === paraId
  );
  if (!paragraph)
    throw new OfficeEditError(
      "unavailable_target",
      "the paragraph is no longer in the document"
    );
  return { story, paragraph };
}
type PptxRun = StorySnapshot["paragraphs"][number]["runs"][number];
interface PptxParagraph {
  id: string;
  start: number;
  length: number;
  text: string;
  runs: PptxRun[];
}
/** Style of the run holding the character at `offset`, else of the last run. */
function pptxStyleAt(
  runs: PptxRun[],
  offset: number
): Partial<PptxRun["style"]> {
  let end = 0;
  for (const run of runs) {
    end += run.text.length;
    if (offset < end) return run.style;
  }
  return runs.at(-1)?.style ?? {};
}
function pptxStoryParagraphs(
  doc: PptxDocument,
  storyId: string
): PptxParagraph[] {
  const story = JSON.parse(
    doc.storyJson(JSON.stringify({ storyId }))
  ) as StorySnapshot;
  let cursor = 0;
  return story.paragraphs.map((paragraph) => {
    const text = paragraph.runs.map((run) => run.text).join("");
    const item = {
      id: paragraph.id,
      start: cursor,
      length: text.length,
      text,
      runs: paragraph.runs,
    };
    cursor += text.length + 1;
    return item;
  });
}
function pptxTextEntries(
  doc: PptxDocument
): Array<OfficeEntry & { storyId: string }> {
  const deck = JSON.parse(doc.snapshotJson()) as DeckSnapshot;
  const entries: Array<OfficeEntry & { storyId: string }> = [];
  deck.slides.forEach((slide, slideIndex) => {
    const visit = (shapes: ShapeSnapshot[]): void =>
      shapes.forEach((shape) => {
        for (const story of shape.textStories)
          story.paragraphs.forEach((paragraph, index) => {
            entries.push({
              id: paragraph.id,
              label: `Slide ${slideIndex + 1}, ${shape.name}`,
              value: paragraph.runs.map((run) => run.text).join(""),
              position: `${story.id}:${index}`,
              storyId: story.id,
            });
          });
        visit(shape.children);
      });
    visit(slide.shapes);
  });
  return entries;
}
function pptxParagraph(
  doc: PptxDocument,
  storyOf: Map<string, string>,
  id: string
) {
  const storyId = storyOf.get(id);
  const paragraph =
    storyId === undefined
      ? undefined
      : pptxStoryParagraphs(doc, storyId).find((item) => item.id === id);
  if (storyId === undefined || !paragraph)
    throw new OfficeEditError(
      "unavailable_target",
      "the paragraph is no longer in the deck"
    );
  return { storyId, paragraph };
}
/** The projection's sheets with their cells only (checkpointCellsJson). */
interface XlsxCells {
  sheets: Array<{
    id: string;
    name: string;
    cells: Array<Omit<XlsxProjection["sheets"][number]["cells"][number], "format">>;
  }>;
}
function xlsxCellValue(cell: XlsxCells["sheets"][number]["cells"][number]) {
  if (cell.formula !== null) return `=${cell.formula}`;
  if (cell.value.kind === "empty" || cell.value.value === undefined) return "";
  return typeof cell.value.value === "string"
    ? cell.value.value
    : canonical(cell.value.value);
}
function xlsxCells(doc: XlsxDocument): XlsxCells {
  return JSON.parse(doc.checkpointCellsJson()) as XlsxCells;
}
function xlsxSheet(
  projection: XlsxCells,
  sheet: string
): { index: number; sheet: XlsxCells["sheets"][number] } {
  const wanted = sheet.trim().toLowerCase();
  const index = projection.sheets.findIndex(
    (item, position) =>
      item.id === sheet ||
      item.name.toLowerCase() === wanted ||
      String(position) === wanted
  );
  if (index < 0)
    throw new OfficeEditError("unavailable_target", "unknown sheet");
  return { index, sheet: projection.sheets[index] };
}
function a1(cell: string): { row: number; col: number; address: string } {
  const match = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(cell.trim());
  if (!match)
    throw new OfficeEditError("invalid_input", "cell must be an A1 address");
  const letters = match[1].toUpperCase();
  let col = 0;
  for (const letter of letters) col = col * 26 + letter.charCodeAt(0) - 64;
  const row = Number(match[2]);
  if (row < 1)
    throw new OfficeEditError("invalid_input", "cell must be an A1 address");
  return { row: row - 1, col: col - 1, address: `${letters}${row}` };
}
function assertCheckpoint(
  format: OfficeFormat,
  baseBytes: Uint8Array,
  checkpoint?: OfficeCheckpoint
) {
  if (!["docx", "xlsx", "pptx"].includes(format))
    throw new TypeError("Unsupported Office format");
  if (!(baseBytes instanceof Uint8Array) || !baseBytes.length)
    throw new TypeError("Expected nonempty Office base bytes");
  if (
    checkpoint &&
    (checkpoint.schemaVersion !== 1 ||
      checkpoint.format !== format ||
      checkpoint.baseSha256 !== hash(baseBytes) ||
      !(checkpoint.state instanceof Uint8Array) ||
      !checkpoint.state.length)
  )
    throw new Error(
      "Office checkpoint does not match the exact base package and schema"
    );
}
async function open(
  format: OfficeFormat,
  baseBytes: Uint8Array,
  checkpoint?: OfficeCheckpoint
): Promise<Session> {
  assertCheckpoint(format, baseBytes, checkpoint);
  await initialize(format);
  const clientId = randomInt(1, 0x1fffffffffff);
  if (format === "docx") {
    const session = await createYrsSession({ clientId });
    try {
      session.openDocx(baseBytes, !checkpoint);
      if (checkpoint) session.loadState(checkpoint.state);
      const base = session.materializeDocx();
      if (!base) throw new Error("DOCX source package was not attached");
      const stories = new Set<string>();
      const embeds = new Map<string, unknown>();
      let projected: ReturnType<typeof yrsToDocument> | undefined;
      const project = () =>
        (projected ??= yrsToDocument(session, base, {
          onStory: (story) => stories.add(story),
          onEmbed: (story, offset, content) =>
            embeds.set(`${story}:${offset}`, content),
        }));
      return {
        state: () => session.encodeState(),
        entries: () => {
          project();
          return docxEntries(session, stories, embeds, base.package.media);
        },
        editable: () => {
          project();
          return [...stories].flatMap((storyId) =>
            docxParagraphs(session, storyId)
              .map((paragraph, index) => ({
                id: `${storyId}:paragraph:${paragraph.paraId}`,
                label: `${storyId}, paragraph ${index + 1}`,
                value: paragraph.text,
                position: `${storyId}:${index}`,
                plain: paragraph.plain,
              }))
              .filter((entry) => entry.plain)
              .map(({ plain: _plain, ...entry }) => entry)
          );
        },
        apply: (command) => {
          if (command.type !== "replace_text")
            throw new OfficeEditError(
              "unsupported_operation",
              "DOCX sources support replace_text only"
            );
          const { story, paragraph } = docxParagraph(
            session,
            command.targetId
          );
          if (!paragraph.plain)
            throw new OfficeEditError(
              "unavailable_target",
              "the paragraph holds objects or tracked changes"
            );
          const text = checkReplacement(paragraph.text, command, "DOCX");
          const paraId = paragraph.paraId;
          const span = changedSpan(paragraph.text, text);
          if (span.start < span.end || span.text)
            session.replaceRange(
              {
                story,
                start: { paraId, offset: paragraph.lead + span.start },
                end: { paraId, offset: paragraph.lead + span.end },
              },
              span.text
            );
          projected = undefined;
          return {
            id: command.targetId,
            inverse: {
              type: "replace_text",
              targetId: command.targetId,
              expectedText: text,
              text: paragraph.text,
            },
          };
        },
        locate: (id) => {
          const { story, paragraph } = docxParagraph(session, id);
          return {
            id,
            path: ["stories", story],
            range: [paragraph.start, paragraph.start + paragraph.length],
          };
        },
        exportBytes: async (determinism) =>
          new Uint8Array(
            (
              await writeDocumentWithRust(
                project(),
                exactBuffer(baseBytes),
                { updateModifiedDate: false },
                undefined,
                determinism
              )
            ).buffer
          ),
        dispose: () => session.destroy(),
      };
    } catch (error) {
      session.destroy();
      throw error;
    }
  }
  if (format === "xlsx") {
    const doc = XlsxDocument.openCollaborative(baseBytes, clientId);
    try {
      if (checkpoint) doc.applyUpdateJson(checkpoint.state);
      const cellAt = (
        sheet: XlsxCells["sheets"][number],
        address: string
      ) => sheet.cells.find((cell) => cell.address === address);
      return {
        state: () => doc.encodeStateAsUpdate(),
        entries: () => xlsxEntries(doc),
        editable: () =>
          xlsxCells(doc).sheets.flatMap((sheet) =>
            sheet.cells.map((cell) => ({
              id: cell.id,
              label: `${sheet.name}!${cell.address}`,
              value: xlsxCellValue(cell),
              position: `${sheet.id}:${cell.address}`,
            }))
          ),
        apply: (command) => {
          if (command.type !== "set_cell")
            throw new OfficeEditError(
              "unsupported_operation",
              "XLSX sources support set_cell only"
            );
          const { index, sheet } = xlsxSheet(xlsxCells(doc), command.sheet);
          const { row, col, address } = a1(command.cell);
          const at = JSON.stringify({ sheet: index, row, col });
          const readInput = () =>
            (JSON.parse(doc.cellJson(at)) as { input: string }).input;
          const before = cellAt(sheet, address);
          const input = readInput();
          if (
            command.expectedValue !== input &&
            command.expectedValue !== (before ? xlsxCellValue(before) : "")
          )
            throw new OfficeEditError(
              "stale_target",
              "the cell differs from expected_value"
            );
          if (!before && !command.value)
            throw new OfficeEditError("invalid_input", "the cell is already empty");
          const result = JSON.parse(
            doc.editCellJson(
              JSON.stringify({ sheet: index, row, col, input: command.value })
            )
          ) as { applied?: boolean };
          if (result.applied === false)
            throw new OfficeEditError(
              "invalid_input",
              "the workbook rejected this cell value"
            );
          const cell =
            before ?? cellAt(xlsxSheet(xlsxCells(doc), sheet.id).sheet, address);
          if (!cell) throw new Error("edited cell is missing from the projection");
          return {
            id: cell.id,
            inverse: {
              type: "set_cell",
              sheet: sheet.id,
              cell: address,
              expectedValue: readInput(),
              value: input,
            },
          };
        },
        locate: (id) => {
          const marker = id.indexOf(":[");
          if (marker <= 0)
            throw new OfficeEditError(
              "unavailable_target",
              "target_id is not an XLSX cell"
            );
          const sheetId = id.slice(0, marker);
          if (!xlsxCells(doc).sheets.some((item) => item.id === sheetId))
            throw new OfficeEditError(
              "unavailable_target",
              "the sheet is no longer in the workbook"
            );
          return {
            id,
            path: ["xlsx:sheets", sheetId, "contents", id.slice(marker + 1)],
          };
        },
        exportBytes: async (determinism) =>
          doc.saveBytesAt(Date.parse(determinism.now) / 86_400_000 + 25569),
        effects: () => JSON.parse(doc.pendingEffectsJson()) as NetEffect[],
        dispose: () => doc.free(),
      };
    } catch (error) {
      doc.free();
      throw error;
    }
  }
  const doc = checkpoint
    ? PptxDocument.openCollaborativeFromUpdate(
        checkpoint.state,
        clientId,
        baseBytes
      )
    : PptxDocument.openCollaborative(baseBytes, clientId);
  // A session serves one call, and replace_text never moves a paragraph to another story.
  let stories: Map<string, string> | undefined;
  const storyOf = () =>
    (stories ??= new Map(
      pptxTextEntries(doc).map((entry) => [entry.id, entry.storyId])
    ));
  return {
    state: () => doc.encodeStateAsUpdate(),
    entries: () => pptxEntries(doc),
    editable: () =>
      pptxTextEntries(doc).map(({ storyId: _storyId, ...entry }) => entry),
    apply: (command) => {
      if (command.type !== "replace_text")
        throw new OfficeEditError(
          "unsupported_operation",
          "PPTX sources support replace_text only"
        );
      const { storyId, paragraph } = pptxParagraph(
        doc,
        storyOf(),
        command.targetId
      );
      const text = checkReplacement(paragraph.text, command, "PPTX");
      const span = changedSpan(paragraph.text, text);
      // A pure insertion replaces nothing; it takes the style of the character before it.
      const style = pptxStyleAt(
        paragraph.runs,
        span.start < span.end ? span.start : Math.max(span.start - 1, 0)
      );
      const start = paragraph.start + span.start;
      if (span.start < span.end)
        doc.deleteTextJson(
          JSON.stringify({ storyId, start, end: paragraph.start + span.end })
        );
      if (span.text)
        doc.insertTextJson(
          JSON.stringify({ storyId, index: start, text: span.text, style })
        );
      return {
        id: command.targetId,
        inverse: {
          type: "replace_text",
          targetId: command.targetId,
          expectedText: text,
          text: paragraph.text,
        },
      };
    },
    locate: (id) => {
      const { storyId, paragraph } = pptxParagraph(doc, storyOf(), id);
      return {
        id,
        path: ["pptx:stories", storyId],
        range: [paragraph.start, paragraph.start + paragraph.length],
      };
    },
    exportBytes: async () => doc.saveBytes(),
    dispose: () => doc.free(),
  };
}
function exactBuffer(bytes: Uint8Array): ArrayBuffer {
  return Uint8Array.from(bytes).buffer;
}

/*
 * Replicas: an XLSX room's source opened for reading pending effects
 * (XlsxEffectsReader), kept between saves so a save does not reopen the
 * workbook. The reader never applies a state: each save's state is checked
 * and projected beside it, so a replica stays the opened source and serves
 * any state of its base. Only xlsxPendingEffects takes a room; edits,
 * inspection, exports and rebases always open their own session. A call that
 * fails drops its room's replica. Under the budget, a new replica pushes out
 * the least recently used ones only once they have been idle for
 * REPLICA_IDLE_MS; otherwise it is not kept.
 *
 * DOCX and PPTX keep none. A DOCX open is a small part of its baseline next
 * to the projection. A PPTX replica saved about 0.3 s per save of the 24 MB
 * deck on the production box while holding about 90 MB, and under a shared
 * budget PPTX replicas pushed out the XLSX ones.
 */
interface Replica {
  baseSha256: string;
  reader: XlsxEffectsReader;
  bytes: number;
  /** performance.now() of its last use. */
  usedAt: number;
}
// Map order is recency: the first entry is the least recently used.
const replicas = new Map<string, Replica>();
let replicaBudget = 0;
let replicaBytes = 0;
const replicaCounts = { hits: 0, misses: 0, evictions: 0 };
/**
 * WASM heap an XLSX replica holds per byte of its unzipped package, measured
 * as linear memory growth per reader kept (2026-10-05): 11-12× for
 * course-guide.xlsx and 15.2× for the 16,000-row gradebook, rounded up. An
 * editing replica held 15.6× and 19.6×.
 */
const XLSX_HEAP_PER_UNZIPPED_BYTE = 16;
/**
 * How long a replica must have gone unused before a new one may push it
 * out. A room being edited saves every 30 s at most (Capy's longest source
 * save debounce) plus its wait in the engine queue, so a replica idle this
 * long belongs to a room nobody edits, or to a worker already saturated.
 * Without it, rooms that save in turn and need more than the budget push
 * out each replica just before its next use and every save misses (measured
 * at 30 rooms: no hits at all).
 */
const REPLICA_IDLE_MS = 120_000;

/**
 * Turns replicas on with a budget in estimated WASM heap bytes (0 turns them
 * off and drops every replica). The host picks the budget; there is none by
 * default.
 */
export function configureOfficeReplicas(budgetBytes: number) {
  if (!Number.isSafeInteger(budgetBytes) || budgetBytes < 0)
    throw new TypeError("Expected a nonnegative replica budget in bytes");
  replicaBudget = budgetBytes;
  evictReplicas();
}
/** Drops a room's replica (the room unloaded). */
export function dropOfficeReplica(room: string) {
  const replica = replicas.get(room);
  if (!replica) return;
  replicas.delete(room);
  replicaBytes -= replica.bytes;
  replica.reader.free();
}
export function officeReplicaStats() {
  return {
    replicas: replicas.size,
    replicaBytes,
    ...replicaCounts,
    wasmBytes: xlsxMemory?.buffer.byteLength ?? 0,
  };
}
/**
 * Drops least recently used replicas until `needed` more bytes fit; with
 * `idleMs`, only those unused for that long. Returns whether they fit.
 */
function evictReplicas(needed = 0, idleMs?: number) {
  const now = performance.now();
  for (const [room, replica] of replicas) {
    if (replicaBytes + needed <= replicaBudget) break;
    if (idleMs !== undefined && now - replica.usedAt < idleMs) break;
    dropOfficeReplica(room);
    replicaCounts.evictions += 1;
  }
  return replicaBytes + needed <= replicaBudget;
}

/** Sum of the ZIP entries' uncompressed sizes; undefined for ZIP64 or a damaged directory. */
function unzippedBytes(bytes: Uint8Array): number | undefined {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  // The end-of-central-directory record: 22 bytes plus a comment of up to 64 KiB.
  for (
    let end = bytes.length - 22;
    end >= Math.max(0, bytes.length - 22 - 0xffff);
    end--
  ) {
    if (view.getUint32(end, true) !== 0x06054b50) continue;
    const count = view.getUint16(end + 10, true);
    let at = view.getUint32(end + 16, true);
    if (count === 0xffff || at === 0xffffffff) return undefined;
    let total = 0;
    for (let index = 0; index < count; index++) {
      if (at + 46 > bytes.length || view.getUint32(at, true) !== 0x02014b50)
        return undefined;
      const size = view.getUint32(at + 24, true);
      if (size === 0xffffffff) return undefined;
      total += size;
      at +=
        46 +
        view.getUint16(at + 28, true) +
        view.getUint16(at + 30, true) +
        view.getUint16(at + 32, true);
    }
    return total;
  }
  return undefined;
}

/**
 * `room`'s replica for this base, opened on a miss and kept when it fits
 * without pushing out replicas still in use; undefined when not kept.
 */
function roomReplica(
  room: string,
  baseBytes: Uint8Array,
  baseSha256: string
): Replica | undefined {
  const held = replicas.get(room);
  if (held?.baseSha256 === baseSha256) {
    held.usedAt = performance.now();
    replicas.delete(room);
    replicas.set(room, held);
    replicaCounts.hits += 1;
    return held;
  }
  if (held) dropOfficeReplica(room);
  replicaCounts.misses += 1;
  const unzipped = unzippedBytes(baseBytes);
  if (unzipped === undefined) return undefined;
  const bytes = Math.ceil(unzipped * XLSX_HEAP_PER_UNZIPPED_BYTE);
  if (bytes > replicaBudget || !evictReplicas(bytes, REPLICA_IDLE_MS))
    return undefined;
  const replica = {
    baseSha256,
    reader: new XlsxEffectsReader(baseBytes, randomInt(1, 0x1fffffffffff)),
    bytes,
    usedAt: performance.now(),
  };
  replicas.set(room, replica);
  replicaBytes += bytes;
  return replica;
}
/**
 * The checkpoint's effects off `room`'s replica, or off a reader opened for
 * this call. A state the reader would not adopt whole is applied to a fresh
 * session, as an editor would.
 */
async function readXlsxEffects(
  baseBytes: Uint8Array,
  checkpoint: OfficeCheckpoint,
  room: string | undefined
): Promise<NetEffect[]> {
  assertCheckpoint("xlsx", baseBytes, checkpoint);
  await initialize("xlsx");
  const replica =
    room && replicaBudget
      ? roomReplica(room, baseBytes, checkpoint.baseSha256)
      : undefined;
  const reader =
    replica?.reader ??
    new XlsxEffectsReader(baseBytes, randomInt(1, 0x1fffffffffff));
  let effects: string | undefined;
  try {
    effects = reader.pendingEffectsJson(checkpoint.state);
  } catch (error) {
    if (replica) dropOfficeReplica(room!);
    throw error;
  } finally {
    if (!replica) reader.free();
  }
  if (effects !== undefined) return JSON.parse(effects) as NetEffect[];
  const session = await open("xlsx", baseBytes, checkpoint);
  try {
    return session.effects!();
  } finally {
    session.dispose();
  }
}
export async function seedOffice(
  format: OfficeFormat,
  baseBytes: Uint8Array
): Promise<OfficeCheckpoint> {
  const session = await open(format, baseBytes);
  try {
    return {
      format,
      schemaVersion: 1,
      baseSha256: hash(baseBytes),
      state: session.state(),
    };
  } finally {
    session.dispose();
  }
}
/** Semantic comparison data only; media is represented by its content hash. */
export async function officeBaseline(
  baseBytes: Uint8Array,
  checkpoint: OfficeCheckpoint,
): Promise<OfficeBaselineEntry[]> {
  const session = await open(checkpoint.format, baseBytes, checkpoint);
  try {
    return session.entries().map(({ asset, ...entry }) => ({
      ...entry,
      value:
        entry.kind === "visual" ? hash(Buffer.from(entry.value)) : entry.value,
      ...(asset ? { imageSHA256: asset.sha256 } : {}),
    }));
  } finally {
    session.dispose();
  }
}

export function compareBaselines(
  from: OfficeBaselineEntry[],
  to: OfficeBaselineEntry[],
): NetEffect[] {
  const before = new Map(from.map((entry) => [entry.id, entry]));
  const after = new Map(to.map((entry) => [entry.id, entry]));
  const effects: NetEffect[] = [];
  for (const id of [...new Set([...before.keys(), ...after.keys()])].sort()) {
    const old = before.get(id),
      next = after.get(id);
    if (
      old &&
      next &&
      old.value === next.value &&
      old.position === next.position
    )
      continue;
    const entry = next ?? old!;
    const effect: NetEffect = {
      id,
      kind: entry.kind,
      operation: !old
        ? "add"
        : !next
          ? "remove"
          : old.value === next.value
            ? "move"
            : "replace",
      label: entry.label,
    };
    if (entry.kind === "text") {
      if (old) effect.before = old.value;
      if (next) effect.after = next.value;
    }
    if (next?.assetRef) effect.assetRef = next.assetRef;
    if (next?.imageSHA256) effect.imageSHA256 = next.imageSHA256;
    effects.push(effect);
  }
  return effects;
}

export async function compare(
  baseBytes: Uint8Array,
  fromCheckpoint: OfficeCheckpoint,
  toCheckpoint: OfficeCheckpoint,
): Promise<NetEffect[]> {
  if (fromCheckpoint.format !== toCheckpoint.format)
    throw new Error("Cannot compare different Office formats");
  return compareBaselines(
    await officeBaseline(baseBytes, fromCheckpoint),
    await officeBaseline(baseBytes, toCheckpoint),
  );
}
/**
 * Lands the edits saved after the capture on seed(export) and returns that
 * state with its effects against the export's derived baseline. Fails when the
 * rebased state does not carry the same changes as latest − captured, or when
 * it would need content the export dropped.
 */
export async function rebaseOffice(
  baseBytes: Uint8Array,
  captured: OfficeCheckpoint,
  latest: OfficeCheckpoint,
  exportedSource: Uint8Array
): Promise<{
  state: Uint8Array;
  effects: NetEffect[];
}> {
  const format = captured.format;
  if (
    !["docx", "xlsx", "pptx"].includes(format) ||
    !(baseBytes instanceof Uint8Array) ||
    !baseBytes.length ||
    !(exportedSource instanceof Uint8Array) ||
    !exportedSource.length
  )
    throw new TypeError(
      "Expected supported Office format and nonempty source bytes"
    );
  for (const checkpoint of [captured, latest]) {
    if (
      checkpoint.format !== format ||
      checkpoint.schemaVersion !== 1 ||
      checkpoint.baseSha256 !== hash(baseBytes) ||
      !(checkpoint.state instanceof Uint8Array) ||
      !checkpoint.state.length
    )
      throw new Error(
        "Office checkpoint does not match the exact base package and schema"
      );
  }
  await initialize(format);
  if (format === "xlsx") {
    let state: Uint8Array;
    try {
      state = XlsxDocument.rebaseCheckpoint(
        baseBytes,
        captured.state,
        latest.state,
        exportedSource,
        randomInt(1, 0x1fffffffffff)
      );
    } catch (error) {
      // The engine reports a refusal as a thrown string; a trap or a glue
      // error (an Error) is not a refusal and stays retryable.
      if (typeof error !== "string") throw error;
      throw new RebaseError(`Office rebase: ${error}`);
    }
    const checkpoint = {
      format,
      schemaVersion: 1 as const,
      baseSha256: hash(exportedSource),
      state,
    };
    return {
      state,
      effects: await xlsxPendingEffects(exportedSource, checkpoint),
    };
  }
  const exported = {
    format,
    schemaVersion: 1 as const,
    baseSha256: hash(exportedSource),
  };
  const seed = (await seedOffice(format, exportedSource)).state;
  let ids: Map<string, string>;
  if (format === "docx") {
    validateExportBacking(
      unzipContainer(baseBytes),
      unzipContainer(exportedSource)
    );
    ids = docxIds(captured.state, seed);
  } else {
    const deck = (bytes: Uint8Array, state?: Uint8Array) => {
      const doc = state
        ? PptxDocument.openCollaborativeFromUpdate(
            state,
            randomInt(1, 0x1fffffffffff),
            bytes
          )
        : PptxDocument.openCollaborative(bytes, randomInt(1, 0x1fffffffffff));
      try {
        return JSON.parse(doc.snapshotJson()) as DeckSnapshot;
      } finally {
        doc.free();
      }
    };
    ids = pptxIds(deck(baseBytes, captured.state), deck(exportedSource));
  }
  const clientId = randomInt(1, 0x1fffffffffff);
  const rebased = transplant(
    format === "docx" ? DOCX_LINEAGE : PPTX_LINEAGE,
    captured.state,
    latest.state,
    seed,
    ids,
    clientId
  );
  const state = rebased.state;
  if (format === "docx") {
    await assertDocxRestorations(
      baseBytes,
      latest.state,
      exportedSource,
      state,
      clientId,
      rebased.ids
    );
    await assertDocxRenders(exportedSource, state);
  } else assertPptxRestorations(seed, state);
  const effects = compareBaselines(
    await officeBaseline(exportedSource, { ...exported, state: seed }),
    await officeBaseline(exportedSource, { ...exported, state })
  );
  const saved = compareBaselines(
    await officeBaseline(baseBytes, captured),
    await officeBaseline(baseBytes, latest)
  );
  // DOCX formatting passes through an export as its own representation, so
  // unchanged formatting may read differently, and so does the structure of
  // a comment body (the export adds its reference run); text and images may
  // not.
  const docxComment = (effect: NetEffect, value: string | undefined) =>
    format === "docx" && effect.id.startsWith("comment:") && value
      ? commentText(value)
      : value;
  const change = (list: NetEffect[]) =>
    list
      .filter((effect) => format === "pptx" || effect.kind !== "visual")
      .map((effect) =>
        JSON.stringify([
          effect.operation,
          effect.kind,
          docxComment(effect, effect.before),
          docxComment(effect, effect.after),
          effect.imageSHA256,
        ])
      )
      .sort();
  if (JSON.stringify(change(saved)) !== JSON.stringify(change(effects)))
    throw new RebaseError(
      "Office rebase: the rebased state does not carry the saved edits"
    );
  return { state, effects };
}

/** A DOCX comment entry's author and visible text. */
function commentText(value: string): string {
  const texts: string[] = [];
  const walk = (node: unknown): void => {
    if (Array.isArray(node)) node.forEach(walk);
    else if (node && typeof node === "object")
      for (const [key, child] of Object.entries(node))
        if (key === "text" && typeof child === "string") texts.push(child);
        else walk(child);
  };
  const { author, body } = JSON.parse(value) as {
    author: unknown;
    body: unknown;
  };
  walk(body);
  return JSON.stringify([author, texts.join("")]);
}

/** The export keeps every package part a later edit may still name: it rewrites owned XML only. */
function validateExportBacking(
  oldParts: Record<string, Uint8Array>,
  newParts: Record<string, Uint8Array>
): void {
  const decoder = new TextDecoder();
  const targetPath = (owner: string, target: string) => {
    const url = new URL(target, `https://docx.invalid/${owner}`);
    if (url.origin !== "https://docx.invalid" || url.search || url.hash)
      throw new RebaseError(
        "Office rebase: DOCX needs an internal package target"
      );
    return decodeURIComponent(url.pathname.slice(1));
  };
  const owned = new Set([
    "[Content_Types].xml",
    "word/document.xml",
    "word/footnotes.xml",
    "word/endnotes.xml",
    "word/comments.xml",
    "word/commentsExtended.xml",
    "word/commentsIds.xml",
    "word/commentsExtensible.xml",
    "word/numbering.xml",
    "docProps/core.xml",
  ]);
  for (const [path, bytes] of Object.entries(oldParts)) {
    if (!path.endsWith(".rels")) continue;
    const next = newParts[path];
    if (!next)
      throw new RebaseError(
        `Office rebase: the export lost relationship part ${path}`
      );
    const previous = parseRelationshipsXmlWithRust(decoder.decode(bytes), path);
    const current = parseRelationshipsXmlWithRust(decoder.decode(next), path);
    for (const [id, relationship] of previous) {
      const replacement = current.get(id);
      if (
        !replacement ||
        replacement.type !== relationship.type ||
        replacement.target !== relationship.target ||
        replacement.targetMode !== relationship.targetMode
      )
        throw new RebaseError(
          `Office rebase: the export changed source relationship ${path}:${id}`
        );
      if (
        relationship.type.endsWith("/header") ||
        relationship.type.endsWith("/footer")
      )
        owned.add(
          targetPath(
            path.replace("/_rels/", "/").replace(/\.rels$/, ""),
            relationship.target
          )
        );
    }
  }
  for (const [path, bytes] of Object.entries(oldParts)) {
    if (owned.has(path) || path.endsWith(".rels")) continue;
    const next = newParts[path];
    if (
      !next ||
      next.length !== bytes.length ||
      !next.every((byte, i) => byte === bytes[i])
    )
      throw new RebaseError(
        `Office rebase: the export changed or lost unmodeled source part ${path}`
      );
  }
}

/** The raw inline markup each projected DOCX paragraph takes from its source paragraph, by story and paragraph id. */
async function docxRawInlines(
  base: Uint8Array,
  state: Uint8Array
): Promise<Map<string, string>> {
  const session = await createYrsSession({
    clientId: randomInt(1, 0x1fffffffffff),
  });
  try {
    session.openDocx(base, false);
    session.loadState(state);
    const document = session.materializeDocx();
    if (!document) throw new Error("DOCX source package was not attached");
    const paragraphs = new Map<string, string>();
    yrsToDocument(session, document, {
      onParagraph: (story, _offset, paragraph, id) => {
        const key = `${story}\u0000${id}`;
        // A repeated id cannot name one source paragraph.
        paragraphs.set(
          key,
          paragraphs.has(key)
            ? "repeated"
            : JSON.stringify(
                paragraph.content.filter((child) => child.type === "rawXml")
              )
        );
      },
    });
    return paragraphs;
  } finally {
    session.destroy();
  }
}

/**
 * The editor can open the rebased state: its render bridge takes every story.
 * It refuses, say, text or a field ahead of a table in one paragraph slot,
 * which the editor never writes but concurrent edits and an export can.
 */
async function assertDocxRenders(
  source: Uint8Array,
  state: Uint8Array
): Promise<void> {
  const session = await createYrsSession({
    clientId: randomInt(1, 0x1fffffffffff),
  });
  try {
    session.openDocx(source, false);
    session.loadState(state);
    for (const story of session.storyIds())
      try {
        session.yrsBlocksForStory(story);
      } catch (error) {
        throw new RebaseError(
          `Office rebase: the rebased story ${story} does not render: ${
            error instanceof Error ? error.message : String(error)
          }`
        );
      }
  } finally {
    session.destroy();
  }
}

/**
 * Every paragraph mark the rebase wrote takes the same raw inline markup from
 * the export as it took from the old source. One the later edits brought back
 * (an Undo of a deletion made before the capture) takes it from a source
 * paragraph the export dropped.
 */
async function assertDocxRestorations(
  baseBytes: Uint8Array,
  latest: Uint8Array,
  exportedSource: Uint8Array,
  state: Uint8Array,
  clientId: number,
  ids: Map<string, string>
): Promise<void> {
  const written = docxWrittenParagraphs(state, clientId);
  if (!written.length) return;
  const back = new Map([...ids].map(([from, to]) => [to, from]));
  const before = await docxRawInlines(baseBytes, latest);
  const after = await docxRawInlines(exportedSource, state);
  for (const [story, id] of written)
    if (
      before.get(`${back.get(story) ?? story}\u0000${back.get(id) ?? id}`) !==
      after.get(`${story}\u0000${id}`)
    )
      throw new RebaseError(
        `Office rebase: DOCX paragraph ${id} needs content the export dropped`
      );
}

/** Pending XLSX effects against the base, read off the checkpoint's overrides; XLSX keeps no stored baseline. */
export async function xlsxPendingEffects(
  baseBytes: Uint8Array,
  checkpoint: OfficeCheckpoint,
  room?: string
): Promise<NetEffect[]> {
  if (checkpoint.format !== "xlsx")
    throw new TypeError("Expected an XLSX checkpoint");
  return readXlsxEffects(baseBytes, checkpoint, room);
}

export async function resolveAsset(
  baseBytes: Uint8Array,
  checkpoint: OfficeCheckpoint,
  objectRef: OfficeObjectRef
): Promise<OfficeAsset> {
  if (objectRef.format !== checkpoint.format || objectRef.kind !== "image")
    throw new Error("Image reference format mismatch");
  const session = await open(checkpoint.format, baseBytes, checkpoint);
  try {
    const entry = session
      .entries()
      .find(
        (entry) =>
          entry.assetRef && canonical(entry.assetRef) === canonical(objectRef)
      );
    if (!entry?.asset)
      throw new Error("Image object is absent from the checkpoint");
    return entry.asset;
  } finally {
    session.dispose();
  }
}
export async function exportOffice(
  baseBytes: Uint8Array,
  checkpoint: OfficeCheckpoint,
  determinism: ExportDeterminism
): Promise<Uint8Array> {
  if (
    !/^[0-9a-f]{64}$/.test(determinism.seed) ||
    !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(determinism.now)
  )
    throw new TypeError(
      "Expected deterministic SHA-256 seed and ISO UTC millisecond time"
    );
  const session = await open(checkpoint.format, baseBytes, checkpoint);
  try {
    return await session.exportBytes(determinism);
  } finally {
    session.dispose();
  }
}
export async function inspectOffice(
  baseBytes: Uint8Array,
  checkpoint: OfficeCheckpoint
): Promise<OfficeEntry[]> {
  const session = await open(checkpoint.format, baseBytes, checkpoint);
  try {
    return session.editable();
  } finally {
    session.dispose();
  }
}
/** Applies content commands in order; the inverse list is returned in undo order and targets describe the post-edit state. */
export async function applyOfficeCommands(
  baseBytes: Uint8Array,
  checkpoint: OfficeCheckpoint,
  commands: OfficeCommand[]
): Promise<OfficeCommandResult> {
  if (!Array.isArray(commands) || !commands.length)
    throw new OfficeEditError("invalid_input", "no commands");
  const session = await open(checkpoint.format, baseBytes, checkpoint);
  try {
    const inverse: OfficeCommand[] = [];
    const ids = new Set<string>();
    for (const command of commands) {
      const applied = session.apply(command);
      inverse.unshift(applied.inverse);
      ids.add(applied.id);
    }
    return {
      state: session.state(),
      inverse,
      targets: [...ids].map((id) => session.locate(id)),
    };
  } finally {
    session.dispose();
  }
}
export async function locateOfficeTargets(
  baseBytes: Uint8Array,
  checkpoint: OfficeCheckpoint,
  ids: string[]
): Promise<OfficeTarget[]> {
  const session = await open(checkpoint.format, baseBytes, checkpoint);
  try {
    return ids.map((id) => session.locate(id));
  } finally {
    session.dispose();
  }
}
export async function runtimeManifest(): Promise<Record<string, string>> {
  return Object.fromEntries(
    await Promise.all(
      Object.entries(assetPaths).map(async ([name, path]) => [
        name,
        hash(await bytesAt(path)),
      ])
    )
  );
}
