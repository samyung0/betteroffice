import { createHash, randomInt } from "node:crypto";
import { readFile } from "node:fs/promises";
import { createYrsSession, type YrsSession } from "../packages/docx/src/yrs";
import { yrsToDocument } from "../packages/docx/src/yrs/yrsToDocument";
import { preloadEditWasm } from "../packages/docx/src/wasm/edit";
import { preloadParseWasm } from "../packages/docx/src/wasm/parse";
import { preloadOpcWasm } from "../packages/docx/src/wasm/opc";
import { writeDocumentWithRust } from "../packages/docx/src/docx/rustSaveFacade";
import initXlsx, {
  XlsxDocument,
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
interface Entry {
  id: string;
  kind: NetEffect["kind"];
  label: string;
  value: string;
  position: string;
  assetRef?: OfficeObjectRef;
  asset?: OfficeAsset;
}
interface Session {
  state(): Uint8Array;
  entries(): Entry[];
  editable(): OfficeEntry[];
  apply(command: OfficeCommand): { id: string; inverse: OfficeCommand };
  locate(id: string): OfficeTarget;
  exportBytes(determinism: ExportDeterminism): Promise<Uint8Array>;
  dispose(): void;
}
const initialized = new Map<OfficeFormat | "opc", Promise<void>>();
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
        await initXlsx({ module_or_path: await bytesAt(assetPaths.xlsx) });
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
function docxEntries(
  session: YrsSession,
  stories: ReadonlySet<string>,
  embeds: ReadonlyMap<string, unknown>
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
        const asset = assetFromDataUrl(segment.payload.src);
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
interface DocxParagraph {
  paraId: string;
  start: number;
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
  let cursor = 0;
  let text = "";
  let plain = true;
  for (const segment of segments) {
    if (segment.kind === "text") {
      if (segment.attributes.del) plain = false;
      text += segment.text;
      cursor += segment.text.length;
      continue;
    }
    if (segment.kind === "pilcrow") {
      paragraphs.push({
        paraId: segment.paraId,
        start,
        length: cursor - start,
        text,
        plain,
      });
      cursor++;
      start = cursor;
      text = "";
      plain = true;
      continue;
    }
    plain = false;
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
interface PptxParagraph {
  id: string;
  start: number;
  length: number;
  text: string;
  style: StorySnapshot["paragraphs"][number]["runs"][number]["style"];
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
      style: paragraph.runs[0]?.style ?? {},
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
function pptxParagraph(doc: PptxDocument, id: string) {
  const entry = pptxTextEntries(doc).find((item) => item.id === id);
  if (!entry)
    throw new OfficeEditError(
      "unavailable_target",
      "the paragraph is no longer in the deck"
    );
  const paragraph = pptxStoryParagraphs(doc, entry.storyId).find(
    (item) => item.id === id
  );
  if (!paragraph)
    throw new OfficeEditError(
      "unavailable_target",
      "the paragraph is no longer in the deck"
    );
  return { storyId: entry.storyId, paragraph };
}
function xlsxCellValue(cell: XlsxProjection["sheets"][number]["cells"][number]) {
  if (cell.formula !== null) return `=${cell.formula}`;
  if (cell.value.kind === "empty" || cell.value.value === undefined) return "";
  return typeof cell.value.value === "string"
    ? cell.value.value
    : canonical(cell.value.value);
}
function xlsxProjection(doc: XlsxDocument): XlsxProjection {
  return JSON.parse(doc.checkpointProjectionJson()) as XlsxProjection;
}
function xlsxSheet(
  projection: XlsxProjection,
  sheet: string
): { index: number; sheet: XlsxProjection["sheets"][number] } {
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
async function open(
  format: OfficeFormat,
  baseBytes: Uint8Array,
  checkpoint?: OfficeCheckpoint
): Promise<Session> {
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
          return docxEntries(session, stories, embeds);
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
          if (text)
            session.insertText(
              { story, paraId, offset: paragraph.length },
              text
            );
          if (paragraph.length)
            session.deleteRange({
              story,
              start: { paraId, offset: 0 },
              end: { paraId, offset: paragraph.length },
            });
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
      const paths = new Map<string, string[]>();
      const cellAt = (
        sheet: XlsxProjection["sheets"][number],
        address: string
      ) => sheet.cells.find((cell) => cell.address === address);
      return {
        state: () => doc.encodeStateAsUpdate(),
        entries: () => xlsxEntries(doc),
        editable: () =>
          xlsxProjection(doc).sheets.flatMap((sheet) =>
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
          const { index, sheet } = xlsxSheet(xlsxProjection(doc), command.sheet);
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
            before ?? cellAt(xlsxSheet(xlsxProjection(doc), sheet.id).sheet, address);
          if (!cell) throw new Error("edited cell is missing from the projection");
          paths.set(cell.id, [
            "xlsx:sheets",
            sheet.id,
            "contents",
            cell.id.slice(sheet.id.length + 1),
          ]);
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
          const path = paths.get(id);
          if (!path) throw new Error("cell was not edited in this session");
          return { id, path };
        },
        exportBytes: async (determinism) =>
          doc.saveBytesAt(Date.parse(determinism.now) / 86_400_000 + 25569),
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
      const { storyId, paragraph } = pptxParagraph(doc, command.targetId);
      const text = checkReplacement(paragraph.text, command, "PPTX");
      const end = paragraph.start + paragraph.length;
      if (text)
        doc.insertTextJson(
          JSON.stringify({ storyId, index: end, text, style: paragraph.style })
        );
      if (paragraph.length)
        doc.deleteTextJson(
          JSON.stringify({ storyId, start: paragraph.start, end })
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
      const { storyId, paragraph } = pptxParagraph(doc, id);
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
export async function compare(
  baseBytes: Uint8Array,
  fromCheckpoint: OfficeCheckpoint,
  toCheckpoint: OfficeCheckpoint
): Promise<NetEffect[]> {
  if (fromCheckpoint.format !== toCheckpoint.format)
    throw new Error("Cannot compare different Office formats");
  const from = await open(fromCheckpoint.format, baseBytes, fromCheckpoint);
  try {
    const to = await open(toCheckpoint.format, baseBytes, toCheckpoint);
    try {
      const before = new Map(from.entries().map((entry) => [entry.id, entry]));
      const after = new Map(to.entries().map((entry) => [entry.id, entry]));
      const effects: NetEffect[] = [];
      for (const id of [
        ...new Set([...before.keys(), ...after.keys()]),
      ].sort()) {
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
        if (next?.asset) effect.imageSHA256 = next.asset.sha256;
        effects.push(effect);
      }
      return effects;
    } finally {
      to.dispose();
    }
  } finally {
    from.dispose();
  }
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
