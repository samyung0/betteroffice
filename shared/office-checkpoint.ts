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
import type { DeckSnapshot, ShapeSnapshot } from "../packages/pptx/src/types";

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
      return {
        state: () => doc.encodeStateAsUpdate(),
        entries: () => xlsxEntries(doc),
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
