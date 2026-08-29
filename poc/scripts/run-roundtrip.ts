import { createHash } from "node:crypto";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import process from "node:process";

import { parseDocx, repackDocx } from "@betteroffice/docx";
import {
  createYrsSession,
  yrsToDocument,
  type YrsSession,
} from "@betteroffice/docx/yrs";
import {
  initWasm as initPptxWasm,
  openPresentation,
  type DeckSnapshot,
  type ShapeSnapshot,
  type StorySnapshot,
} from "@betteroffice/pptx";
import { initWasm as initXlsxWasm, openWorkbook } from "@betteroffice/xlsx";
import JSZip from "jszip";

const root = path.resolve(import.meta.dir, "../..");
const fixtureDirectory = path.join(root, "poc/fixtures");
const outputDirectory = path.join(root, "poc/output");

const markers = {
  docx: ["EVO_EDIT_MARKER_DOCX", "EVO_EDIT_MARKER_DOCX_EDITED"],
  xlsx: ["EVO_EDIT_MARKER_XLSX", "EVO_EDITED_XLSX"],
  pptx: ["EVO_EDIT_MARKER_PPTX", "EVO_EDITED_PPTX"],
} as const;

interface PartDiff {
  inputParts: number;
  outputParts: number;
  added: string[];
  missing: string[];
  changed: string[];
  unchanged: number;
}

interface PhaseMetrics {
  milliseconds: number;
  rssDeltaBytes: number;
}

interface FormatResult {
  sourceBytes: number;
  noopBytes: number;
  editedBytes: number;
  view: Record<string, unknown>;
  open: PhaseMetrics;
  noOpSave: PhaseMetrics;
  editAndSave: PhaseMetrics;
  noOpParts: PartDiff;
  editedParts: PartDiff;
}

function bytes(file: Buffer): Uint8Array {
  return Uint8Array.from(file);
}

async function timed<T>(
  operation: () => T | Promise<T>
): Promise<[T, PhaseMetrics]> {
  const rssBefore = process.memoryUsage.rss();
  const startedAt = performance.now();
  const value = await operation();
  return [
    value,
    {
      milliseconds: Number((performance.now() - startedAt).toFixed(2)),
      rssDeltaBytes: process.memoryUsage.rss() - rssBefore,
    },
  ];
}

function digest(value: Uint8Array): string {
  return createHash("sha256").update(value).digest("hex");
}

async function zipParts(value: Uint8Array): Promise<Map<string, string>> {
  const archive = await JSZip.loadAsync(value);
  const entries = Object.values(archive.files).filter((entry) => !entry.dir);
  const parts = await Promise.all(
    entries.map(
      async (entry) =>
        [entry.name, digest(await entry.async("uint8array"))] as const
    )
  );
  return new Map(parts);
}

async function diffParts(
  input: Uint8Array,
  output: Uint8Array
): Promise<PartDiff> {
  const [before, after] = await Promise.all([
    zipParts(input),
    zipParts(output),
  ]);
  const added = [...after.keys()].filter((name) => !before.has(name)).sort();
  const missing = [...before.keys()].filter((name) => !after.has(name)).sort();
  const changed = [...before.entries()]
    .filter(([name, hash]) => after.has(name) && after.get(name) !== hash)
    .map(([name]) => name)
    .sort();
  return {
    inputParts: before.size,
    outputParts: after.size,
    added,
    missing,
    changed,
    unchanged: before.size - missing.length - changed.length,
  };
}

async function assertEmptySharedStringsWhenRemoved(
  source: Uint8Array,
  diff: PartDiff
): Promise<void> {
  if (!diff.missing.includes("xl/sharedStrings.xml")) return;
  const archive = await JSZip.loadAsync(source);
  const sharedStrings = await archive
    .file("xl/sharedStrings.xml")
    ?.async("text");
  assert(
    sharedStrings,
    "XLSX removed a shared-string part it could not inspect"
  );
  assert(
    !/<(?:[A-Za-z_][\w.-]*:)?si(?:\s|>)/.test(sharedStrings),
    "XLSX removed a non-empty shared-string table"
  );
}

function assert(condition: unknown, message: string): asserts condition {
  if (!condition) throw new Error(message);
}

function assertEquivalent(
  left: unknown,
  right: unknown,
  message: string
): void {
  assert(JSON.stringify(left) === JSON.stringify(right), message);
}

function bodyText(session: YrsSession): string {
  return session
    .paragraphs("body")
    .map((paragraph) => paragraph.text)
    .join("\n");
}

function replaceDocxMarker(
  session: YrsSession,
  from: string,
  to: string
): void {
  const paragraph = session
    .paragraphs("body")
    .find((candidate) => candidate.text.includes(from));
  assert(paragraph, `DOCX marker not found: ${from}`);
  const start = paragraph.text.indexOf(from);
  if (to.startsWith(from)) {
    session.insertText(
      { story: "body", paraId: paragraph.paraId, offset: start + from.length },
      to.slice(from.length)
    );
    return;
  }
  session.deleteRange({
    story: "body",
    start: { paraId: paragraph.paraId, offset: start },
    end: { paraId: paragraph.paraId, offset: start + from.length },
  });
  session.insertText(
    { story: "body", paraId: paragraph.paraId, offset: start },
    to
  );
}

async function saveDocx(
  source: Uint8Array,
  edit: boolean
): Promise<Uint8Array> {
  const parsed = await parseDocx(source.buffer as ArrayBuffer, {
    preloadFonts: false,
  });
  const session = await createYrsSession({ clientId: edit ? 81002 : 81001 });
  try {
    session.seedFromDocx(source);
    if (edit) replaceDocxMarker(session, markers.docx[0], markers.docx[1]);
    return new Uint8Array(await repackDocx(yrsToDocument(session, parsed)));
  } finally {
    session.destroy();
  }
}

async function verifyDocx(
  value: Uint8Array,
  marker: string
): Promise<Record<string, unknown>> {
  const parsed = await parseDocx(value.buffer as ArrayBuffer, {
    preloadFonts: false,
  });
  const session = await createYrsSession({ clientId: 81003 });
  try {
    session.seedFromDocx(value);
    const text = bodyText(session);
    assert(text.includes(marker), `DOCX did not reopen with ${marker}`);
    return {
      stories: session.storyIds().length,
      bodyParagraphs: session.paragraphs("body").length,
      sections: parsed.package.document.sections?.length ?? 0,
      headers: parsed.package.headers?.size ?? 0,
      footers: parsed.package.footers?.size ?? 0,
    };
  } finally {
    session.destroy();
  }
}

async function runDocx(): Promise<FormatResult> {
  const source = bytes(
    await readFile(path.join(fixtureDirectory, "feature-rich.docx"))
  );
  const [view, open] = await timed(() => verifyDocx(source, markers.docx[0]));
  const [noop, noOpSave] = await timed(() => saveDocx(source, false));
  await verifyDocx(noop, markers.docx[0]);
  const [edited, editAndSave] = await timed(() => saveDocx(source, true));
  await verifyDocx(edited, markers.docx[1]);
  await Promise.all([
    writeFile(path.join(outputDirectory, "feature-rich.noop.docx"), noop),
    writeFile(path.join(outputDirectory, "feature-rich.edited.docx"), edited),
  ]);
  return {
    sourceBytes: source.byteLength,
    noopBytes: noop.byteLength,
    editedBytes: edited.byteLength,
    view,
    open,
    noOpSave,
    editAndSave,
    noOpParts: await diffParts(source, noop),
    editedParts: await diffParts(source, edited),
  };
}

async function runXlsx(): Promise<FormatResult> {
  const source = bytes(
    await readFile(path.join(fixtureDirectory, "feature-rich.xlsx"))
  );
  await initXlsxWasm(
    bytes(
      await readFile(
        path.join(root, "packages/xlsx/dist/generated/xlsx_wasm_bg.wasm")
      )
    )
  );
  const [view, open] = await timed(() => {
    const handle = openWorkbook(source);
    try {
      const sheetInfo = handle.sheetInfo();
      const displayList = handle.displayList({
        x: 0,
        y: 0,
        width: 1400,
        height: 900,
      });
      assert(
        handle.cell(0, 8, 0).input === markers.xlsx[0],
        "XLSX marker cell is wrong"
      );
      assert(
        (displayList.charts?.length ?? 0) > 0,
        "XLSX chart did not reach the view model"
      );
      return {
        sheets: sheetInfo.sheetNames,
        commands: displayList.commands.length,
        charts: displayList.charts?.length ?? 0,
        formula: handle.cell(0, 4, 3).input,
      };
    } finally {
      handle.dispose();
    }
  });
  const sourceLogical = xlsxLogicalSnapshot(source);
  const [noop, noOpSave] = await timed(() => {
    const handle = openWorkbook(source);
    try {
      return handle.save();
    } finally {
      handle.dispose();
    }
  });
  const noOpHandle = openWorkbook(noop);
  try {
    assert(
      noOpHandle.cell(0, 8, 0).input === markers.xlsx[0],
      "XLSX no-op save lost marker"
    );
  } finally {
    noOpHandle.dispose();
  }
  assertEquivalent(
    xlsxLogicalSnapshot(noop),
    sourceLogical,
    "XLSX no-op save changed the logical workbook"
  );
  const [edited, editAndSave] = await timed(() => {
    const handle = openWorkbook(source);
    try {
      assert(
        handle.editCell(0, 8, 0, markers.xlsx[1]).applied,
        "XLSX edit was rejected"
      );
      return handle.save();
    } finally {
      handle.dispose();
    }
  });
  const editedHandle = openWorkbook(edited);
  try {
    assert(
      editedHandle.cell(0, 8, 0).input === markers.xlsx[1],
      "XLSX edit did not reopen"
    );
  } finally {
    editedHandle.dispose();
  }
  await Promise.all([
    writeFile(path.join(outputDirectory, "feature-rich.noop.xlsx"), noop),
    writeFile(path.join(outputDirectory, "feature-rich.edited.xlsx"), edited),
  ]);
  const [noOpParts, editedParts] = await Promise.all([
    diffParts(source, noop),
    diffParts(source, edited),
  ]);
  await assertEmptySharedStringsWhenRemoved(source, noOpParts);
  return {
    sourceBytes: source.byteLength,
    noopBytes: noop.byteLength,
    editedBytes: edited.byteLength,
    view,
    open,
    noOpSave,
    editAndSave,
    noOpParts,
    editedParts,
  };
}

function xlsxLogicalSnapshot(value: Uint8Array): unknown {
  const handle = openWorkbook(value);
  try {
    const sheetNames = handle.sheetInfo().sheetNames;
    return sheetNames.map((name, index) => {
      handle.setActiveSheet(index);
      const range = index === 0 ? "A1:H9" : "A1:D6";
      const info = handle.sheetInfo();
      const displayList = handle.displayList({
        x: 0,
        y: 0,
        width: 1400,
        height: 900,
      });
      return {
        name,
        cells: handle.rangeCells(index, range),
        mergedRanges: handle.mergedRanges(index, range),
        frozenRows: info.frozenRows,
        frozenCols: info.frozenCols,
        charts:
          displayList.charts?.map((chart) => ({
            id: chart.id,
            rect: chart.rect,
          })) ?? [],
        representativeFormatting: [
          handle.selectionFormatting(index, index === 0 ? "A1:H2" : "A1:D1"),
          handle.selectionFormatting(index, index === 0 ? "A4:D4" : "A2:D6"),
        ],
      };
    });
  } finally {
    handle.dispose();
  }
}

function flattenShapes(shapes: ShapeSnapshot[]): ShapeSnapshot[] {
  return shapes.flatMap((shape) => [shape, ...flattenShapes(shape.children)]);
}

function storyText(story: StorySnapshot): string {
  return story.paragraphs
    .map((paragraph) => paragraph.runs.map((run) => run.text).join(""))
    .join("\n");
}

function findPresentationMarker(
  snapshot: DeckSnapshot,
  marker: string
): StorySnapshot {
  for (const slide of snapshot.slides) {
    for (const shape of flattenShapes(slide.shapes)) {
      const story = shape.textStories.find((candidate) =>
        storyText(candidate).includes(marker)
      );
      if (story) return story;
    }
  }
  throw new Error(`PPTX marker not found: ${marker}`);
}

async function runPptx(): Promise<FormatResult> {
  const source = bytes(
    await readFile(path.join(fixtureDirectory, "feature-rich.pptx"))
  );
  const font = bytes(
    await readFile(
      path.join(
        root,
        "crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
      )
    )
  );
  await initPptxWasm(
    bytes(
      await readFile(
        path.join(root, "packages/pptx/dist/generated/pptx_wasm_bg.wasm")
      )
    )
  );
  const [view, open] = await timed(() => {
    const handle = openPresentation(source, {
      clientId: 82001,
      fonts: [{ family: "Liberation Sans", bytes: font }],
    });
    try {
      const snapshot = handle.snapshot();
      const frames = snapshot.slides.map((_, index) =>
        handle.layoutSlide(index)
      );
      findPresentationMarker(snapshot, markers.pptx[0]);
      assert(
        frames.every((frame) => frame.primitives.length > 0),
        "PPTX view model is empty"
      );
      return {
        slides: snapshot.slides.length,
        shapes: snapshot.slides.reduce(
          (count, slide) => count + flattenShapes(slide.shapes).length,
          0
        ),
        primitives: frames.reduce(
          (count, frame) => count + frame.primitives.length,
          0
        ),
      };
    } finally {
      handle.dispose();
    }
  });
  const [noop, noOpSave] = await timed(() => {
    const handle = openPresentation(source, { clientId: 82002 });
    try {
      return handle.save();
    } finally {
      handle.dispose();
    }
  });
  const noOpHandle = openPresentation(noop, { clientId: 82003 });
  try {
    findPresentationMarker(noOpHandle.snapshot(), markers.pptx[0]);
  } finally {
    noOpHandle.dispose();
  }
  const [edited, editAndSave] = await timed(() => {
    const handle = openPresentation(source, { clientId: 82004 });
    try {
      const story = findPresentationMarker(handle.snapshot(), markers.pptx[0]);
      const text = storyText(story);
      const start = text.indexOf(markers.pptx[0]);
      handle.deleteText(story.id, start, start + markers.pptx[0].length);
      handle.insertText(story.id, start, markers.pptx[1]);
      return handle.save();
    } finally {
      handle.dispose();
    }
  });
  const editedHandle = openPresentation(edited, { clientId: 82005 });
  try {
    findPresentationMarker(editedHandle.snapshot(), markers.pptx[1]);
  } finally {
    editedHandle.dispose();
  }
  await Promise.all([
    writeFile(path.join(outputDirectory, "feature-rich.noop.pptx"), noop),
    writeFile(path.join(outputDirectory, "feature-rich.edited.pptx"), edited),
  ]);
  return {
    sourceBytes: source.byteLength,
    noopBytes: noop.byteLength,
    editedBytes: edited.byteLength,
    view,
    open,
    noOpSave,
    editAndSave,
    noOpParts: await diffParts(source, noop),
    editedParts: await diffParts(source, edited),
  };
}

await mkdir(outputDirectory, { recursive: true });
const generatedAt = new Date().toISOString();
const docx = await runDocx();
const xlsx = await runXlsx();
const pptx = await runPptx();
const report = {
  generatedAt,
  runtime: { bun: Bun.version, platform: process.platform, arch: process.arch },
  docx,
  xlsx,
  pptx,
};
await writeFile(
  path.join(outputDirectory, "roundtrip-report.json"),
  `${JSON.stringify(report, null, 2)}\n`
);

const allowedNoOpRemovals: Record<string, string[]> = {
  docx: [],
  xlsx: ["xl/sharedStrings.xml"],
  pptx: [],
};

for (const [format, result] of Object.entries({ docx, xlsx, pptx })) {
  const unexpectedMissingParts = result.noOpParts.missing.filter(
    (name) => !allowedNoOpRemovals[format].includes(name)
  );
  assert(
    unexpectedMissingParts.length === 0,
    `${format} no-op save removed OOXML parts`
  );
  console.log(
    `${format.toUpperCase()}: open ${result.open.milliseconds} ms; ` +
      `no-op ${result.noOpSave.milliseconds} ms; edit ${result.editAndSave.milliseconds} ms; ` +
      `${result.noOpParts.changed.length}/${result.noOpParts.inputParts} parts changed on no-op` +
      (result.noOpParts.missing.length > 0
        ? `; removed allowed inert part ${result.noOpParts.missing.join(", ")}`
        : "")
  );
}
console.log(`Report: ${path.join(outputDirectory, "roundtrip-report.json")}`);
