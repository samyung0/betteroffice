import { test, expect } from "bun:test";
import { readFile, readdir } from "node:fs/promises";
import {
  seedOffice,
  compare,
  exportOffice,
  resolveAsset,
} from "./office-checkpoint";
import { createYrsSession } from "../packages/docx/src/yrs";
import { XlsxDocument } from "../packages/xlsx/src/wasm/generated/xlsx_wasm.js";
import { PptxDocument } from "../packages/pptx/src/wasm/generated/pptx_wasm.js";
import { unzipContainer, rezipContainer } from "../packages/docx/src/wasm/opc";
const fixed = { seed: "0".repeat(64), now: "2026-09-06T00:00:00.000Z" };
const fixture = (name: string) =>
  readFile(new URL(`../apps/demo/public/${name}`, import.meta.url));

test("XLSX checkpoints restore sparse edits, cancel net effects after undo, and reject a different exact base", async () => {
  const bytes = await fixture("sample.xlsx");
  const before = await seedOffice("xlsx", bytes);
  const doc = XlsxDocument.openCollaborative(bytes, 9991);
  try {
    doc.applyUpdateJson(before.state);
    doc.editCellJson(
      JSON.stringify({ sheet: 0, row: 0, col: 0, input: "checkpoint value" })
    );
    const after = { ...before, state: doc.encodeStateAsUpdate() };
    expect(
      (await compare(bytes, before, after)).some((effect) =>
        effect.after?.includes("checkpoint value")
      )
    ).toBe(true);
    const exported = await exportOffice(bytes, after, fixed);
    expect(await exportOffice(bytes, after, fixed)).toEqual(exported);
    const reopened = XlsxDocument.open(exported);
    try {
      expect(
        reopened.cellJson(JSON.stringify({ sheet: 0, row: 0, col: 0 }))
      ).toContain("checkpoint value");
    } finally {
      reopened.free();
    }
    doc.undoJson();
    expect(
      await compare(bytes, before, {
        ...before,
        state: doc.encodeStateAsUpdate(),
      })
    ).toEqual([]);
    const changedBase = Uint8Array.from(bytes);
    changedBase[changedBase.length - 1] ^= 1;
    await expect(exportOffice(changedBase, after, fixed)).rejects.toThrow(
      "exact base"
    );
  } finally {
    doc.free();
  }
});

test("PPTX checkpoint exports preserve opaque parts and mismatched source restoration fails immediately", async () => {
  const bytes = await fixture("betteroffice-demo.pptx");
  const before = await seedOffice("pptx", bytes);
  expect(await compare(bytes, before, before)).toEqual([]);
  const exported = await exportOffice(bytes, before, fixed);
  const wrong = Uint8Array.from(bytes);
  wrong[wrong.length - 1] ^= 1;
  expect(() =>
    PptxDocument.openCollaborativeFromUpdate(before.state, 9992, wrong)
  ).toThrow();
  const restored = PptxDocument.openCollaborativeFromUpdate(
    before.state,
    9993,
    bytes
  );
  try {
    expect(restored.saveBytes()).toEqual(exported);
  } finally {
    restored.free();
  }
});

test("DOCX inserted image bytes resolve by stable placement and survive fresh checkpoint export", async () => {
  const bytes = await fixture("betteroffice-demo.docx");
  const before = await seedOffice("docx", bytes);
  const doc = await createYrsSession({ clientId: 9994 });
  try {
    doc.openDocx(bytes, false);
    doc.loadState(before.state);
    const paragraph = doc.paragraphs("body")[0];
    doc.insertImage(
      { story: "body", paraId: paragraph.paraId, offset: 0 },
      {
        src: "data:image/png;base64,AQID",
        rId: "rId_img_temporary",
        width: 20,
        height: 20,
      }
    );
    const after = { ...before, state: doc.encodeState() };
    const effects = await compare(bytes, before, after);
    const effect = effects.find(
      (effect) => effect.kind === "image" && effect.operation === "add"
    );
    expect(effect?.assetRef).toBeDefined();
    const asset = await resolveAsset(bytes, after, effect!.assetRef!);
    expect([...asset.bytes]).toEqual([1, 2, 3]);
    const exported = await exportOffice(bytes, after, fixed);
    const parts = unzipContainer(exported);
    expect(
      Object.entries(parts).some(
        ([path, data]) =>
          path.startsWith("word/media/") &&
          data.length === 3 &&
          data[0] === 1 &&
          data[2] === 3
      )
    ).toBe(true);
    expect(await exportOffice(bytes, after, fixed)).toEqual(exported);
    doc.applyRawOps("body", [
      {
        op: "setEmbedAttr",
        index: 0,
        key: "src",
        value: "data:image/png;base64,BAUG",
      },
    ]);
    const replaced = { ...before, state: doc.encodeState() };
    const replacement = (await compare(bytes, after, replaced)).find(
      (effect) => effect.kind === "image"
    );
    expect(replacement?.operation).toBe("replace");
    expect(replacement?.id).toBe(effect!.id);
    expect(replacement?.assetRef).toEqual(effect!.assetRef);
    expect(replacement?.imageSHA256).not.toBe(effect!.imageSHA256);
    expect(replacement?.imageSHA256).toMatch(/^[a-f0-9]{64}$/);
    const newAsset = await resolveAsset(
      bytes,
      replaced,
      replacement!.assetRef!
    );
    expect([...newAsset.bytes]).toEqual([4, 5, 6]);
    expect(newAsset.sha256).toBe(replacement!.imageSHA256!);
    const replacementParts = unzipContainer(
      await exportOffice(bytes, replaced, fixed)
    );
    expect(
      Object.entries(replacementParts).some(
        ([path, data]) =>
          path.startsWith("word/media/") &&
          data.length === 3 &&
          data[0] === 4 &&
          data[2] === 6
      )
    ).toBe(true);
    doc.applyRawOps("body", [{ op: "delete", index: 0, len: 1 }]);
    const removal = (
      await compare(bytes, replaced, { ...before, state: doc.encodeState() })
    ).find((effect) => effect.kind === "image");
    expect(removal?.operation).toBe("remove");
    expect(removal?.assetRef).toBeUndefined();
    expect(removal?.imageSHA256).toBeUndefined();
  } finally {
    doc.destroy();
  }
});

test("PPTX text edits and image removal survive headless restore without changing unrelated media", async () => {
  const bytes = await fixture("betteroffice-demo.pptx");
  const before = await seedOffice("pptx", bytes);
  const doc = PptxDocument.openCollaborativeFromUpdate(
    before.state,
    9995,
    bytes
  );
  try {
    const deck = JSON.parse(
      doc.snapshotJson()
    ) as import("../packages/pptx/src/types").DeckSnapshot;
    const shapes = deck.slides.flatMap((slide) =>
      slide.shapes.map((shape) => ({ slide, shape }))
    );
    const text = shapes.find(({ shape }) => shape.textStories.length > 0)!;
    const picture = shapes.find(({ shape }) => shape.mediaPartPath)!;
    expect(picture).toBeDefined();
    const ref = {
      format: "pptx" as const,
      kind: "image" as const,
      slideId: picture.slide.id,
      id: picture.shape.id,
    };
    const sourceAsset = await resolveAsset(bytes, before, ref);
    expect(sourceAsset.bytes.length).toBeGreaterThan(0);
    doc.insertTextJson(
      JSON.stringify({
        storyId: text.shape.textStories[0].id,
        index: 0,
        text: "Fresh checkpoint text. ",
      })
    );
    const edited = { ...before, state: doc.encodeStateAsUpdate() };
    expect(
      (await compare(bytes, before, edited)).some(
        (effect) =>
          effect.kind === "text" &&
          effect.after?.includes("Fresh checkpoint text.")
      )
    ).toBe(true);
    const output = await exportOffice(bytes, edited, fixed);
    const originalParts = unzipContainer(bytes);
    const outputParts = unzipContainer(output);
    expect(outputParts[picture.shape.mediaPartPath!]).toEqual(
      originalParts[picture.shape.mediaPartPath!]
    );
    const fresh = PptxDocument.openCollaborative(output, 9996);
    try {
      expect(fresh.snapshotJson()).toContain("Fresh checkpoint text.");
    } finally {
      fresh.free();
    }
    doc.removeShapeJson(
      JSON.stringify({ slideId: picture.slide.id, shapeId: picture.shape.id })
    );
    const removed = { ...before, state: doc.encodeStateAsUpdate() };
    expect(
      (await compare(bytes, edited, removed)).some(
        (effect) => effect.kind === "image" && effect.operation === "remove"
      )
    ).toBe(true);
    await expect(resolveAsset(bytes, removed, ref)).rejects.toThrow();
    expect(await exportOffice(bytes, removed, fixed)).toEqual(
      await exportOffice(bytes, removed, fixed)
    );
  } finally {
    doc.free();
  }
});

test("XLSX source drawing images resolve from the current sheet placement and preserve exact media", async () => {
  const original = await fixture("sample.xlsx");
  await seedOffice("xlsx", original);
  const { rezipContainer } = await import("../packages/docx/src/wasm/opc");
  const parts = unzipContainer(original);
  const encode = (text: string) => new TextEncoder().encode(text);
  parts["xl/worksheets/sheet1.xml"] = encode(
    new TextDecoder()
      .decode(parts["xl/worksheets/sheet1.xml"])
      .replace(
        "</worksheet>",
        '<drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdImage"/></worksheet>'
      )
  );
  parts["xl/worksheets/_rels/sheet1.xml.rels"] = encode(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>'
  );
  parts["xl/drawings/drawing1.xml"] = encode(
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Source image"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>'
  );
  parts["xl/drawings/_rels/drawing1.xml.rels"] = encode(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/source.png"/></Relationships>'
  );
  parts["xl/media/source.png"] = Uint8Array.of(1, 2, 3);
  parts["[Content_Types].xml"] = encode(
    new TextDecoder()
      .decode(parts["[Content_Types].xml"])
      .replace(
        "</Types>",
        '<Default Extension="png" ContentType="image/png"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>'
      )
  );
  const bytes = rezipContainer(parts);
  const checkpoint = await seedOffice("xlsx", bytes);
  const doc = XlsxDocument.openCollaborative(bytes, 9997);
  try {
    doc.applyUpdateJson(checkpoint.state);
    const projection = JSON.parse(doc.checkpointProjectionJson()) as {
      sheets: Array<{ id: string; images: Array<{ id: string }> }>;
    };
    const sheet = projection.sheets.find((sheet) => sheet.images.length)!;
    const asset = await resolveAsset(bytes, checkpoint, {
      format: "xlsx",
      kind: "image",
      sheetId: sheet.id,
      id: sheet.images[0].id,
    });
    expect([...asset.bytes]).toEqual([1, 2, 3]);
    expect(
      unzipContainer(await exportOffice(bytes, checkpoint, fixed))[
        "xl/media/source.png"
      ]
    ).toEqual(parts["xl/media/source.png"]);
  } finally {
    doc.free();
  }
});

async function docxFixture(name: string): Promise<Uint8Array> {
  const source = await fixture("betteroffice-demo.docx");
  await seedOffice("docx", source);
  const parts = unzipContainer(source);
  const root = new URL(
    `../packages/docx/src/yrs/__fixtures__/seed-parity/${name}/`,
    import.meta.url
  );
  const overlay = async (directory: URL, prefix = ""): Promise<void> => {
    await Promise.all(
      (
        await readdir(directory, { withFileTypes: true })
      ).map(async (entry) => {
        const path = new URL(
          entry.name + (entry.isDirectory() ? "/" : ""),
          directory
        );
        if (entry.isDirectory()) await overlay(path, prefix + entry.name + "/");
        else parts[prefix + entry.name] = await readFile(path);
      })
    );
  };
  await overlay(root);
  return rezipContainer(parts);
}

function embedLocation(
  doc: Awaited<ReturnType<typeof createYrsSession>>,
  story: string,
  matches: (segment: ReturnType<typeof doc.storySegments>[number]) => boolean
) {
  let offset = 0;
  for (const segment of doc.storySegments(story)) {
    if (matches(segment)) {
      for (const paragraph of doc.paragraphs(story)) {
        const span = doc.locateParagraph(story, paragraph.paraId);
        if (span.start <= offset && span.end > offset)
          return {
            absolute: offset,
            at: {
              story,
              paraId: paragraph.paraId,
              offset: offset - span.start,
            },
            segment,
          };
      }
      throw new Error("Embed has no containing paragraph");
    }
    offset += segment.kind === "text" ? segment.text.length : 1;
  }
  throw new Error("Embed was not found");
}

test("deleting every DOCX comment clears source and reply parts across export and fresh seeding", async () => {
  const bytes = await docxFixture("comments");
  const before = await seedOffice("docx", bytes);
  const doc = await createYrsSession({ clientId: 10001 });
  const fresh = await createYrsSession({ clientId: 10002 });
  try {
    doc.openDocx(bytes, false);
    doc.loadState(before.state);
    const comments = doc.listComments();
    expect(comments.length).toBeGreaterThan(0);
    expect(comments.some((comment) => !!comment.parentId)).toBe(true);
    doc.applyRawOps(
      "body",
      comments.map((comment) => ({
        op: "removeComment" as const,
        id: comment.id,
      }))
    );
    const after = { ...before, state: doc.encodeState() };
    expect(doc.listComments()).toEqual([]);
    expect(
      (await compare(bytes, before, after)).filter(
        (effect) => effect.kind === "text" && effect.operation === "remove"
      )
    ).toHaveLength(comments.length);
    const output = await exportOffice(bytes, after, fixed);
    const parts = unzipContainer(output);
    for (const part of [
      "comments.xml",
      "commentsExtended.xml",
      "commentsIds.xml",
      "commentsExtensible.xml",
    ]) {
      const xml = new TextDecoder().decode(parts[`word/${part}`]);
      expect(xml).not.toContain("paraId=");
      expect(xml).not.toContain("durableId=");
      expect(xml).not.toContain("w:id=");
    }
    const reseeded = await seedOffice("docx", output);
    fresh.openDocx(output, false);
    fresh.loadState(reseeded.state);
    expect(fresh.listComments()).toEqual([]);
  } finally {
    doc.destroy();
    fresh.destroy();
  }
});

test("DOCX dropdown, date and plaintext evidence uses the exact text exported by the native projection", async () => {
  const bytes = await docxFixture("content-controls");
  const before = await seedOffice("docx", bytes);
  const doc = await createYrsSession({ clientId: 10003 });
  try {
    doc.openDocx(bytes, false);
    doc.loadState(before.state);
    let previous = before;
    for (const change of [
      {
        tag: "choice",
        before: "Beta",
        after: "Alpha",
        value: { kind: "dropdown" as const, value: "alpha" },
      },
      {
        tag: "when",
        before: "4 March 2026",
        after: "6 September 2027",
        value: { kind: "date" as const, date: "2027-09-06" },
      },
      {
        tag: "inlineText",
        before: "inline plain text value",
        after: "  New plain text — exact value  ",
        value: undefined,
      },
    ]) {
      const location = embedLocation(
        doc,
        "body",
        (segment) =>
          segment.kind === "embed" && segment.payload.tag === change.tag
      );
      if (change.value) doc.setContentControlValueAt(location.at, change.value);
      else
        doc.applyRawOps("body", [
          {
            op: "setEmbedAttr",
            index: location.absolute,
            key: "content",
            value: [{ kind: "text", text: change.after, attrs: {} }],
          },
        ]);
      const current = { ...before, state: doc.encodeState() };
      const effects = (await compare(bytes, previous, current)).filter(
        (effect) => effect.kind === "text"
      );
      expect(effects).toHaveLength(1);
      expect(effects[0].operation).toBe("replace");
      expect(effects[0].before).toBe(change.before);
      expect(effects[0].after).toBe(change.after);
      const output = await exportOffice(bytes, current, fixed);
      const fresh = await createYrsSession({ clientId: 10004 });
      try {
        fresh.openDocx(output, true);
        const saved = fresh
          .storySegments("body")
          .find(
            (segment) =>
              segment.kind === "embed" && segment.payload.tag === change.tag
          );
        expect(saved?.kind).toBe("embed");
        if (saved?.kind !== "embed") throw new Error("Missing saved control");
        const content = saved.payload.content as Array<{
          kind: string;
          text?: string;
        }>;
        expect(
          content
            .filter((item) => item.kind === "text")
            .map((item) => item.text ?? "")
            .join("")
        ).toBe(change.after);
      } finally {
        fresh.destroy();
      }
      previous = current;
    }
  } finally {
    doc.destroy();
  }
});

test("DOCX table range deletion removes reachable text and image evidence while Undo restores it", async () => {
  const bytes = await fixture("betteroffice-demo.docx");
  const seed = await seedOffice("docx", bytes);
  for (const mode of ["range", "command"] as const) {
    const doc = await createYrsSession({ clientId: 10005 });
    try {
      doc.openDocx(bytes, false);
      doc.loadState(seed.state);
      const child = doc
        .storyIds()
        .find((story) => story.startsWith("body:t0:"))!;
      const paragraph = doc.paragraphs(child)[0];
      doc.insertImage(
        { story: child, paraId: paragraph.paraId, offset: 0 },
        {
          src: "data:image/png;base64,AQID",
          rId: "rId_nested",
          width: 20,
          height: 20,
        }
      );
      const before = { ...seed, state: doc.encodeState() };
      const image = (await compare(bytes, seed, before)).find(
        (effect) => effect.kind === "image"
      )!;
      expect(await resolveAsset(bytes, before, image.assetRef!)).toMatchObject({
        bytes: Uint8Array.of(1, 2, 3),
      });
      const cellText = doc
        .storyIds()
        .filter((story) => story.startsWith("body:t0:"))
        .flatMap((story) =>
          doc.paragraphs(story).map((paragraph) => paragraph.text)
        );
      if (mode === "range") {
        const location = embedLocation(
          doc,
          "body",
          (segment) => segment.kind === "embed" && segment.embedKind === "table"
        );
        doc.deleteRange({
          story: "body",
          start: location.at,
          end: { ...location.at, offset: location.at.offset + 1 },
        });
        expect(doc.storyIds()).toContain(child);
      } else doc.deleteTable({ story: "body", tableIndex: 0 });
      const after = { ...seed, state: doc.encodeState() };
      const effects = await compare(bytes, before, after);
      expect(
        effects
          .filter(
            (effect) => effect.kind === "text" && effect.operation === "remove"
          )
          .map((effect) => effect.before)
          .sort()
      ).toEqual([...cellText].sort());
      expect(
        effects.some(
          (effect) => effect.kind === "image" && effect.operation === "remove"
        )
      ).toBe(true);
      await expect(
        resolveAsset(bytes, after, image.assetRef!)
      ).rejects.toThrow();
      const output = unzipContainer(await exportOffice(bytes, after, fixed));
      expect(
        new TextDecoder().decode(output["word/document.xml"])
      ).not.toContain("<w:tbl>");
      doc.undo();
      const restored = { ...seed, state: doc.encodeState() };
      expect(
        (await compare(bytes, before, restored)).filter(
          (effect) => effect.kind === "text"
        )
      ).toEqual([]);
      const restoredImage = (await compare(bytes, after, restored)).find(
        (effect) => effect.kind === "image"
      )!;
      if (mode === "range")
        expect(restoredImage.assetRef).toEqual(image.assetRef);
      expect(
        (await resolveAsset(bytes, restored, restoredImage.assetRef!)).bytes
      ).toEqual(Uint8Array.of(1, 2, 3));
    } finally {
      doc.destroy();
    }
  }
});
