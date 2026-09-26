import { test, expect } from "bun:test";
import { readFile, readdir } from "node:fs/promises";
import {
  applyOfficeCommands,
  locateOfficeTargets,
  seedOffice,
  officeBaseline,
  compareBaselines,
  compare,
  exportOffice,
  inspectOffice,
  rebaseOffice,
  resolveAsset,
  xlsxPendingEffects,
} from "./office-checkpoint";
import { createYrsSession } from "../packages/docx/src/yrs";
import { XlsxDocument } from "../packages/xlsx/src/wasm/generated/xlsx_wasm.js";
import { PptxDocument } from "../packages/pptx/src/wasm/generated/pptx_wasm.js";
import { unzipContainer, rezipContainer } from "../packages/docx/src/wasm/opc";
import { Window } from "happy-dom";
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
    expect(await xlsxPendingEffects(bytes, after)).toEqual([
      expect.objectContaining({
        kind: "text",
        after: '{"kind":"text","value":"checkpoint value"}',
      }),
    ]);
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
      await xlsxPendingEffects(bytes, {
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

test("XLSX publications rebase later edits as overrides and report them as effects", async () => {
  const bytes = await fixture("sample.xlsx");
  const seeded = await seedOffice("xlsx", bytes);
  const doc = XlsxDocument.openCollaborative(bytes, 9990);
  const edit = (row: number, input: string) =>
    doc.editCellJson(JSON.stringify({ sheet: 0, row, col: 0, input }));
  try {
    doc.applyUpdateJson(seeded.state);
    edit(0, "captured");
    const captured = { ...seeded, state: doc.encodeStateAsUpdate() };
    const published = await exportOffice(bytes, captured, fixed);
    edit(1, "later");
    const latest = { ...seeded, state: doc.encodeStateAsUpdate() };
    const rebased = await rebaseOffice(bytes, captured, latest, published);
    expect(rebased.baseline).toEqual([]);
    expect(rebased.effects).toEqual([
      expect.objectContaining({ after: '{"kind":"text","value":"later"}' }),
    ]);
    const base = await seedOffice("xlsx", published);
    const reopened = XlsxDocument.open(
      await exportOffice(published, { ...base, state: rebased.state }, fixed)
    );
    try {
      for (const [row, input] of [
        [0, "captured"],
        [1, "later"],
      ] as const)
        expect(
          reopened.cellJson(JSON.stringify({ sheet: 0, row, col: 0 }))
        ).toContain(input);
    } finally {
      reopened.free();
    }
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
    expectWellFormedWithoutComments(output);
    const reseeded = await seedOffice("docx", output);
    fresh.openDocx(output, false);
    fresh.loadState(reseeded.state);
    expect(fresh.listComments()).toEqual([]);
  } finally {
    doc.destroy();
    fresh.destroy();
  }
});

const xmlParser = new new Window().DOMParser();
/** Every XML part parses, and no comment part, relationship or override is left. */
function expectWellFormedWithoutComments(docx: Uint8Array) {
  const parts = unzipContainer(docx);
  const text = (path: string) => new TextDecoder().decode(parts[path]);
  for (const path of Object.keys(parts).filter((path) => /\.(xml|rels)$/.test(path))) {
    // happy-dom rejects a valid single-quoted XML declaration; the elements are what is checked.
    const xml = text(path).replace(/^<\?xml[^?]*\?>/, "");
    const parsed = xmlParser.parseFromString(xml, "application/xml");
    expect([path, parsed.getElementsByTagName("parsererror").length]).toEqual([path, 0]);
  }
  expect(Object.keys(parts).filter((path) => path.startsWith("word/comments"))).toEqual([]);
  expect(text("[Content_Types].xml")).not.toContain("comments");
  expect(text("word/_rels/document.xml.rels")).not.toContain("comments");
}

test("an edited DOCX without comments exports no comment parts and only well-formed XML", async () => {
  const bytes = new Uint8Array(
    await readFile(new URL("../poc/fixtures/feature-rich.docx", import.meta.url))
  );
  const before = await seedOffice("docx", bytes);
  const doc = await createYrsSession({ clientId: 10003 });
  try {
    doc.openDocx(bytes, false);
    doc.loadState(before.state);
    const [first] = doc.paragraphs("body");
    doc.insertText({ story: "body", paraId: first.paraId, offset: 0 }, "Edited ");
    expectWellFormedWithoutComments(
      await exportOffice(bytes, { ...before, state: doc.encodeState() }, fixed)
    );
  } finally {
    doc.destroy();
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

test("agent edits replace one DOCX paragraph, locate it for guards and undo through the inverse", async () => {
  const bytes = await fixture("betteroffice-demo.docx");
  const before = await seedOffice("docx", bytes);
  const entries = await inspectOffice(bytes, before);
  const target = entries.find((entry) => entry.value.length > 0)!;
  expect(target.id).toContain(":paragraph:");
  const edit = await applyOfficeCommands(bytes, before, [
    {
      type: "replace_text",
      targetId: target.id,
      expectedText: target.value,
      text: "Edited by the agent",
    },
  ]);
  const after = { ...before, state: edit.state };
  expect(edit.targets).toEqual([
    {
      id: target.id,
      path: ["stories", target.id.split(":paragraph:")[0]],
      range: [expect.any(Number), expect.any(Number)],
    },
  ]);
  expect(edit.targets[0].range![1] - edit.targets[0].range![0]).toBe(
    "Edited by the agent".length
  );
  const effects = await compare(bytes, before, after);
  expect(effects.find((effect) => effect.id === target.id)).toMatchObject({
    operation: "replace",
    before: target.value,
    after: "Edited by the agent",
  });
  await expect(
    applyOfficeCommands(bytes, after, [
      {
        type: "replace_text",
        targetId: target.id,
        expectedText: target.value,
        text: "stale",
      },
    ])
  ).rejects.toThrow("stale_target");
  const undone = await applyOfficeCommands(bytes, after, edit.inverse);
  expect(
    (await inspectOffice(bytes, { ...before, state: undone.state })).find(
      (entry) => entry.id === target.id
    )?.value
  ).toBe(target.value);
});

test("agent replace_text rewrites only the changed DOCX span and refuses tracked insertions", async () => {
  const source = await fixture("betteroffice-demo.docx");
  await seedOffice("docx", source);
  const parts = unzipContainer(source);
  const run = (text: string, bold = false) =>
    `<w:r>${
      bold ? "<w:rPr><w:b/></w:rPr>" : ""
    }<w:t xml:space="preserve">${text}</w:t></w:r>`;
  parts["word/document.xml"] = new TextEncoder().encode(
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>` +
      `<w:p w14:paraId="0000000A">${run("I like ")}${run("cats", true)}${run(
        " a lot"
      )}</w:p>` +
      `<w:p w14:paraId="0000000B">${run(
        "Kept "
      )}<w:ins w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z">${run(
        "inserted"
      )}</w:ins></w:p>` +
      `<w:sectPr/></w:body></w:document>`
  );
  const bytes = rezipContainer(parts);
  const before = await seedOffice("docx", bytes);
  const entries = await inspectOffice(bytes, before);
  const liked = entries.find((entry) => entry.value === "I like cats a lot")!;
  expect(entries.some((entry) => entry.value.includes("inserted"))).toBe(false);
  const edit = await applyOfficeCommands(bytes, before, [
    {
      type: "replace_text",
      targetId: liked.id,
      expectedText: liked.value,
      text: "I like dogs a lot",
    },
  ]);
  const xml = new TextDecoder().decode(
    unzipContainer(
      await exportOffice(bytes, { ...before, state: edit.state }, fixed)
    )["word/document.xml"]
  );
  const runs = [...xml.matchAll(/<w:r>(.*?)<\/w:r>/g)].map(([, body]) => ({
    bold: body.includes("<w:b/>"),
    text: body.replace(/<[^>]+>/g, ""),
  }));
  expect(runs.filter((item) => item.bold).map((item) => item.text)).toEqual([
    "dogs",
  ]);
  expect(runs.map((item) => item.text).join("")).toContain("I like dogs a lot");
  await expect(
    applyOfficeCommands(bytes, before, [
      {
        type: "replace_text",
        targetId: liked.id.replace(/:paragraph:.*/, ":paragraph:0000000B"),
        expectedText: "Kept inserted",
        text: "Kept",
      },
    ])
  ).rejects.toThrow("unavailable_target");
});

test("DOCX charts and unmodeled VML drawings survive seeding, an edit and repeated export", async () => {
  const xmlOf = (bytes: Uint8Array, part = "word/document.xml") =>
    new TextDecoder().decode(unzipContainer(bytes)[part]);
  const charts = (xml: string) =>
    [...xml.matchAll(/<c:chart [^>]*r:id="([^"]+)"/g)].map(([, id]) => id);
  const vml = (xml: string) =>
    xml.match(/<w:pict><v:rect [\s\S]*?<\/w:pict>/g) ?? [];
  for (const name of ["exchange-plan.docx", "opaque-objects.docx"]) {
    const source = new Uint8Array(
      await readFile(new URL(`../poc/fixtures/${name}`, import.meta.url))
    );
    expect(charts(xmlOf(source))).toHaveLength(2);
    let bytes = source;
    for (const round of [1, 2]) {
      const seeded = await seedOffice("docx", bytes);
      const target = (await inspectOffice(bytes, seeded)).find(
        (entry) => entry.value.length > 20
      )!;
      const edit = await applyOfficeCommands(bytes, seeded, [
        {
          type: "replace_text",
          targetId: target.id,
          expectedText: target.value,
          text: `${target.value} (round ${round})`,
        },
      ]);
      bytes = await exportOffice(
        bytes,
        { ...seeded, state: edit.state },
        fixed
      );
      const xml = xmlOf(bytes);
      expect(xml).toContain(`(round ${round})`);
      expect(charts(xml)).toEqual(charts(xmlOf(source)));
      for (const id of charts(xml))
        expect(xmlOf(bytes, "word/_rels/document.xml.rels")).toContain(
          `Id="${id}"`
        );
      expect(vml(xml)).toEqual(vml(xmlOf(source)));
    }
  }
});

test("DOCX OLE objects and text boxes export their authored XML until edited", async () => {
  const bytes = new Uint8Array(
    await readFile(
      new URL("../poc/fixtures/opaque-objects.docx", import.meta.url)
    )
  );
  const xmlOf = (docx: Uint8Array, part = "word/document.xml") =>
    new TextDecoder().decode(unzipContainer(docx)[part]);
  const alternates = (xml: string) =>
    xml.match(/<mc:AlternateContent>[\s\S]*?<\/mc:AlternateContent>/g) ?? [];
  const source = xmlOf(bytes);
  expect(alternates(source)).toHaveLength(2);
  const seeded = await seedOffice("docx", bytes);
  const target = (await inspectOffice(bytes, seeded)).find(
    (entry) => entry.value.length > 20
  )!;
  const edit = await applyOfficeCommands(bytes, seeded, [
    {
      type: "replace_text",
      targetId: target.id,
      expectedText: target.value,
      text: `${target.value} edited`,
    },
  ]);
  const unedited = await exportOffice(
    bytes,
    { ...seeded, state: edit.state },
    fixed
  );
  const xml = xmlOf(unedited);
  expect(xml.match(/<w:object [\s\S]*?<\/w:object>/g)).toEqual(
    source.match(/<w:object [\s\S]*?<\/w:object>/g)
  );
  expect(xml).toContain('<o:OLEObject Type="Embed"');
  expect(xmlOf(unedited, "word/_rels/document.xml.rels")).toContain(
    'Id="rIdOle"'
  );
  expect(alternates(xml)).toEqual(alternates(source));

  const session = await createYrsSession({ clientId: 9990 });
  try {
    session.openDocx(bytes, false);
    session.loadState(edit.state);
    let offset = 0;
    for (const segment of session.storySegments("body")) {
      if (segment.kind === "embed" && segment.embedKind === "shape") break;
      offset += segment.kind === "text" ? segment.text.length : 1;
    }
    session.applyRawOps("body", [
      { op: "setEmbedAttr", index: offset, key: "width", value: 200 },
    ]);
    const resized = xmlOf(
      await exportOffice(
        bytes,
        { ...seeded, state: session.encodeState() },
        fixed
      )
    );
    expect(alternates(resized)).toEqual(alternates(source).slice(1));
    expect(resized.match(/<wps:txbx>/g)).toHaveLength(2);
    expect(resized.match(/<w:object /g)).toHaveLength(1);
  } finally {
    session.destroy();
  }
});

test("agent edits set XLSX cells by sheet name and address and clear them through the inverse", async () => {
  const bytes = await fixture("sample.xlsx");
  const before = await seedOffice("xlsx", bytes);
  const entries = await inspectOffice(bytes, before);
  const sheetName = entries[0].label.split("!")[0];
  const edit = await applyOfficeCommands(bytes, before, [
    {
      type: "set_cell",
      sheet: sheetName,
      cell: "ZZ100",
      expectedValue: "",
      value: "agent value",
    },
  ]);
  const after = { ...before, state: edit.state };
  const added = (await inspectOffice(bytes, after)).find(
    (entry) => entry.label === `${sheetName}!ZZ100`
  )!;
  expect(added.value).toBe("agent value");
  expect(edit.targets).toEqual([
    { id: added.id, path: ["xlsx:sheets", expect.any(String), "contents", expect.any(String)] },
  ]);
  expect(edit.targets[0].path[1] + ":" + edit.targets[0].path[3]).toBe(added.id);
  expect(edit.inverse).toEqual([
    {
      type: "set_cell",
      sheet: edit.targets[0].path[1],
      cell: "ZZ100",
      expectedValue: "agent value",
      value: "",
    },
  ]);
  const undone = await applyOfficeCommands(bytes, after, edit.inverse);
  expect(
    await compare(bytes, before, { ...before, state: undone.state })
  ).toEqual([]);
  await expect(
    applyOfficeCommands(bytes, before, [
      { type: "set_cell", sheet: "Nope", cell: "A1", expectedValue: "", value: "x" },
    ])
  ).rejects.toThrow("unavailable_target");
});

test("agent edits replace one PPTX paragraph and keep the paragraph identity for guards", async () => {
  const bytes = await fixture("betteroffice-demo.pptx");
  const before = await seedOffice("pptx", bytes);
  const target = (await inspectOffice(bytes, before)).find(
    (entry) => entry.value.length > 0
  )!;
  const edit = await applyOfficeCommands(bytes, before, [
    {
      type: "replace_text",
      targetId: target.id,
      expectedText: target.value,
      text: "Agent headline",
    },
  ]);
  const after = { ...before, state: edit.state };
  expect(edit.targets[0]).toEqual({
    id: target.id,
    path: [
      "pptx:stories",
      target.position.slice(0, target.position.lastIndexOf(":")),
    ],
    range: [expect.any(Number), expect.any(Number)],
  });
  expect(
    (await compare(bytes, before, after)).find((effect) => effect.id === target.id)
  ).toMatchObject({ operation: "replace", after: "Agent headline" });
  const undone = await applyOfficeCommands(bytes, after, edit.inverse);
  expect(
    await compare(bytes, before, { ...before, state: undone.state })
  ).toEqual([]);
  await expect(
    applyOfficeCommands(bytes, before, [
      {
        type: "replace_text",
        targetId: target.id,
        expectedText: target.value,
        text: "two\nparagraphs",
      },
    ])
  ).rejects.toThrow("invalid_input");
});

test("agent replace_text keeps PPTX runs outside the changed span and styles it from the first replaced character", async () => {
  const bytes = await fixture("betteroffice-demo.pptx");
  const seeded = await seedOffice("pptx", bytes);
  const target = (await inspectOffice(bytes, seeded)).find(
    (entry) => entry.value.length > 0
  )!;
  const storyId = target.position.slice(0, target.position.lastIndexOf(":"));
  const runsOf = (state: Uint8Array) => {
    const doc = PptxDocument.openCollaborativeFromUpdate(state, 9997, bytes);
    try {
      const story = JSON.parse(doc.storyJson(JSON.stringify({ storyId }))) as {
        paragraphs: Array<{
          id: string;
          runs: Array<{ text: string; style: { bold?: boolean | null } }>;
        }>;
      };
      return story.paragraphs
        .find((paragraph) => paragraph.id === target.id)!
        .runs.map((run) => ({ text: run.text, bold: run.style.bold === true }));
    } finally {
      doc.free();
    }
  };
  const doc = PptxDocument.openCollaborativeFromUpdate(
    seeded.state,
    9998,
    bytes
  );
  let mixed: Uint8Array;
  try {
    const story = JSON.parse(doc.storyJson(JSON.stringify({ storyId }))) as {
      paragraphs: Array<{ id: string; runs: Array<{ text: string }> }>;
    };
    let start = 0;
    for (const paragraph of story.paragraphs) {
      if (paragraph.id === target.id) break;
      start += paragraph.runs.map((run) => run.text).join("").length + 1;
    }
    doc.deleteTextJson(
      JSON.stringify({ storyId, start, end: start + target.value.length })
    );
    doc.insertTextJson(
      JSON.stringify({ storyId, index: start, text: "Plain " })
    );
    doc.insertTextJson(
      JSON.stringify({
        storyId,
        index: start + 6,
        text: "Bold",
        style: { bold: true },
      })
    );
    mixed = doc.encodeStateAsUpdate();
  } finally {
    doc.free();
  }
  expect(runsOf(mixed)).toEqual([
    { text: "Plain ", bold: false },
    { text: "Bold", bold: true },
  ]);
  const edit = await applyOfficeCommands(bytes, { ...seeded, state: mixed }, [
    {
      type: "replace_text",
      targetId: target.id,
      expectedText: "Plain Bold",
      text: "Plain Cold",
    },
    {
      type: "replace_text",
      targetId: target.id,
      expectedText: "Plain Cold",
      text: "Plainly Cold",
    },
  ]);
  expect(runsOf(edit.state)).toEqual([
    { text: "Plainly ", bold: false },
    { text: "Cold", bold: true },
  ]);
});

test("agent edit targets re-locate by stable id after earlier paragraphs move and count astral characters in UTF-16", async () => {
  const bytes = await fixture("betteroffice-demo.docx");
  const before = await seedOffice("docx", bytes);
  const entries = (await inspectOffice(bytes, before)).filter(
    (entry) => entry.value.length > 0
  );
  const [first, second] = entries;
  const edited = await applyOfficeCommands(bytes, before, [
    {
      type: "replace_text",
      targetId: second.id,
      expectedText: second.value,
      text: "Emoji 😀 paragraph",
    },
  ]);
  const shifted = await applyOfficeCommands(
    bytes,
    { ...before, state: edited.state },
    [
      {
        type: "replace_text",
        targetId: first.id,
        expectedText: first.value,
        text: `${first.value} plus a longer first paragraph`,
      },
    ]
  );
  const checkpoint = { ...before, state: shifted.state };
  const [located] = await locateOfficeTargets(bytes, checkpoint, [second.id]);
  const offset = "plus a longer first paragraph".length + 1;
  expect(located.range).toEqual([
    edited.targets[0].range![0] + offset,
    edited.targets[0].range![1] + offset,
  ]);
  expect(located.range![1] - located.range![0]).toBe("Emoji 😀 paragraph".length);
  const undone = await applyOfficeCommands(bytes, checkpoint, edited.inverse);
  expect(
    (await inspectOffice(bytes, { ...before, state: undone.state })).find(
      (entry) => entry.id === second.id
    )?.value
  ).toBe(second.value);
  await expect(
    locateOfficeTargets(bytes, checkpoint, ["body:paragraph:missing"])
  ).rejects.toThrow("unavailable_target");
  const workbook = await fixture("sample.xlsx");
  const seeded = await seedOffice("xlsx", workbook);
  const [cell] = await inspectOffice(workbook, seeded);
  const [cellTarget] = await locateOfficeTargets(workbook, seeded, [cell.id]);
  expect(cellTarget.path[0]).toBe("xlsx:sheets");
  expect(`${cellTarget.path[1]}:${cellTarget.path[3]}`).toBe(cell.id);
});

test("PPTX and XLSX agent edits count astral characters in UTF-16, and a cleared cell still locates while a missing sheet does not", async () => {
  const deck = await fixture("betteroffice-demo.pptx");
  const seededDeck = await seedOffice("pptx", deck);
  const paragraph = (await inspectOffice(deck, seededDeck)).find(
    (entry) => entry.value.length > 0
  )!;
  const edited = await applyOfficeCommands(deck, seededDeck, [
    {
      type: "replace_text",
      targetId: paragraph.id,
      expectedText: paragraph.value,
      text: "Slide 😀 title",
    },
  ]);
  const [located] = await locateOfficeTargets(
    deck,
    { ...seededDeck, state: edited.state },
    [paragraph.id]
  );
  expect(located.range![1] - located.range![0]).toBe("Slide 😀 title".length);
  expect(
    (await inspectOffice(deck, { ...seededDeck, state: edited.state })).find(
      (entry) => entry.id === paragraph.id
    )?.value
  ).toBe("Slide 😀 title");

  const workbook = await fixture("sample.xlsx");
  const seededBook = await seedOffice("xlsx", workbook);
  const [cell] = await inspectOffice(workbook, seededBook);
  const sheetName = cell.label.split("!")[0];
  const written = await applyOfficeCommands(workbook, seededBook, [
    {
      type: "set_cell",
      sheet: sheetName,
      cell: "ZZ200",
      expectedValue: "",
      value: "cell 😀 value",
    },
  ]);
  const withCell = { ...seededBook, state: written.state };
  const added = (await inspectOffice(workbook, withCell)).find(
    (entry) => entry.label === `${sheetName}!ZZ200`
  )!;
  expect(added.value).toBe("cell 😀 value");
  await expect(
    locateOfficeTargets(workbook, withCell, [added.id])
  ).resolves.toHaveLength(1);
  const cleared = await applyOfficeCommands(workbook, withCell, written.inverse);
  await expect(
    locateOfficeTargets(workbook, { ...seededBook, state: cleared.state }, [added.id])
  ).resolves.toHaveLength(1);
  await expect(
    locateOfficeTargets(workbook, withCell, [`missing-sheet:${added.id.slice(added.id.indexOf(":[") + 1)}`])
  ).rejects.toThrow("unavailable_target");
});


test("semantic baselines roundtrip without embedded media and detect edits and reversals", async () => {
  for (const [format, name] of [
    ["docx", "betteroffice-demo.docx"],
    ["xlsx", "sample.xlsx"],
    ["pptx", "betteroffice-demo.pptx"],
  ] as const) {
    const bytes = await fixture(name);
    const checkpoint = await seedOffice(format, bytes);
    const baseline = await officeBaseline(bytes, checkpoint);
    const encoded = JSON.stringify(baseline);
    expect(encoded).not.toContain('"bytes":');
    expect(encoded).not.toContain("data:image/");
    const restored = JSON.parse(encoded) as typeof baseline;
    expect(compareBaselines(restored, baseline)).toEqual([]);
    const text = baseline.find((entry) => entry.kind === "text");
    expect(text).toBeDefined();
    const edited = baseline.map((entry) =>
      entry.id === text!.id ? { ...entry, value: "Changed lesson" } : entry,
    );
    expect(compareBaselines(restored, edited)).toEqual([
      {
        id: text!.id,
        kind: "text",
        label: text!.label,
        operation: "replace",
        before: text!.value,
        after: "Changed lesson",
      },
    ]);
    expect(compareBaselines(restored, baseline)).toEqual([]);
    for (const entry of baseline.filter((entry) => entry.kind === "visual"))
      expect(entry.value).toMatch(/^[a-f0-9]{64}$/);
  }
});
