import { test, expect } from "bun:test";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import {
  exportOffice,
  inspectOffice,
  officeBaseline,
  rebaseOffice,
  resolveAsset,
  seedOffice,
  type OfficeCheckpoint,
} from "./office-checkpoint";
import { createYrsSession, type YrsSession } from "../packages/docx/src/yrs";
import { unzipContainer } from "../packages/docx/src/wasm/opc";

const fixed = { seed: "0".repeat(64), now: "2026-09-06T00:00:00.000Z" };
const fixture = async (name: string) =>
  new Uint8Array(
    await readFile(new URL(`../poc/fixtures/${name}`, import.meta.url))
  );
const sha256 = (bytes: Uint8Array) =>
  createHash("sha256").update(bytes).digest("hex");
const holds = (bytes: Uint8Array, text: string) =>
  Buffer.from(bytes).includes(text);

async function openState(bytes: Uint8Array, state: Uint8Array) {
  const session = await createYrsSession({ clientId: 9101 });
  session.openDocx(bytes, false);
  session.loadState(state);
  return session;
}

function images(session: YrsSession, story: string) {
  const found: Array<{ index: number; src: string }> = [];
  let offset = 0;
  for (const segment of session.storySegments(story)) {
    if (segment.kind === "embed" && segment.embedKind === "image")
      found.push({ index: offset, src: String(segment.payload.src) });
    offset += segment.kind === "text" ? segment.text.length : 1;
  }
  return found;
}

test("DOCX source images reference their parts and export as they did with inline bytes", async () => {
  for (const name of ["exchange-plan.docx", "opaque-objects.docx"]) {
    const bytes = await fixture(name);
    const seed = await seedOffice("docx", bytes);
    expect(holds(seed.state, "data:image")).toBe(false);
    expect(holds(seed.state, "media:word/media/")).toBe(true);

    const media = Object.entries(unzipContainer(bytes)).filter(([path]) =>
      path.startsWith("word/media/")
    );
    const partHashes = new Set(media.map(([, data]) => sha256(data)));
    const entries = (await officeBaseline(bytes, seed)).filter(
      (entry) => entry.kind === "image"
    );
    expect(entries.length).toBeGreaterThan(0);
    for (const entry of entries)
      expect(partHashes.has(entry.imageSHA256!)).toBe(true);
    const asset = await resolveAsset(bytes, seed, entries[0].assetRef!);
    expect(asset.sha256).toBe(entries[0].imageSHA256!);

    const exported = unzipContainer(await exportOffice(bytes, seed, fixed));
    expect(
      Object.keys(exported).filter((path) => path.startsWith("word/media/"))
    ).toEqual(media.map(([path]) => path));
    for (const [path, data] of media) expect(exported[path]).toEqual(data);
  }

  // Inline data URLs and references export the same package.
  const bytes = await fixture("exchange-plan.docx");
  const seed = await seedOffice("docx", bytes);
  const inline = await openState(bytes, seed.state);
  try {
    const media = inline.materializeDocx()!.package.media!;
    for (const story of inline.storyIds()) {
      const ops = images(inline, story).map(({ index, src }) => ({
        op: "setEmbedAttr" as const,
        index,
        key: "src",
        value: media.get(src.slice("media:".length))!.dataUrl!,
      }));
      if (ops.length) inline.applyRawOps(story, ops);
    }
    expect(holds(inline.encodeState(), "data:image")).toBe(true);
    expect(await exportOffice(bytes, seed, fixed)).toEqual(
      await exportOffice(bytes, { ...seed, state: inline.encodeState() }, fixed)
    );
  } finally {
    inline.destroy();
  }
});

test("a DOCX publication rebases a later edit and turns an inserted image into a reference", async () => {
  const bytes = await fixture("exchange-plan.docx");
  const seed = await seedOffice("docx", bytes);
  const doc = await openState(bytes, seed.state);
  try {
    const [first, second] = doc.paragraphs("body");
    doc.insertImage(
      { story: "body", paraId: first.paraId, offset: 0 },
      {
        src: "data:image/png;base64,AQID",
        rId: "rId_img_temporary",
        width: 20,
        height: 20,
      }
    );
    const captured = { ...seed, state: doc.encodeState() };
    doc.insertText({ story: "body", paraId: second.paraId, offset: 0 }, "Later ");
    const latest = { ...seed, state: doc.encodeState() };
    const published = await exportOffice(bytes, captured, fixed);
    const rebased = await rebaseOffice(bytes, captured, latest, published);
    expect(holds(rebased.state, "data:image")).toBe(false);

    const next: OfficeCheckpoint = {
      format: "docx",
      schemaVersion: 1,
      baseSha256: sha256(published),
      state: rebased.state,
    };
    expect(
      (await inspectOffice(published, next)).some((entry) =>
        entry.value.startsWith("Later ")
      )
    ).toBe(true);
    expect(
      rebased.effects.filter(
        (effect) => effect.kind !== "visual" && effect.operation !== "move"
      )
    ).toEqual([
      expect.objectContaining({
        kind: "text",
        operation: "replace",
        after: expect.stringMatching(/^Later /),
      }),
    ]);
    const inserted = sha256(Uint8Array.from([1, 2, 3]));
    const part = Object.entries(unzipContainer(published)).find(
      ([path, data]) => path.startsWith("word/media/") && sha256(data) === inserted
    )![0];
    const rebasedDoc = await openState(published, rebased.state);
    try {
      expect(images(rebasedDoc, "body").map(({ src }) => src)).toContain(
        `media:${part}`
      );
    } finally {
      rebasedDoc.destroy();
    }
    expect(
      (await officeBaseline(published, next)).some(
        (entry) => entry.imageSHA256 === inserted
      )
    ).toBe(true);
  } finally {
    doc.destroy();
  }
});

test("a DOCX image reference to a missing part fails the baseline and the export", async () => {
  const bytes = await fixture("exchange-plan.docx");
  const seed = await seedOffice("docx", bytes);
  const doc = await openState(bytes, seed.state);
  try {
    const [image] = images(doc, "body");
    doc.applyRawOps("body", [
      {
        op: "setEmbedAttr",
        index: image.index,
        key: "src",
        value: "media:word/media/missing.png",
      },
    ]);
    const broken = { ...seed, state: doc.encodeState() };
    await expect(officeBaseline(bytes, broken)).rejects.toThrow(
      "word/media/missing.png is absent from the source package"
    );
    await expect(exportOffice(bytes, broken, fixed)).rejects.toThrow(
      "missing package part word/media/missing.png"
    );
  } finally {
    doc.destroy();
  }
});
