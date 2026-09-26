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
import { rezipContainer, unzipContainer } from "../packages/docx/src/wasm/opc";

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

test("a DOCX state above the 64 MiB update cap loads and exports but is refused as one update", async () => {
  const bytes = await fixture("feature-rich.docx");
  const seed = await seedOffice("docx", bytes);
  const doc = await openState(bytes, seed.state);
  let state: Uint8Array;
  try {
    const [first] = doc.paragraphs("body");
    const photo = Buffer.alloc(9 * 1024 * 1024, 7);
    for (let index = 0; index < 6; index += 1)
      doc.insertImage(
        { story: "body", paraId: first.paraId, offset: 0 },
        {
          src: `data:image/png;base64,${photo.toString("base64")}`,
          rId: `rId_img_photo${index}`,
          width: 20,
          height: 20,
        }
      );
    state = doc.encodeState();
  } finally {
    doc.destroy();
  }
  expect(state.length).toBeGreaterThan(64 * 1024 * 1024);
  const exported = unzipContainer(await exportOffice(bytes, { ...seed, state }, fixed));
  expect(
    Object.entries(exported).some(
      ([path, data]) => path.startsWith("word/media/") && data.length === 9 * 1024 * 1024
    )
  ).toBe(true);

  const peer = await createYrsSession({ clientId: 9102 });
  try {
    peer.openDocx(bytes, false);
    let refusal: unknown;
    try {
      peer.applyUpdate(state);
    } catch (error) {
      refusal = error;
    }
    expect(refusal).toBeInstanceOf(Error);
    expect((refusal as Error).message).toContain("update exceeds 67108864 bytes");
  } finally {
    peer.destroy();
  }
}, 120_000);

test("Word-like DOCX runs merge on export and keep per-character formatting and tracked changes", async () => {
  const run = (text: string, rsid: string, properties = "") =>
    `<w:r w:rsidR="${rsid}">${properties && `<w:rPr>${properties}</w:rPr>`}` +
    `<w:t xml:space="preserve">${text}</w:t></w:r>`;
  const tracked =
    '<w:i/><w:rPrChange w:id="7" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"><w:rPr/></w:rPrChange>';
  const parts = unzipContainer(await fixture("feature-rich.docx"));
  parts["word/document.xml"] = new TextEncoder().encode(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
      `<w:p>${run("Split ", "00A1")}${run("by ", "00B2")}${run("bold", "00C3", "<w:b/>")}${run(" rsids", "00D4")}</w:p>` +
      `<w:p>${run("Tracked ", "00A1")}${run("change", "00B2", tracked)}${run(" kept", "00C3")}</w:p>` +
      "</w:body></w:document>"
  );
  const bytes = rezipContainer(parts);
  const seed = await seedOffice("docx", bytes);
  const doc = await openState(bytes, seed.state);
  try {
    const [merged, kept] = doc.paragraphs("body");
    expect(merged.properties._originalRunBoundaries).toBeUndefined();
    expect(kept.properties._originalRunBoundaries).toHaveLength(3);
  } finally {
    doc.destroy();
  }
  const xml = new TextDecoder().decode(
    unzipContainer(await exportOffice(bytes, seed, fixed))["word/document.xml"]
  );
  const runs = [...xml.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].map(([paragraph]) =>
    [...paragraph.matchAll(/<w:r>([\s\S]*?)<\/w:r>/g)].map(([, body]) => ({
      text: [...body.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(([, text]) => text).join(""),
      properties: /<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(body)?.[1] ?? "",
    }))
  );
  expect(runs[0].map((item) => item.text)).toEqual(["Split by ", "bold", " rsids"]);
  expect(runs[0].map((item) => item.properties.includes("<w:b/>"))).toEqual([false, true, false]);
  expect(runs[1].map((item) => item.text)).toEqual(["Tracked ", "change", " kept"]);
  expect(runs[1].map((item) => item.properties.includes("<w:rPrChange"))).toEqual([false, true, false]);
  expect(runs[1][1].properties).toContain("<w:i/>");
});

// A seed change fails here and needs a maintenance window (record 23); the
// budget is the measured size plus 2%, so an updated hash cannot hide growth.
const goldenSeeds: Array<[string, string, number]> = [
  ["exchange-plan.docx", "595d4a2b6469f7e6d27ef1947b5a16f822209539e1a23e7658ddf08ea099731e", 301_100],
  ["opaque-objects.docx", "675194d3d0f387d294337fa4c15353953452afb29e9d99cee24697de80b4cd18", 47_300],
  ["book-30p.docx", "a31ad9336c4356c003c7851c6f91211057376e4863f8a1a5ac540e485f442546", 248_600],
  ["images-10.docx", "e73f3f87accec1052ce911256c09ca15f3e4662f9086e379f802221d1ac9cc00", 64_900],
  ["lecture.pptx", "62cdeb13ca369f0bd510a8fd32a10e5c3067922459551d7720dc25519d2cf2e8", 111_300],
  ["deck-50.pptx", "02d15d15807bf30eb01e862fc707951d8ed5e94dcd6855da69d457d558951928", 164_500],
];

test.each(goldenSeeds)(
  "%s seeds to its golden bytes within its size budget",
  async (name, golden, budget) => {
    const bytes = await fixture(name);
    const format = name.endsWith(".pptx") ? "pptx" : "docx";
    const seed = await seedOffice(format, bytes);
    expect((await seedOffice(format, bytes)).state).toEqual(seed.state);
    expect(seed.state.length).toBeLessThanOrEqual(budget);
    expect(sha256(seed.state)).toBe(golden);
    if (format === "docx")
      for (const entry of await officeBaseline(bytes, seed))
        expect(entry.id).not.toMatch(/<[1-9]\d*#/);
  },
  30_000
);
