import { beforeAll, expect, test } from "bun:test";
import { readFile } from "node:fs/promises";
import * as Y from "yjs";
import {
  createYrsSession,
  type YrsLoc,
  type YrsSession,
} from "../packages/docx/src/yrs";
import { rezipContainer } from "../packages/docx/src/wasm/opc";
import { PptxDocument } from "../packages/pptx/src/wasm/generated/pptx_wasm.js";
import { RebaseError, transplant } from "./office-rebase";
import {
  exportOffice,
  officeBaseline,
  rebaseOffice,
  seedOffice,
  type OfficeCheckpoint,
} from "./office-checkpoint";

const fixed = { seed: "0".repeat(64), now: "2026-09-28T00:00:00.000Z" };
// Seeding loads the engines the sessions below use directly.
beforeAll(async () => {
  await seedOffice(
    "docx",
    await fixture("apps/demo/public/betteroffice-demo.docx")
  );
  await seedOffice(
    "pptx",
    await fixture("apps/demo/public/betteroffice-demo.pptx")
  );
});
const fixture = async (path: string) =>
  new Uint8Array(await readFile(new URL(`../${path}`, import.meta.url)));
const checkpoint = (
  format: "docx" | "pptx",
  base: Uint8Array,
  state: Uint8Array
): OfficeCheckpoint => ({
  format,
  schemaVersion: 1,
  baseSha256: new Bun.CryptoHasher("sha256").update(base).digest("hex"),
  state,
});

function random(seed: number) {
  let state = seed;
  const next = () => {
    state ^= state << 13;
    state ^= state >>> 17;
    state ^= state << 5;
    return (state >>> 0) / 0x100000000;
  };
  return {
    int: (lo: number, hi: number) => lo + Math.floor(next() * (hi - lo + 1)),
    pick: <T>(items: T[]) => items[Math.floor(next() * items.length)],
    choose: (weights: Record<string, number>) => {
      let r = next() * Object.values(weights).reduce((a, b) => a + b, 0);
      for (const [key, weight] of Object.entries(weights))
        if ((r -= weight) < 0) return key;
      return Object.keys(weights)[0];
    },
  };
}
type Random = ReturnType<typeof random>;
const WORDS = "the notes study quiz recall chapter answer review".split(" ");

/** One DOCX edit of the kinds a writer makes; false when the pick does not apply. */
function editDocx(session: YrsSession, kind: string, rng: Random): boolean {
  const story = "body";
  const paragraphs = session
    .paragraphs(story)
    .filter((p) => p.text.length > 0 && !p.text.includes("￼"));
  if (!paragraphs.length) return false;
  const p = rng.pick(paragraphs);
  const at = (offset: number) => ({ story, paraId: p.paraId, offset });
  try {
    if (kind === "type")
      session.insertText(at(rng.int(0, p.text.length)), ` ${rng.pick(WORDS)}`);
    else if (kind === "delete") {
      const start = rng.int(0, Math.max(0, p.text.length - 2));
      session.deleteRange({
        story,
        start: at(start),
        end: at(Math.min(p.text.length, start + rng.int(1, 12))),
      });
    } else if (kind === "split") {
      const receipt = session.splitParagraph(at(p.text.length));
      session.insertText(
        { story, paraId: receipt.secondParaId, offset: 0 },
        `${rng.pick(WORDS)} ${rng.pick(WORDS)}.`
      );
    } else if (kind === "merge")
      session.mergeParagraphs(story, p.paraId, "forward");
    else if (kind === "bold") {
      const start = rng.int(0, p.text.length - 1);
      session.toggleMark(
        {
          story,
          start: at(start),
          end: at(Math.min(p.text.length, start + 8)),
        },
        { type: "bold" }
      );
    } else if (kind === "color") {
      const start = rng.int(0, p.text.length - 1);
      session.formatRange(
        {
          story,
          start: at(start),
          end: at(Math.min(p.text.length, start + 8)),
        },
        { color: { rgb: "C00000" } }
      );
    } else if (kind === "align")
      session.setParagraphAttrs(
        { story, start: at(0), end: at(p.text.length) },
        { alignment: rng.pick(["center", "right"]) }
      );
    else return false;
    return true;
  } catch {
    return false;
  }
}

/** One PPTX edit of the kinds a presenter makes; false when the pick does not apply. */
function editPptx(doc: PptxDocument, kind: string, rng: Random): boolean {
  const deck = JSON.parse(doc.snapshotJson()) as {
    slides: Array<{
      id: string;
      shapes: Array<{
        id: string;
        textStories: Array<{
          id: string;
          paragraphs: Array<{ runs: Array<{ text: string }> }>;
        }>;
      }>;
    }>;
  };
  const stories = deck.slides.flatMap((slide) =>
    slide.shapes.flatMap((shape) =>
      shape.textStories.map((story) => ({
        slide: slide.id,
        shape: shape.id,
        story,
      }))
    )
  );
  const text = (story: (typeof stories)[number]["story"]) =>
    story.paragraphs.map((p) => p.runs.map((r) => r.text).join("")).join("\n");
  const withText = stories.filter((entry) => text(entry.story).length > 0);
  if (!withText.length) return false;
  const { slide, shape, story } = rng.pick(withText);
  const length = text(story).length;
  const storyId = story.id;
  try {
    if (kind === "type")
      doc.insertTextJson(
        JSON.stringify({
          storyId,
          index: rng.int(0, length),
          text: ` ${rng.pick(WORDS)}`,
        })
      );
    else if (kind === "delete") {
      const start = rng.int(0, Math.max(0, length - 2));
      doc.deleteTextJson(
        JSON.stringify({
          storyId,
          start,
          end: Math.min(length, start + rng.int(1, 12)),
        })
      );
    } else if (kind === "split") {
      doc.insertParagraphBreakJson(JSON.stringify({ storyId, index: length }));
      doc.insertTextJson(
        JSON.stringify({
          storyId,
          index: length + 1,
          text: `${rng.pick(WORDS)}.`,
        })
      );
    } else if (kind === "bold") {
      const start = rng.int(0, length - 1);
      doc.formatTextJson(
        JSON.stringify({
          storyId,
          start,
          end: Math.min(length, start + 8),
          patch: { bold: true },
        })
      );
    } else if (kind === "slide") {
      const receipt = JSON.parse(
        doc.insertSlideJson(
          JSON.stringify({
            index: rng.int(0, deck.slides.length),
            layoutPartPath: null,
          })
        )
      );
      doc.addTextBoxJson(
        JSON.stringify({
          slideId: receipt.slideId,
          draft: {
            name: "Note",
            rect: { x: 914400, y: 914400, width: 6096000, height: 1371600 },
            text: rng.pick(WORDS),
            style: {},
          },
        })
      );
    } else if (kind === "move")
      doc.moveShapeJson(
        JSON.stringify({
          slideId: slide,
          shapeId: shape,
          x: rng.int(0, 5_000_000),
          y: rng.int(0, 3_000_000),
        })
      );
    else return false;
    return true;
  } catch {
    return false;
  }
}

const MIXED = {
  type: 40,
  delete: 15,
  split: 10,
  merge: 5,
  bold: 10,
  color: 5,
  align: 5,
  slide: 5,
  move: 5,
};
const STRUCTURAL = {
  split: 30,
  merge: 20,
  delete: 20,
  type: 15,
  bold: 5,
  slide: 20,
  move: 20,
};

/** An editing session over a state: `edit` applies n edits of the weighted kinds. */
async function open(
  format: "docx" | "pptx",
  base: Uint8Array,
  state: Uint8Array,
  rng: Random
) {
  if (format === "docx") {
    const session = await createYrsSession({ clientId: rng.int(1, 2 ** 31) });
    session.openDocx(base, false);
    session.loadState(state);
    return {
      edit(n: number, weights: Record<string, number>) {
        for (let i = 0; i < n; i++)
          for (
            let attempt = 0;
            attempt < 8 && !editDocx(session, rng.choose(weights), rng);
            attempt++
          );
      },
      state: () => session.encodeState(),
      close: () => session.destroy(),
    };
  }
  const doc = PptxDocument.openCollaborativeFromUpdate(
    state,
    rng.int(1, 2 ** 31),
    base
  );
  return {
    edit(n: number, weights: Record<string, number>) {
      for (let i = 0; i < n; i++)
        for (
          let attempt = 0;
          attempt < 8 && !editPptx(doc, rng.choose(weights), rng);
          attempt++
        );
    },
    state: () => doc.encodeStateAsUpdate(),
    close: () => doc.free(),
  };
}

const values = async (base: Uint8Array, state: OfficeCheckpoint) =>
  (await officeBaseline(base, state))
    .filter((entry) => entry.kind === "text")
    .map((entry) => entry.value)
    .sort();

for (const [format, path, seed] of [
  ["docx", "poc/fixtures/feature-rich.docx", 7],
  ["docx", "apps/demo/public/betteroffice-demo.docx", 11],
  ["pptx", "poc/fixtures/deck-50.pptx", 7],
  ["pptx", "apps/demo/public/betteroffice-demo.pptx", 11],
] as const)
  test(`${format} publications land later edits on seed(export) over three cycles (${path})`, async () => {
    const rng = random(seed);
    let base: Uint8Array = await fixture(path);
    let state: Uint8Array = (await seedOffice(format, base)).state;
    for (let cycle = 0; cycle < 3; cycle++) {
      const session = await open(format, base, state, rng);
      try {
        session.edit(20, MIXED);
        const captured = session.state();
        const exported = await exportOffice(
          base,
          checkpoint(format, base, captured),
          fixed
        );
        session.edit(20, STRUCTURAL);
        const latest = session.state();
        const rebased = await rebaseOffice(
          base,
          checkpoint(format, base, captured),
          checkpoint(format, base, latest),
          exported
        );
        const exportSeed = (await seedOffice(format, exported)).state;
        const doc = new Y.Doc();
        Y.applyUpdate(doc, rebased.state);
        const change = Y.encodeStateAsUpdate(
          doc,
          Y.encodeStateVectorFromUpdate(exportSeed)
        );
        const rebuilt = new Y.Doc();
        rebuilt.transact(() => {
          Y.applyUpdate(rebuilt, exportSeed);
          Y.applyUpdate(rebuilt, change);
        });
        expect(Y.encodeStateAsUpdate(rebuilt)).toEqual(
          Y.encodeStateAsUpdate(doc)
        );
        // The change is about the size of the later edits, not the document.
        const later = Y.diffUpdate(
          latest,
          Y.encodeStateVectorFromUpdate(captured)
        );
        expect(change.length).toBeLessThan(3 * later.length + 512);
        expect(
          await values(exported, checkpoint(format, exported, rebased.state))
        ).toEqual(await values(base, checkpoint(format, base, latest)));
        expect(rebased.effects.length).toBeGreaterThan(0);
        base = exported;
        state = rebased.state;
      } finally {
        session.close();
      }
    }
  }, 120_000);

const encoder = new TextEncoder();
function docx(body: string): Uint8Array {
  return rezipContainer(
    Object.fromEntries(
      Object.entries({
        "[Content_Types].xml":
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
        "_rels/.rels":
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
        "word/document.xml": `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
      }).map(([path, xml]) => [path, encoder.encode(xml)])
    )
  );
}
const paragraph = (id: string, text: string, extra = "") =>
  `<w:p w14:paraId="${id}"><w:r><w:t>${text}</w:t></w:r>${extra}</w:p>`;
const cell = (text: string) =>
  `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const table = (rows: string[][]) =>
  `<w:tbl>${rows
    .map((row) => `<w:tr>${row.map(cell).join("")}</w:tr>`)
    .join("")}</w:tbl>`;

async function docxSession(base: Uint8Array, state?: Uint8Array) {
  const session = await createYrsSession({ clientId: 4101 });
  session.openDocx(base, !state);
  if (state) session.loadState(state);
  return session;
}
const text = (session: YrsSession, story = "body") =>
  session.paragraphs(story).map((p) => p.text);

test("a DOCX table inserted after the capture keeps its cells when the export renumbered the tables", async () => {
  const base = docx(
    `${table([
      ["a", "b"],
      ["c", "d"],
    ])}${paragraph("11111111", "Middle")}${table([
      ["e", "f"],
      ["g", "h"],
    ])}${paragraph("22222222", "End")}`
  );
  const session = await docxSession(base);
  try {
    // With the first table deleted, the export numbers the second one t0,
    // the id a table inserted later takes as well.
    session.deleteTable({ story: "body", tableIndex: 0 });
    const captured = session.encodeState();
    const exported = await exportOffice(
      base,
      checkpoint("docx", base, captured),
      fixed
    );
    const [middle] = session.paragraphs("body");
    const inserted = session.insertTable(
      { story: "body", paraId: middle.paraId, offset: 0 },
      2,
      2
    );
    session.applyRawOps(inserted.createdStoryIds[0], [
      { op: "insert", index: 0, text: "new cell" },
    ]);
    const latest = session.encodeState();
    const rebased = await rebaseOffice(
      base,
      checkpoint("docx", base, captured),
      checkpoint("docx", base, latest),
      exported
    );
    const current = await docxSession(exported, rebased.state);
    try {
      const renamed = current.storyIds().filter((id) => id.includes("~"));
      expect(renamed).toEqual(
        inserted.createdStoryIds.map((id) => `${id}~1`).sort()
      );
      expect(text(current, `${inserted.createdStoryIds[0]}~1`)).toEqual([
        "new cell",
      ]);
      expect(text(current, "body:t0:r0c0")).toEqual(["e"]);
    } finally {
      current.destroy();
    }
  } finally {
    session.destroy();
  }
});

test("a DOCX comment added after the capture keeps its range", async () => {
  const base = docx(
    `${paragraph("11111111", "First paragraph")}${paragraph(
      "22222222",
      "Second paragraph"
    )}`
  );
  const session = await docxSession(base);
  try {
    const [first] = session.paragraphs("body");
    session.insertText(
      { story: "body", paraId: first.paraId, offset: 0 },
      "Captured "
    );
    const captured = session.encodeState();
    const exported = await exportOffice(
      base,
      checkpoint("docx", base, captured),
      fixed
    );
    const [, second] = session.paragraphs("body");
    session.insertText(
      { story: "body", paraId: first.paraId, offset: 0 },
      "Later "
    );
    const { commentId } = session.addComment(
      [
        {
          story: "body",
          start: { paraId: second.paraId, offset: 0 },
          end: { paraId: second.paraId, offset: 6 },
        },
      ],
      "Reviewer",
      "2026-09-28T00:00:00Z",
      [{ type: "paragraph", content: [] }]
    );
    const latest = session.encodeState();
    const rebased = await rebaseOffice(
      base,
      checkpoint("docx", base, captured),
      checkpoint("docx", base, latest),
      exported
    );
    const current = await docxSession(exported, rebased.state);
    try {
      expect(current.resolveComment(commentId)).toEqual(
        session.resolveComment(commentId)
      );
      expect(text(current)).toEqual(text(session));
    } finally {
      current.destroy();
    }
  } finally {
    session.destroy();
  }
});

/** A capture, its export, the later edits and their rebase onto the export. */
async function publishDocx(
  base: Uint8Array,
  before: (session: YrsSession) => void,
  after: (session: YrsSession) => void
) {
  const session = await docxSession(base);
  try {
    before(session);
    const captured = session.encodeState();
    const exported = await exportOffice(
      base,
      checkpoint("docx", base, captured),
      fixed
    );
    after(session);
    const latest = session.encodeState();
    const rebase = () =>
      rebaseOffice(
        base,
        checkpoint("docx", base, captured),
        checkpoint("docx", base, latest),
        exported
      );
    return { exported, latest, rebase };
  } finally {
    session.destroy();
  }
}
/** Story units as text: `¶` a pilcrow, `[kind]` any other embed. */
const units = (session: YrsSession, story = "body") =>
  session
    .storySegments(story)
    .map((segment) =>
      segment.kind === "text"
        ? segment.text
        : segment.kind === "pilcrow"
        ? "¶"
        : `[${segment.embedKind}]`
    )
    .join("");
const pageBreakParagraphs = docx(
  `${paragraph(
    "11111111",
    "Before"
  )}<w:p w14:paraId="22222222"><w:r><w:br w:type="page"/></w:r></w:p>${paragraph(
    "44444444",
    "Chapter"
  )}`
);

test("a DOCX page break whose paragraph was edited before the capture rebases exactly", async () => {
  const { exported, latest, rebase } = await publishDocx(
    pageBreakParagraphs,
    // The export used to drop the break of an edited paragraph.
    (session) =>
      session.insertText(
        { story: "body", paraId: "22222222", offset: 0 },
        "Typed"
      ),
    // Backspace at the start of Chapter deletes the break.
    (session) =>
      session.deleteRange({
        story: "body",
        start: { paraId: "44444444", offset: 0 },
        end: { paraId: "44444444", offset: 1 },
      })
  );
  const current = await docxSession(exported, (await rebase()).state);
  const later = await docxSession(pageBreakParagraphs, latest);
  try {
    expect(units(current)).toBe(units(later));
    expect(units(current)).toBe("Before¶Typed¶Chapter¶");
  } finally {
    current.destroy();
    later.destroy();
  }
});

/**
 * `Before` merged into `After` on one replica while another opens After's
 * slot with `block`: together they put text ahead of that block in one slot,
 * which neither editor allows on its own.
 */
async function blockAfterText(
  base: Uint8Array,
  block: (session: YrsSession, at: YrsLoc) => void
): Promise<Uint8Array> {
  const left = await docxSession(base);
  const right = await createYrsSession({ clientId: 4102 });
  try {
    right.openDocx(base, false);
    right.loadState(left.encodeState());
    const [before, after] = left.paragraphs("body");
    block(left, { story: "body", paraId: after.paraId, offset: 0 });
    right.mergeParagraphs("body", before.paraId, "forward");
    const merged = new Y.Doc();
    Y.applyUpdate(merged, left.encodeState());
    Y.applyUpdate(merged, right.encodeState());
    return Y.encodeStateAsUpdate(merged);
  } finally {
    left.destroy();
    right.destroy();
  }
}
const beforeAfter = docx(
  `${paragraph("11111111", "Before")}${paragraph("22222222", "After")}`
);

test("DOCX text that concurrent edits put ahead of a page break in its slot stays ahead of it, and edits to it rebase", async () => {
  const concurrent = await blockAfterText(beforeAfter, (session, at) =>
    session.insertPageBreak(at)
  );
  const { exported, latest, rebase } = await publishDocx(
    beforeAfter,
    (session) => session.loadState(concurrent),
    (session) =>
      session.deleteRange({
        story: "body",
        start: { paraId: session.paragraphs("body")[0].paraId, offset: 1 },
        end: { paraId: session.paragraphs("body")[0].paraId, offset: 3 },
      })
  );
  const current = await docxSession(exported, (await rebase()).state);
  const later = await docxSession(beforeAfter, latest);
  try {
    // The export keeps the break between the texts; seeded again it closes
    // the paragraph, where the seed places every break after visible text.
    expect(text(current)).toEqual(text(later));
    expect(units(later)).toBe("Bore[pageBreak]After¶");
    expect(units(current)).toBe("BoreAfter¶[pageBreak]");
  } finally {
    current.destroy();
    later.destroy();
  }
});

test("a DOCX table that concurrent edits put after text in its slot follows that text in the export and refuses later edits", async () => {
  const concurrent = await blockAfterText(beforeAfter, (session, at) =>
    session.insertTable(at, 1, 1)
  );
  const { rebase } = await publishDocx(
    beforeAfter,
    (session) => {
      session.loadState(concurrent);
      expect(units(session)).toBe("Before[table]After¶");
    },
    // The export writes the text after the table, so the rebase lands edits
    // to the text but not to the table it no longer matches.
    (session) => session.deleteTable({ story: "body", tableIndex: 0 })
  );
  await expect(rebase()).rejects.toThrow("has no counterpart in the export");
});

/** The units a comment covers, `[kind]` for an embed. */
function covered(session: YrsSession, commentId: string): string {
  const [anchor] = session.resolveComment(commentId);
  const all = session
    .storySegments("body")
    .flatMap((segment) =>
      segment.kind === "text"
        ? [...segment.text]
        : [segment.kind === "pilcrow" ? "¶" : `[${segment.embedKind}]`]
    );
  return all.slice(anchor.start, anchor.end).join("");
}
const comment = (
  session: YrsSession,
  paraId: string,
  start: number,
  end: number
) =>
  session.addComment(
    [
      {
        story: "body",
        start: { paraId, offset: start },
        end: { paraId, offset: end },
      },
    ],
    "Reviewer",
    "2026-09-28T00:00:00Z",
    [{ type: "paragraph", content: [] }]
  ).commentId;

test("a DOCX comment added after the capture covers the same text where the export moved a table", async () => {
  const concurrent = await blockAfterText(beforeAfter, (session, at) =>
    session.insertTable(at, 1, 1)
  );
  let commentId = "";
  let refused = "";
  const { exported, latest, rebase } = await publishDocx(
    beforeAfter,
    (session) => session.loadState(concurrent),
    (session) => {
      const [slot] = session.paragraphs("body");
      commentId = comment(session, slot.paraId, 0, 4);
      // A comment over the table the export moved has nowhere to land.
      refused = comment(session, slot.paraId, 5, 8);
      expect(covered(session, commentId)).toBe("Befo");
      expect(covered(session, refused)).toBe("e[table]A");
    }
  );
  await expect(rebase()).rejects.toThrow(
    "a comment anchor covers content the export wrote differently"
  );
  // Without it, the other lands on the text it covered.
  const later = await docxSession(beforeAfter, latest);
  try {
    later.applyRawOps("body", [{ op: "removeComment", id: refused }]);
    const rebased = await rebaseOffice(
      beforeAfter,
      checkpoint("docx", beforeAfter, concurrent),
      checkpoint("docx", beforeAfter, later.encodeState()),
      exported
    );
    const current = await docxSession(exported, rebased.state);
    try {
      expect(units(current)).toBe("[table]BeforeAfter¶");
      expect(covered(current, commentId)).toBe("Befo");
    } finally {
      current.destroy();
    }
  } finally {
    later.destroy();
  }
});

test("a DOCX comment made before the capture can be changed and removed after it", async () => {
  const base = docx(
    `${paragraph("11111111", "First paragraph")}${paragraph(
      "22222222",
      "Second paragraph"
    )}`
  );
  for (const change of ["patch", "remove"] as const) {
    let commentId = "";
    const { exported, rebase } = await publishDocx(
      base,
      (session) => {
        commentId = comment(session, "11111111", 0, 5);
      },
      (session) => {
        // The export saves the comment under a numeric id of its own.
        session.applyRawOps("body", [
          change === "patch"
            ? {
                op: "patchComment",
                id: commentId,
                fields: {
                  done: true,
                  body: [
                    {
                      type: "paragraph",
                      content: [
                        {
                          type: "run",
                          content: [{ type: "text", text: "Edited" }],
                        },
                      ],
                    },
                  ],
                },
              }
            : { op: "removeComment", id: commentId },
        ]);
      }
    );
    const current = await docxSession(exported, (await rebase()).state);
    try {
      const comments = current.listComments();
      if (change === "remove") expect(comments).toEqual([]);
      else {
        expect(comments).toHaveLength(1);
        expect(comments[0].done).toBe(true);
        expect(JSON.stringify(comments[0].body)).toContain("Edited");
        expect(covered(current, comments[0].id)).toBe("First");
      }
    } finally {
      current.destroy();
    }
  }
});

test("edits around a DOCX comment made before the capture land beside the reference mark the export adds", async () => {
  const base = docx(
    `${paragraph("11111111", "Intro paragraph")}${paragraph(
      "22222222",
      "alpha beta gamma"
    )}`
  );
  const edits: Array<[string, (session: YrsSession) => void]> = [
    [
      "typing in an earlier paragraph",
      (session) =>
        session.insertText(
          { story: "body", paraId: "11111111", offset: 0 },
          "New "
        ),
    ],
    [
      "typing at the start of the commented text",
      (session) =>
        session.insertText(
          { story: "body", paraId: "22222222", offset: 0 },
          "X"
        ),
    ],
    [
      "typing right after the commented text",
      (session) =>
        session.insertText(
          { story: "body", paraId: "22222222", offset: 5 },
          "X"
        ),
    ],
    [
      "a delete across the comment's end",
      (session) =>
        session.deleteRange({
          story: "body",
          start: { paraId: "22222222", offset: 3 },
          end: { paraId: "22222222", offset: 8 },
        }),
    ],
    [
      "a comment over the old one's end",
      (session) => comment(session, "22222222", 2, 10),
    ],
  ];
  for (const [name, edit] of edits) {
    const { exported, latest, rebase } = await publishDocx(
      base,
      (session) => comment(session, "22222222", 0, 5),
      edit
    );
    const seeded = await docxSession(exported);
    // The export writes the comment's reference mark after "alpha".
    expect(units(seeded)).toContain("alpha[field] beta");
    seeded.destroy();
    const current = await docxSession(exported, (await rebase()).state);
    const later = await docxSession(base, latest);
    try {
      expect(text(current), name).toEqual(text(later));
    } finally {
      current.destroy();
      later.destroy();
    }
  }
});

test("a DOCX edit inside a run with a complex-script theme font rebases with that font", async () => {
  const base = docx(
    `${paragraph(
      "11111111",
      "Before"
    )}<w:p w14:paraId="22222222"><w:r><w:rPr><w:rFonts w:cstheme="majorBidi"/></w:rPr><w:t>Theme font</w:t></w:r></w:p>`
  );
  const themed = (session: YrsSession) =>
    session
      .storySegments("body")
      .filter((segment) => segment.kind === "text")
      .map((segment) => [
        segment.text,
        (segment.attributes.fontFamily as { csTheme?: string } | undefined)
          ?.csTheme,
      ]);
  const { exported, latest, rebase } = await publishDocx(
    base,
    (session) =>
      session.insertText(
        { story: "body", paraId: "11111111", offset: 6 },
        " edited"
      ),
    (session) =>
      session.insertText(
        { story: "body", paraId: "22222222", offset: 5 },
        " typed"
      )
  );
  const current = await docxSession(exported, (await rebase()).state);
  const later = await docxSession(base, latest);
  try {
    expect(themed(later)).toContainEqual(["Theme typed font", "majorBidi"]);
    expect(themed(current)).toEqual(themed(later));
  } finally {
    current.destroy();
    later.destroy();
  }
});

test("a DOCX paragraph restored after the capture fails the rebase when its raw markup left with the export", async () => {
  const base = docx(
    `${paragraph("11111111", "Kept")}${paragraph(
      "22222222",
      "Marked",
      '<x:mark xmlns:x="urn:raw-inline"/>'
    )}${paragraph("33333333", "Tail")}`
  );
  const session = await docxSession(base);
  try {
    const [kept, marked] = session.paragraphs("body");
    session.deleteRange({
      story: "body",
      start: { paraId: kept.paraId, offset: kept.text.length },
      end: { paraId: marked.paraId, offset: marked.text.length },
    });
    const captured = session.encodeState();
    const exported = await exportOffice(
      base,
      checkpoint("docx", base, captured),
      fixed
    );
    expect(session.undo()).toBe(true);
    await expect(
      rebaseOffice(
        base,
        checkpoint("docx", base, captured),
        checkpoint("docx", base, session.encodeState()),
        exported
      )
    ).rejects.toThrow("needs content the export dropped");
  } finally {
    session.destroy();
  }
});

test("a PPTX slide restored after the capture fails the rebase: its part left with the export", async () => {
  const base = await fixture("apps/demo/public/betteroffice-demo.pptx");
  const doc = PptxDocument.openCollaborative(base, 4201);
  try {
    const [slide] = JSON.parse(doc.snapshotJson()).slides as Array<{
      id: string;
    }>;
    doc.deleteSlideJson(JSON.stringify({ slideId: slide.id }));
    const captured = doc.encodeStateAsUpdate();
    const exported = await exportOffice(
      base,
      checkpoint("pptx", base, captured),
      fixed
    );
    doc.undoJson();
    await expect(
      rebaseOffice(
        base,
        checkpoint("pptx", base, captured),
        checkpoint("pptx", base, doc.encodeStateAsUpdate()),
        exported
      )
    ).rejects.toThrow("bound to source XML the export dropped");
  } finally {
    doc.free();
  }
});

test("deleting after the capture the paragraph an editor comment ends in takes the export's reference mark with it", async () => {
  const breakParagraph = `<w:p w14:paraId="33333333"><w:r><w:br w:type="page"/></w:r></w:p>`;
  const removals: Array<[string, (session: YrsSession) => void]> = [
    [
      "the text, then Delete",
      (session) => {
        session.deleteRange({
          story: "body",
          start: { paraId: "11111111", offset: 0 },
          end: { paraId: "11111111", offset: 5 },
        });
        session.mergeParagraphs("body", "11111111", "forward");
      },
    ],
    [
      "the paragraph with its mark",
      (session) =>
        session.deleteRange({
          story: "body",
          start: { paraId: "11111111", offset: 0 },
          end: { paraId: session.paragraphs("body")[1].paraId, offset: 0 },
        }),
    ],
  ];
  for (const block of [table([["cell"]]), breakParagraph])
    for (const [name, remove] of removals) {
      const base = docx(
        `${paragraph("11111111", "alpha")}${block}${paragraph(
          "22222222",
          "After"
        )}`
      );
      const { exported, latest, rebase } = await publishDocx(
        base,
        (session) => comment(session, "11111111", 0, 5),
        remove
      );
      const current = await docxSession(exported, (await rebase()).state);
      const later = await docxSession(base, latest);
      try {
        expect(units(current), name).toBe(units(later));
        expect(() => current.yrsBlocksForStory("body")).not.toThrow();
      } finally {
        current.destroy();
        later.destroy();
      }
    }
});

test("a DOCX comment removed after the capture takes its reference mark, so the next publication rebases", async () => {
  const base = docx(
    `${paragraph("11111111", "alpha beta gamma")}${paragraph(
      "22222222",
      "delta epsilon"
    )}`
  );
  const first = await publishDocx(
    base,
    (session) => comment(session, "11111111", 0, 5),
    (session) =>
      session.applyRawOps("body", [
        { op: "removeComment", id: session.listComments()[0].id },
      ])
  );
  const rebased = (await first.rebase()).state;
  // The next cycle starts from that publication and deletes across the mark.
  const session = await docxSession(first.exported, rebased);
  try {
    expect(units(session)).toBe("alpha beta gamma¶delta epsilon¶");
    session.insertText(
      {
        story: "body",
        paraId: session.paragraphs("body")[1].paraId,
        offset: 0,
      },
      "c2 "
    );
    const captured = session.encodeState();
    const exported = await exportOffice(
      first.exported,
      checkpoint("docx", first.exported, captured),
      fixed
    );
    const [paragraphId] = session.paragraphs("body").map((p) => p.paraId);
    session.deleteRange({
      story: "body",
      start: { paraId: paragraphId, offset: 2 },
      end: { paraId: paragraphId, offset: 9 },
    });
    const next = await rebaseOffice(
      first.exported,
      checkpoint("docx", first.exported, captured),
      checkpoint("docx", first.exported, session.encodeState()),
      exported
    );
    const current = await docxSession(exported, next.state);
    try {
      expect(text(current)).toEqual(text(session));
    } finally {
      current.destroy();
    }
  } finally {
    session.destroy();
  }
});

test("a DOCX rebase refuses a state the editor cannot render: the export's reference mark ahead of a table", async () => {
  const base = docx(
    `${paragraph("11111111", "alpha")}${paragraph("22222222", "After")}`
  );
  const right = await createYrsSession({ clientId: 4102 });
  try {
    const { rebase } = await publishDocx(
      base,
      (session) => {
        // A comment whose text is gone: the export still writes its mark.
        comment(session, "11111111", 0, 5);
        session.deleteRange({
          story: "body",
          start: { paraId: "11111111", offset: 0 },
          end: { paraId: "11111111", offset: 5 },
        });
      },
      (session) => {
        // Concurrent edits: one editor opens After's slot with a table while
        // another deletes the empty paragraph before it.
        right.openDocx(base, false);
        right.loadState(session.encodeState());
        session.insertTable(
          { story: "body", paraId: "22222222", offset: 0 },
          1,
          1
        );
        right.mergeParagraphs("body", "11111111", "forward");
        session.loadState(right.encodeState());
        expect(units(session)).toBe("[table]After¶");
      }
    );
    await expect(rebase()).rejects.toThrow("does not render");
  } finally {
    right.destroy();
  }
});

test("any failure inside a rebase refuses it as a RebaseError", () => {
  const doc = new Y.Doc();
  doc.getMap("stories").set("body", new Y.Text("x"));
  const state = Y.encodeStateAsUpdate(doc);
  doc.destroy();
  const lineage = {
    maps: ["stories"],
    arrays: [],
    keyed: new Set(["stories"]),
    settle() {
      throw new TypeError("undefined is not an object");
    },
  };
  expect(() => transplant(lineage, state, state, state, new Map(), 7)).toThrow(
    RebaseError
  );
});
