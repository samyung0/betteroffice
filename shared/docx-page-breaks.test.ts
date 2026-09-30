import { beforeAll, describe, expect, test } from "bun:test";
import { createHash } from "node:crypto";
import { createYrsSession, type YrsSession } from "../packages/docx/src/yrs";
import { rezipContainer, unzipContainer } from "../packages/docx/src/wasm/opc";
import {
  applyOfficeCommands,
  exportOffice,
  inspectOffice,
  rebaseOffice,
  seedOffice,
} from "./office-checkpoint";

const fixed = { seed: "0".repeat(64), now: "2026-09-29T00:00:00.000Z" };
const W =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const OFFICE = "application/vnd.openxmlformats-officedocument.wordprocessingml";
const REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const p = (id: string, xml: string) => `<w:p w14:paraId="${id}">${xml}</w:p>`;
const BR = `<w:r><w:br w:type="page"/></w:r>`;
const COL = `<w:r><w:br w:type="column"/></w:r>`;
const tracked = (tag: "ins" | "del", xml: string) =>
  `<w:${tag} w:id="90" w:author="A" w:date="2026-09-01T00:00:00Z">${xml}</w:${tag}>`;
const table = (xml: string) =>
  `<w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc>${xml}</w:tc></w:tr></w:tbl>`;
const tail = p("22222222", run("tail"));

/** Each story kind: its id in the session and the part it saves to. */
const STORY = {
  body: ["body", "word/document.xml"],
  control: ["body:sdt0", "word/document.xml"],
  cell: ["body:t0:r0c0", "word/document.xml"],
  header: ["hf:rId20", "word/header1.xml"],
  footer: ["hf:rId21", "word/footer1.xml"],
  footnote: ["fn:1", "word/footnotes.xml"],
  nested: ["body:t0:r0c0:t0:r0c0", "word/document.xml"],
  headerCell: ["hf:rId20:t0:r0c0", "word/header1.xml"],
  endnote: ["en:1", "word/endnotes.xml"],
} as const;
type Where = keyof typeof STORY;
const OUTSIDE: Where[] = ["cell", "header", "footer", "footnote"];
const EVERY: Where[] = ["body", "control", ...OUTSIDE];
const NINE: Where[] = [...EVERY, "nested", "headerCell", "endnote"];

/** A package holding `content` in story `where`, with `tail` in the body. */
function docx(where: Where, content: string): Uint8Array {
  const body =
    where === "body"
      ? content
      : where === "control"
      ? `<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`
      : where === "cell"
      ? table(content)
      : where === "nested"
      ? table(`${table(content)}${p("5B5B5B5B", "")}`)
      : where === "footnote"
      ? p("11111111", `${run("x")}<w:r><w:footnoteReference w:id="1"/></w:r>`)
      : where === "endnote"
      ? p("11111111", `${run("x")}<w:r><w:endnoteReference w:id="1"/></w:r>`)
      : "";
  const notes =
    where === "endnote"
      ? {
          type: `<Override PartName="/word/endnotes.xml" ContentType="${OFFICE}.endnotes+xml"/>`,
          rel: `<Relationship Id="rId24" Type="${REL}/endnotes" Target="endnotes.xml"/>`,
        }
      : { type: "", rel: "" };
  const parts: Record<string, string> = {
    "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="${OFFICE}.comments+xml"/><Override PartName="/word/header1.xml" ContentType="${OFFICE}.header+xml"/><Override PartName="/word/footer1.xml" ContentType="${OFFICE}.footer+xml"/><Override PartName="/word/footnotes.xml" ContentType="${OFFICE}.footnotes+xml"/><Override PartName="/word/numbering.xml" ContentType="${OFFICE}.numbering+xml"/>${notes.type}</Types>`,
    "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
    "word/_rels/document.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="${REL}/comments" Target="comments.xml"/><Relationship Id="rId20" Type="${REL}/header" Target="header1.xml"/><Relationship Id="rId21" Type="${REL}/footer" Target="footer1.xml"/><Relationship Id="rId22" Type="${REL}/footnotes" Target="footnotes.xml"/><Relationship Id="rId23" Type="${REL}/numbering" Target="numbering.xml"/>${notes.rel}</Relationships>`,
    "word/document.xml": `<w:document ${W}><w:body>${body}${tail}<w:sectPr><w:headerReference w:type="default" r:id="rId20"/><w:footerReference w:type="default" r:id="rId21"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
    // One comment keeps the comments part, which a rebase requires the export to keep.
    "word/comments.xml": `<w:comments ${W}><w:comment w:id="1" w:author="R" w:date="2026-09-01T00:00:00Z"><w:p><w:r><w:t>c1</w:t></w:r></w:p></w:comment></w:comments>`,
    "word/numbering.xml": `<w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
    "word/header1.xml": `<w:hdr ${W}>${
      where === "header"
        ? content
        : where === "headerCell"
        ? `${table(content)}${p("66666665", "")}`
        : p("66666666", run("head"))
    }</w:hdr>`,
    "word/footer1.xml": `<w:ftr ${W}>${
      where === "footer" ? content : p("77777777", run("foot"))
    }</w:ftr>`,
    "word/footnotes.xml": `<w:footnotes ${W}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1">${
      where === "footnote" ? content : p("88888888", run("note"))
    }</w:footnote></w:footnotes>`,
    ...(where === "endnote" && {
      "word/endnotes.xml": `<w:endnotes ${W}><w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="1">${content}</w:endnote></w:endnotes>`,
    }),
  };
  return rezipContainer(
    Object.fromEntries(
      Object.entries(parts).map(([path, part]) => [
        path,
        new TextEncoder().encode(part),
      ])
    )
  );
}

beforeAll(() => seedOffice("docx", docx("body", p("11111111", run("x")))));
let clientId = 9800;
async function open(bytes: Uint8Array): Promise<YrsSession> {
  const session = await createYrsSession({ clientId: (clientId += 1) });
  session.openDocx(bytes, true);
  return session;
}
const checkpoint = (base: Uint8Array, state: Uint8Array) => ({
  format: "docx" as const,
  schemaVersion: 1 as const,
  baseSha256: createHash("sha256").update(base).digest("hex"),
  state,
});
const publish = (bytes: Uint8Array, session: YrsSession) =>
  exportOffice(bytes, checkpoint(bytes, session.encodeState()), fixed);

/** A segment as a unit: text, ¶, [kind] per embed ([pageBreak^] when it leads its paragraph), +/- for tracked ones. */
function units_(
  segment: ReturnType<YrsSession["storySegments"]>[number]
): string {
  if (segment.kind === "text") return segment.text;
  if (segment.kind === "pilcrow") return "¶";
  if (segment.payload.modelKind === "commentReference") return "";
  const mark = segment.attributes.ins ? "+" : segment.attributes.del ? "-" : "";
  return `${mark}[${segment.embedKind}${segment.payload.leading ? "^" : ""}]`;
}

/** A story's units (see `units_`). */
function units(session: YrsSession, story: string): string {
  return session.storySegments(story).map(units_).join("");
}

/** The story's part as text: [PB]/[CB] breaks, ¶ per paragraph, +{…}/-{…} tracked changes, T[…] tables. */
function view(bytes: Uint8Array, part: string): string {
  const xml = new TextDecoder().decode(unzipContainer(bytes)[part]);
  const from = Math.max(
    xml.indexOf("<w:body>"),
    xml.indexOf('<w:footnote w:id="1"'),
    xml.indexOf('<w:endnote w:id="1"'),
    0
  );
  return [
    ...xml
      .slice(from)
      .matchAll(/<(\/?)w:(p|br|t|delText|ins|del|tbl)\b([^>]*?)\/?>([^<]*)/g),
  ]
    .map(([, close, tag, attrs, text]) => {
      if (tag === "t" || tag === "delText") return close ? "" : text;
      if (tag === "br")
        return /"page"/.test(attrs)
          ? "[PB]"
          : /"column"/.test(attrs)
          ? "[CB]"
          : "";
      if (tag === "p") return close ? "¶" : "";
      if (tag === "tbl") return close ? "]" : "T[";
      return close ? "}" : tag === "ins" ? "+{" : "-{";
    })
    .join("");
}

const breaks = (bytes: Uint8Array, part: string, type = "page") =>
  new TextDecoder()
    .decode(unzipContainer(bytes)[part])
    .split(`<w:br w:type="${type}"/>`).length - 1;

/** The toolbar's page break: split at the caret, then the break at the new paragraph's start. */
function toolbarBreak(
  session: YrsSession,
  story: string,
  paraId: string,
  offset: number
): string {
  const { secondParaId } = session.splitParagraph({ story, paraId, offset });
  session.insertPageBreak({ story, paraId: secondParaId, offset: 0 });
  return secondParaId;
}

const paraOf = (session: YrsSession, story: string, text: string) =>
  session.paragraphs(story).find((paragraph) => paragraph.text === text)!
    .paraId;

/**
 * Publishes after `edit`, then twice more: the story's units (without the
 * body's tail paragraph) before and after each publication, and each export.
 */
async function publications(
  bytes: Uint8Array,
  story: string,
  edit: (session: YrsSession) => void
) {
  let session = await open(bytes);
  edit(session);
  const own = (text: string) => text.replace(/tail¶$/, "");
  const seen = [own(units(session, story))];
  const exports: Uint8Array[] = [];
  for (let publication = 0; publication < 3; publication += 1) {
    bytes = await publish(bytes, session);
    exports.push(bytes);
    session.destroy();
    session = await open(bytes);
    seen.push(own(units(session, story)));
  }
  session.destroy();
  return { seen, exports };
}

/** Captures after `before`, publishes, applies `after` and rebases onto the export's seed. */
async function rebase(
  bytes: Uint8Array,
  story: string,
  before: (session: YrsSession) => void,
  after: (session: YrsSession) => void
) {
  const session = await open(bytes);
  before(session);
  const captured = session.encodeState();
  const exported = await exportOffice(
    bytes,
    checkpoint(bytes, captured),
    fixed
  );
  after(session);
  const latest = units(session, story);
  const { state } = await rebaseOffice(
    bytes,
    checkpoint(bytes, captured),
    checkpoint(bytes, session.encodeState()),
    exported
  );
  session.destroy();
  const rebased = await createYrsSession({ clientId: (clientId += 1) });
  rebased.openDocx(exported, false);
  rebased.loadState(state);
  const result = {
    latest,
    rebased: units(rebased, story),
    next: await publish(exported, rebased),
  };
  rebased.destroy();
  return result;
}

describe("the toolbar's page break (split, then the break at the new paragraph's start)", () => {
  test.each(EVERY)(
    "in %s saves in its place and seeds back as the same unit",
    async (where) => {
      const [story, part] = STORY[where];
      const { seen, exports } = await publications(
        docx(where, p("44444444", run("abcdef"))),
        story,
        (session) => toolbarBreak(session, story, "44444444", 3)
      );
      expect(seen).toEqual(Array(4).fill("abc¶[pageBreak]def¶"));
      for (const bytes of exports)
        expect(view(bytes, part)).toContain("abc[PB]¶def¶");
    }
  );

  test.each(EVERY)(
    "at the start of %s leads its paragraph from the first publication on",
    async (where) => {
      const [story, part] = STORY[where];
      const { seen, exports } = await publications(
        docx(where, p("44444444", run("abcdef"))),
        story,
        (session) =>
          session.insertPageBreak({ story, paraId: "44444444", offset: 0 })
      );
      expect(seen).toEqual([
        "[pageBreak]abcdef¶",
        ...Array(3).fill("[pageBreak^]abcdef¶"),
      ]);
      for (const bytes of exports)
        expect(view(bytes, part)).toContain("[PB]abcdef¶");
    }
  );

  test.each(EVERY)(
    "in %s stays single through Enter after a publication",
    async (where) => {
      const [story, part] = STORY[where];
      let bytes = docx(where, p("44444444", run("abcdef")));
      let session = await open(bytes);
      toolbarBreak(session, story, "44444444", 3);
      bytes = await publish(bytes, session);
      session.destroy();
      session = await open(bytes);
      session.splitParagraph({
        story,
        paraId: paraOf(session, story, "def"),
        offset: 3,
      });
      bytes = await publish(bytes, session);
      session.destroy();
      expect(view(bytes, part)).toContain("abc[PB]¶de¶f¶");
    }
  );
});

// Offsets in the break's paragraph count the break that opens its slot.
const after: Array<[string, (session: YrsSession, story: string) => void]> = [
  [
    "Enter mid-text after the break",
    (s, story) =>
      s.splitParagraph({ story, paraId: paraOf(s, story, "def"), offset: 3 }),
  ],
  [
    "a second break",
    (s, story) => toolbarBreak(s, story, paraOf(s, story, "def"), 3),
  ],
  [
    "deleting the text after the break",
    (s, story) => {
      const paraId = paraOf(s, story, "def");
      s.deleteRange({
        story,
        start: { paraId, offset: 1 },
        end: { paraId, offset: 4 },
      });
    },
  ],
  [
    "Enter right after the break",
    (s, story) =>
      s.splitParagraph({ story, paraId: paraOf(s, story, "def"), offset: 1 }),
  ],
  [
    "Backspace over the break",
    (s, story) =>
      s.deleteAt(
        { story, paraId: paraOf(s, story, "def"), offset: 1 },
        "backward"
      ),
  ],
  [
    "typing above the break",
    (s, story) => s.insertText({ story, paraId: "44444444", offset: 1 }, "Q"),
  ],
];
describe.each(EVERY)(
  "an edit after a publication's capture lands beside a break in %s",
  (where) => {
    test.each(after)("%s", async (_, edit) => {
      const [story, part] = STORY[where];
      const { latest, rebased, next } = await rebase(
        docx(where, p("44444444", run("abcdef"))),
        story,
        (session) => toolbarBreak(session, story, "44444444", 3),
        (session) => edit(session, story)
      );
      expect(rebased).toBe(latest);
      expect(breaks(next, part)).toBe(latest.split("[pageBreak").length - 1);
    });
  }
);

describe.each(EVERY)("breaks from the file in %s", (where) => {
  test.each<[string, string, string, string]>([
    [
      "a trailing page break",
      `${p("44444444", `${run("abc")}${BR}`)}${p("55555555", run("def"))}`,
      "abc¶[pageBreak]def¶",
      "abcX[PB]¶def¶",
    ],
    [
      "a leading page break",
      `${p("33333333", run("prev"))}${p("44444444", `${BR}${run("abc")}`)}`,
      "prev¶[pageBreak^]abc¶",
      "prev¶[PB]abcX¶",
    ],
    [
      "a page break mid-paragraph",
      p("44444444", `${run("ab")}${BR}${run("cd")}`),
      "abcd¶[pageBreak]",
      "abcdX[PB]¶",
    ],
    [
      "two leading page breaks",
      p("44444444", `${BR}${BR}${run("abc")}`),
      "[pageBreak^][pageBreak^]abc¶",
      "[PB][PB]abcX¶",
    ],
    [
      "a leading column break",
      `${p("33333333", run("prev"))}${p("44444444", `${COL}${run("abc")}`)}`,
      "prev¶[columnBreak^]abc¶",
      "prev¶[CB]abcX¶",
    ],
    [
      "a break-only paragraph",
      `${p("33333333", run("prev"))}${p("44444444", BR)}${p(
        "55555555",
        run("def")
      )}`,
      "prev¶¶[pageBreak]def¶",
      "prev¶X[PB]¶def¶",
    ],
  ])(
    "%s seeds as units and survives an edit of its paragraph",
    async (_, xml, seeded, saved) => {
      const [story, part] = STORY[where];
      const bytes = docx(where, xml);
      const session = await open(bytes);
      expect(units(session, story)).toStartWith(seeded);
      const end = session
        .paragraphSpans(story)
        .find(({ paraId }) => paraId === "44444444")!.length;
      session.insertText({ story, paraId: "44444444", offset: end }, "X");
      const out = await publish(bytes, session);
      session.destroy();
      expect(view(out, part)).toContain(saved);
    }
  );
});

describe.each(EVERY)(
  "a page break leading a paragraph of the file in %s, edited as in Word",
  (where) => {
    const xml = `${p("33333333", run("prev"))}${p(
      "44444444",
      `${BR}${run("abcdef")}`
    )}`;
    test.each<[string, (session: YrsSession, story: string) => void, string]>([
      [
        "Enter before it adds an empty paragraph before the break",
        (s, story) =>
          s.splitParagraph({ story, paraId: "44444444", offset: 0 }),
        "prev¶¶[PB]abcdef¶",
      ],
      [
        "Enter mid-text keeps it before the first half",
        (s, story) =>
          s.splitParagraph({ story, paraId: "44444444", offset: 4 }),
        "prev¶[PB]abc¶def¶",
      ],
      [
        "Enter right after it leaves it before an empty paragraph",
        (s, story) =>
          s.splitParagraph({ story, paraId: "44444444", offset: 1 }),
        "prev[PB]¶¶abcdef¶",
      ],
      [
        "Backspace right after it removes it alone",
        (s, story) =>
          s.deleteAt({ story, paraId: "44444444", offset: 1 }, "backward"),
        "prev¶abcdef¶",
      ],
      [
        "Delete at the previous paragraph's end removes it alone",
        (s, story) =>
          s.deleteAt({ story, paraId: "33333333", offset: 4 }, "forward"),
        "prev¶abcdef¶",
      ],
      [
        "deleting the text after it leaves it before an empty paragraph",
        (s, story) =>
          s.deleteRange({
            story,
            start: { paraId: "44444444", offset: 1 },
            end: { paraId: "44444444", offset: 7 },
          }),
        "prev[PB]¶¶",
      ],
    ])("%s", async (_, edit, saved) => {
      const [story, part] = STORY[where];
      const { seen, exports } = await publications(
        docx(where, xml),
        story,
        (session) => edit(session, story)
      );
      expect(view(exports[0], part)).toContain(saved);
      expect(seen.slice(2)).toEqual([seen[1], seen[1]]);
    });
  }
);

describe("tracked breaks", () => {
  const author = { name: "Reviewer", date: "2026-09-29T00:00:00Z" };
  test.each(
    EVERY.flatMap((where) =>
      (
        [
          [
            "a deleted leading break",
            p("44444444", `${tracked("del", BR)}${run("abc")}`),
            "-[pageBreak^]abc¶",
            "-{[PB]}abcX¶",
          ],
          [
            "an inserted leading break",
            p("44444444", `${tracked("ins", BR)}${run("abc")}`),
            "+[pageBreak^]abc¶",
            "+{[PB]}abcX¶",
          ],
          [
            "an inserted trailing break",
            `${p("44444444", `${run("abc")}${tracked("ins", BR)}`)}${p(
              "55555555",
              run("def")
            )}`,
            "abc¶+[pageBreak]def¶",
            "abcX+{[PB]}¶def¶",
          ],
        ] as const
      ).map(
        ([name, xml, seeded, saved]) =>
          [where, name, xml, seeded, saved] as const
      )
    )
  )(
    "%s: %s keeps its tracking through an edit",
    async (where, _, xml, seeded, saved) => {
      const [story, part] = STORY[where];
      const bytes = docx(where, xml);
      const session = await open(bytes);
      expect(units(session, story)).toStartWith(seeded);
      const end = session
        .paragraphSpans(story)
        .find(({ paraId }) => paraId === "44444444")!.length;
      session.insertText({ story, paraId: "44444444", offset: end }, "X");
      const out = await publish(bytes, session);
      session.destroy();
      expect(view(out, part)).toContain(saved);
    }
  );

  test("a suggested deletion of a break saves as w:del, and accepting it removes the break", async () => {
    const bytes = docx(
      "body",
      `${p("33333333", `${run("abc")}${BR}`)}${p("44444444", run("def"))}`
    );
    const session = await open(bytes);
    session.deleteAt(
      { story: "body", paraId: "33333333", offset: 3 },
      "forward",
      author
    );
    const suggested = await publish(bytes, session);
    expect(view(suggested, STORY.body[1])).toContain("abc-{[PB]}¶def¶");
    // The break opens the second paragraph's slot.
    session.acceptChange({
      story: "body",
      start: { paraId: "44444444", offset: 0 },
      end: { paraId: "44444444", offset: 1 },
    });
    expect(view(await publish(bytes, session), STORY.body[1])).toContain(
      "abc¶def¶"
    );
    session.destroy();
  });

  test("the toolbar's break in suggesting mode saves as w:ins", async () => {
    const bytes = docx("body", p("44444444", run("abcdef")));
    const session = await open(bytes);
    const { secondParaId } = session.splitParagraph(
      { story: "body", paraId: "44444444", offset: 3 },
      author
    );
    session.insertPageBreak(
      { story: "body", paraId: secondParaId, offset: 0 },
      author
    );
    expect(units(session, "body")).toStartWith("abc¶+[pageBreak]def¶");
    const out = await publish(bytes, session);
    session.destroy();
    expect(view(out, STORY.body[1])).toContain("abc+{[PB]}¶def¶");
    const reopened = await open(out);
    expect(units(reopened, "body")).toStartWith("abc¶+[pageBreak]def¶");
    reopened.destroy();
  });
});

test.each(["cell", "header"] as Where[])(
  "a comment after a break before a nested table in a %s keeps its range",
  async (where) => {
    const [story] = STORY[where];
    let bytes = docx(
      where,
      `${p("33333333", run("before"))}${table(p("66666666", run("inner")))}${p(
        "44444444",
        run("xyzw")
      )}`
    );
    let session = await open(bytes);
    session.insertPageBreak({ story, paraId: "44444444", offset: 0 });
    // The break and the table open the paragraph's slot: "yz" is 3..5.
    session.addComment(
      [
        {
          story,
          start: { paraId: "44444444", offset: 3 },
          end: { paraId: "44444444", offset: 5 },
        },
      ],
      "R",
      "2026-09-29T00:00:00Z",
      [
        {
          type: "paragraph",
          content: [{ type: "run", content: [{ type: "text", text: "c" }] }],
        },
      ]
    );
    const covered = (s: YrsSession) =>
      s.listComments().flatMap(({ id }) =>
        s.resolveComment(id).map(({ start, end }) =>
          s
            .storySegments(story)
            .flatMap((segment) =>
              segment.kind === "text"
                ? [...segment.text]
                : segment.kind === "pilcrow"
                ? ["¶"]
                : [segment.payload.modelKind === "commentReference" ? "" : "#"]
            )
            .slice(start, end)
            .join("")
        )
      );
    const seen = [covered(session)];
    for (let publication = 0; publication < 3; publication += 1) {
      bytes = await publish(bytes, session);
      session.destroy();
      session = await open(bytes);
      seen.push(covered(session));
    }
    session.destroy();
    expect(seen).toEqual(Array(4).fill(["yz"]));
  }
);

describe("a break with no paragraph before it and no text to lead", () => {
  const start = (where: Where) =>
    docx(
      where,
      `${table(p("66666666", run("inner")))}${p("44444444", run("xyzw"))}`
    );
  test.each(EVERY)(
    "in %s saves as its own paragraph, which the next seed holds",
    async (where) => {
      const [story, part] = STORY[where];
      const { seen, exports } = await publications(
        start(where),
        story,
        (session) =>
          session.insertPageBreak({ story, paraId: "44444444", offset: 0 })
      );
      expect(seen[0]).toStartWith("[pageBreak][table]xyzw¶");
      expect(seen.slice(1)).toEqual(Array(3).fill(seen[1]));
      expect(seen[1]).toStartWith("¶[pageBreak][table]xyzw¶");
      expect(view(exports[2], part)).toContain("[PB]¶T[inner¶]xyzw¶");
    }
  );

  test("an insertion made there after a capture refuses the rebase; one elsewhere lands", async () => {
    const bytes = start("cell");
    const [story] = STORY.cell;
    const breakFirst = (session: YrsSession) =>
      session.insertPageBreak({ story, paraId: "44444444", offset: 0 });
    await expect(
      rebase(bytes, story, breakFirst, (session) =>
        session.splitParagraph({ story, paraId: "44444444", offset: 0 })
      )
    ).rejects.toThrow("Office rebase:");
    const { latest, rebased } = await rebase(
      bytes,
      story,
      breakFirst,
      (session) =>
        session.insertText({ story, paraId: "44444444", offset: 3 }, "Q")
    );
    expect(rebased).toBe(`¶${latest}`);
  });
});

const S = (id: number) => `<w:commentRangeStart w:id="${id}"/>`;
const E = (id: number) => `<w:commentRangeEnd w:id="${id}"/>`;
const ref = (id: number) =>
  `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`;
/** What each comment covers: its story units, `#` for an embed. */
const covered = (session: YrsSession, story: string) =>
  session.listComments().flatMap(({ id }) =>
    session.resolveComment(id).map(({ start, end }) =>
      session
        .storySegments(story)
        .flatMap((segment) =>
          segment.kind === "text"
            ? [...segment.text]
            : segment.kind === "pilcrow"
            ? ["¶"]
            : [segment.payload.modelKind === "commentReference" ? "" : "#"]
        )
        .slice(start, end)
        .join("")
    )
  );

/** An editor comment in `story` from `start` to `end` (paragraph id, offset). */
const editorComment =
  (story: string, start: [string, number], end: [string, number]) =>
  (session: YrsSession) =>
    void session.addComment(
      [
        {
          story,
          start: { paraId: start[0], offset: start[1] },
          end: { paraId: end[0], offset: end[1] },
        },
      ],
      "R",
      "2026-09-29T00:00:00Z",
      [
        {
          type: "paragraph",
          content: [{ type: "run", content: [{ type: "text", text: "c" }] }],
        },
      ]
    );

describe.each(EVERY)(
  "Enter at the start of a break's paragraph, then Delete or Backspace, in %s",
  (where) => {
    const xmls = [
      [
        "a break leading the paragraph's text",
        `${p("33333333", run("prev"))}${p(
          "44444444",
          `${BR}${run("abcdef")}`
        )}`,
      ],
      [
        "a break closing the paragraph before",
        `${p("33333333", `${run("prev")}${BR}`)}${p(
          "44444444",
          run("abcdef")
        )}`,
      ],
    ] as const;
    const author = { name: "Reviewer", date: "2026-09-29T00:00:00Z" };
    test.each(
      xmls.flatMap(([name, xml]) =>
        (["forward", "backward"] as const).flatMap((direction) =>
          (["plain", "suggesting"] as const).map(
            (mode) => [name, direction, mode, xml] as const
          )
        )
      )
    )("restores the document: %s, %s, %s", async (_, direction, mode, xml) => {
      const [story] = STORY[where];
      const session = await open(docx(where, xml));
      const before = units(session, story);
      const by = mode === "suggesting" ? author : undefined;
      const { firstParaId } = session.splitParagraph(
        { story, paraId: "44444444", offset: 0 },
        by
      );
      if (direction === "forward")
        session.deleteAt(
          { story, paraId: firstParaId, offset: 0 },
          "forward",
          by
        );
      else
        session.deleteAt(
          { story, paraId: "44444444", offset: 0 },
          "backward",
          by
        );
      expect(units(session, story)).toBe(before);
      // Suggesting mode withdraws the author's own pending paragraph.
      expect(
        session
          .paragraphs(story)
          .filter(({ properties }) => properties.pPrIns || properties.pPrDel)
      ).toEqual([]);
      session.destroy();
    });
  }
);

describe.each(["body", "control", "cell"] as Where[])(
  "the editor keeps space-before after a break exactly where the save does, in %s",
  (where) => {
    const flags = (session: YrsSession, story: string) =>
      (
        session.yrsBlocksForStory(story) as Array<{
          kind: string;
          keepsLeadingSpacing?: boolean;
        }>
      )
        .filter(({ kind }) => kind === "pageBreak")
        .map(({ keepsLeadingSpacing }) => keepsLeadingSpacing === true);
    const lead = `${p("33333333", run("prev"))}${p(
      "44444444",
      `${BR}${run("abc")}`
    )}`;
    test.each<
      [string, string, (session: YrsSession, story: string) => void, boolean[]]
    >([
      ["a break leading text", lead, () => {}, [true]],
      [
        "the text after a leading break deleted",
        lead,
        (s, story) =>
          s.deleteRange({
            story,
            start: { paraId: "44444444", offset: 1 },
            end: { paraId: "44444444", offset: 4 },
          }),
        [false],
      ],
      [
        // Its reference mark, which the save writes, is the text the break leads.
        "an editor comment over a leading break, the text after it deleted",
        lead,
        (s, story) => {
          s.addComment(
            [
              {
                story,
                start: { paraId: "44444444", offset: 0 },
                end: { paraId: "44444444", offset: 1 },
              },
            ],
            "R",
            "2026-09-29T00:00:00Z",
            [
              {
                type: "paragraph",
                content: [
                  { type: "run", content: [{ type: "text", text: "c" }] },
                ],
              },
            ]
          );
          s.deleteRange({
            story,
            start: { paraId: "44444444", offset: 1 },
            end: { paraId: "44444444", offset: 4 },
          });
        },
        [true],
      ],
      [
        // Its reference mark goes after the break, which leads it.
        "an editor comment ending before a leading break, the text after it deleted",
        lead,
        (s, story) => {
          editorComment(story, ["33333333", 0], ["44444444", 0])(s);
          s.deleteRange({
            story,
            start: { paraId: "44444444", offset: 1 },
            end: { paraId: "44444444", offset: 4 },
          });
        },
        [true],
      ],
      [
        "Enter right after a leading break",
        lead,
        (s, story) =>
          s.splitParagraph({ story, paraId: "44444444", offset: 1 }),
        [false],
      ],
      [
        "the toolbar's break",
        p("44444444", run("abcdef")),
        (s, story) => toolbarBreak(s, story, "44444444", 3),
        [false],
      ],
      [
        "a break opening the story before text",
        p("44444444", run("abc")),
        (s, story) =>
          s.insertPageBreak({ story, paraId: "44444444", offset: 0 }),
        [true],
      ],
      [
        "a break opening the story before an empty paragraph",
        `${p("44444444", "")}${p("55555555", run("abc"))}`,
        (s, story) =>
          s.insertPageBreak({ story, paraId: "44444444", offset: 0 }),
        [false],
      ],
      [
        "a break opening the story before a table",
        `${table(p("66666666", run("inner")))}${p("44444444", run("xyzw"))}`,
        (s, story) =>
          s.insertPageBreak({ story, paraId: "44444444", offset: 0 }),
        [false],
      ],
    ])("%s", async (_, xml, edit, expected) => {
      const [story] = STORY[where];
      const bytes = docx(where, xml);
      const session = await open(bytes);
      edit(session, story);
      const editor = flags(session, story);
      const reopened = await open(await publish(bytes, session));
      session.destroy();
      expect(editor).toEqual(expected);
      expect(flags(reopened, story)).toEqual(editor);
      reopened.destroy();
    });
  }
);

describe.each(EVERY)(
  "comments around a break leading a paragraph in %s",
  (where) => {
    test.each([
      [
        "opening before the break",
        `${S(1)}${BR}${run("abc")}${E(1)}${ref(1)}${run("def")}`,
        "#abc",
      ],
      ["over only the break", `${S(1)}${BR}${E(1)}${ref(1)}${run("abc")}`, "#"],
      [
        "opening after the break",
        `${BR}${S(1)}${run("abc")}${E(1)}${ref(1)}`,
        "abc",
      ],
      [
        // An empty range resolves to no range.
        "closing before the break",
        `${S(1)}${E(1)}${ref(1)}${BR}${run("abc")}`,
        null,
      ],
    ])(
      "a comment %s keeps its range across publications",
      async (_, xml, range) => {
        const [story] = STORY[where];
        let bytes = docx(
          where,
          `${p("33333333", run("prev"))}${p("44444444", xml)}`
        );
        const seen: string[][] = [];
        for (let publication = 0; publication < 3; publication += 1) {
          const session = await open(bytes);
          seen.push(covered(session, story));
          session.insertText(
            { story: "body", paraId: "22222222", offset: 0 },
            "x"
          );
          bytes = await publish(bytes, session);
          session.destroy();
        }
        expect(seen).toEqual(Array(3).fill(range === null ? [] : [range]));
      }
    );
  }
);

test.each(EVERY)(
  "an editor comment starting at the toolbar's break in %s saves starting after the break",
  async (where) => {
    const [story] = STORY[where];
    let bytes = docx(where, p("44444444", run("abcdef")));
    let session = await open(bytes);
    const second = toolbarBreak(session, story, "44444444", 3);
    session.addComment(
      [
        {
          story,
          start: { paraId: second, offset: 0 },
          end: { paraId: second, offset: 3 },
        },
      ],
      "R",
      "2026-09-29T00:00:00Z",
      [
        {
          type: "paragraph",
          content: [{ type: "run", content: [{ type: "text", text: "c" }] }],
        },
      ]
    );
    const seen = [covered(session, story).filter((range) => range !== "")];
    for (let publication = 0; publication < 2; publication += 1) {
      bytes = await publish(bytes, session);
      session.destroy();
      session = await open(bytes);
      seen.push(covered(session, story).filter((range) => range !== ""));
    }
    session.destroy();
    expect(seen).toEqual([["#de"], ["de"], ["de"]]);
  }
);

test("an empty list item before a paragraph a break leads keeps its number", async () => {
  const num = `<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>`;
  const session = await open(
    docx(
      "body",
      `<w:p>${num}${run("one")}</w:p><w:p>${num}</w:p>${p(
        "44444444",
        `${BR}${run("abc")}`
      )}`
    )
  );
  const markers = (
    session.yrsBlocksForStory("body") as Array<{
      kind: string;
      attrs?: { listMarker?: unknown; listMarkerHidden?: boolean };
    }>
  )
    .filter(({ attrs }) => attrs?.listMarker)
    .map(({ attrs }) => attrs?.listMarkerHidden === true);
  session.destroy();
  expect(markers).toEqual([false, false]);
});

/** The part's view without tracked-change marks. */
const plain = (bytes: Uint8Array, part: string) =>
  view(bytes, part).replace(/[+-]\{|\}/g, "");
const field = (result: string) =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DATE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${result}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
describe.each(EVERY)("a page break between a paragraph's %s", (where) => {
  test.each(
    [
      [
        "simple field and text",
        `<w:fldSimple w:instr=" DATE ">${run("2026")}</w:fldSimple>`,
      ],
      ["complex field and text", field(run("2026"))],
      [
        "hyperlink and text",
        `<w:hyperlink w:anchor="target">${run("2026")}</w:hyperlink>`,
      ],
      ["tracked insertion and text", tracked("ins", run("2026"))],
      [
        "simple field in a tracked insertion and text",
        tracked(
          "ins",
          `<w:fldSimple w:instr=" DATE ">${run("2026")}</w:fldSimple>`
        ),
      ],
      [
        "simple field in a hyperlink and text",
        `<w:hyperlink w:anchor="target"><w:fldSimple w:instr=" DATE ">${run(
          "2026"
        )}</w:fldSimple></w:hyperlink>`,
      ],
      [
        "hyperlink in a tracked insertion and text",
        tracked(
          "ins",
          `<w:hyperlink w:anchor="target">${run("2026")}</w:hyperlink>`
        ),
      ],
    ].flatMap(([name, before]) => [
      [name, "after text", `${run("a ")}${before}`],
      [name, "opening the paragraph", before],
    ])
  )(
    "%s, %s, stays in place until the paragraph's text changes",
    async (_, __, before) => {
      const [story, part] = STORY[where];
      const bytes = docx(where, p("44444444", `${before}${BR}${run("x")}`));
      const session = await open(bytes);
      session.insertText({ story: "body", paraId: "22222222", offset: 0 }, "z");
      const untouched = await publish(bytes, session);
      const end = session
        .paragraphSpans(story)
        .find(({ paraId }) => paraId === "44444444")!.length;
      session.insertText({ story, paraId: "44444444", offset: end }, "X");
      const edited = await publish(bytes, session);
      session.destroy();
      expect(plain(untouched, part)).toContain("2026[PB]x¶");
      expect(plain(edited, part)).toContain("2026xX[PB]¶");
      const reopened = await open(untouched);
      expect(plain(await publish(untouched, reopened), part)).toContain(
        "2026[PB]x¶"
      );
      reopened.destroy();
    }
  );
});

test.each(EVERY)(
  "a comment from the paragraph before, closing at a break leading the next, keeps its range in %s",
  async (where) => {
    const [story] = STORY[where];
    let bytes = docx(
      where,
      `${p("33333333", `${S(1)}${run("prev")}`)}${p(
        "44444444",
        `${E(1)}${BR}${run("abc")}${ref(1)}`
      )}`
    );
    const seen: string[][] = [];
    for (let publication = 0; publication < 3; publication += 1) {
      const session = await open(bytes);
      seen.push(covered(session, story));
      session.insertText({ story: "body", paraId: "22222222", offset: 0 }, "x");
      bytes = await publish(bytes, session);
      session.destroy();
    }
    expect(seen).toEqual(Array(3).fill(["prev¶"]));
  }
);

test("a heading retyped after its break's paragraph was published empty keeps no space-before, in the editor and the save", async () => {
  const bytes = docx(
    "body",
    `${p("33333333", run("prev"))}${p("44444444", `${BR}${run("abc")}`)}`
  );
  const flags = (session: YrsSession) =>
    (
      session.yrsBlocksForStory("body") as Array<{
        kind: string;
        keepsLeadingSpacing?: boolean;
      }>
    )
      .filter(({ kind }) => kind === "pageBreak")
      .map(({ keepsLeadingSpacing }) => keepsLeadingSpacing === true);
  const retype = (session: YrsSession) =>
    session.insertText({ story: "body", paraId: "44444444", offset: 1 }, "New");
  const session = await open(bytes);
  session.deleteRange({
    story: "body",
    start: { paraId: "44444444", offset: 1 },
    end: { paraId: "44444444", offset: 4 },
  });
  const captured = session.encodeState();
  const exported = await exportOffice(
    bytes,
    checkpoint(bytes, captured),
    fixed
  );
  retype(session);
  // Without the publication the break still leads the heading.
  expect(flags(session)).toEqual([true]);
  expect(view(await publish(bytes, session), "word/document.xml")).toContain(
    "prev¶[PB]New¶"
  );
  const { state } = await rebaseOffice(
    bytes,
    checkpoint(bytes, captured),
    checkpoint(bytes, session.encodeState()),
    exported
  );
  session.destroy();
  const rebased = await createYrsSession({ clientId: (clientId += 1) });
  rebased.openDocx(exported, false);
  rebased.loadState(state);
  expect(flags(rebased)).toEqual([false]);
  expect(view(await publish(exported, rebased), "word/document.xml")).toContain(
    "prev[PB]¶New¶"
  );
  rebased.destroy();
});

describe("an empty paragraph that ends a section keeps its section break", () => {
  const sect = `<w:pPr><w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr>`;
  const sections = (bytes: Uint8Array) =>
    new TextDecoder()
      .decode(unzipContainer(bytes)["word/document.xml"])
      .split("<w:sectPr").length - 1;
  const author = { name: "Reviewer", date: "2026-09-29T00:00:00Z" };
  test.each(
    (
      [
        ["page break", `${p("44444444", `${BR}${run("abc")}`)}`, "prev¶¶abc¶"],
        [
          "column break",
          `${p("44444444", `${COL}${run("abc")}`)}`,
          "prev¶¶abc¶",
        ],
        [
          "table",
          `${table(p("66666666", run("cell")))}${p("44444444", run("abc"))}`,
          null,
        ],
      ] as const
    ).flatMap(([name, slot, removed]) =>
      (["forward", "backward"] as const).flatMap((direction) =>
        (["plain", "suggesting"] as const).map(
          (mode) => [name, direction, mode, slot, removed] as const
        )
      )
    )
  )("before a %s: %s, %s", async (_, direction, mode, slot, removed) => {
    const bytes = docx(
      "body",
      `${p(
        "33333333",
        run("prev")
      )}<w:p w14:paraId="55555555">${sect}</w:p>${slot}`
    );
    const session = await open(bytes);
    const before = units(session, "body");
    const by = mode === "suggesting" ? author : undefined;
    if (direction === "forward")
      session.deleteAt(
        { story: "body", paraId: "55555555", offset: 0 },
        "forward",
        by
      );
    else
      session.deleteAt(
        { story: "body", paraId: "44444444", offset: 0 },
        "backward",
        by
      );
    // Before a break, Delete or Backspace removes the break; before a table, nothing.
    const expected =
      removed === null
        ? before
        : mode === "plain"
        ? `${removed}tail¶`
        : before.replace("[", "-[");
    expect(units(session, "body")).toBe(expected);
    const out = await publish(bytes, session);
    session.destroy();
    expect(sections(out)).toBe(2);
  });
});

describe.each(["body", "cell", "header", "footnote"] as Where[])(
  "in a file without w14:paraId, a page break beside content in %s",
  (where) => {
    test.each(
      (
        [
          [
            "a simple field",
            `<w:fldSimple w:instr=" DATE ">${run("2026")}</w:fldSimple>`,
          ],
          [
            "a hyperlink",
            `<w:hyperlink w:anchor="target">${run("2026")}</w:hyperlink>`,
          ],
          ["a tracked insertion", tracked("ins", run("2026"))],
        ] as const
      ).flatMap(([name, before]) =>
        (
          ["Enter in a paragraph above", "a paragraph above deleted"] as const
        ).map((edit) => [name, edit, before] as const)
      )
    )("after %s stays in place after %s", async (_, edit, before) => {
      const [story, part] = STORY[where];
      const xml = `<w:p>${run("one")}</w:p><w:p>${run("above")}</w:p><w:p>${run(
        "a "
      )}${before}${BR}${run("x")}</w:p>`;
      const bytes = docx(where, xml);
      const session = await open(bytes);
      const id = (text: string) =>
        session.paragraphs(story).find((paragraph) => paragraph.text === text)!
          .paraId;
      if (edit === "Enter in a paragraph above")
        session.splitParagraph({ story, paraId: id("above"), offset: 2 });
      else
        session.deleteRange({
          story,
          start: { paraId: id("one"), offset: 3 },
          end: { paraId: id("above"), offset: 5 },
        });
      const out = await publish(bytes, session);
      session.destroy();
      expect(plain(out, part)).toContain("a 2026[PB]x¶");
    });
  }
);

/** The paragraph `paraId`'s content in order: B break, S/E bookmark start/end, F simple field, T text. */
function order(bytes: Uint8Array, part: string, paraId: string): string {
  const xml = new TextDecoder().decode(unzipContainer(bytes)[part]);
  const at = xml.indexOf(`w14:paraId="${paraId}"`);
  const paragraph = xml.slice(at, xml.indexOf("</w:p>", at));
  return [
    ...paragraph.matchAll(/<w:(br|bookmarkStart|bookmarkEnd|fldSimple|t)\b/g),
  ]
    .map(
      ([, tag]) =>
        ({
          br: "B",
          bookmarkStart: "S",
          bookmarkEnd: "E",
          fldSimple: "F",
          t: "T",
        }[tag!])
    )
    .join("")
    .replace(/T+/g, "T");
}

describe.each(EVERY)(
  "bookmarks and empty fields around a leading break in %s",
  (where) => {
    const bookmark = (id: number, kind: "Start" | "End") =>
      kind === "Start"
        ? `<w:bookmarkStart w:id="${id}" w:name="_Toc${id}"/>`
        : `<w:bookmarkEnd w:id="${id}"/>`;
    const seq = `<w:fldSimple w:instr=" SEQ Figure "/>`;
    test.each([
      [
        "a bookmark after a break with no text",
        `${BR}${bookmark(5, "Start")}${bookmark(5, "End")}`,
        "BSE",
        "BSE",
      ],
      ["an empty field after a break with no text", `${BR}${seq}`, "BF", "BFT"],
      ["an empty field before a break", `${seq}${BR}${run("x")}`, "FBT", "FTB"],
      [
        "a bookmark opening before a leading break",
        `${bookmark(5, "Start")}${BR}${run("Chapter")}${bookmark(5, "End")}`,
        "SBTE",
        "SBTET",
      ],
      [
        "a bookmark opening after a leading break",
        `${BR}${bookmark(5, "Start")}${run("Chapter")}${bookmark(5, "End")}`,
        "BSTE",
        "BSTET",
      ],
    ])("%s keeps its place", async (_, content, untouched, edited) => {
      const [story, part] = STORY[where];
      let bytes = docx(
        where,
        `${p("33333333", run("prev"))}${p("44444444", content)}`
      );
      const seen: string[] = [];
      for (let publication = 0; publication < 3; publication += 1) {
        const session = await open(bytes);
        session.insertText(
          { story: "body", paraId: "22222222", offset: 0 },
          "z"
        );
        bytes = await publish(bytes, session);
        session.destroy();
        seen.push(order(bytes, part, "44444444"));
      }
      expect(seen).toEqual(Array(3).fill(untouched));
      // Typing at the paragraph's end moves a break after text to its end, as for any text.
      const session = await open(bytes);
      const end = session
        .paragraphSpans(story)
        .find(({ paraId }) => paraId === "44444444")!.length;
      if (end > 1)
        session.insertText({ story, paraId: "44444444", offset: end }, "Q");
      const out = await publish(bytes, session);
      session.destroy();
      expect(order(out, part, "44444444")).toBe(end > 1 ? edited : untouched);
    });
  }
);

describe.each(["body", "cell"] as Where[])(
  "a field Accept All resolves in a paragraph a page break's slot opens, in %s",
  (where) => {
    const link = (xml: string) =>
      `<w:hyperlink w:anchor="target">${xml}</w:hyperlink>`;
    // Accepting the first field's insertion uncovers a link it then projects.
    const fields = `${field(
      `${run("20")}${tracked("ins", link(run("26")))}`
    )}${field(link(run("30")))}`;
    test.each([
      [
        "a break closing the paragraph before",
        `${p("33333333", `${run("prev")}${BR}`)}${p("44444444", fields)}`,
      ],
      [
        "a break leading the paragraph",
        `${p("33333333", run("prev"))}${p("44444444", `${BR}${fields}`)}`,
      ],
      [
        "a table before the paragraph",
        `${table(p("66666666", run("in")))}${p("44444444", fields)}`,
      ],
    ])(
      "after %s, numbers each field's projection as its export seeds it",
      async (_, xml) => {
        const [story, part] = STORY[where];
        // Units with the projection ids that pair a field with its children.
        const projections = (session: YrsSession) =>
          session
            .storySegments(story)
            .map((segment) => {
              const child = segment.attributes.fieldResult as
                | { id?: number }
                | undefined;
              const owner =
                segment.kind === "embed"
                  ? segment.payload.resultProjection
                  : undefined;
              return `${units_(segment)}${child ? `<${child.id}` : ""}${
                owner ? `>${(owner as { id?: number }).id}` : ""
              }`;
            })
            .join("");
        const bytes = docx(where, xml);
        const session = await open(bytes);
        session.acceptChange({ all: true });
        const resolved = projections(session);
        const out = await publish(bytes, session);
        session.destroy();
        const reopened = await open(out);
        expect(projections(reopened)).toBe(resolved);
        reopened.destroy();
        expect(plain(out, part)).toContain("2026");
        expect(new TextDecoder().decode(unzipContainer(out)[part])).toMatch(
          /<w:fldChar w:fldCharType="separate"\/><\/w:r><w:hyperlink w:anchor="target"><w:r><w:t[^>]*>30</
        );
      }
    );
  }
);

describe.each(EVERY)(
  "removing an empty paragraph that holds a bookmark, in %s",
  (where) => {
    const bookmark = `<w:bookmarkStart w:id="5" w:name="Target"/><w:bookmarkEnd w:id="5"/>`;
    const holder = `<w:p w14:paraId="55555555">${bookmark}</w:p>`;
    const slots = [
      ["a page break", `${p("44444444", `${BR}${run("abc")}`)}`, "SEBT"],
      ["a column break", `${p("44444444", `${COL}${run("abc")}`)}`, "SEBT"],
      [
        "a table",
        `${table(p("66666666", run("in")))}${p("44444444", run("abc"))}`,
        "SET",
      ],
    ] as const;
    test.each(
      slots.flatMap(([name, slot, order]) =>
        (["forward", "backward"] as const).map(
          (direction) => [name, direction, slot, order] as const
        )
      )
    )(
      "before %s moves it to the paragraph that stays: %s",
      async (_, direction, slot, expected) => {
        const [story, part] = STORY[where];
        const bytes = docx(
          where,
          `${p("33333333", run("prev"))}${holder}${slot}`
        );
        const session = await open(bytes);
        if (direction === "forward")
          session.deleteAt({ story, paraId: "55555555", offset: 0 }, "forward");
        else
          session.deleteAt(
            { story, paraId: "44444444", offset: 0 },
            "backward"
          );
        let out = await publish(bytes, session);
        session.destroy();
        for (let publication = 0; publication < 2; publication += 1) {
          expect(
            new TextDecoder()
              .decode(unzipContainer(out)[part])
              .split('w:name="Target"').length - 1
          ).toBe(1);
          expect(order(out, part, "44444444")).toBe(expected);
          const reopened = await open(out);
          reopened.insertText(
            { story: "body", paraId: "22222222", offset: 0 },
            "z"
          );
          out = await publish(out, reopened);
          reopened.destroy();
        }
      }
    );

    test("Enter at a break's slot, then Delete, leaves one copy of the paragraph's bookmark", async () => {
      const [story, part] = STORY[where];
      const bytes = docx(
        where,
        `${p("33333333", run("prev"))}${p(
          "44444444",
          `${bookmark}${BR}${run("abc")}`
        )}`
      );
      const session = await open(bytes);
      const { firstParaId } = session.splitParagraph({
        story,
        paraId: "44444444",
        offset: 0,
      });
      session.deleteAt({ story, paraId: firstParaId, offset: 0 }, "forward");
      const out = await publish(bytes, session);
      session.destroy();
      expect(
        new TextDecoder()
          .decode(unzipContainer(out)[part])
          .split('w:name="Target"').length - 1
      ).toBe(1);
      expect(order(out, part, "44444444")).toBe("SEBT");
    });
  }
);

describe.each(EVERY)(
  "a column break after a bookmark in a paragraph without text, in %s",
  (where) => {
    const bookmark = `<w:bookmarkStart w:id="5" w:name="_Toc5"/><w:bookmarkEnd w:id="5"/>`;
    test.each([
      [
        "after another paragraph",
        (content: string) => `${p("33333333", run("prev"))}${content}`,
        "SEB",
      ],
      ["opening the story", (content: string) => content, "SEB"],
      [
        "with a page break before it",
        (content: string) => `${p("33333333", run("prev"))}${content}`,
        "SEBB",
      ],
    ] as const)("%s stays in its paragraph", async (name, around, expected) => {
      const [story, part] = STORY[where];
      const breaks =
        name === "with a page break before it" ? `${BR}${COL}` : COL;
      let bytes = docx(where, around(p("44444444", `${bookmark}${breaks}`)));
      const session = await open(bytes);
      const seeded = units(session, story);
      session.destroy();
      for (let publication = 0; publication < 2; publication += 1) {
        const reopened = await open(bytes);
        reopened.insertText(
          { story: "body", paraId: "22222222", offset: 0 },
          "z"
        );
        bytes = await publish(bytes, reopened);
        reopened.destroy();
        expect(order(bytes, part, "44444444")).toBe(expected);
        const again = await open(bytes);
        // The body's tail paragraph takes the typing.
        expect(units(again, story).replace(/z*tail¶$/, "")).toBe(
          seeded.replace(/tail¶$/, "")
        );
        again.destroy();
      }
    });
  }
);

describe.each(EVERY)(
  "a paragraph without text whose breaks end in a column break, in %s",
  (where) => {
    test.each([
      ["page then column", `${BR}${COL}`, "[PB][CB]"],
      ["two pages then column", `${BR}${BR}${COL}`, "[PB][PB][CB]"],
      [
        "tracked page then column",
        `${tracked("ins", BR)}${COL}`,
        "+{[PB]}[CB]",
      ],
      ["tracked column alone", tracked("ins", COL), "+{[CB]}"],
    ] as const)(
      "keeps its breaks in one paragraph across publications: %s",
      async (_, breaks, shown) => {
        const [story, part] = STORY[where];
        for (const [around, expected] of [
          [
            (own: string) => `${own}${p("45454545", run("abc"))}`,
            `${shown}¶abc¶`,
          ],
          [
            (own: string) => `${p("33333333", run("prev"))}${own}`,
            `prev¶${shown}¶`,
          ],
        ] as const) {
          let bytes = docx(where, around(p("44444444", breaks)));
          const seen: string[] = [];
          for (let publication = 0; publication < 3; publication += 1) {
            const session = await open(bytes);
            seen.push(units(session, story).replace(/z*tail¶$/, ""));
            session.insertText(
              { story: "body", paraId: "22222222", offset: 0 },
              "z"
            );
            bytes = await publish(bytes, session);
            session.destroy();
            expect(view(bytes, part)).toContain(expected);
          }
          expect(seen).toEqual(Array(3).fill(seen[0]));
        }
      }
    );

    test("keeps one paragraph once a story-opening heading's text after them is deleted", async () => {
      const [story, part] = STORY[where];
      let bytes = docx(
        where,
        `${p("44444444", `${BR}${COL}${run("abc")}`)}${p(
          "45454545",
          run("next")
        )}`
      );
      let session = await open(bytes);
      session.deleteRange({
        story,
        start: { paraId: "44444444", offset: 2 },
        end: { paraId: "44444444", offset: 5 },
      });
      const edited = units(session, story)
        .replace(/\^/g, "")
        .replace(/z*tail¶$/, "");
      for (let publication = 0; publication < 3; publication += 1) {
        bytes = await publish(bytes, session);
        session.destroy();
        expect(view(bytes, part)).toContain("[PB][CB]¶next¶");
        session = await open(bytes);
        expect(units(session, story).replace(/z*tail¶$/, "")).toBe(edited);
        session.insertText(
          { story: "body", paraId: "22222222", offset: 0 },
          "z"
        );
      }
      session.destroy();
    });
  }
);

describe.each(EVERY)(
  "handing an empty paragraph's bookmark to a text-less paragraph that ends in a column break, in %s",
  (where) => {
    const holder = `<w:p w14:paraId="55555555"><w:bookmarkStart w:id="5" w:name="Target"/><w:bookmarkEnd w:id="5"/></w:p>`;
    const shapes = [
      ["a column break", COL],
      ["page then column", `${BR}${COL}`],
      ["two pages then column", `${BR}${BR}${COL}`],
      ["tracked page then column", `${tracked("ins", BR)}${COL}`],
      ["column, page, column", `${COL}${BR}${COL}`],
    ] as const;
    test.each(
      shapes.flatMap(([name, breaks]) =>
        (["after a paragraph", "at the story's start"] as const).flatMap(
          (place) =>
            (["forward", "backward"] as const).map(
              (direction) => [name, place, direction, breaks] as const
            )
        )
      )
    )("keeps the breaks: %s, %s, %s", async (_, place, direction, breaks) => {
      const [story, part] = STORY[where];
      const before =
        place === "after a paragraph" ? p("33333333", run("prev")) : "";
      const bytes = docx(
        where,
        `${before}${holder}${p("44444444", breaks)}${p(
          "45454545",
          run("next")
        )}`
      );
      const session = await open(bytes);
      if (direction === "forward")
        session.deleteAt({ story, paraId: "55555555", offset: 0 }, "forward");
      else
        session.deleteAt({ story, paraId: "44444444", offset: 0 }, "backward");
      const edited = units(session, story).replace(/\^/g, "");
      let out = await publish(bytes, session);
      session.destroy();
      for (let publication = 0; publication < 2; publication += 1) {
        const xml = new TextDecoder().decode(unzipContainer(out)[part]);
        expect([breaksIn(xml, "page"), breaksIn(xml, "column")]).toEqual([
          breaks.split('w:type="page"').length - 1,
          breaks.split('w:type="column"').length - 1,
        ]);
        expect(xml.split('w:name="Target"').length - 1).toBe(1);
        const reopened = await open(out);
        expect(
          units(reopened, story)
            .replace(/\^/g, "")
            .replace(/z*tail¶$/, "")
        ).toBe(edited.replace(/tail¶$/, ""));
        reopened.insertText(
          { story: "body", paraId: "22222222", offset: 0 },
          "z"
        );
        out = await publish(out, reopened);
        reopened.destroy();
      }
    });
  }
);

const breaksIn = (xml: string, type: string) =>
  xml.split(`<w:br w:type="${type}"/>`).length - 1;

describe.each(EVERY)(
  "a comment boundary beside a text-less page and column break paragraph, in %s",
  (where) => {
    test.each([
      [
        "ending after the breaks",
        (prev: string) =>
          `${prev.replace("PREV", `${S(1)}${run("prev")}`)}${p(
            "44444444",
            `${BR}${COL}${E(1)}`
          )}${p("45454545", `${run("abc")}${ref(1)}`)}`,
        "prev¶##",
      ],
      [
        "opening after the breaks",
        (prev: string) =>
          `${prev.replace("PREV", run("prev"))}${p(
            "44444444",
            `${BR}${COL}${S(1)}`
          )}${p("45454545", `${run("abc")}${E(1)}${ref(1)}`)}`,
        "¶abc",
      ],
      [
        "opening after the breaks at the story's start",
        () =>
          `${p("44444444", `${BR}${COL}${S(1)}`)}${p(
            "45454545",
            `${run("abc")}${E(1)}${ref(1)}`
          )}`,
        "¶abc",
      ],
    ] as const)(
      "%s keeps its range across publications",
      async (_, xml, range) => {
        const [story] = STORY[where];
        let bytes = docx(where, xml(p("33333333", "PREV")));
        const seen: string[][] = [];
        for (let publication = 0; publication < 3; publication += 1) {
          const session = await open(bytes);
          seen.push(covered(session, story));
          session.insertText(
            { story: "body", paraId: "22222222", offset: 0 },
            "z"
          );
          bytes = await publish(bytes, session);
          session.destroy();
        }
        expect(seen).toEqual(Array(3).fill([range]));
      }
    );
  }
);

describe.each(EVERY)(
  "a story-opening bookmark handed to a text-less column-break paragraph, in %s",
  (where) => {
    const holder = `<w:p w14:paraId="55555555"><w:bookmarkStart w:id="5" w:name="Target"/><w:bookmarkEnd w:id="5"/></w:p>`;
    const shapes = [
      ["a column break", COL],
      ["page then column", `${BR}${COL}`],
      ["two pages then column", `${BR}${BR}${COL}`],
      ["tracked page then column", `${tracked("ins", BR)}${COL}`],
      ["column, page, column", `${COL}${BR}${COL}`],
    ] as const;
    test.each(
      shapes.flatMap(([name, breaks]) =>
        (["forward", "backward"] as const).flatMap((direction) =>
          (["end", "start"] as const).map(
            (at) => [name, direction, at, breaks] as const
          )
        )
      )
    )(
      "saves the same with or without a publication before the typing: %s, %s, typing at the %s",
      async (_, direction, at, breaks) => {
        const [story, part] = STORY[where];
        const bytes = docx(
          where,
          `${holder}${p("44444444", breaks)}${p("45454545", run("next"))}`
        );
        const handOff = (session: YrsSession) =>
          direction === "forward"
            ? void session.deleteAt(
                { story, paraId: "55555555", offset: 0 },
                "forward"
              )
            : void session.deleteAt(
                { story, paraId: "44444444", offset: 0 },
                "backward"
              );
        const type = (session: YrsSession) => {
          const end = session
            .paragraphSpans(story)
            .find(({ paraId }) => paraId === "44444444")!.length;
          session.insertText(
            { story, paraId: "44444444", offset: at === "end" ? end : 0 },
            "Q"
          );
        };
        const session = await open(bytes);
        handOff(session);
        type(session);
        const direct = order(await publish(bytes, session), part, "44444444");
        session.destroy();
        const { next } = await rebase(bytes, story, handOff, type);
        expect(order(next, part, "44444444")).toBe(direct);
        expect(direct).toMatch(/^B+SET$/);
      }
    );
  }
);

describe.each(EVERY)(
  "an editor comment over a heading a page break opens, its text removed after a capture, in %s",
  (where) => {
    const [story, part] = STORY[where];
    const heading = p("44444444", `${BR}${run("Heading")}`);
    /** Each place of the heading: its package, and the story's text a save keeps. */
    const layouts = {
      "after a paragraph": {
        bytes: docx(where, `${p("33333333", run("prev"))}${heading}`),
        saved: /^(T\[)?prev¶\[PB\]¶\]? #$/,
      },
      // The reference mark is content for the break to lead: no break-only paragraph.
      "at the story's start": {
        bytes: docx(where, `${heading}${p("45454545", run("next"))}`),
        saved: /^(T\[)?\[PB\]¶next¶\]? #$/,
      },
    };
    type Layout = keyof typeof layouts;
    /** An editor comment from the break to `end` in the heading's paragraph. */
    const comment = (end: number) => (session: YrsSession) =>
      void session.addComment(
        [
          {
            story,
            start: { paraId: "44444444", offset: 0 },
            end: { paraId: "44444444", offset: end },
          },
        ],
        "R",
        "2026-09-29T00:00:00Z",
        [
          {
            type: "paragraph",
            content: [{ type: "run", content: [{ type: "text", text: "c" }] }],
          },
        ]
      );
    const comments = [
      ["over the break and the heading", comment(8)],
      ["over the break", comment(1)],
    ] as const;
    /** The heading's offset in its paragraph: after the break, and after a reference mark a publication put there. */
    const headingAt = (session: YrsSession) => {
      const units = session
        .storySegments(story)
        .flatMap((segment) =>
          segment.kind === "text" ? [...segment.text] : ["#"]
        )
        .join("");
      return units.indexOf("Heading") - (units.startsWith("prev") ? 5 : 0);
    };
    const removals = [
      [
        "a range delete",
        (session: YrsSession) => {
          const at = headingAt(session);
          session.deleteRange({
            story,
            start: { paraId: "44444444", offset: at },
            end: { paraId: "44444444", offset: at + 7 },
          });
        },
      ],
      [
        "Backspace",
        (session: YrsSession) => {
          const at = headingAt(session);
          for (let offset = at + 7; offset > at; offset -= 1)
            session.deleteAt({ story, paraId: "44444444", offset }, "backward");
        },
      ],
    ] as const;
    /** The story's breaks and text and what the comment covers, in this publication and two more. */
    const republished = async (published: Uint8Array) => {
      const seen: string[] = [];
      for (let publication = 0; publication < 3; publication += 1) {
        const session = await open(published);
        seen.push(
          `${view(published, part).replace(/z*tail¶/, "")} ${covered(
            session,
            story
          )}`
        );
        session.insertText(
          { story: "body", paraId: "22222222", offset: 0 },
          "z"
        );
        published = await publish(published, session);
        session.destroy();
      }
      return seen;
    };
    /** The comment keeps covering the break, which still opens the heading's paragraph. */
    const keeps = (layout: Layout, seen: string[]) => {
      expect(seen[0]).toMatch(layouts[layout].saved);
      expect(new Set(seen).size).toBe(1);
    };

    test.each([
      ...comments.flatMap(([name, cover]) =>
        removals.map(
          ([way, remove]) =>
            ["after a paragraph", name, way, cover, remove] as const
        )
      ),
      ...comments.map(
        ([name, cover]) =>
          [
            "at the story's start",
            name,
            removals[0][0],
            cover,
            removals[0][1],
          ] as const
      ),
    ])(
      "%s, a comment %s, the text removed by %s: the direct, publish-then-edit and rebased saves keep the comment over the break",
      async (layout, _, __, cover, remove) => {
        const { bytes } = layouts[layout];
        const session = await open(bytes);
        cover(session);
        const published = await publish(bytes, session);
        remove(session);
        const direct = await republished(await publish(bytes, session));
        session.destroy();
        const reopened = await open(published);
        remove(reopened);
        const sequential = await republished(
          await publish(published, reopened)
        );
        reopened.destroy();
        const { next } = await rebase(bytes, story, cover, remove);
        keeps(layout, direct);
        expect(sequential).toEqual(direct);
        expect(await republished(next)).toEqual(direct);
      }
    );

    test.each(
      (Object.keys(layouts) as Layout[]).flatMap((layout) =>
        comments.map(([name, cover]) => [layout, name, cover] as const)
      )
    )(
      "%s, a comment %s, the text removed by the agent's replace_text: the direct and rebased saves keep the comment over the break",
      async (layout, _, cover) => {
        const { bytes } = layouts[layout];
        const session = await open(bytes);
        cover(session);
        const captured = checkpoint(bytes, session.encodeState());
        session.destroy();
        const exported = await exportOffice(bytes, captured, fixed);
        const target = (await inspectOffice(bytes, captured)).find(({ id }) =>
          id.endsWith(":paragraph:44444444")
        )!;
        const { state: edited } = await applyOfficeCommands(bytes, captured, [
          {
            type: "replace_text",
            targetId: target.id,
            expectedText: "Heading",
            text: "",
          },
        ]);
        const latest = { ...captured, state: edited };
        const direct = await republished(
          await exportOffice(bytes, latest, fixed)
        );
        const { state } = await rebaseOffice(bytes, captured, latest, exported);
        const seeded = await seedOffice("docx", exported);
        const next = await exportOffice(exported, { ...seeded, state }, fixed);
        keeps(layout, direct);
        expect(await republished(next)).toEqual(direct);
      }
    );
  }
);

/** Paragraph XML as text, [PB] breaks, <c c> comment range marks, R reference marks and ¶ per paragraph. */
const marked = (xml: string) =>
  [
    ...xml.matchAll(
      /<(\/?)w:(p|br|t|commentRangeStart|commentRangeEnd|commentReference)\b([^>]*?)\/?>([^<]*)/g
    ),
  ]
    .map(([, close, tag, attrs, text]) =>
      tag === "t"
        ? close
          ? ""
          : text
        : tag === "br"
        ? /"page"/.test(attrs!)
          ? "[PB]"
          : ""
        : tag === "p"
        ? close
          ? "¶"
          : ""
        : tag === "commentRangeStart"
        ? "<c"
        : tag === "commentRangeEnd"
        ? "c>"
        : "R"
    )
    .join("");

describe.each(EVERY)(
  "a paragraph holding a page break and a comment's reference mark saves unchanged, in %s",
  (where) => {
    test.each([
      [
        "a comment over the break",
        `${p("33333333", run("prev"))}${p(
          "44444444",
          `${S(1)}${BR}${E(1)}${ref(1)}`
        )}`,
      ],
      [
        "a comment over the paragraph before",
        `${p("33333333", `${S(1)}${run("prev")}${E(1)}`)}${p(
          "44444444",
          `${BR}${ref(1)}`
        )}`,
      ],
      [
        "a comment over the break at the story's start",
        `${p("44444444", `${S(1)}${BR}${E(1)}${ref(1)}`)}${p(
          "45454545",
          run("next")
        )}`,
      ],
      [
        "a comment from the break into the next paragraph",
        `${p("33333333", run("prev"))}${p(
          "44444444",
          `${S(1)}${BR}${ref(1)}`
        )}${p("45454545", `${run("nx")}${E(1)}`)}`,
      ],
    ])("%s", async (_, xml) => {
      const [story, part] = STORY[where];
      // An edit after them has the story saved again.
      let bytes = docx(where, `${xml}${p("5A5A5A5A", run("edit"))}`);
      const text = (bytes: Uint8Array) => {
        const saved = new TextDecoder().decode(unzipContainer(bytes)[part]);
        const start = saved.search(/<w:p w14:paraId="(33333333|44444444)"/);
        return saved.slice(start, saved.indexOf('<w:p w14:paraId="5A5A5A5A"'));
      };
      const source = text(bytes);
      const seen: string[] = [];
      for (let publication = 0; publication < 3; publication += 1) {
        const session = await open(bytes);
        session.insertText({ story, paraId: "5A5A5A5A", offset: 0 }, "z");
        bytes = await publish(bytes, session);
        session.destroy();
        seen.push(text(bytes));
      }
      // The save writes the source's paragraphs in its own form, then the same bytes.
      expect(seen.map(marked)).toEqual(Array(3).fill(marked(source)));
      expect(new Set(seen).size).toBe(1);
    });
  }
);

/**
 * The story's units (without their breaks' flags unless `flags`) and what
 * its comments cover: after `edit`, then after each of three publications.
 */
async function kept(
  bytes: Uint8Array,
  story: string,
  edit: (session: YrsSession) => void,
  flags = true
) {
  let session = await open(bytes);
  edit(session);
  const state = () => {
    const seen = units(session, story);
    return `${flags ? seen : seen.replaceAll("^", "")} ${covered(
      session,
      story
    )}`;
  };
  const seen = [state()];
  for (let publication = 0; publication < 3; publication += 1) {
    bytes = await publish(bytes, session);
    session.destroy();
    session = await open(bytes);
    seen.push(state());
  }
  session.destroy();
  return seen;
}

/** Types `text` at the end of paragraph `paraId`. */
const typeAtEnd =
  (story: string, paraId: string, text: string) => (session: YrsSession) =>
    session.insertText(
      {
        story,
        paraId,
        offset: session
          .paragraphSpans(story)
          .find((span) => span.paraId === paraId)!.length,
      },
      text
    );

/**
 * The comment's breaks and range stay as they are over three publications,
 * and each edit after a capture lands as it saves directly.
 */
async function keepsBreaks(
  bytes: Uint8Array,
  story: string,
  cover: (session: YrsSession) => void,
  edits: Array<(session: YrsSession) => void>,
  flags = true
) {
  const commented = await kept(bytes, story, cover, flags);
  expect(commented).toEqual(Array(4).fill(commented[0]));
  for (const edit of edits) {
    const direct = await kept(
      bytes,
      story,
      (session) => {
        cover(session);
        edit(session);
      },
      flags
    );
    const { next } = await rebase(bytes, story, cover, edit);
    expect((await kept(next, story, () => {}, flags)).slice(0, 3)).toEqual(
      direct.slice(1)
    );
  }
}

/** Removes every comment. */
const removeComments = (session: YrsSession) => {
  for (const { id } of session.listComments())
    session.applyRawOps("body", [{ op: "removeComment", id }]);
};

describe.each(NINE)(
  "an editor comment ending at a break a paragraph opens with keeps every break, in %s",
  (where) => {
    const [story] = STORY[where];
    const heading = (breaks: string) =>
      `${p("33333333", run("prev"))}${p(
        "44444444",
        `${breaks}${run("Heading")}`
      )}`;
    test.each([
      [
        "from the paragraph before to the page break",
        heading(BR),
        ["33333333", 0],
        ["44444444", 0],
      ],
      [
        "from inside the paragraph before to the column break",
        heading(COL),
        ["33333333", 2],
        ["44444444", 0],
      ],
      [
        "from the paragraph before to two breaks",
        heading(`${BR}${COL}`),
        ["33333333", 0],
        ["44444444", 0],
      ],
      [
        "over the page break of two",
        heading(`${BR}${COL}`),
        ["44444444", 0],
        ["44444444", 1],
      ],
      [
        "over the column break of two at the story's start",
        `${p("44444444", `${COL}${BR}${run("Heading")}`)}${p(
          "45454545",
          run("next")
        )}`,
        ["44444444", 0],
        ["44444444", 1],
      ],
    ] as const)("%s", async (_, xml, start, end) => {
      await keepsBreaks(
        docx(where, xml),
        story,
        editorComment(story, [...start], [...end]),
        [typeAtEnd(story, "44444444", "Q")]
      );
    });

    test("from the paragraph before to the page break, the heading's text removed", async () => {
      // The reference mark the save writes after the break is content for it
      // to lead, so it stays in the emptied heading's paragraph.
      await keepsBreaks(
        docx(where, heading(BR)),
        story,
        (session) => {
          editorComment(story, ["33333333", 0], ["44444444", 0])(session);
          session.deleteRange({
            story,
            start: { paraId: "44444444", offset: 1 },
            end: { paraId: "44444444", offset: 8 },
          });
        },
        [typeAtEnd(story, "33333333", "Q")]
      );
    });
  }
);

describe.each(NINE)(
  "a comment over a paragraph holding only a column break keeps the break its own, in %s",
  (where) => {
    const [story] = STORY[where];
    const only = (breaks: string) =>
      `${p("33333333", run("prev"))}${p("44444444", breaks)}${p(
        "45454545",
        run("next")
      )}`;
    test.each([
      [
        "an editor comment over the break",
        only(COL),
        editorComment(story, ["44444444", 0], ["44444444", 1]),
      ],
      [
        "an editor comment over a page and a column break",
        only(`${BR}${COL}`),
        editorComment(story, ["44444444", 0], ["44444444", 2]),
      ],
      [
        "an editor comment over the column break after a page break",
        only(`${BR}${COL}`),
        editorComment(story, ["44444444", 1], ["44444444", 2]),
      ],
      [
        "a comment of the file with no reference mark",
        only(`${S(1)}${COL}${E(1)}`),
        () => {},
      ],
    ] as const)("%s", async (_, xml, cover) => {
      // The breaks read back as leading the comment's reference mark, as
      // they did before the reference counted as content.
      await keepsBreaks(
        docx(where, xml),
        story,
        cover,
        [typeAtEnd(story, "45454545", "Q")],
        false
      );
    });
  }
);

describe.each(NINE)(
  "a comment ending at the text a break opens keeps covering the break, in %s",
  (where) => {
    const [story] = STORY[where];
    const opens = (breaks: string, text = run("Heading")) =>
      `${p("33333333", run("prev"))}${p("44444444", `${breaks}${text}`)}${p(
        "45454545",
        run("next")
      )}`;
    const upTo = (offset: number) =>
      editorComment(story, ["33333333", 0], ["44444444", offset]);
    const type = typeAtEnd(story, "44444444", "Q");
    // A column break alone in its paragraph reads back as leading the
    // comment's reference mark, as it did before (see the tests above), and
    // text typed after it following a capture refuses the rebase (below).
    test.each([
      ["a page break", opens(BR), upTo(1), [type, removeComments], true],
      ["a column break", opens(COL), upTo(1), [type, removeComments], true],
      [
        "a page and a column break",
        opens(`${BR}${COL}`),
        upTo(2),
        [type, removeComments],
        true,
      ],
      [
        "a column break alone in its paragraph",
        opens(COL, ""),
        upTo(1),
        [removeComments],
        false,
      ],
      [
        "a page break, a comment of the file with no reference mark",
        `${p("33333333", `${S(1)}${run("prev")}`)}${p(
          "44444444",
          `${BR}${E(1)}${run("Heading")}`
        )}${p("45454545", run("next"))}`,
        () => {},
        [type, removeComments],
        true,
      ],
    ] as const)("%s", async (_, xml, cover, edits, flags) => {
      await keepsBreaks(docx(where, xml), story, cover, [...edits], flags);
    });
  }
);

describe.each(NINE)(
  "text typed after a capture into a paragraph holding only breaks and a comment's end, in %s",
  (where) => {
    const [story] = STORY[where];
    const only = (breaks: string, prev = run("prev")) =>
      `${p("33333333", prev)}${p("44444444", breaks)}${p(
        "45454545",
        run("next")
      )}`;
    const upTo = (start: [string, number], end: number) =>
      editorComment(story, start, ["44444444", end]);
    const type = typeAtEnd(story, "44444444", "Q");
    // The capture's save writes the comment's reference after the breaks,
    // so the seed reads them as leading it; the latest state does not.
    test.each([
      [
        "a column break, a comment ending before it",
        only(COL),
        upTo(["33333333", 0], 0),
      ],
      [
        "a column break, two comments ending before it",
        only(COL),
        (session: YrsSession) => {
          upTo(["33333333", 2], 0)(session);
          upTo(["33333333", 0], 0)(session);
        },
      ],
      [
        "a column break, a comment over it",
        only(COL),
        upTo(["44444444", 0], 1),
      ],
      [
        "a column break, a comment from the paragraph before over it",
        only(COL),
        upTo(["33333333", 0], 1),
      ],
      [
        "a page and a column break, a comment ending before them",
        only(`${BR}${COL}`),
        upTo(["33333333", 0], 0),
      ],
      [
        "a page and a column break, a comment over them",
        only(`${BR}${COL}`),
        upTo(["44444444", 0], 2),
      ],
      [
        "a column break, a comment of the file with no reference mark ending before it",
        only(`${E(1)}${COL}`, `${S(1)}${run("prev")}`),
        () => {},
      ],
    ] as const)("%s: the rebase refuses", async (_, xml, cover) => {
      await expect(
        rebase(docx(where, xml), story, cover, type)
      ).rejects.toThrow("Office rebase:");
    });

    test.each([
      ["a column break", COL],
      ["a page and a column break", `${BR}${COL}`],
    ])(
      "%s with no comment: the typing lands as saved directly",
      async (_, breaks) => {
        await keepsBreaks(docx(where, only(breaks)), story, () => {}, [type]);
      }
    );
  }
);
