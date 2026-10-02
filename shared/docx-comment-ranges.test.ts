import { beforeAll, expect, test } from "bun:test";
import { createHash } from "node:crypto";
import { createYrsSession, type YrsSession } from "../packages/docx/src/yrs";
import { rezipContainer, unzipContainer } from "../packages/docx/src/wasm/opc";
import { exportOffice, rebaseOffice, seedOffice } from "./office-checkpoint";
import { addComment as matrixComment, blocks, bm, CB as matrixCB, covered as matrixCovered, docx as matrixDocx, len as matrixLen, none, open as matrixOpen, orders, p as matrixParagraph, PB, prime, publish as matrixPublish, run as matrixRun, runRow, sig, textStart as matrixTextStart, type Edit as MatrixEdit, type Where as MatrixWhere } from "./matrix/lib";
import { parseDocx } from "../packages/docx/src/docx";
import { repackDocx } from "../packages/docx/src/docx/rezip";
import { injectReplyRangeMarkers } from "../packages/docx/src/docx/injectReplyRangeMarkers";
import { yrsToDocument } from "../packages/docx/src/yrs/yrsToDocument";

const fixed = { seed: "0".repeat(64), now: "2026-09-29T00:00:00.000Z" };
const W =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const OFFICE = "application/vnd.openxmlformats-officedocument.wordprocessingml";
const REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PNG =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const p = (id: string, xml: string) => `<w:p w14:paraId="${id}">${xml}</w:p>`;
const ref = (id: number) =>
  `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`;
const S = (id: number) => `<w:commentRangeStart w:id="${id}"/>`;
const E = (id: number) => `<w:commentRangeEnd w:id="${id}"/>`;
const table = (...cells: string[]) =>
  `<w:tbl><w:tr>${cells
    .map((cell) => `<w:tc>${cell}</w:tc>`)
    .join("")}</w:tr></w:tbl>`;
const field = (result: string) =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DATE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${result}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
const range = (id: number, xml: string) =>
  `<w:commentRangeStart w:id="${id}"/>${xml}<w:commentRangeEnd w:id="${id}"/>${ref(
    id
  )}`;
const image = `<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="95250" cy="95250"/><wp:docPr id="1" name="Picture 1"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="95250" cy="95250"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

/**
 * A package whose body is `body`, with a Word comment for each id in
 * `comments` and, when given, a default header (story `hf:rId20`).
 */
function docx(body: string, comments = [1], header = ""): Uint8Array {
  const headerType = `<Override PartName="/word/header1.xml" ContentType="${OFFICE}.header+xml"/>`;
  const headerRel = `<Relationship Id="rId20" Type="${REL}/header" Target="header1.xml"/>`;
  const headerRef = `<w:headerReference w:type="default" r:id="rId20"/>`;
  const parts: Record<string, string | Uint8Array> = {
    "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="${OFFICE}.comments+xml"/>${
      header && headerType
    }</Types>`,
    "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
    "word/_rels/document.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="${REL}/comments" Target="comments.xml"/><Relationship Id="rIdImage" Type="${REL}/image" Target="media/image1.png"/>${
      header && headerRel
    }</Relationships>`,
    "word/media/image1.png": Buffer.from(PNG, "base64"),
    ...(header ? { "word/header1.xml": `<w:hdr ${W}>${header}</w:hdr>` } : {}),
    "word/document.xml": `<w:document ${W}><w:body>${body}<w:sectPr>${
      header && headerRef
    }<w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
    "word/comments.xml": `<w:comments ${W}>${comments
      .map(
        (id) =>
          `<w:comment w:id="${id}" w:author="Reviewer" w:date="2026-09-01T00:00:00Z"><w:p><w:r><w:t>c${id}</w:t></w:r></w:p></w:comment>`
      )
      .join("")}</w:comments>`,
  };
  return rezipContainer(
    Object.fromEntries(
      Object.entries(parts).map(([path, part]) => [
        path,
        typeof part === "string" ? new TextEncoder().encode(part) : part,
      ])
    )
  );
}

// Seeding loads the engines the sessions below use directly.
beforeAll(() => seedOffice("docx", docx(p("11111111", run("x")))));
let clientId = 9400;
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

type Span = [story: string, paraId: string, start: number, end: number];
/** Adds a comment named `name` over `span`, ending in paragraph `endPara` when given. */
function comment(
  session: YrsSession,
  name: string,
  [story, paraId, start, end]: Span,
  endPara = paraId
): string {
  return session.addComment(
    [
      {
        story,
        start: { paraId, offset: start },
        end: { paraId: endPara, offset: end },
      },
    ],
    "Reviewer",
    "2026-09-29T00:00:00Z",
    [
      {
        type: "paragraph",
        content: [{ type: "run", content: [{ type: "text", text: name }] }],
      },
    ]
  ).commentId;
}

/**
 * What each comment covers, by the name in its body: the story's units under
 * its anchors, with the reference marks an export adds left out.
 */
function covered(session: YrsSession): Record<string, string> {
  const result: Record<string, string> = {};
  for (const { id, body } of session.listComments()) {
    const name = JSON.stringify(body).match(/"text":"([^"]+)"/)![1];
    result[name] = session
      .resolveComment(id)
      .map(({ story, start, end }) =>
        session
          .storySegments(story)
          .flatMap((segment) =>
            segment.kind === "text"
              ? [...segment.text]
              : segment.kind === "pilcrow"
              ? ["¶"]
              : [
                  segment.payload.modelKind === "commentReference"
                    ? ""
                    : `[${segment.embedKind}]`,
                ]
          )
          .slice(start, end)
          .join("")
      )
      .join("|");
  }
  return result;
}

/** Publishes `bytes` with `edit` and republishes twice; returns what the comments covered each time. */
async function publications(
  bytes: Uint8Array,
  edit: (session: YrsSession) => void
) {
  let session = await open(bytes);
  edit(session);
  const seen = [covered(session)];
  for (let publication = 0; publication < 2; publication += 1) {
    bytes = await publish(bytes, session);
    session.destroy();
    session = await open(bytes);
    seen.push(covered(session));
    session.insertText({ story: "body", paraId: "22222222", offset: 0 }, "x");
  }
  session.destroy();
  return { seen, bytes };
}

const tail = p("22222222", run("tail"));

test("an emptied comment has the same publication behavior in every reported story location", async () => {
  await prime();
  for (const where of ["body", "cell", "headerCell", "endnote"] as const) {
    const bytes = matrixDocx(where, matrixParagraph("33333333", matrixRun("prev")) + matrixParagraph("44444444", PB + matrixRun("Heading")));
    const row = await runRow({
      id: `empty comment in ${where}`,
      bytes,
      where,
      before: none,
      after: (session, story) => {
        matrixComment(session, story, ["44444444", 1], ["44444444", 4]);
        session.deleteRange({ story, start: { paraId: "44444444", offset: 1 }, end: { paraId: "44444444", offset: 8 } });
      },
    });
    expect(row.cls).toBe("exact+unstable");
  }
});
const body = "body";

// "alpha" and one story unit lead " beta", which the new comment covers.
const alpha = run("alpha");
test.each([
  ["an earlier comment's reference mark", `${range(1, alpha)}`],
  ["a line break", `${alpha}<w:r><w:br/></w:r>`],
  ["a symbol", `${alpha}<w:r><w:sym w:font="Wingdings" w:char="F0E0"/></w:r>`],
  [
    "a field",
    `${alpha}<w:fldSimple w:instr=" DATE "><w:r><w:t>2026-09-29</w:t></w:r></w:fldSimple>`,
  ],
])(
  "a comment after %s keeps its range across publications",
  async (_, lead) => {
    const { seen } = await publications(
      docx(p("11111111", `${lead}${run(" beta gamma")}`) + tail),
      (session) => comment(session, "new", [body, "11111111", 6, 11])
    );
    expect(seen.map((ranges) => ranges.new)).toEqual([
      " beta",
      " beta",
      " beta",
    ]);
  }
);

const text = p("11111111", run("alpha beta gamma"));
test.each<[string, string, (session: YrsSession) => void]>([
  [
    "two comments ending at one point",
    text + tail,
    (session) => {
      comment(session, "a", [body, "11111111", 0, 5]);
      comment(session, "b", [body, "11111111", 2, 5]);
    },
  ],
  [
    "a reply on its parent's range",
    text + tail,
    (session) => {
      const parent = comment(session, "a", [body, "11111111", 0, 5]);
      const id = comment(session, "b", [body, "11111111", 0, 5]);
      session.applyRawOps(body, [
        { op: "patchComment", id, fields: { parentId: parent } },
      ]);
    },
  ],
  [
    "a nested comment",
    text + tail,
    (session) => {
      comment(session, "a", [body, "11111111", 0, 16]);
      comment(session, "b", [body, "11111111", 6, 10]);
    },
  ],
  [
    "overlapping comments",
    text + tail,
    (session) => {
      comment(session, "a", [body, "11111111", 0, 10]);
      comment(session, "b", [body, "11111111", 6, 16]);
    },
  ],
  [
    "a comment across paragraphs",
    text + p("33333333", run("delta")) + tail,
    (session) => comment(session, "a", [body, "11111111", 6, 3], "33333333"),
  ],
  [
    "a comment over an image",
    p("11111111", `${alpha}${image}${run(" beta")}`) + tail,
    (session) => comment(session, "a", [body, "11111111", 0, 6]),
  ],
  [
    "a comment in a paragraph after a table",
    `<w:tbl><w:tr><w:tc>${p(
      "44444444",
      run("cell")
    )}</w:tc></w:tr></w:tbl>${text}${tail}`,
    // The paragraph's offsets count the table that opens its slot.
    (session) => comment(session, "a", [body, "11111111", 7, 11]),
  ],
  [
    "a comment in a table cell",
    `<w:tbl><w:tr><w:tc>${p(
      "44444444",
      run("cell text")
    )}</w:tc></w:tr></w:tbl>${tail}`,
    (session) => comment(session, "a", ["body:t0:r0c0", "44444444", 5, 9]),
  ],
  [
    "a comment in a block content control",
    `<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtContent>${p(
      "44444444",
      run("sdt text")
    )}</w:sdtContent></w:sdt>${tail}`,
    (session) => {
      const story = session.storyIds().find((id) => id.includes(":sdt"))!;
      comment(session, "a", [story, "44444444", 4, 8]);
    },
  ],
])("%s keeps its range across publications", async (_, xml, edit) => {
  const { seen } = await publications(docx(xml, []), edit);
  expect(seen[1]).toEqual(seen[0]);
  expect(seen[2]).toEqual(seen[0]);
});

/** The toolbar's page break: split at the caret, then the break at the new paragraph's start. */
function pageBreak(
  session: YrsSession,
  story: string,
  paraId: string,
  offset: number
): string {
  const { secondParaId } = session.splitParagraph({ story, paraId, offset });
  session.insertPageBreak({ story, paraId: secondParaId, offset: 0 });
  return secondParaId;
}

// A page break is a story unit in every story, saved in its place.
const breakStories: Array<[string, string, string, string]> = [
  [
    "a table cell",
    "body:t0:r0c0",
    `${table(p("44444444", run("abcdef")))}${tail}`,
    "",
  ],
  ["a header", "hf:rId20", tail, p("44444444", run("abcdef"))],
  ["the body", body, p("44444444", run("abcdef")) + tail, ""],
];

test.each(breakStories)(
  "the comments beside a page break in %s keep their ranges across publications",
  async (_, story, xml, header) => {
    const { seen, bytes } = await publications(
      docx(xml, [], header),
      (session) => {
        const next = pageBreak(session, story, "44444444", 2);
        comment(session, "before", [story, "44444444", 0, 2]);
        comment(session, "after", [story, next, 2, 4]);
      }
    );
    expect(seen).toEqual(Array(3).fill({ before: "ab", after: "de" }));
    const part =
      story === "hf:rId20" ? "word/header1.xml" : "word/document.xml";
    const xmlOut = new TextDecoder().decode(unzipContainer(bytes)[part]);
    expect(xmlOut.match(/<w:br w:type="page"\/>/g) ?? []).toHaveLength(1);
  }
);

// An emptied range keeps its start before its end, then saves as its reference mark alone.
test.each<
  [string, string, string, (session: YrsSession, story: string) => void]
>([
  [
    "its text deleted",
    body,
    p("44444444", run("abcdef")) + tail,
    (session, story) => {
      comment(session, "a", [story, "44444444", 2, 4]);
      session.deleteRange({
        story,
        start: { paraId: "44444444", offset: 2 },
        end: { paraId: "44444444", offset: 4 },
      });
    },
  ],
  ...breakStories
    .slice(0, 2)
    .map(
      ([where, story, xml]): [
        string,
        string,
        string,
        (session: YrsSession, story: string) => void
      ] => [
        `only a page break in ${where}`,
        story,
        xml,
        (session) => {
          const next = pageBreak(session, story, "44444444", 2);
          comment(session, "a", [story, next, 0, 1]);
        },
      ]
    ),
])(
  "a comment over %s stays empty across publications",
  async (where, story, xml, edit) => {
    const header = where.endsWith("a header")
      ? p("44444444", run("abcdef"))
      : "";
    const { seen, bytes } = await publications(
      docx(xml, [], header),
      (session) => edit(session, story)
    );
    expect(seen.slice(1)).toEqual([{ a: "" }, { a: "" }]);
    const part =
      story === "hf:rId20" ? "word/header1.xml" : "word/document.xml";
    const order = [
      ...new TextDecoder()
        .decode(unzipContainer(bytes)[part])
        .matchAll(/<w:comment(RangeStart|RangeEnd|Reference) /g),
    ].map(([, kind]) => kind);
    expect(order).toEqual(["Reference"]);
  }
);

test("a comment ending inside a hyperlink widens to the link, and a later comment keeps its range", async () => {
  const link = `<w:hyperlink w:anchor="target"><w:r><w:t>linktext</w:t></w:r></w:hyperlink>`;
  const { seen } = await publications(
    docx(p("11111111", `${run("a ")}${link}${run(" b and more")}`) + tail, []),
    (session) => {
      comment(session, "a", [body, "11111111", 0, 6]);
      comment(session, "b", [body, "11111111", 13, 16]);
    }
  );
  expect(seen).toEqual([
    { a: "a link", b: "and" },
    { a: "a linktext", b: "and" },
    { a: "a linktext", b: "and" },
  ]);
});

test("Word's comment layouts seed exactly", async () => {
  const link = (xml: string) =>
    `<w:hyperlink w:anchor="target">${xml}</w:hyperlink>`;
  const tracked = (tag: string, xml: string) =>
    `<w:${tag} w:id="90" w:author="A" w:date="2026-09-01T00:00:00Z">${xml}</w:${tag}>`;
  const bytes = docx(
    // A reply's markers sit beside its parent's; Word's first comment id is 0.
    p(
      "11111111",
      `${S(0)}${S(1)}${alpha}${E(0)}${ref(0)}${E(1)}${ref(1)}${run(" beta")}`
    ) +
      p(
        "33333333",
        range(2, `${run("outer ")}${range(3, run("inner"))}${run(" end")}`)
      ) +
      p(
        "55555555",
        `${run("one ")}${S(4)}${run("two ")}${S(5)}${run("three")}${E(4)}${ref(
          4
        )}${run(" four")}${E(5)}${ref(5)}`
      ) +
      p("66666666", `${S(6)}${run("across")}`) +
      p("77777777", `${run("paragraphs")}${E(6)}${ref(6)}`) +
      table(p("44444444", range(7, run("cell")))) +
      // Markers inside a container move to its edges.
      p(
        "88888888",
        `${run("a ")}${link(`${run("li")}${S(8)}${run("nk")}`)}${run(" b")}${E(
          8
        )}${ref(8)}`
      ) +
      p(
        "99999999",
        `${tracked("ins", `${run("ins")}${S(9)}${run("erted")}${E(9)}`)}${ref(
          9
        )}${run(" c")}`
      ) +
      p(
        "AAAAAAAA",
        `${run("d ")}${S(10)}${tracked(
          "del",
          `<w:r><w:delText>de</w:delText></w:r>${E(
            10
          )}<w:r><w:delText>leted</w:delText></w:r>`
        )}${ref(10)}`
      ) +
      p(
        "BBBBBBBB",
        `<w:sdt><w:sdtPr><w:id w:val="8"/></w:sdtPr><w:sdtContent>${run(
          "ct"
        )}${S(11)}${run("rl")}</w:sdtContent></w:sdt>${run(" e")}${E(11)}${ref(
          11
        )}`
      ) +
      p(
        "CCCCCCCC",
        `${run("f ")}${field(`${run("20")}${S(12)}${run("26")}${E(12)}`)}${ref(
          12
        )}`
      ) +
      tail,
    [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
  );
  const seed = await seedOffice("docx", bytes);
  const session = await open(bytes);
  expect(covered(session)).toEqual({
    c0: "alpha",
    c1: "alpha",
    c2: "outer inner end",
    c3: "inner",
    c4: "two three",
    c5: "three four",
    c6: "across¶paragraphs",
    c7: "cell",
    c8: "link b",
    c9: "inserted",
    c10: "deleted",
    c11: "[sdt] e",
    c12: "[field]",
  });
  session.destroy();
  // A changed hash means the seed changed: the pin bump needs a maintenance window.
  expect(createHash("sha256").update(seed.state).digest("hex")).toBe(
    "bfffae2178380812a1df8c546b15b85d731c7e8d7827c16724b182658e9eb572"
  );
});

/** Each comment's range markers and reference marks in document.xml, in order. */
const marks = (bytes: Uint8Array) =>
  [
    ...new TextDecoder()
      .decode(unzipContainer(bytes)["word/document.xml"])
      .matchAll(/<w:comment(RangeStart|RangeEnd|Reference) w:id="(\d+)"/g),
  ].map(([, kind, id]) => `${kind === "Reference" ? "R" : kind[5]}${id}`);

test.each([
  [
    "into the next cell",
    table(
      p("44444444", `${run("A ")}${S(5)}${run("one")}`),
      p("55555555", `${run("two")}${E(5)}${ref(5)}`)
    ),
  ],
  [
    "out of a cell",
    table(p("44444444", `${run("in ")}${S(5)}${run("cell")}`)) +
      p("11111111", `${run("after")}${E(5)}${ref(5)}`),
  ],
  [
    "into a cell",
    p("11111111", `${run("pre ")}${S(5)}${run("body")}`) +
      table(p("44444444", `${run("in")}${E(5)}${ref(5)}${run(" cell")}`)),
  ],
  [
    "without an end",
    p("11111111", `${run("a ")}${S(5)}${run("b")}`) +
      p("33333333", run("next")),
  ],
])(
  "a Word range running %s saves one start, end and reference and then holds",
  async (_, xml) => {
    let bytes = docx(xml + tail, [5]);
    const seen: Array<Record<string, string>> = [];
    for (let publication = 0; publication < 2; publication += 1) {
      const session = await open(bytes);
      bytes = await publish(bytes, session);
      session.destroy();
      expect(marks(bytes).sort()).toEqual(["E5", "R5", "S5"]);
      const reopened = await open(bytes);
      seen.push(covered(reopened));
      reopened.destroy();
    }
    expect(seen[1]).toEqual(seen[0]);
  }
);

const inserted = (xml: string) =>
  `<w:ins w:id="90" w:author="A" w:date="2026-09-01T00:00:00Z">${xml}</w:ins>`;
test.each([
  [
    "a field result",
    field(`${run("20")}${S(5)}${run("26")}${E(5)}`),
    "[field]",
  ],
  [
    "a simple field",
    `<w:fldSimple w:instr=" DATE ">${run("20")}${S(5)}${run("26")}${E(
      5
    )}</w:fldSimple>`,
    "[field]",
  ],
  [
    "a change in a field result",
    field(`${inserted(`${S(5)}${run("20")}${E(5)}`)}${run("26")}`),
    "[field]",
  ],
])(
  "a Word range inside %s holds the whole field across publications",
  async (_, xml, whole) => {
    let bytes = docx(
      p("11111111", `${run("a ")}${xml}${ref(5)}${run(" b")}`) + tail,
      [5]
    );
    const seen: Array<Record<string, string>> = [];
    for (let publication = 0; publication < 3; publication += 1) {
      const session = await open(bytes);
      seen.push(covered(session));
      bytes = await publish(bytes, session);
      session.destroy();
      expect(marks(bytes).sort()).toEqual(["E5", "R5", "S5"]);
    }
    expect(seen).toEqual(Array(3).fill({ c5: whole }));
  }
);

test("a bookmark after a line break keeps its text across publications", async () => {
  const { bytes } = await publications(
    docx(
      p(
        "11111111",
        `${alpha}<w:r><w:br/></w:r>${run(
          " "
        )}<w:bookmarkStart w:id="5" w:name="mark"/>${run(
          "beta"
        )}<w:bookmarkEnd w:id="5"/>${run(" gamma")}`
      ) + tail,
      []
    ),
    () => {}
  );
  const xml = new TextDecoder().decode(
    unzipContainer(bytes)["word/document.xml"]
  );
  const marked =
    /<w:bookmarkStart [^>]*\/>([\s\S]*?)<w:bookmarkEnd [^>]*\/>/.exec(xml)![1];
  expect(
    [...marked.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(([, t]) => t).join("")
  ).toBe("beta");
});

/**
 * Gives the source comment's paragraph Word's `w14:paraId`, and with
 * `companions` the commentsExtended and commentsIds parts Word writes too.
 */
function withCommentIds(bytes: Uint8Array, companions: boolean): Uint8Array {
  const parts = unzipContainer(bytes);
  const text = (path: string) => new TextDecoder().decode(parts[path]);
  const set = (path: string, xml: string) =>
    (parts[path] = new TextEncoder().encode(xml));
  set(
    "word/comments.xml",
    text("word/comments.xml").replace(
      '<w:comment w:id="1" w:author="Reviewer" w:date="2026-09-01T00:00:00Z"><w:p>',
      '<w:comment w:id="1" w:author="Reviewer" w:date="2026-09-01T00:00:00Z"><w:p w14:paraId="0ABC1234">'
    )
  );
  if (companions) {
    set(
      "word/commentsExtended.xml",
      `<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="0ABC1234" w15:done="0"/></w15:commentsEx>`
    );
    set(
      "word/commentsIds.xml",
      `<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"><w16cid:commentId w16cid:paraId="0ABC1234" w16cid:durableId="7D1A2B3C"/></w16cid:commentsIds>`
    );
    set(
      "[Content_Types].xml",
      text("[Content_Types].xml").replace(
        "</Types>",
        `<Override PartName="/word/commentsExtended.xml" ContentType="${OFFICE}.commentsExtended+xml"/><Override PartName="/word/commentsIds.xml" ContentType="${OFFICE}.commentsIds+xml"/></Types>`
      )
    );
    set(
      "word/_rels/document.xml.rels",
      text("word/_rels/document.xml.rels").replace(
        "</Relationships>",
        `<Relationship Id="rId10" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/><Relationship Id="rId11" Type="http://schemas.microsoft.com/office/2016/09/relationships/commentsIds" Target="commentsIds.xml"/></Relationships>`
      )
    );
  }
  return rezipContainer(parts);
}

/**
 * Each comment's paragraph id and durable id, and the paragraph id each
 * reply names as its parent, by the comment's w:id.
 */
function threadIds(bytes: Uint8Array) {
  const parts = unzipContainer(bytes);
  const text = (path: string) =>
    parts[path] ? new TextDecoder().decode(parts[path]) : "";
  const paraIds = new Map(
    [
      ...text("word/comments.xml").matchAll(
        /<w:comment w:id="(-?\d+)"[^>]*><w:p w14:paraId="([0-9A-F]+)"/g
      ),
    ].map(([, id, paraId]) => [id!, paraId!])
  );
  const durable = new Map(
    [
      ...text("word/commentsIds.xml").matchAll(
        /w16cid:paraId="([0-9A-F]+)" w16cid:durableId="([0-9A-F]+)"/g
      ),
    ].map(([, paraId, id]) => [paraId!, id!])
  );
  const parents = new Map(
    [
      ...text("word/commentsExtended.xml").matchAll(
        /w15:paraId="([0-9A-F]+)"[^>]*w15:paraIdParent="([0-9A-F]+)"/g
      ),
    ].map(([, paraId, parent]) => [paraId!, parent!])
  );
  return [...paraIds].map(([id, paraId]) => ({
    id,
    paraId,
    durable: durable.get(paraId),
    parent: parents.get(paraId),
  }));
}

test.each([
  ["without w14:paraId", (bytes: Uint8Array) => bytes],
  ["with w14:paraId", (bytes: Uint8Array) => withCommentIds(bytes, false)],
  [
    "with w14:paraId, commentsExtended and commentsIds",
    (bytes: Uint8Array) => withCommentIds(bytes, true),
  ],
] as const)(
  "a comment and reply added after a capture get their own ids, the source comment's paragraph %s",
  async (_, shape) => {
    let bytes = shape(
      docx(
        p("11111111", `${range(1, run("prev"))}`) +
          p("33333333", run("next")) +
          tail
      )
    );
    const session = await open(bytes);
    const captured = session.encodeState();
    const exported = await exportOffice(
      bytes,
      checkpoint(bytes, captured),
      fixed
    );
    const parent = comment(session, "added", [body, "33333333", 0, 4]);
    session.applyRawOps(body, [
      {
        op: "patchComment",
        id: `reply-${parent}`,
        fields: {
          author: "Reviewer",
          date: "2026-09-29T00:00:00Z",
          body: [],
          parentId: parent,
        },
      },
    ]);
    const latest = session.encodeState();
    const direct = await exportOffice(bytes, checkpoint(bytes, latest), fixed);
    session.destroy();
    const { state } = await rebaseOffice(
      bytes,
      checkpoint(bytes, captured),
      checkpoint(bytes, latest),
      exported
    );
    bytes = await exportOffice(exported, checkpoint(exported, state), fixed);
    const ids = threadIds(bytes);
    expect(ids).toEqual(threadIds(direct));
    expect(ids).toHaveLength(3);
    for (const key of ["paraId", "durable"] as const)
      expect(new Set(ids.map((entry) => entry[key])).size).toBe(ids.length);
    const [reply] = ids.filter((entry) => entry.parent);
    expect(reply!.parent).toBe(
      ids.find((entry) => entry.id !== "1" && !entry.parent)!.paraId
    );
    for (let publication = 0; publication < 2; publication += 1) {
      const next = await open(bytes);
      next.insertText({ story: body, paraId: "22222222", offset: 0 }, "x");
      bytes = await publish(bytes, next);
      next.destroy();
      expect(threadIds(bytes)).toEqual(ids);
    }
  }
);

test("a generated comment paragraph id avoids an existing document paragraph id", async () => {
  const probeBytes = docx(p("11111111", run("text")) + tail, []);
  const probe = await open(probeBytes);
  comment(probe, "probe", [body, "11111111", 0, 4]);
  const occupied = threadIds(await publish(probeBytes, probe))[0]!.paraId;
  probe.destroy();

  let bytes = docx(p(occupied, run("text")) + tail, []);
  const session = await open(bytes);
  comment(session, "kept", [body, occupied, 0, 4]);
  bytes = await publish(bytes, session);
  session.destroy();
  const ids = threadIds(bytes);
  expect(ids[0]!.paraId).not.toBe(occupied);
  for (let publication = 0; publication < 2; publication += 1) {
    const next = await open(bytes);
    next.insertText({ story: body, paraId: "22222222", offset: 0 }, "x");
    bytes = await publish(bytes, next);
    next.destroy();
    expect(threadIds(bytes)).toEqual(ids);
  }
});

const flowBreak = (kind: "page" | "column") => `<w:r><w:br w:type="${kind}"/></w:r>`;
const commentStories = ["body", "cell", "header"] as const;
function commentStory(where: typeof commentStories[number], xml: string) {
  return {
    story: where === "cell" ? "body:t0:r0c0" : where === "header" ? "hf:rId20" : body,
    bytes: docx((where === "body" ? xml : where === "cell" ? table(xml) : "") + tail, [1], where === "header" ? xml : ""),
  };
}

for (const where of commentStories) {
  test(`a draft reply puts its reference beside the parent's after a leading break in ${where}`, async () => {
    const { bytes, story } = commentStory(where, p("33333333", run("prev")) + p("44444444", flowBreak("page") + run("Heading")));
    const session = await open(bytes);
    const parent = comment(session, "a", [story, "33333333", 0, 0], "44444444");
    session.applyRawOps(body, [{ op: "patchComment", id: `reply-${parent}`, fields: { author: "Reviewer", date: fixed.now, body: [], parentId: parent } }]);
    const base = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });
    const projected = yrsToDocument(session, base);
    const paragraphs = where === "header" ? projected.package.headers!.get("rId20")!.content : projected.package.document.content;
    injectReplyRangeMarkers(paragraphs, projected.package.document.comments ?? []);
    const saved = new Uint8Array(await repackDocx(projected));
    session.destroy();
    const xml = sig(saved, where === "header" ? "word/header1.xml" : "word/document.xml");
    expect(xml).toContain("[PB]RRHeading");
    expect(xml.indexOf("R")).toBeGreaterThan(xml.indexOf("[PB]"));
  });

  test(`a comment on a text-less page-and-column-break paragraph keeps the same spacing before and after save in ${where}`, async () => {
    const { bytes, story } = commentStory(where, p("33333333", run("prev")) + p("44444444", flowBreak("page") + flowBreak("column")) + p("45454545", run("next")));
    const session = await open(bytes);
    comment(session, "a", [story, "33333333", 0, 0], "44444444");
    const before = blocks(session, story);
    const saved = await publish(bytes, session);
    session.destroy();
    const reopened = await open(saved);
    expect(before).toContain("PB^");
    expect(blocks(reopened, story)).toBe(before);
    reopened.destroy();
  });

  test(`a Word reference before a leading break keeps the break ahead of text in ${where}`, async () => {
    const { bytes, story } = commentStory(where, p("33333333", S(1) + run("prev")) + p("44444444", E(1) + ref(1) + flowBreak("page") + run("Heading")));
    const session = await open(bytes);
    const parts = session.storySegments(story);
    const at = parts.findIndex((part) => part.kind === "embed" && part.embedKind === "pageBreak");
    const text = parts.findIndex((part) => part.kind === "text" && part.text.includes("Heading"));
    expect(at).toBeGreaterThanOrEqual(0);
    expect(at).toBeLessThan(text);
    session.destroy();
    const { seen } = await publications(bytes, () => {});
    expect(seen).toEqual(Array(3).fill({ c1: "prev¶" }));
  });

  test(`a source range after a leading break stays there when its reference is in the next paragraph in ${where}`, async () => {
    const { bytes } = commentStory(where, p("33333333", run("prev")) + p("44444444", flowBreak("page") + S(1) + run("Heading") + E(1)) + p("45454545", ref(1) + run("next")));
    const { seen } = await publications(bytes, () => {});
    expect(seen).toEqual(Array(3).fill({ c1: "Heading" }));
  });

  test(`an editor range at the text start after a leading break stays there in ${where}`, async () => {
    const { bytes, story } = commentStory(where, p("33333333", run("prev")) + p("44444444", flowBreak("page") + run("Heading")));
    const { seen } = await publications(bytes, (session) => comment(session, "a", [story, "44444444", 1, 4]));
    expect(seen.map((value) => value.a)).toEqual(Array(3).fill("Hea"));
  });

  for (const kind of ["page", "column"] as const) {
    test(`a comment added after capture ends before a trailing ${kind} break in ${where}`, async () => {
      const { bytes, story } = commentStory(where, p("33333333", run("prev") + flowBreak(kind)) + p("44444444", run("Heading")));
      const session = await open(bytes);
      comment(session, "B", [story, "33333333", 0, 2]);
      const captured = session.encodeState();
      const exported = await publish(bytes, session);
      comment(session, "C", [story, "33333333", 2, 4]);
      const latest = session.encodeState();
      const direct = await publish(bytes, session);
      session.destroy();
      const { state } = await rebaseOffice(bytes, checkpoint(bytes, captured), checkpoint(bytes, latest), exported);
      const rebased = await exportOffice(exported, checkpoint(exported, state), fixed);
      const part = where === "header" ? "word/header1.xml" : "word/document.xml";
      expect(sig(rebased, part)).toBe(sig(direct, part));
      for (const saved of [direct, rebased]) {
        const { seen } = await publications(saved, () => {});
        expect(seen.map((value) => value.C)).toEqual(Array(3).fill("ev"));
      }
    });
  }
}

/** Every comment range in `bytes` opens before it closes. */
const ordered = (bytes: Uint8Array) => {
  const xml = new TextDecoder().decode(unzipContainer(bytes)["word/document.xml"]);
  return [...xml.matchAll(/<w:commentRange(End) w:id="(\d+)"/g)].every(
    ({ index, 2: id }) => xml.slice(0, index).includes(`<w:commentRangeStart w:id="${id}"`)
  );
};
// 8a: a comment ending where an emptied one sits saves both in order, so direct and rebased saves agree.
test.each([
  ["Word's shape", `${E(1)}${ref(1)}${pageBreak}`],
  ["no reference", `${E(1)}${pageBreak}`],
])("an emptied comment beside another's end saves alike directly and rebased, %s", async (_, lead) => {
  const bytes = docx(p("33333333", `${S(1)}${run("prev")}`) + p("44444444", `${lead}${run("Heading")}`) + tail);
  const later = (session: YrsSession) => {
    comment(session, "before", [body, "33333333", 0, 0], "44444444");
    comment(session, "over", [body, "44444444", 0, 1]);
    session.deleteAt({ story: body, paraId: "44444444", offset: 1 }, "backward");
  };
  for (let attempt = 0; attempt < 6; attempt += 1) {
    const session = await open(bytes);
    const captured = session.encodeState();
    const exported = await exportOffice(bytes, checkpoint(bytes, captured), fixed);
    later(session);
    const latest = session.encodeState();
    session.destroy();
    const direct = await exportOffice(bytes, checkpoint(bytes, latest), fixed);
    const { state } = await rebaseOffice(bytes, checkpoint(bytes, captured), checkpoint(bytes, latest), exported);
    const next = await exportOffice(exported, checkpoint(exported, state), fixed);
    const saves = [];
    for (const saved of [direct, next]) {
      expect(ordered(saved)).toBe(true);
      const reopened = await open(saved);
      saves.push(covered(reopened));
      reopened.destroy();
    }
    expect(saves[1]).toEqual(saves[0]!);
    expect(saves[0]!.over).toBe("");
  }
});

// 8b: text typed after a text-less column-break paragraph keeps the break in its paragraph while a comment
// boundary sits before it, so the comment neither grows over the break nor loses it.
const columnBreak = `<w:r><w:br w:type="column"/></w:r>`;
test.each([
  ["ending before it", [body, "33333333", 0, 0] as Span, "44444444", "prev¶"],
  ["over it", [body, "44444444", 0, 1] as Span, undefined, "[columnBreak]"],
])("a comment %s keeps its range when text is typed after a column-break paragraph", async (_, span, endPara, text) => {
  const { seen } = await publications(
    docx(p("33333333", run("prev")) + p("44444444", columnBreak) + p("45454545", run("next")) + tail),
    (session) => {
      comment(session, "c", span, endPara);
      session.insertText({ story: body, paraId: "44444444", offset: 1 }, "Q");
    }
  );
  expect(seen.map((ranges) => ranges.c)).toEqual([text, text, text]);
});

test.each([
  ["ending before it", [body, "33333333", 0, 0] as Span, "44444444", "prev¶"],
  ["over it", [body, "44444444", 0, 1] as Span, undefined, "[columnBreak]"],
])("a comment %s keeps its range over a column break closing the paragraph before", async (_, span, endPara, text) => {
  const { seen } = await publications(
    docx(p("33333333", run("prev") + columnBreak) + p("44444444", run("Heading")) + tail),
    (session) => comment(session, "c", span, endPara)
  );
  expect(seen.map((ranges) => ranges.c)).toEqual([text, text, text]);
});

// Round 2 H: a source comment's reference opens a heading; editor comments end before it and cover it. Deleting the
// heading's first units after a capture empties the covering comment: the rebase keeps it as the direct save does,
// or refuses.
test("an emptied editor comment over a source reference survives the rebase or refuses, in every story", async () => {
  await prime();
  for (const where of ["body", "control", "cell", "header", "footer", "footnote", "headerCell", "endnote"] as const) {
    const bytes = matrixDocx(
      where,
      matrixParagraph("33333333", `${S(1)}${matrixRun("prev")}`) + matrixParagraph("44444444", `${E(1)}${ref(1)}${matrixRun("Heading")}`)
    );
    const row = await runRow({
      id: `emptied over a source reference in ${where}`,
      bytes,
      where,
      before: (session, story) => {
        matrixComment(session, story, ["33333333", 0], ["44444444", 0]);
        matrixComment(session, story, ["44444444", 0], ["44444444", 1]);
      },
      after: (session, story) =>
        session.deleteRange({ story, start: { paraId: "44444444", offset: 0 }, end: { paraId: "44444444", offset: 2 } }),
    });
    expect(row.cls).not.toStartWith("silent");
  }
});

// Round 2 review B: breaks a comment boundary or bookmark sits at, with edits after a capture. J: unrelated
// edits land; N: a rebase that would move the breaks, bookmarks or comment differently refuses.
const bmMark = `<w:bookmarkStart w:id="5" w:name="m5"/><w:bookmarkEnd w:id="5"/>`;
const reviewB: Record<string, string> = {
  "prev[CB]¶[CB]¶next": matrixParagraph("33333333", matrixRun("prev") + matrixCB) + matrixParagraph("44444444", matrixCB) + matrixParagraph("45454545", matrixRun("next")),
  "prev¶<bm/>[PB][CB]¶next": matrixParagraph("33333333", matrixRun("prev")) + matrixParagraph("44444444", bmMark + PB + matrixCB) + matrixParagraph("45454545", matrixRun("next")),
  "prev¶<bm>[CB]</bm>¶next": matrixParagraph("33333333", matrixRun("prev")) + matrixParagraph("44444444", `<w:bookmarkStart w:id="5" w:name="m5"/>${matrixCB}<w:bookmarkEnd w:id="5"/>`) + matrixParagraph("45454545", matrixRun("next")),
};
const endOf44: (s: Parameters<MatrixEdit>[0], st: string) => { story: string; paraId: string; offset: number } = (s, st) =>
  ({ story: st, paraId: "44444444", offset: matrixLen(s, st, "44444444") });
const reviewBEdits: Record<string, MatrixEdit> = {
  "type Q": (s, st) => void s.insertText(endOf44(s, st), "Q"),
  "join next": (s, st) => void s.deleteAt(endOf44(s, st), "forward"),
  "type Q, Backspace before Q": (s, st) => {
    s.insertText(endOf44(s, st), "Q");
    s.deleteAt({ story: st, paraId: "44444444", offset: matrixTextStart(s, st, "44444444") }, "backward");
  },
};
test.each(["body", "cell", "header"] as MatrixWhere[])(
  "edits after a capture next to breaks a comment ends at or a bookmark sits at land exactly or refuse, in %s",
  async (where) => {
    await prime();
    for (const [name, xml] of Object.entries(reviewB))
      for (const [edit, after] of Object.entries(reviewBEdits))
        for (const order of orders(`${where} | ${name} | ${edit}`, (s, st) => void matrixComment(s, st, ["33333333", 0], ["44444444", 0]), after)) {
          const row = await runRow({ id: order.id, bytes: matrixDocx(where, xml), where, before: order.before, after: order.after });
          expect(`${order.id}: ${row.cls}`).not.toMatch(/: silent|: error/);
          // Edits made before the capture, then typing elsewhere: the breaks read alike, so the rebase lands.
          if (order.id.endsWith("edit<cap<tail")) expect(`${order.id}: ${row.cls}`).toMatch(/: exact/);
        }
  }
);

// Round 2 review B, L: the coverage refusal fires only when an edit after the capture touches the comment's range or
// what it covers; typing elsewhere lands with the accepted timing difference.
test("an edit after a capture away from a comment the export settled lands", async () => {
  await prime();
  for (const where of ["body", "cell", "header"] as MatrixWhere[]) {
    const bytes = matrixDocx(
      where,
      matrixParagraph("33333333", matrixRun("prev")) +
        matrixParagraph("44444444", `<w:commentRangeStart w:id="1"/>${PB}${matrixRun("Heading")}<w:commentRangeEnd w:id="1"/>${ref(1)}`)
    );
    const row = await runRow({
      id: `src comment over [PB]Heading in ${where}`,
      bytes,
      where,
      before: (session, story) =>
        void matrixComment(session, story, ["44444444", matrixTextStart(session, story, "44444444")], ["44444444", matrixLen(session, story, "44444444")]),
      after: none,
    });
    expect(row.cls).not.toStartWith("refused");
    expect(row.cls).not.toStartWith("silent");
  }
});

// Round 2 review C, S: a comment range typing left reversed before the capture is no edit after it: typing elsewhere
// after the capture lands.
test("a comment range reversed before a capture lets unrelated edits after it land", async () => {
  await prime();
  for (const where of ["body", "cell", "header"] as MatrixWhere[]) {
    const row = await runRow({
      id: `reversed range in ${where}`,
      bytes: matrixDocx(where, reviewB["prev¶<bm/>[PB][CB]¶next"]!),
      where,
      before: (session, story) => {
        matrixComment(session, story, ["44444444", 0], ["44444444", matrixLen(session, story, "44444444")]);
        session.insertText(endOf44(session, story), "Q");
      },
      after: (session) => void session.insertText({ story: "body", paraId: "22222222", offset: 0 }, "z"),
    });
    expect(row.cls).not.toStartWith("refused");
    expect(row.cls).not.toStartWith("silent");
  }
});

// Round 2 review D, X: a range reversed before a capture is touched when a later edit touches what lies between
// its ends: those rebases refuse instead of landing silently.
test("a comment range reversed before a capture refuses edits after it that touch its content", async () => {
  await prime();
  const files = {
    "prev¶<bm/>[PB][CB]¶next": reviewB["prev¶<bm/>[PB][CB]¶next"]!,
    "prev¶abc¶next": matrixParagraph("33333333", matrixRun("prev")) + matrixParagraph("44444444", matrixRun("abc")) + matrixParagraph("45454545", matrixRun("next")),
  };
  const setups: Record<string, MatrixEdit> = {
    "a comment over 44": (session, story) => {
      matrixComment(session, story, ["44444444", 0], ["44444444", matrixLen(session, story, "44444444")]);
      session.insertText(endOf44(session, story), "Q");
    },
    "an empty comment at its end": (session, story) => {
      matrixComment(session, story, ["44444444", matrixLen(session, story, "44444444")], ["44444444", matrixLen(session, story, "44444444")]);
      session.insertText(endOf44(session, story), "Q");
    },
  };
  const afters: Record<string, MatrixEdit> = {
    "Backspace before Q": (s, st) => void s.deleteAt({ story: st, paraId: "44444444", offset: matrixTextStart(s, st, "44444444") }, "backward"),
    "type at 44 start": (s, st) => void s.insertText({ story: st, paraId: "44444444", offset: 0 }, "S"),
    "delete 44 content": (s, st) => void s.deleteRange({ story: st, start: { paraId: "44444444", offset: 0 }, end: { paraId: "44444444", offset: matrixLen(s, st, "44444444") } }),
  };
  for (const where of ["body", "cell", "header"] as MatrixWhere[])
    for (const [file, xml] of Object.entries(files))
      for (const [setup, before] of Object.entries(setups))
        for (const [edit, after] of Object.entries(afters)) {
          const id = `${where} | ${file} | ${setup} | ${edit}`;
          const row = await runRow({ id, bytes: matrixDocx(where, xml), where, before, after });
          expect(`${id}: ${row.cls}`).not.toMatch(/: silent|: error/);
        }
});

// Follow-up item 5: a comment range typing reversed before a capture covers what lies between its ends, not the units
// beside them, so a join or Backspace beside it after the capture lands with the accepted timing difference instead
// of refusing, and never silently (the reviewer's rv4-s rows).
test("a join or Backspace beside a comment range reversed before a capture lands with the timing difference", async () => {
  await prime();
  const textAt44 = (s: YrsSession, st: string) => ({ story: st, paraId: "44444444", offset: matrixLen(s, st, "44444444") });
  const typeQ = (comment: (s: YrsSession, st: string) => void): MatrixEdit => (s, st) => {
    comment(s, st);
    s.insertText(textAt44(s, st), "Q");
  };
  const emptyAtEnd = typeQ((s, st) => matrixComment(s, st, ["44444444", matrixLen(s, st, "44444444")], ["44444444", matrixLen(s, st, "44444444")]));
  const over44 = typeQ((s, st) => matrixComment(s, st, ["44444444", 0], ["44444444", matrixLen(s, st, "44444444")]));
  const backspaceBeforeQ: MatrixEdit = (s, st) => void s.deleteAt({ story: st, paraId: "44444444", offset: matrixTextStart(s, st, "44444444") }, "backward");
  const joinNext: MatrixEdit = (s, st) => void s.deleteAt(textAt44(s, st), "forward");
  const joinPrev: MatrixEdit = (s, st) => void s.deleteAt({ story: st, paraId: "44444444", offset: 0 }, "backward");
  const three = (x: string) => matrixParagraph("33333333", matrixRun("prev")) + matrixParagraph("44444444", x) + matrixParagraph("45454545", matrixRun("next"));
  const bmMark = `<w:bookmarkStart w:id="5" w:name="m5"/><w:bookmarkEnd w:id="5"/>`;
  const rows: Array<[string, string, MatrixEdit, MatrixEdit]> = [
    ["prev[CB]¶[CB]¶next | empty, Backspace before Q", matrixParagraph("33333333", matrixRun("prev") + matrixCB) + matrixParagraph("44444444", matrixCB) + matrixParagraph("45454545", matrixRun("next")), emptyAtEnd, backspaceBeforeQ],
    ["prev[CB]¶[CB]¶next | empty, join next", matrixParagraph("33333333", matrixRun("prev") + matrixCB) + matrixParagraph("44444444", matrixCB) + matrixParagraph("45454545", matrixRun("next")), emptyAtEnd, joinNext],
    ["prev¶<bm/>[PB][CB]¶next | empty, Backspace before Q", three(bmMark + PB + matrixCB), emptyAtEnd, backspaceBeforeQ],
    ["prev¶<bm/>[PB][CB]¶next | empty, join prev", three(bmMark + PB + matrixCB), emptyAtEnd, joinPrev],
    ["prev¶<bm/>[PB][CB]¶next | over, Backspace before Q", three(bmMark + PB + matrixCB), over44, backspaceBeforeQ],
    ["prev¶<bm/>[PB][CB]¶next | over, join prev", three(bmMark + PB + matrixCB), over44, joinPrev],
    ["prev¶[PB][CB]¶next | empty, join next", three(PB + matrixCB), emptyAtEnd, joinNext],
    ["prev¶abc¶next | empty, join next", three(matrixRun("abc")), emptyAtEnd, joinNext],
  ];
  for (const where of ["body", "cell", "header"] as MatrixWhere[])
    for (const [name, xml, before, after] of rows) {
      const id = `${where} | ${name}`;
      const row = await runRow({ id, bytes: matrixDocx(where, xml), where, before, after });
      expect(`${id}: ${row.cls}`).toMatch(/: timing/);
    }
});

// Decided 2026-10-02: Undo and Redo re-anchor the comment ranges and bookmarks in text they restore, so another
// replica, a peer and the save place them where the editor shows them (they followed the restored text only in the
// session that pressed Undo), and a bookmark in text that Redo deletes again comes back with the next Undo.
// "hello world" holds an editor comment or a bookmark over "lo wo".
const undoP = "44444444";
const undoRange = (start: number, end: number) => ({ story: "body", start: { paraId: undoP, offset: start }, end: { paraId: undoP, offset: end } });
const undoFlows: Record<string, [(session: YrsSession) => void, string]> = {
  "delete all, Undo": [(session) => (session.deleteRange(undoRange(0, 11)), session.undo()), "lo wo"],
  "delete the middle, Undo": [(session) => (session.deleteRange(undoRange(2, 9)), session.undo()), "lo wo"],
  "type in it, delete all, Undo": [
    (session) => {
      session.insertText({ story: "body", paraId: undoP, offset: 6 }, "QQ");
      session.addUndoBoundary();
      session.deleteRange(undoRange(0, 13));
      session.undo();
    },
    "lo QQwo",
  ],
  "delete all, Undo, Redo, Undo": [
    (session) => {
      session.deleteRange(undoRange(0, 11));
      session.undo();
      session.redo();
      session.undo();
    },
    "lo wo",
  ],
};
test.each(Object.keys(undoFlows).flatMap((flow) => (["comment", "bookmark"] as const).map((shape) => [shape, flow] as const)))(
  "a %s over text survives %s in the editor, a peer and the save",
  async (shape, flow) => {
    await prime();
    const bytes = matrixDocx("body", matrixParagraph(undoP, shape === "comment" ? matrixRun("hello world") : matrixRun("hel") + bm(matrixRun("lo wo")) + matrixRun("rld")));
    const [act, expected] = undoFlows[flow]!;
    const session = await matrixOpen(bytes);
    if (shape === "comment") matrixComment(session, "body", [undoP, 3], [undoP, 8]);
    session.addUndoBoundary();
    const peer = await matrixOpen(bytes);
    peer.applyUpdate(session.encodeStateAsUpdate(peer.encodeStateVector()));
    act(session);
    peer.applyUpdate(session.encodeStateAsUpdate(peer.encodeStateVector()));
    const saved = await matrixPublish(bytes, session.encodeState());
    if (shape === "comment") {
      expect(matrixCovered(session, "body")).toEqual([expected]);
      expect(matrixCovered(peer, "body")).toEqual([expected]);
      expect(sig(saved, "word/document.xml")).toStartWith(`hel<c1${expected}c1>Rrld¶`);
      const reopened = await matrixOpen(saved);
      expect(matrixCovered(reopened, "body")).toEqual([expected]);
      reopened.destroy();
    } else {
      expect(sig(saved, "word/document.xml")).toStartWith(`helB5${expected}E5rld¶`);
      expect(sig(await matrixPublish(bytes, peer.encodeState()), "word/document.xml")).toStartWith(`helB5${expected}E5rld¶`);
    }
    session.destroy();
    peer.destroy();
  }
);

// Two peers: B types at the paragraph's start while A deletes the text and undoes (a comment), or deletes, undoes,
// redoes and undoes (a bookmark). Both peers and the reopened save keep the marker over "lo wo" and every character.
test.each(["comment", "bookmark"] as const)("a %s over text keeps its place on both peers when one undoes beside the other's typing", async (shape) => {
  await prime();
  const bytes = matrixDocx("body", matrixParagraph(undoP, shape === "comment" ? matrixRun("hello world") : matrixRun("hel") + bm(matrixRun("lo wo")) + matrixRun("rld")));
  const A = await matrixOpen(bytes);
  if (shape === "comment") matrixComment(A, "body", [undoP, 3], [undoP, 8]);
  A.addUndoBoundary();
  const B = await matrixOpen(bytes);
  B.applyUpdate(A.encodeStateAsUpdate(B.encodeStateVector()));
  A.deleteRange(undoRange(0, 11));
  A.undo();
  if (shape === "bookmark") {
    A.redo();
    A.undo();
  }
  B.insertText({ story: "body", paraId: undoP, offset: 0 }, "Z");
  const [toB, toA] = [A.encodeStateAsUpdate(B.encodeStateVector()), B.encodeStateAsUpdate(A.encodeStateVector())];
  A.applyUpdate(toA);
  B.applyUpdate(toB);
  const saved = await matrixPublish(bytes, A.encodeState());
  const reopened = await matrixOpen(saved);
  const text = sig(saved, "word/document.xml").split("¶")[0]!;
  expect([...text.replace(/<c1|c1>R|B5|E5/g, "")].sort().join("")).toBe([..."Zhello world"].sort().join(""));
  if (shape === "comment") {
    for (const session of [A, B, reopened]) expect(matrixCovered(session, "body")).toEqual(["lo wo"]);
    expect(text).toContain("<c1lo woc1>R");
  } else {
    expect(text).toContain("B5lo woE5");
    expect(sig(await matrixPublish(bytes, B.encodeState()), "word/document.xml").split("¶")[0]).toBe(text);
  }
  for (const session of [A, B, reopened]) session.destroy();
});
