import { beforeAll, expect, test } from "bun:test";
import { createHash } from "node:crypto";
import { createYrsSession, type YrsSession } from "../packages/docx/src/yrs";
import { rezipContainer, unzipContainer } from "../packages/docx/src/wasm/opc";
import { exportOffice, seedOffice } from "./office-checkpoint";

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
