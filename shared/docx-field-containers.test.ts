import { beforeAll, expect, test } from "bun:test";
import { createHash } from "node:crypto";
import { createYrsSession, type YrsSession } from "../packages/docx/src/yrs";
import { rezipContainer, unzipContainer } from "../packages/docx/src/wasm/opc";
import { exportOffice, rebaseOffice, seedOffice } from "./office-checkpoint";
import { RebaseError } from "./office-rebase";

const fixed = { seed: "0".repeat(64), now: "2026-09-29T00:00:00.000Z" };
const W =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const OFFICE = "application/vnd.openxmlformats-officedocument.wordprocessingml";
const REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const deleted = (text: string) =>
  `<w:r><w:delText xml:space="preserve">${text}</w:delText></w:r>`;
const p = (id: string, xml: string) => `<w:p w14:paraId="${id}">${xml}</w:p>`;
const ref = (id: number) =>
  `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`;
const change = (tag: string, id: number) => (xml: string) =>
  `<w:${tag} w:id="${id}" w:author="A" w:date="2026-09-01T00:00:00Z">${xml}</w:${tag}>`;
const ins = change("ins", 90);
const del = change("del", 91);
const moveFrom = change("moveFrom", 92);
const moveTo = change("moveTo", 93);
const fs = (xml: string, instruction = " DATE ") =>
  `<w:fldSimple w:instr="${instruction}">${xml}</w:fldSimple>`;
const link = (xml: string) =>
  `<w:hyperlink w:anchor="target">${xml}</w:hyperlink>`;
const sdt = (xml: string) =>
  `<w:sdt><w:sdtPr><w:id w:val="8"/></w:sdtPr><w:sdtContent>${xml}</w:sdtContent></w:sdt>`;
const bm = (xml: string) =>
  `<w:bookmarkStart w:id="5" w:name="mark"/>${xml}<w:bookmarkEnd w:id="5"/>`;
const char = (type: string) =>
  `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
const instr = (text: string) =>
  `<w:r><w:instrText xml:space="preserve">${text}</w:instrText></w:r>`;
const field = (result: string, instruction = " DATE ") =>
  `${char("begin")}${instr(instruction)}${char("separate")}${result}${char("end")}`;
const foreign = `<x:ext xmlns:x="urn:x" x:v="1"/>`;

function docx(body: string): Uint8Array {
  const parts: Record<string, string> = {
    "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="${OFFICE}.comments+xml"/></Types>`,
    "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
    "word/_rels/document.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="${REL}/comments" Target="comments.xml"/></Relationships>`,
    "word/document.xml": `<w:document ${W}><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
    "word/comments.xml": `<w:comments ${W}><w:comment w:id="5" w:author="Reviewer" w:date="2026-09-01T00:00:00Z"><w:p><w:r><w:t>c5</w:t></w:r></w:p></w:comment></w:comments>`,
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

// Seeding loads the engines the sessions below use directly.
beforeAll(() => seedOffice("docx", docx(p("11111111", run("x")))));
let clientId = 9700;
/** A session over `bytes`: seeded, or holding `state` over them. */
async function open(
  bytes: Uint8Array,
  state?: Uint8Array
): Promise<YrsSession> {
  const session = await createYrsSession({ clientId: (clientId += 1) });
  session.openDocx(bytes, !state);
  if (state) session.loadState(state);
  return session;
}

const checkpoint = (base: Uint8Array, state: Uint8Array) => ({
  format: "docx" as const,
  schemaVersion: 1 as const,
  baseSha256: createHash("sha256").update(base).digest("hex"),
  state,
});

const publish = (bytes: Uint8Array, state: Uint8Array) =>
  exportOffice(bytes, checkpoint(bytes, state), fixed);

const documentXml = (bytes: Uint8Array, path = "word/document.xml") =>
  new TextDecoder().decode(unzipContainer(bytes)[path]);

/**
 * A paragraph's markup as text: F<> simple field, [«code»|result] complex
 * field, H() link, +{} -{} M+{} M-{} tracked changes, S{} content control,
 * {bm bm} bookmark, X foreign markup.
 */
function view(bytes: Uint8Array, paraId = "11111111", path?: string): string {
  const paragraph = documentXml(bytes, path).match(
    new RegExp(`<w:p [^>]*"${paraId}"[\\s\\S]*?</w:p>`)
  )![0];
  const tags =
    /<(\/?)(?:w:(t|delText|instrText|fldChar|ins|del|moveFrom|moveTo|fldSimple|hyperlink|bookmarkStart|bookmarkEnd|sdt)|x:(ext))\b([^>]*?)(\/?)>([^<]*)/g;
  const opening: Record<string, string> = {
    fldSimple: "F<",
    hyperlink: "H(",
    ins: "+{",
    del: "-{",
    moveTo: "M+{",
    moveFrom: "M-{",
    sdt: "S{",
  };
  return [...paragraph.matchAll(tags)]
    .map(([, close, tag, ext, attrs, , text]) => {
      if (ext) return "X";
      if (tag === "t" || tag === "delText") return close ? "" : text;
      if (tag === "instrText") return close ? "" : `«${text.trim()}»`;
      if (tag === "fldChar")
        return { begin: "[", separate: "|", end: "]" }[
          attrs.match(/fldCharType="(\w+)"/)![1]
        ];
      if (tag === "bookmarkStart") return "{bm";
      if (tag === "bookmarkEnd") return "bm}";
      if (close) return tag === "fldSimple" ? ">" : tag === "hyperlink" ? ")" : "}";
      return opening[tag];
    })
    .join("");
}

/** The comment markers and references the body writes, as S5, E5, R5. */
const marks = (bytes: Uint8Array) =>
  [
    ...documentXml(bytes).matchAll(
      /<w:comment(RangeStart|RangeEnd|Reference) w:id="(\d+)"/g
    ),
  ]
    .map(([, kind, id]) => `${{ RangeStart: "S", RangeEnd: "E" }[kind] ?? "R"}${id}`)
    .sort();

/** The body's units: characters, ¶ and [embed], with reference marks as "". */
const units = (session: YrsSession) =>
  session
    .storySegments("body")
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
    );

/** What the one comment covers. */
function covered(session: YrsSession): string {
  const [{ id }] = session.listComments();
  const [{ start, end }] = session.resolveComment(id);
  return units(session).slice(start, end).join("");
}

const tail = p("22222222", run("tail"));
const paragraph = (xml: string) =>
  docx(p("11111111", `${run("a ")}${xml}${run(" b")}`) + tail);
const edit = (session: YrsSession, paraId: string, text: string) =>
  session.insertText({ story: "body", paraId, offset: 0 }, text);

/**
 * Each container as Word writes it with its innermost run passed through
 * `wrap`, the first paragraph's view once saved, and the text the editor's
 * field shows: its result, insertions in and deletions out.
 */
type Build = (wrap: (xml: string) => string) => string;
const cases: Array<[string, Build, string, string]> = [
  ["a bookmark inside w:fldSimple", (wrap) => fs(bm(wrap(run("2026")))), "a F<{bm2026bm}> b", "a 2026 b"],
  ["w:fldSimple inside a hyperlink", (wrap) => link(fs(wrap(run("2026")))), "a H(F<2026>) b", "a 2026 b"],
  ["w:fldSimple inside w:ins", (wrap) => ins(fs(wrap(run("2026")))), "a +{F<2026>} b", "a 2026 b"],
  ["w:ins inside w:fldSimple", (wrap) => fs(ins(wrap(run("2026")))), "a F<+{2026}> b", "a 2026 b"],
  ["w:sdt inside w:fldSimple", (wrap) => fs(sdt(wrap(run("2026")))), "a F<S{2026}> b", "a 2026 b"],
  [
    "w:ins inside a complex field result",
    (wrap) => field(`${run("20")}${ins(wrap(run("26")))}`),
    "a [«DATE»|20+{26}] b",
    "a 2026 b",
  ],
  [
    "w:del inside a complex field result",
    (wrap) => field(`${run("20")}${del(wrap(deleted("26")))}`),
    "a [«DATE»|20-{26}] b",
    "a 20 b",
  ],
  [
    "w:moveFrom and w:moveTo inside a complex field result",
    (wrap) => field(`${run("20")}${moveFrom(deleted("x"))}${moveTo(wrap(run("26")))}`),
    "a [«DATE»|20M-{x}M+{26}] b",
    "a 2026 b",
  ],
  [
    "w:sdt inside a complex field result",
    (wrap) => field(`${run("20")}${sdt(wrap(run("26")))}`),
    "a [«DATE»|20S{26}] b",
    "a 2026 b",
  ],
  [
    "w:ins beside a hyperlink in a complex field result",
    (wrap) => field(`${link(run("20"))}${ins(wrap(run("26")))}`),
    "a [«DATE»|H(20)+{26}] b",
    "a 2026 b",
  ],
  [
    "w:sdt beside a simple field in a complex field result",
    (wrap) => field(`${fs(run("20"), " PAGE ")}${sdt(wrap(run("26")))}`),
    "a [«DATE»|F<20>S{26}] b",
    "a 2026 b",
  ],
  [
    "foreign markup inside a complex field result",
    (wrap) => field(`${run("20")}${foreign}${wrap(run("26"))}`),
    "a [«DATE»|20X26] b",
    "a 2026 b",
  ],
  [
    "a complex field inside w:fldSimple",
    (wrap) => fs(`${run("20")}${field(wrap(run("26")), " PAGE ")}`),
    "a F<20[«PAGE»|26]> b",
    "a 2026 b",
  ],
  [
    "w:ins in the code of a complex field inside w:fldSimple",
    (wrap) => fs(`${char("begin")}${ins(wrap(instr(" PAGE ")))}${char("separate")}${run("26")}${char("end")}`),
    "a F<[+{«PAGE»}|26]> b",
    "a 26 b",
  ],
];
const plain = (xml: string) => xml;
/** The first paragraph as the editor shows it: its text and its fields' texts. */
function shownText(session: YrsSession): string {
  let text = "";
  for (const segment of session.storySegments("body")) {
    if (segment.kind === "pilcrow") break;
    if (segment.kind === "text") text += segment.text;
    else if (segment.embedKind === "field")
      text += String(segment.payload.displayText ?? "");
  }
  return text;
}

test.each(cases)(
  "%s round-trips through three publications, its paragraph untouched",
  async (_, build, expected, shown) => {
    let bytes = paragraph(build(plain));
    let saved = "";
    for (let publication = 0; publication < 3; publication += 1) {
      const session = await open(bytes);
      expect(shownText(session)).toBe(shown);
      edit(session, "22222222", "x");
      bytes = await publish(bytes, session.encodeState());
      session.destroy();
      expect(view(bytes)).toBe(expected);
      const first = documentXml(bytes).match(
        /<w:p [^>]*"11111111"[\s\S]*?<\/w:p>/
      )![0];
      if (publication > 0) expect(first).toBe(saved);
      saved = first;
    }
  }
);

test.each(cases)(
  "%s keeps its place when its paragraph is edited",
  async (_, build, expected) => {
    let bytes = paragraph(build(plain));
    for (let publication = 1; publication <= 3; publication += 1) {
      const session = await open(bytes);
      edit(session, "11111111", "x");
      bytes = await publish(bytes, session.encodeState());
      session.destroy();
      expect(view(bytes)).toBe(`${"x".repeat(publication)}${expected}`);
    }
  }
);

/**
 * Captures after `before`, publishes, makes an edit to the field's paragraph
 * after the capture and rebases it onto the export; returns the paragraph's
 * next saved view once the rebased units match the latest ones.
 */
async function rebased(bytes: Uint8Array, before: (session: YrsSession) => void) {
  const session = await open(bytes);
  before(session);
  edit(session, "22222222", "x");
  const captured = session.encodeState();
  const exported = await publish(bytes, captured);
  edit(session, "11111111", "y");
  const latest = session.encodeState();
  const latestUnits = units(session).join("");
  session.destroy();
  const { state } = await rebaseOffice(
    bytes,
    checkpoint(bytes, captured),
    checkpoint(bytes, latest),
    exported
  );
  const rebased = await open(exported, state);
  expect(units(rebased).join("")).toBe(latestUnits);
  const next = await publish(exported, rebased.encodeState());
  rebased.destroy();
  return view(next);
}

test.each(cases)(
  "%s rebases an edit to its paragraph made after the capture",
  async (_, build, expected) => {
    expect(await rebased(paragraph(build(plain)), () => {})).toBe(`y${expected}`);
  }
);

/** A projected field's units: its projected children, then the field. */
const projectedUnits: Record<string, string> = {
  "w:ins beside a hyperlink in a complex field result": "20[field]",
  "w:sdt beside a simple field in a complex field result": "[field][field]",
};
test.each(cases)(
  "a Word comment range inside %s holds the whole field across publications",
  async (name, build, expected) => {
    const wrap = (xml: string) =>
      `<w:commentRangeStart w:id="5"/>${xml}<w:commentRangeEnd w:id="5"/>`;
    let bytes = docx(
      p("11111111", `${run("a ")}${build(wrap)}${ref(5)}${run(" b")}`) + tail
    );
    for (let publication = 0; publication < 3; publication += 1) {
      const session = await open(bytes);
      expect(covered(session)).toBe(projectedUnits[name] ?? "[field]");
      bytes = await publish(bytes, session.encodeState());
      session.destroy();
      expect(marks(bytes)).toEqual(["E5", "R5", "S5"]);
      expect(view(bytes)).toBe(expected);
    }
  }
);

/** The view and field texts Accept All and Reject All leave, where they differ from the source. */
const resolvedCases: Record<string, Record<"accept" | "reject", [string, string]>> = {
  "w:fldSimple inside w:ins": {
    accept: ["a F<2026> b", "a 2026 b"],
    reject: ["a  b", "a  b"],
  },
  "w:ins inside w:fldSimple": {
    accept: ["a F<2026> b", "a 2026 b"],
    reject: ["a F<> b", "a  b"],
  },
  "w:ins inside a complex field result": {
    accept: ["a [«DATE»|2026] b", "a 2026 b"],
    reject: ["a [«DATE»|20] b", "a 20 b"],
  },
  "w:del inside a complex field result": {
    accept: ["a [«DATE»|20] b", "a 20 b"],
    reject: ["a [«DATE»|2026] b", "a 2026 b"],
  },
  "w:moveFrom and w:moveTo inside a complex field result": {
    accept: ["a [«DATE»|2026] b", "a 2026 b"],
    reject: ["a [«DATE»|20x] b", "a 20x b"],
  },
  "w:ins beside a hyperlink in a complex field result": {
    accept: ["a [«DATE»|H(20)26] b", "a 2026 b"],
    reject: ["a [«DATE»|H(20)] b", "a 20 b"],
  },
  "w:ins in the code of a complex field inside w:fldSimple": {
    accept: ["a F<[«PAGE»|26]> b", "a 26 b"],
    reject: ["a F<[«»|26]> b", "a 26 b"],
  },
};
const resolveAll = (mode: "accept" | "reject") => (session: YrsSession) =>
  mode === "accept"
    ? session.acceptChange({ all: true })
    : session.rejectChange({ all: true });

test.each(
  cases.flatMap(([name, build, expected, shown]) =>
    (["accept", "reject"] as const).map(
      (mode) =>
        [name, mode, build, ...(resolvedCases[name]?.[mode] ?? [expected, shown])] as const
    )
  )
)(
  "%s: %s all resolves the changes it keeps, across publications and a rebase",
  async (name, mode, build, expected, shown) => {
    const bytes = paragraph(build(plain));
    const session = await open(bytes);
    // Changes inside a field stay out of the revision list.
    expect(session.listRevisions().length).toBe(name === "w:fldSimple inside w:ins" ? 1 : 0);
    resolveAll(mode)(session);
    expect(session.listRevisions()).toEqual([]);
    expect(shownText(session)).toBe(shown);
    let saved = await publish(bytes, session.encodeState());
    session.destroy();
    for (let publication = 0; publication < 3; publication += 1) {
      expect(view(saved)).toBe(expected);
      const reopened = await open(saved);
      edit(reopened, "22222222", "x");
      saved = await publish(saved, reopened.encodeState());
      reopened.destroy();
    }
    expect(await rebased(bytes, resolveAll(mode))).toBe(`y${expected}`);
  }
);

test.each([
  [
    "w:del holding its field's end",
    field(run("20")).replace(char("end"), del(`${deleted("26")}${char("end")}`)),
  ],
  ["w:ins holding its field's end", field(run("2026")).replace(char("end"), ins(char("end")))],
  [
    "w:ins opening a nested field",
    field(`${run("20")}${ins(`${char("begin")}${instr(" PAGE ")}${char("separate")}`)}${run("26")}${char("end")}`),
  ],
])("a %s leaves its field balanced and byte-stable", async (_, xml) => {
  let bytes = paragraph(xml);
  let saved = "";
  for (let publication = 0; publication < 3; publication += 1) {
    const session = await open(bytes);
    edit(session, "22222222", "x");
    bytes = await publish(bytes, session.encodeState());
    session.destroy();
    const fields = view(bytes).replace(/[^[\]]/g, "");
    expect(fields.split("[").length).toBe(fields.split("]").length);
    if (publication > 0) expect(view(bytes)).toBe(saved);
    saved = view(bytes);
  }
});

test.each([
  ["a complex field result", (xml: string) => field(`${run("20")}${ins(xml)}`)],
  ["w:fldSimple", (xml: string) => fs(ins(xml))],
])(
  "a Word comment reference inside a kept change in %s is written once and goes with its comment",
  async (_, build) => {
    const bytes = docx(
      p(
        "11111111",
        `${run("a ")}<w:commentRangeStart w:id="5"/>${build(
          `${run("26")}<w:commentRangeEnd w:id="5"/>${ref(5)}`
        )}${run(" b")}`
      ) + tail
    );
    let saved = bytes;
    for (let publication = 0; publication < 3; publication += 1) {
      const session = await open(saved);
      saved = await publish(saved, session.encodeState());
      session.destroy();
      expect(marks(saved)).toEqual(["E5", "R5", "S5"]);
    }
    expect(await rebased(bytes, () => {})).toStartWith("ya ");
    const session = await open(bytes);
    session.applyRawOps("body", [
      { op: "removeComment", id: session.listComments()[0].id },
    ]);
    const removed = await publish(bytes, session.encodeState());
    session.destroy();
    expect(marks(removed)).toEqual([]);
  }
);

const toc = (first: string) =>
  docx(
    p("11111111", `${char("begin")}${instr(" TOC \\o ")}${char("separate")}${first}`) +
      p("33333333", `${run("Entry2 2")}${char("end")}`) +
      tail
  );
test.each([
  ["w:del", `${del(deleted("Old"))}${run("Entry1 1")}`, "-{Old}Entry1 1", "Entry1 1", "OldEntry1 1"],
  ["w:ins", `${ins(run("Entry1"))}${run(" 1")}`, "+{Entry1} 1", "Entry1 1", " 1"],
  ["w:ins around a hyperlink", ins(link(run("Entry1 1"))), "+{H(Entry1 1)}", "H(Entry1 1)", ""],
  ["w:sdt", `${run("En")}${sdt(run("try1"))}${run(" 1")}`, "EnS{try1} 1", "EnS{try1} 1", "EnS{try1} 1"],
])(
  "a field result spanning paragraphs keeps %s in its first paragraph, and Accept or Reject All resolves it",
  async (_, first, kept, accepted, rejected) => {
    const bytes = toc(first);
    const result = (saved: Uint8Array) => view(saved).match(/\|(.*)\]/)?.[1];
    for (const [resolve, expected] of [
      [() => {}, kept],
      [resolveAll("accept"), accepted],
      [resolveAll("reject"), rejected],
    ] as const) {
      let saved = bytes;
      for (let publication = 0; publication < 3; publication += 1) {
        const session = await open(saved);
        if (publication === 0) resolve(session);
        edit(session, "22222222", "x");
        saved = await publish(saved, session.encodeState());
        session.destroy();
        expect(result(saved)).toBe(expected);
      }
    }
  }
);

/** The first paragraph as the story holds it: text, and [kind] per embed with a field's text. */
const unitsView = (session: YrsSession) =>
  session
    .storySegments("body")
    .map((segment) =>
      segment.kind === "text"
        ? segment.text
        : segment.kind === "pilcrow"
        ? "¶"
        : segment.embedKind === "field"
        ? `[${segment.payload.displayText}]`
        : `[${segment.embedKind}]`
    )
    .join("")
    .split("¶")[0];

test.each([
  ["a hyperlink Accept All uncovers", field(`${run("20")}${ins(link(run("26")))}`), "accept"],
  ["a simple field Accept All uncovers", field(`${run("20")}${ins(fs(run("26"), " PAGE "))}`), "accept"],
  ["a hyperlink Reject All restores", field(`${run("20")}${del(link(deleted("26")))}`), "reject"],
  ["a nested field Accept All uncovers", field(`${run("20")}${ins(field(run("26"), " PAGE "))}`), "accept"],
  ["a nested field whose change Reject All removes", field(`${run("20")}${field(`${run("2")}${ins(run("6"))}`, " PAGE ")}`), "reject"],
  ["a projected simple field's change Accept All resolves", field(`${run("20")}${fs(ins(run("26")), " PAGE ")}`), "accept"],
] as const)(
  "%s in a field result leaves the units its export seeds, and later edits rebase",
  async (_, xml, mode) => {
    const bytes = paragraph(xml);
    const session = await open(bytes);
    resolveAll(mode)(session);
    expect(session.hasFieldChanges()).toBe(false);
    const resolved = unitsView(session);
    const out = await publish(bytes, session.encodeState());
    session.destroy();
    const reopened = await open(out);
    expect(unitsView(reopened)).toBe(resolved);
    reopened.destroy();
    // The field embed: the paragraph's last embed.
    const fieldAt = (session: YrsSession) => {
      let at = 0;
      let offset = 0;
      for (const segment of session.storySegments("body")) {
        if (segment.kind === "pilcrow") break;
        if (segment.kind === "embed") at = offset;
        offset += segment.kind === "text" ? segment.text.length : 1;
      }
      return at;
    };
    for (const later of [
      (session: YrsSession) => edit(session, "11111111", "y"),
      // Right after the children it projects, which the typing joins.
      (session: YrsSession) =>
        session.insertText({ story: "body", paraId: "11111111", offset: fieldAt(session) }, "y"),
      (session: YrsSession) =>
        session.deleteRange({
          story: "body",
          start: { paraId: "11111111", offset: fieldAt(session) },
          end: { paraId: "11111111", offset: fieldAt(session) + 1 },
        }),
    ]) {
      const session = await open(bytes);
      resolveAll(mode)(session);
      edit(session, "22222222", "x");
      const captured = session.encodeState();
      const exported = await publish(bytes, captured);
      later(session);
      const direct = view(await publish(bytes, session.encodeState()));
      const { state } = await rebaseOffice(bytes, checkpoint(bytes, captured), checkpoint(bytes, session.encodeState()), exported);
      session.destroy();
      const rebased = await open(exported, state);
      expect(view(await publish(exported, rebased.encodeState()))).toBe(direct);
      rebased.destroy();
    }
  }
);

test("a projected link edited before Accept All keeps its edit beside the link it uncovers", async () => {
  const bytes = paragraph(field(`${link(run("20"))}${ins(link(run("26")))}`));
  const session = await open(bytes);
  edit(session, "11111111", "");
  session.insertText({ story: "body", paraId: "11111111", offset: 3 }, "x");
  resolveAll("accept")(session);
  const out = await publish(bytes, session.encodeState());
  session.destroy();
  expect(view(out)).toBe("a [«DATE»|H(2x0)H(26)] b");
});

test.each([
  [
    "a change",
    `<w:r><w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> DATE \\@ yyyy </w:instrText></w:r>${ins(instr(" \\* MERGEFORMAT "))}`,
    "a [«DATE \\@ yyyy»+{«\\* MERGEFORMAT»}|2026] b",
  ],
  [
    "foreign markup",
    `<w:r><w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> DATE \\@ yyyy </w:instrText></w:r>${foreign}`,
    "a [«DATE \\@ yyyy»X|2026] b",
  ],
  [
    "a nested field (IF { PAGE })",
    `${char("begin")}${instr(" IF ")}${field(run("1"), " PAGE ")}${instr(" = 1 yes no ")}`,
    "a [«IF»[«PAGE»|1]«= 1 yes no»|2026] b",
  ],
])("field code holding %s keeps its instruction and nodes across publications", async (_, code, expected) => {
  let bytes = paragraph(`${code}${char("separate")}${run("2026")}${char("end")}`);
  for (let publication = 0; publication < 2; publication += 1) {
    const session = await open(bytes);
    edit(session, "22222222", "x");
    bytes = await publish(bytes, session.encodeState());
    session.destroy();
    expect(view(bytes)).toBe(expected);
  }
});

test.each(["accept", "reject"] as const)(
  "%s all keeps a nested field in field code as a field",
  async (mode) => {
    let bytes = paragraph(
      `${char("begin")}${instr(" IF ")}${field(run("1"), " PAGE ")}${ins(instr(" = 1 yes no "))}${char("separate")}${run("a")}${char("end")}`
    );
    const expected = mode === "accept" ? "a [«IF»[«PAGE»|1]«= 1 yes no»|a] b" : "a [«IF»[«PAGE»|1]|a] b";
    for (let publication = 0; publication < 2; publication += 1) {
      const session = await open(bytes);
      if (publication === 0) resolveAll(mode)(session);
      edit(session, "22222222", "x");
      bytes = await publish(bytes, session.encodeState());
      session.destroy();
      expect(view(bytes)).toBe(expected);
    }
  }
);

test("a comment reference sharing a run with text in a kept change is written once", async () => {
  const bytes = docx(
    p(
      "11111111",
      `${run("a ")}<w:commentRangeStart w:id="5"/>${field(
        `${run("20")}${ins('<w:commentRangeEnd w:id="5"/><w:r><w:t>26</w:t><w:commentReference w:id="5"/></w:r>')}`
      )}${run(" b")}`
    ) + tail
  );
  const session = await open(bytes);
  const out = await publish(bytes, session.encodeState());
  session.destroy();
  expect(marks(out)).toEqual(["E5", "R5", "S5"]);
});

test("Accept All of a formatted kept run settles the field's formatting in one publication", async () => {
  let bytes = paragraph(field(`${run("20")}${ins('<w:r><w:rPr><w:b/></w:rPr><w:t>26</w:t></w:r>')}`));
  const saved: string[] = [];
  for (let publication = 0; publication < 2; publication += 1) {
    const session = await open(bytes);
    if (publication === 0) resolveAll("accept")(session);
    edit(session, "22222222", "x");
    bytes = await publish(bytes, session.encodeState());
    session.destroy();
    saved.push(documentXml(bytes).match(/<w:p [^>]*"11111111"[\s\S]*?<\/w:p>/)![0]);
  }
  expect(saved[1]).toBe(saved[0]);
});

/** The offset of the first projected-child text holding `text` in the first paragraph. */
function childAt(session: YrsSession, text: string): number {
  let offset = 0;
  for (const segment of session.storySegments("body")) {
    if (segment.kind === "pilcrow") break;
    if (segment.kind === "text" && segment.attributes.fieldResult && segment.text.includes(text))
      return offset + segment.text.indexOf(text);
    offset += segment.kind === "text" ? segment.text.length : 1;
  }
  throw new Error(`no projected child holds ${text}`);
}

type Edit = (session: YrsSession) => void;
/**
 * Makes `before` ahead of a capture and `after` behind it, and returns the
 * rebased next publication beside the latest state's direct publication,
 * as `show` reads them.
 */
async function landed(bytes: Uint8Array, before: Edit, after: Edit, show: (bytes: Uint8Array) => string = view) {
  const session = await open(bytes);
  before(session);
  edit(session, "22222222", "x");
  const captured = session.encodeState();
  const exported = await publish(bytes, captured);
  after(session);
  const latest = session.encodeState();
  session.destroy();
  const direct = show(await publish(bytes, latest));
  const { state } = await rebaseOffice(bytes, checkpoint(bytes, captured), checkpoint(bytes, latest), exported);
  const rebased = await open(exported, state);
  const next = show(await publish(exported, rebased.encodeState()));
  rebased.destroy();
  return { next, direct };
}
const typeIn = (text: string): Edit => (session) =>
  session.insertText({ story: "body", paraId: "11111111", offset: childAt(session, text) + 1 }, "Z");

const ref5 = `<w:commentRangeStart w:id="5"/>${run("c ")}<w:commentRangeEnd w:id="5"/>${ref(5)}`;
const pageBreak = `<w:r><w:br w:type="page"/></w:r>`;
const splitRuns = ["00A1", "00B2", "00C3"]
  .map((rsid, index) => `<w:r w:rsidR="${rsid}"><w:t>${"abc"[index]}</w:t></w:r>`)
  .join("");
const refField = (text: string, bookmark: string) => field(link(run(text)), ` REF ${bookmark} \\h `);
const uncovered = field(`${run("20")}${ins(link(run("26")))}`);
// A rebase never re-pairs a projected child: text typed in one after the
// capture carries the capture's number for its field, and lands only when the
// export's seed numbers the field the same way.
test.each([
  ["nothing before the capture", refField("20", "a"), () => {}, "20"],
  ["Accept All after a Word comment", `${ref5}${uncovered}`, resolveAll("accept"), "26"],
] as const)("text typed in a projected child after the capture lands in its field: %s", async (_, xml, before, text) => {
  const { next, direct } = await landed(paragraph(xml), before, typeIn(text));
  expect(next).toBe(direct);
});
test.each([
  ["Word's split runs before it", `${splitRuns}${refField("20", "a")}${run(" ")}${refField("xy", "b")}`, () => {}, "20"],
  [
    "bold added before the capture",
    refField("20", "a"),
    (session: YrsSession) =>
      session.formatRange({ story: "body", start: { paraId: "11111111", offset: 0 }, end: { paraId: "11111111", offset: 1 } }, { bold: true }),
    "20",
  ],
  [
    "a comment added before the capture",
    refField("20", "a"),
    (session: YrsSession) =>
      session.addComment(
        [{ story: "body", start: { paraId: "11111111", offset: 0 }, end: { paraId: "11111111", offset: 1 } }],
        "Reviewer",
        "2026-09-29T00:00:00Z",
        [{ type: "paragraph", content: [{ type: "run", content: [{ type: "text", text: "c" }] }] }]
      ),
    "20",
  ],
  ["Accept All after a tab", `<w:r><w:t>x</w:t><w:tab/></w:r>${uncovered}`, resolveAll("accept"), "26"],
  ["Accept All after a bookmark", `<w:bookmarkStart w:id="7" w:name="m"/>${run("x")}<w:bookmarkEnd w:id="7"/>${uncovered}`, resolveAll("accept"), "26"],
  ["Accept All beside a mid-paragraph break", `${pageBreak}${uncovered}${refField("30", "b")}`, resolveAll("accept"), "30"],
  ["Reject All of a tracked leading break", `${ins(pageBreak)}${uncovered}${refField("30", "b")}`, resolveAll("reject"), "30"],
] as const)("text typed in a projected child after the capture refuses the rebase: %s", async (_, xml, before, text) => {
  await expect(landed(paragraph(xml), before, typeIn(text))).rejects.toBeInstanceOf(RebaseError);
});

test.each([
  ["a deleted projected link stays deleted", field(`${link(run("20"))}${ins(run("26"))}`), "20", "delete", "a [«DATE»|26] b"],
  ["a deleted projected simple field stays deleted", field(`${fs(run("20"), " PAGE ")}${ins(run("26"))}`), "", "deleteField", "a [«DATE»|26] b"],
  ["typing in a link whose nested field resolves is kept", field(`${link(`${run("20")}${fs(ins(run("1")), " PAGE ")}`)}${run("x")}`), "20", "type", "a [«DATE»|H(2Z0F<1>)x] b"],
] as const)("Accept All resolves changes only: %s", async (_, xml, text, kind, expected) => {
  const bytes = paragraph(xml);
  const session = await open(bytes);
  if (kind === "type") session.insertText({ story: "body", paraId: "11111111", offset: childAt(session, text) + 1 }, "Z");
  else {
    const at = kind === "delete" ? childAt(session, text) : 2;
    session.deleteRange({
      story: "body",
      start: { paraId: "11111111", offset: at },
      end: { paraId: "11111111", offset: at + (kind === "delete" ? text.length : 1) },
    });
  }
  resolveAll("accept")(session);
  const out = await publish(bytes, session.encodeState());
  session.destroy();
  expect(view(out)).toBe(expected);
});

test("typing in a TOC's first-paragraph entry is kept by Accept All, which keeps the entry a link", async () => {
  const bytes = docx(
    p("11111111", `${char("begin")}${instr(" TOC \\o \\h ")}${char("separate")}${link(run("Intro 1"))}${ins(run("X"))}`) +
      p("33333333", `${run("Entry2 2")}${char("end")}`) +
      tail
  );
  const session = await open(bytes);
  session.insertText({ story: "body", paraId: "11111111", offset: childAt(session, "Intro") + 2 }, "Z");
  resolveAll("accept")(session);
  const out = await publish(bytes, session.encodeState());
  session.destroy();
  expect(view(out)).toBe("[«TOC \\o \\h»|H(InZtro 1)X]");
});

const deletedRun = del(deleted("xy"));
test.each([
  ["after a comment", paragraph(`<w:commentRangeStart w:id="5"/>${run("c ")}<w:commentRangeEnd w:id="5"/>${uncovered}`)],
  [
    "after a comment and deleted text before it",
    paragraph(`${deletedRun}<w:commentRangeStart w:id="5"/>${run("c ")}<w:commentRangeEnd w:id="5"/>${ref(5)}${uncovered}`),
  ],
  [
    "below a paragraph with deleted text and a comment",
    docx(p("33333333", `${deletedRun}<w:commentRangeStart w:id="5"/>${run("c")}<w:commentRangeEnd w:id="5"/>${ref(5)}`) + p("11111111", `${uncovered}${refField("30", "b")}`) + tail),
  ],
])("Accept All numbers the fields it projects as its export's seed does %s", async (_, bytes) => {
  const ids = (session: YrsSession) =>
    session.storySegments("body").flatMap((segment) =>
      segment.kind === "embed" && segment.payload.resultProjection
        ? [(segment.payload.resultProjection as { id: number }).id]
        : []
    );
  const session = await open(bytes);
  resolveAll("accept")(session);
  const resolved = ids(session);
  const out = await publish(bytes, session.encodeState());
  session.destroy();
  const reopened = await open(out);
  expect(resolved).toEqual(ids(reopened));
  reopened.destroy();
});

test("an external link Accept All uncovers carries its URL at once", async () => {
  const parts = unzipContainer(
    paragraph(field(`${run("20")}${ins(`<w:hyperlink r:id="rId20">${run("26")}</w:hyperlink>`)}`))
  );
  const rels = new TextDecoder().decode(parts["word/_rels/document.xml.rels"]).replace(
    "</Relationships>",
    `<Relationship Id="rId20" Type="${REL}/hyperlink" Target="https://example.com/x" TargetMode="External"/></Relationships>`
  );
  parts["word/_rels/document.xml.rels"] = new TextEncoder().encode(rels);
  const session = await open(rezipContainer(parts));
  resolveAll("accept")(session);
  const hrefs = session
    .storySegments("body")
    .flatMap((segment) => {
      const link = segment.attributes.hyperlink as { href?: string } | undefined;
      return link ? [link.href] : [];
    });
  session.destroy();
  expect(hrefs).toEqual(["https://example.com/x"]);
});

/**
 * `docx(body)` with a default header (rId10) and footnote 1 holding
 * `header` and `footnote`, and each part's own relationships `rels`.
 */
function withStories(body: string, header: string, footnote: string, rels = { header: "", footnotes: "" }): Uint8Array {
  const parts = unzipContainer(docx(body));
  const text = (path: string) => new TextDecoder().decode(parts[path]);
  const set = (path: string, xml: string) => (parts[path] = new TextEncoder().encode(xml));
  const relationships = (xml: string) =>
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${xml}</Relationships>`;
  set("word/header1.xml", `<w:hdr ${W}>${header}</w:hdr>`);
  set(
    "word/footnotes.xml",
    `<w:footnotes ${W}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="1">${footnote}</w:footnote></w:footnotes>`
  );
  set("word/_rels/header1.xml.rels", relationships(rels.header));
  set("word/_rels/footnotes.xml.rels", relationships(rels.footnotes));
  set(
    "word/_rels/document.xml.rels",
    text("word/_rels/document.xml.rels").replace(
      "</Relationships>",
      `<Relationship Id="rId10" Type="${REL}/header" Target="header1.xml"/><Relationship Id="rId11" Type="${REL}/footnotes" Target="footnotes.xml"/></Relationships>`
    )
  );
  set(
    "[Content_Types].xml",
    text("[Content_Types].xml").replace(
      "</Types>",
      `<Override PartName="/word/header1.xml" ContentType="${OFFICE}.header+xml"/><Override PartName="/word/footnotes.xml" ContentType="${OFFICE}.footnotes+xml"/></Types>`
    )
  );
  set("word/document.xml", text("word/document.xml").replace("<w:sectPr>", `<w:sectPr><w:headerReference w:type="default" r:id="rId10"/>`));
  return rezipContainer(parts);
}
const noteRef = `<w:r><w:footnoteReference w:id="1"/></w:r>`;
const cell = (xml: string) =>
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>${xml}</w:tc></w:tr></w:tbl>`;

/** The story, paragraph and offset in it of the first unit `match` picks. */
function locate(session: YrsSession, match: (segment: ReturnType<YrsSession["storySegments"]>[number]) => boolean) {
  for (const story of session.storyIds()) {
    let offset = 0;
    let found: number | undefined;
    for (const segment of session.storySegments(story)) {
      if (segment.kind === "pilcrow") {
        if (found !== undefined) return { story, paraId: segment.paraId, offset: found };
        offset = 0;
        continue;
      }
      if (found === undefined && match(segment)) found = offset;
      offset += segment.kind === "text" ? segment.text.length : 1;
    }
  }
  throw new Error("no such unit");
}
const fieldAt = (session: YrsSession, instruction: string) =>
  locate(session, (segment) => segment.kind === "embed" && String(segment.payload.instruction ?? "").includes(instruction));
const textAt = (session: YrsSession, text: string) => {
  const at = locate(session, (segment) => segment.kind === "text" && segment.text.includes(text));
  const segment = session.storySegments(at.story).find((entry) => entry.kind === "text" && entry.text.includes(text));
  return { ...at, offset: at.offset + (segment as { text: string }).text.indexOf(text) };
};
/** Backspace just after the field whose instruction holds `instruction`: its embed goes. */
const backspaceField = (instruction: string): Edit => (session) => {
  const { story, paraId, offset } = fieldAt(session, instruction);
  session.deleteAt({ story, paraId, offset: offset + 1 }, "backward");
};
/** Types Z `after` units into the first text holding `text`. */
const typeInText = (text: string, after = 1): Edit => (session) => {
  const { story, paraId, offset } = textAt(session, text);
  session.insertText({ story, paraId, offset: offset + after }, "Z");
};
/** Deletes from inside the text holding `text` through its field's embed and the unit after. */
const deleteAcrossEnd = (text: string, instruction: string): Edit => (session) => {
  const { story, paraId, offset } = textAt(session, text);
  const end = fieldAt(session, instruction).offset + 2;
  session.deleteRange({ story, start: { paraId, offset: offset + 1 }, end: { paraId, offset: end } });
};

const twoFields = `${run("a ")}${refField("20", "a")}${run(" mid ")}${refField("30", "b")}${run(" b")}`;
const orphaned = "a H(20) mid [«REF b \\h»|H(30)] b";
test.each([
  ["Backspace removes its embed", paragraph(`${refField("20", "a")}${run(" mid ")}${refField("30", "b")}`), () => {}, backspaceField("REF a"), orphaned],
  ["a delete across its end", paragraph(`${refField("20", "a")}${run(" mid ")}${refField("30", "b")}`), () => {}, deleteAcrossEnd("20", "REF a"), "a H(2)mid [«REF b \\h»|H(30)] b"],
  ["text is typed in it, its embed gone before the capture", paragraph(`${refField("20", "a")}${run(" mid ")}${refField("30", "b")}`), backspaceField("REF a"), typeInText("20"), "a H(2Z0) mid [«REF b \\h»|H(30)] b"],
  ["the next field follows at once", paragraph(`${refField("20", "a")}${refField("30", "b")}`), () => {}, backspaceField("REF a"), "a H(20)[«REF b \\h»|H(30)] b"],
  ["the next field has no child at its index", paragraph(`${field(`${run("q")}${link(run("20"))}`, " REF a \\h ")}${run(" mid ")}${refField("30", "b")}`), () => {}, backspaceField("REF a"), orphaned],
  ["Accept All uncovered it", paragraph(`${uncovered}${run(" mid ")}${field(`${run("20")}${ins(link(run("27")))}`, " REF b ")}`), resolveAll("accept"), deleteAcrossEnd("26", "DATE"), "a H(2)mid [«REF b»|20H(27)] b"],
] as const)("a field's child stays where it is when %s", async (_, bytes, before, after, expected) => {
  const { next, direct } = await landed(bytes, before, after);
  expect(direct).toBe(expected);
  expect(next).toBe(direct);
});

test.each([
  ["a table cell", withStories(cell(p("66666666", twoFields)) + tail, p("44444444", run("h")), p("55555555", run("n"))), "66666666", undefined],
  ["a header", withStories(tail, p("44444444", twoFields), p("55555555", run("n"))), "44444444", "word/header1.xml"],
  ["a footnote", withStories(p("22222222", `${run("tail")}${noteRef}`), p("44444444", run("h")), p("55555555", twoFields)), "55555555", "word/footnotes.xml"],
])("a field's child in %s stays where it is when its embed goes", async (_, bytes, paraId, path) => {
  const { next, direct } = await landed(bytes, () => {}, backspaceField("REF a"), (out) => view(out, paraId, path));
  expect(direct).toBe(orphaned);
  expect(next).toBe(direct);
});

test("a field's child stays where it is when one user types in it while another deletes its embed", async () => {
  const bytes = paragraph(`${refField("20", "a")}${run(" mid ")}${refField("30", "b")}`);
  const a = await open(bytes);
  edit(a, "22222222", "x");
  const captured = a.encodeState();
  const exported = await publish(bytes, captured);
  const b = await open(bytes, captured);
  typeInText("20")(a);
  backspaceField("REF a")(b);
  a.applyUpdate(b.encodeStateAsUpdate(a.encodeStateVector()));
  const latest = a.encodeState();
  a.destroy();
  b.destroy();
  const { state } = await rebaseOffice(bytes, checkpoint(bytes, captured), checkpoint(bytes, latest), exported);
  const rebased = await open(exported, state);
  const next = view(await publish(exported, rebased.encodeState()));
  rebased.destroy();
  expect(view(await publish(bytes, latest))).toBe("a H(2Z0) mid [«REF b \\h»|H(30)] b");
  expect(next).toBe("a H(2Z0) mid [«REF b \\h»|H(30)] b");
});

test.each([
  ["Accept All", paragraph(field(`${link(run("20"))}${ins(run("26"))}`)), "20", resolveAll("accept"), "a [«DATE»|H(2Z0)26] b"],
  ["Reject All", paragraph(field(`${link(run("20"))}${ins(run("26"))}`)), "20", resolveAll("reject"), "a [«DATE»|H(2Z0)] b"],
  [
    "Reject All of a TOC",
    docx(p("11111111", `${char("begin")}${instr(" TOC \\o \\h ")}${char("separate")}${link(run("Intro 1"))}${del(deleted("Old"))}`) + p("33333333", `${run("Entry2 2")}${char("end")}`) + tail),
    "Intro",
    resolveAll("reject"),
    "[«TOC \\o \\h»|H(IZntro 1)Old]",
  ],
] as const)("%s after a capture lands exactly when a projected child was edited before it", async (_, bytes, text, after, expected) => {
  const { next, direct } = await landed(bytes, typeIn(text), after);
  expect(direct).toBe(expected);
  expect(next).toBe(direct);
});

const hyperlinkTo = (url: string) =>
  `<Relationship Id="rId30" Type="${REL}/hyperlink" Target="${url}" TargetMode="External"/>`;
const partLink = field(`${run("20")}${ins(`<w:hyperlink r:id="rId30">${run("26")}</w:hyperlink>`)}`);
test.each([
  ["a header", withStories(tail, p("44444444", partLink), p("55555555", run("n")), { header: hyperlinkTo("https://part.example/"), footnotes: "" }), "44444444", "word/header1.xml"],
  [
    "a footnote",
    withStories(p("22222222", `${run("tail")}${noteRef}`), p("44444444", run("h")), p("55555555", partLink), { header: "", footnotes: hyperlinkTo("https://part.example/") }),
    "55555555",
    "word/footnotes.xml",
  ],
])("a link Accept All uncovers in %s resolves with its part's relationships", async (_, source, paraId, path) => {
  const parts = unzipContainer(source);
  parts["word/_rels/document.xml.rels"] = new TextEncoder().encode(
    new TextDecoder().decode(parts["word/_rels/document.xml.rels"]).replace("</Relationships>", `${hyperlinkTo("https://document.example/")}</Relationships>`)
  );
  const bytes = rezipContainer(parts);
  const session = await open(bytes);
  resolveAll("accept")(session);
  const { story } = textAt(session, "26");
  const hrefs = session.storySegments(story).flatMap((segment) => {
    const link = segment.attributes.hyperlink as { href?: string } | undefined;
    return link ? [link.href] : [];
  });
  session.destroy();
  expect(hrefs).toEqual(["https://part.example/"]);
  const { next, direct } = await landed(bytes, resolveAll("accept"), typeInText("26"), (out) => view(out, paraId, path));
  expect(direct).toBe("[«DATE»|20H(2Z6)]");
  expect(next).toBe(direct);
});

/** A link to its own bookmark: links out of a field merge only when they share a target. */
const linkTo = (text: string) => `<w:hyperlink w:anchor="${text}">${run(text)}</w:hyperlink>`;
const joinNext = (paraId: string): Edit => (session) => {
  const { length } = session.paragraphSpans("body").find((span) => span.paraId === paraId)!;
  session.deleteAt({ story: "body", paraId, offset: length }, "forward");
};
/** A paragraph opening with `REF a` over link AA, then one opening with `REF b` over `second`. */
const twoParagraphs = (second: string) =>
  docx(p("11111111", refField("AA", "a")) + p("33333333", field(second, " REF b \\h ")) + tail);
/** Publishes `bytes` after `edits` in one session, without a rebase. */
async function directly(bytes: Uint8Array, ...edits: Edit[]) {
  const session = await open(bytes);
  for (const change of edits) change(session);
  const out = await publish(bytes, session.encodeState());
  session.destroy();
  return view(out);
}

test.each([
  ["one link", field(link(run("BB")), " REF b \\h "), "BB"],
  ["two links", field(`${linkTo("AA")}${linkTo("BB")}`, " REF b \\h "), "AA"],
])("text typed at a link's end before the capture refuses a rebase that deletes its field (%s)", async (_, xml, text) => {
  const landing = landed(paragraph(xml), typeInText(text, text.length), backspaceField("REF b"));
  await expect(landing).rejects.toBeInstanceOf(RebaseError);
});

test("text typed at a link's end before the capture lands when the field is left alone", async () => {
  const bytes = paragraph(field(link(run("BB")), " REF b \\h "));
  const { next, direct } = await landed(bytes, typeInText("BB", 2), (session) => edit(session, "22222222", "y"));
  expect(direct).toBe("a [«REF b \\h»|H(BB)Z] b");
  expect(next).toBe(direct);
});

test.each(["accept", "reject"] as const)(
  "%s all keeps where it stands the link of a field deleted after a join, which shares the next field's number",
  async (mode) => {
    const out = await directly(
      twoParagraphs(`${run("x")}${link(run("BB"))}${ins(run("26"))}`),
      joinNext("11111111"),
      backspaceField("REF a"),
      resolveAll(mode)
    );
    expect(out).toBe(mode === "accept" ? "H(AA)[«REF b \\h»|xH(BB)26]" : "H(AA)[«REF b \\h»|xH(BB)]");
  }
);

test.each([
  [
    "Accept All renumbers the next field",
    paragraph(`${refField("AA", "a")}${field(`${link(run("BB"))}${ins(run("26"))}`, " REF b \\h ")}`),
    [backspaceField("REF a"), resolveAll("accept")],
    "a H(AA)[«REF b \\h»|H(BB)26] b",
  ],
  [
    "Accept All renumbers the next field, its links kept in order",
    paragraph(`${field(`${linkTo("AA")}${linkTo("CC")}`, " REF a \\h ")}${field(`${link(run("BB"))}${ins(run("26"))}`, " REF b \\h ")}`),
    [backspaceField("REF a"), resolveAll("accept")],
    "a H(AA)H(CC)[«REF b \\h»|H(BB)26] b",
  ],
  ["a join gives both fields one number", twoParagraphs(link(run("BB"))), [joinNext("11111111"), backspaceField("REF a")], "H(AA)[«REF b \\h»|H(BB)]"],
] as const)("a deleted field's link joins no other field when %s", async (_, bytes, edits, expected) => {
  expect(await directly(bytes, ...edits)).toBe(expected);
});

test.each([
  ["a REF field's link", paragraph(field(linkTo("BB"), " REF b \\h ")), "BB", 1, "a [«REF b \\h»|H(BB)] b"],
  [
    "a TOC's first entry",
    docx(p("11111111", `${char("begin")}${instr(" TOC \\o \\h ")}${char("separate")}${link(run("Introduction 1"))}`) + p("33333333", `${link(run("Entry2 2"))}${char("end")}`) + tail),
    "Introduction",
    2,
    "[«TOC \\o \\h»|H(Introduction 1)]",
  ],
] as const)("Clear formatting on part of %s keeps it in its field", async (_, bytes, text, from, expected) => {
  const clear: Edit = (session) => {
    const at = textAt(session, text);
    session.clearFormatting({
      story: "body",
      start: { paraId: at.paraId, offset: at.offset + from },
      end: { paraId: at.paraId, offset: at.offset + from + 2 },
    });
  };
  expect(await directly(bytes, clear)).toBe(expected);
});

test("undoing a field embed's deletion gives the field its link back", async () => {
  const bytes = paragraph(`${refField("AA", "a")}${run(" mid ")}${refField("BB", "b")}`);
  const session = await open(bytes);
  session.beginUndoCapture();
  backspaceField("REF a")(session);
  session.undo();
  const out = await publish(bytes, session.encodeState());
  session.destroy();
  expect(view(out)).toBe("a [«REF a \\h»|H(AA)] mid [«REF b \\h»|H(BB)] b");
});

const threeLinks = paragraph(field(`${linkTo("AA")}${linkTo("BB")}${linkTo("CC")}`));
const deleteText = (text: string): Edit => (session) => {
  const { story, paraId, offset } = textAt(session, text);
  session.deleteRange({ story, start: { paraId, offset }, end: { paraId, offset: offset + text.length } });
};
test.each([
  [
    "a split between two links before the capture and a join after",
    (session: YrsSession) => session.splitParagraph({ story: "body", paraId: "11111111", offset: textAt(session, "BB").offset }),
    joinNext("11111111"),
  ],
  ["a link deleted before the capture and typing at the end of the next after", deleteText("AA"), typeInText("BB", 2)],
  [
    "a link deleted before the capture and the next retyped after",
    deleteText("AA"),
    (session: YrsSession) => {
      typeInText("BB", 2)(session);
      deleteText("BB")(session);
    },
  ],
] as const)("a rebase refuses to land two of a field's links in one result slot: %s", async (_, before, after) => {
  await expect(landed(threeLinks, before, after)).rejects.toBeInstanceOf(RebaseError);
});

test("a rebase fails as a RebaseError when a child's field stops projecting in the export", async () => {
  const bytes = paragraph(field(`${link(run("AA"))}${link(run("BB"))}`));
  const split: Edit = (session) =>
    void session.splitParagraph({ story: "body", paraId: "11111111", offset: fieldAt(session, "DATE").offset });
  await expect(landed(bytes, split, joinNext("11111111"))).rejects.toBeInstanceOf(RebaseError);
});

// The refusal a UAT journey reproduces: the export writes a run holding
// text and a tab as two runs, so its seed numbers the TOC field otherwise
// than the capture, and typing in the TOC's first entry after the capture
// carries the capture's number.
test("typing in a TOC's first entry after the capture refuses when a text and tab run leads its paragraph", async () => {
  const bytes = docx(
    p("11111111", `<w:r><w:t>Contents</w:t><w:tab/></w:r>${char("begin")}${instr(" TOC \\o \\h ")}${char("separate")}${link(run("Intro 1"))}`) +
      p("33333333", `${link(run("Entry2 2"))}${char("end")}`) +
      tail
  );
  await expect(landed(bytes, () => {}, typeIn("Intro"))).rejects.toBeInstanceOf(RebaseError);
});
