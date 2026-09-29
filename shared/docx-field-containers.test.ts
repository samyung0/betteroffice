import { beforeAll, expect, test } from "bun:test";
import { createHash } from "node:crypto";
import { createYrsSession, type YrsSession } from "../packages/docx/src/yrs";
import { rezipContainer, unzipContainer } from "../packages/docx/src/wasm/opc";
import { exportOffice, rebaseOffice, seedOffice } from "./office-checkpoint";

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
const ins = (xml: string) =>
  `<w:ins w:id="90" w:author="A" w:date="2026-09-01T00:00:00Z">${xml}</w:ins>`;
const del = (xml: string) =>
  `<w:del w:id="91" w:author="A" w:date="2026-09-01T00:00:00Z">${xml}</w:del>`;
const fs = (xml: string) => `<w:fldSimple w:instr=" DATE ">${xml}</w:fldSimple>`;
const link = (xml: string) =>
  `<w:hyperlink w:anchor="target">${xml}</w:hyperlink>`;
const sdt = (xml: string) =>
  `<w:sdt><w:sdtPr><w:id w:val="8"/></w:sdtPr><w:sdtContent>${xml}</w:sdtContent></w:sdt>`;
const bm = (xml: string) =>
  `<w:bookmarkStart w:id="5" w:name="mark"/>${xml}<w:bookmarkEnd w:id="5"/>`;
const field = (result: string) =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DATE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${result}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;

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

const documentXml = (bytes: Uint8Array) =>
  new TextDecoder().decode(unzipContainer(bytes)["word/document.xml"]);

/**
 * The first paragraph's markup as text: F<> simple field, H() link, +{} and
 * -{} tracked changes, S{} content control, {bm bm} bookmark.
 */
function view(bytes: Uint8Array): string {
  const xml = documentXml(bytes);
  const first = xml.slice(xml.indexOf("<w:body>"), xml.indexOf("</w:p>"));
  return [
    ...first.matchAll(
      /<(\/?)w:(t|delText|ins|del|fldSimple|hyperlink|bookmarkStart|bookmarkEnd|sdt)\b([^>]*?)(\/?)>([^<]*)/g
    ),
  ]
    .map(([, close, tag, , , text]) => {
      if (tag === "t" || tag === "delText") return close ? "" : text;
      if (tag === "bookmarkStart") return "{bm";
      if (tag === "bookmarkEnd") return "bm}";
      if (close) return tag === "fldSimple" ? ">" : tag === "hyperlink" ? ")" : "}";
      return { fldSimple: "F<", hyperlink: "H(", ins: "+{", del: "-{", sdt: "S{" }[tag];
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
  [
    "a bookmark inside w:fldSimple",
    (wrap) => fs(bm(wrap(run("2026")))),
    "a F<{bm2026bm}> b",
    "2026",
  ],
  [
    "w:fldSimple inside a hyperlink",
    (wrap) => link(fs(wrap(run("2026")))),
    "a H(F<2026>) b",
    "2026",
  ],
  [
    "w:fldSimple inside w:ins",
    (wrap) => ins(fs(wrap(run("2026")))),
    "a +{F<2026>} b",
    "2026",
  ],
  [
    "w:ins inside w:fldSimple",
    (wrap) => fs(ins(wrap(run("2026")))),
    "a F<+{2026}> b",
    "2026",
  ],
  [
    "w:sdt inside w:fldSimple",
    (wrap) => fs(sdt(wrap(run("2026")))),
    "a F<S{2026}> b",
    "2026",
  ],
  [
    "w:ins inside a complex field result",
    (wrap) => field(`${run("20")}${ins(wrap(run("26")))}`),
    "a 20+{26} b",
    "2026",
  ],
  [
    "w:del inside a complex field result",
    (wrap) => field(`${run("20")}${del(wrap(deleted("26")))}`),
    "a 20-{26} b",
    "20",
  ],
  [
    "w:sdt inside a complex field result",
    (wrap) => field(`${run("20")}${sdt(wrap(run("26")))}`),
    "a 20S{26} b",
    "2026",
  ],
];
const plain = (xml: string) => xml;

test.each(cases)(
  "%s round-trips through three publications, its paragraph untouched",
  async (_, build, expected, shown) => {
    let bytes = paragraph(build(plain));
    let saved = "";
    for (let publication = 0; publication < 3; publication += 1) {
      const session = await open(bytes);
      const fields = session
        .storySegments("body")
        .flatMap((segment) =>
          segment.kind === "embed" ? [segment.payload.displayText] : []
        );
      expect(fields).toEqual([shown]);
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

test.each(cases)(
  "%s rebases an edit to its paragraph made after the capture",
  async (_, build, expected) => {
    const bytes = paragraph(build(plain));
    const session = await open(bytes);
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
    expect(view(next)).toBe(`y${expected}`);
  }
);

test.each(cases)(
  "a Word comment range inside %s holds the whole field across publications",
  async (_, build, expected) => {
    const wrap = (xml: string) =>
      `<w:commentRangeStart w:id="5"/>${xml}<w:commentRangeEnd w:id="5"/>`;
    let bytes = docx(
      p("11111111", `${run("a ")}${build(wrap)}${ref(5)}${run(" b")}`) + tail
    );
    for (let publication = 0; publication < 3; publication += 1) {
      const session = await open(bytes);
      expect(covered(session)).toBe("[field]");
      bytes = await publish(bytes, session.encodeState());
      session.destroy();
      expect(marks(bytes)).toEqual(["E5", "R5", "S5"]);
      expect(view(bytes)).toBe(expected);
    }
  }
);
