import { beforeAll, expect, test } from "bun:test";
import { createHash } from "node:crypto";
import { createYrsSession, type YrsSession } from "../packages/docx/src/yrs";
import { rezipContainer, unzipContainer } from "../packages/docx/src/wasm/opc";
import { exportOffice, rebaseOffice, seedOffice } from "./office-checkpoint";
import { RebaseError } from "./office-rebase";
import { E as commentEnd, S as commentStart, STORY, docx as matrixDocx, fieldAt as matrixFieldAt, len as matrixLen, locate as matrixLocate, orders, parity, partXml, prime, ref as commentRef, runRow, units as matrixUnits, sig, type Edit as MatrixEdit, type Where } from "./matrix/lib";

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
  ["a complex field inside a hyperlink", (wrap) => link(field(wrap(run("2026")), " PAGEREF target \\h ")), "a H([«PAGEREF target \\h»|2026]) b", "a 2026 b"],
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

test.each([
  ["a whole inserted field", ins(field(run("kept"))), ["<w:ins ", "<w:fldChar "]],
  ["a whole deleted field", del(field(deleted("kept"))), ["<w:del ", "<w:fldChar "]],
  ["a content control in a link", link(sdt(run("kept"))), ["<w:hyperlink ", "<w:sdt>"]],
  ["a revision in a link", link(ins(run("kept"))), ["<w:hyperlink ", "<w:ins "]],
  ["a link in a revision", ins(link(run("kept"))), ["<w:hyperlink ", "<w:ins "]],
  ["a revision in a content control", sdt(ins(run("kept"))), ["<w:sdt>", "<w:ins "]],
  ["nested insertions and deletions", ins(del(deleted("kept"))), ["<w:ins ", "<w:del "]],
] as const)("%s keeps its content and wrappers through publications", async (_, content, tags) => {
  let bytes = paragraph(content);
  for (let publication = 0; publication < 3; publication += 1) {
    const session = await open(bytes);
    edit(session, "22222222", "x");
    bytes = await publish(bytes, session.encodeState());
    session.destroy();
    const saved = documentXml(bytes);
    expect(saved).toContain("kept");
    for (const tag of tags) expect(saved).toContain(tag);
    if (tags.some((tag) => tag === "<w:fldChar ")) expect([...saved.matchAll(/<w:fldChar /g)]).toHaveLength(3);
  }
});

test.each([
  ["a content control", (xml: string) => sdt(xml)],
  ["a link in a field", (xml: string) => field(link(xml))],
  ["a control in a field", (xml: string) => field(sdt(xml))],
  ["a field in a control", (xml: string) => sdt(field(xml))],
])("Accept and Reject All resolve revisions inside %s", async (_, wrap) => {
  for (const mode of ["accept", "reject"] as const) {
    let bytes = paragraph(wrap(ins(run("new")) + del(deleted("old")) + ins(del(deleted("both")))));
    const session = await open(bytes);
    expect(session.hasFieldChanges()).toBe(true);
    resolveAll(mode)(session);
    expect(session.hasFieldChanges()).toBe(false);
    bytes = await publish(bytes, session.encodeState());
    session.destroy();
    for (let publication = 0; publication < 3; publication += 1) {
      const xml = documentXml(bytes);
      expect(xml).not.toMatch(/<w:(?:ins|del|moveFrom|moveTo)\b/);
      expect(xml).toContain(mode === "accept" ? "new" : "old");
      expect(xml).not.toContain(mode === "accept" ? "old" : "new");
      expect(xml).not.toContain("both");
      const reopened = await open(bytes);
      bytes = await publish(bytes, reopened.encodeState());
      reopened.destroy();
    }
  }
});

test.each([
  ["typing before the range", (session: YrsSession) => session.insertText({ story: "body", paraId: "11111111", offset: 0 }, "z")],
  ["Enter inside the range", (session: YrsSession) => session.splitParagraph({ story: "body", paraId: "11111111", offset: 3 })],
  ["joining the next paragraph", (session: YrsSession) => session.mergeParagraphs("body", "11111111", "forward")],
])("bookmarks keep one range after %s and publications", async (_, edit) => {
  let bytes = docx(p("11111111", run("A") + bm(run("BCD")) + run("E")) + p("22222222", bm(run("FG")).replaceAll('w:id="5"', 'w:id="6"').replace('w:name="mark"', 'w:name="other"')));
  const session = await open(bytes);
  edit(session);
  bytes = await publish(bytes, session.encodeState());
  session.destroy();
  for (let publication = 0; publication < 3; publication += 1) {
    const xml = documentXml(bytes);
    for (const [id, text] of [[5, "BCD"], [6, "FG"]] as const) {
      expect([...xml.matchAll(new RegExp(`<w:bookmarkStart w:id="${id}"`, "g"))]).toHaveLength(1);
      expect([...xml.matchAll(new RegExp(`<w:bookmarkEnd w:id="${id}"`, "g"))]).toHaveLength(1);
      const range = xml.match(new RegExp(`<w:bookmarkStart w:id="${id}"[^>]*>([\\s\\S]*?)<w:bookmarkEnd w:id="${id}"`))![1]!;
      expect([...range.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((match) => match[1]).join("")).toBe(text);
    }
    const reopened = await open(bytes);
    bytes = await publish(bytes, reopened.encodeState());
    reopened.destroy();
  }
});

test.each([["a link", link], ["a content control", sdt], ["a field link", (xml: string) => field(link(xml))]])("bookmarks inside %s stay in that container", async (_, wrap) => {
  let bytes = paragraph(wrap(run("A") + bm(run("BCD")) + run("E")));
  for (let publication = 0; publication < 3; publication += 1) {
    const session = await open(bytes);
    edit(session, "22222222", "z");
    bytes = await publish(bytes, session.encodeState());
    session.destroy();
    const xml = documentXml(bytes);
    expect(xml).toMatch(/<(?:w:hyperlink|w:sdtContent)[^>]*>[\s\S]*?<w:bookmarkStart w:id="5"[\s\S]*?<w:t[^>]*>BCD<\/w:t>[\s\S]*?<w:bookmarkEnd w:id="5"/);
  }
});

test("accepting a paragraph-mark deletion preserves bookmarks from both paragraphs", async () => {
  let bytes = docx(p("11111111", '<w:pPr><w:rPr><w:del w:id="2" w:author="A"/></w:rPr></w:pPr>' + bm(run("AB"))) + p("22222222", bm(run("CD")).replaceAll('w:id="5"', 'w:id="6"')));
  const session = await open(bytes);
  session.acceptChange({ all: true });
  bytes = await publish(bytes, session.encodeState());
  session.destroy();
  const xml = documentXml(bytes);
  expect([...xml.matchAll(/<w:bookmarkStart\b/g)]).toHaveLength(2);
  expect([...xml.matchAll(/<w:bookmarkEnd\b/g)]).toHaveLength(2);
  expect(view(bytes, "22222222")).toBe("{bmABbm}{bmCDbm}");
});

test("Accept All exposes a field link's bookmarks as anchored boundaries", async () => {
  let bytes = paragraph(field(ins(link(run("A") + bm(run("BCD")) + run("E")))));
  const session = await open(bytes);
  session.acceptChange({ all: true });
  expect(session.storySegments("body").some((part) => part.kind === "embed" && part.embedKind === "bookmark")).toBe(false);
  bytes = await publish(bytes, session.encodeState());
  session.destroy();
  for (let publication = 0; publication < 3; publication += 1) {
    expect(view(bytes, "11111111")).toContain("H(A{bmBCDbm}E)");
    const reopened = await open(bytes);
    edit(reopened, "22222222", "z");
    bytes = await publish(bytes, reopened.encodeState());
    reopened.destroy();
  }
});

test("a field spanning paragraphs closes at its authored position through edits and publications", async () => {
  let bytes = docx(p("11111111", char("begin") + instr(" TOC ") + char("separate") + link(run("first"))) + p("22222222", link(run("second")) + char("end") + run("after")) + p("33333333", run("tail")));
  for (let publication = 0; publication < 3; publication += 1) {
    const session = await open(bytes);
    session.insertText({ story: "body", paraId: "22222222", offset: 0 }, "x");
    bytes = await publish(bytes, session.encodeState());
    session.destroy();
    expect(view(bytes, "11111111")).toBe("[«TOC»|H(first)");
    expect(view(bytes, "22222222")).toBe(`${"x".repeat(publication + 1)}H(second)]after`);
    expect([...documentXml(bytes).matchAll(/w:fldCharType="end"/g)]).toHaveLength(1);
  }
});

test("a continued field keeps its closing anchor when the captured export renumbers its owner", async () => {
  const bytes = docx(p("11111111", run("prefix") + char("begin") + instr(" TOC ") + char("separate") + link(run("first"))) + p("33333333", run("second") + char("end") + run("after")) + tail);
  const { next, direct } = await landed(bytes, (session) => session.deleteRange({
    story: "body", start: { paraId: "11111111", offset: 0 }, end: { paraId: "11111111", offset: 6 },
  }), (session) => session.insertText({ story: "body", paraId: "33333333", offset: 0 }, "Z"),
  (saved) => `${view(saved, "11111111")}|${view(saved, "33333333")}`);
  expect(next).toBe(direct);
  expect(next).toBe("[«TOC»|H(first)|Zsecond]after");
});

test("a field's separate and end characters stay in their later paragraphs", async () => {
  let bytes = docx(p("11111111", char("begin") + instr(" TOC ")) + p("33333333", char("separate") + run("entry")) + p("44444444", char("end") + run("after")) + tail);
  for (let publication = 0; publication < 3; publication++) {
    const session = await open(bytes);
    edit(session, "22222222", "x");
    bytes = await publish(bytes, session.encodeState());
    session.destroy();
    expect(view(bytes, "11111111")).toBe("[«TOC»");
    expect(view(bytes, "33333333")).toBe("|entry");
    expect(view(bytes, "44444444")).toBe("]after");
  }
});

for (const [name, opening, closing, firstMarkers, lastMarkers] of [
  ["nested code", char("begin") + instr(" IF ") + char("begin") + instr(" PAGE "), char("separate") + run("12") + char("end") + char("separate") + run("yes") + char("end"), ["begin", "begin"], ["separate", "end", "separate", "end"]],
  ["nested result", char("begin") + instr(" IF ") + char("separate") + char("begin") + instr(" PAGE "), char("separate") + run("12") + char("end") + char("end"), ["begin", "separate", "begin"], ["separate", "end", "end"]],
  ["closed child", char("begin") + instr(" IF ") + field(run("1"), " PAGE "), char("separate") + run("yes") + char("end"), ["begin", "begin", "separate", "end"], ["separate", "end"]],
] as const) {
  for (const [container, wrap] of [["paragraph", (xml: string) => xml], ["link", link], ["control", sdt]] as const) {
    test(`continued ${name} fields in a ${container} keep their own markers through publications`, async () => {
      let bytes = docx(p("11111111", wrap(opening)) + p("33333333", closing) + tail);
      for (let publication = 0; publication < 3; publication++) {
        const session = await open(bytes);
        edit(session, "22222222", "x");
        bytes = await publish(bytes, session.encodeState());
        session.destroy();
        const paragraphs = [...documentXml(bytes).matchAll(/<w:p\b[^>]*>([\s\S]*?)<\/w:p>/g)];
        const markers = (xml: string) => [...xml.matchAll(/w:fldCharType="([^"]+)"/g)].map((match) => match[1]);
        expect(markers(paragraphs[0][1])).toEqual(firstMarkers);
        expect(markers(paragraphs[1][1])).toEqual(lastMarkers);
        expect(documentXml(bytes)).toContain("PAGE");
      }
    });
    if (name === "closed child") continue;
    test(`continued ${name} fields in a ${container} rebase their renumbered markers`, async () => {
      const bytes = docx(p("11111111", run("prefix") + wrap(opening)) + p("33333333", ins(closing)) + tail);
      const { next, direct } = await landed(bytes, (session) => session.deleteRange({
        story: "body", start: { paraId: "11111111", offset: 0 }, end: { paraId: "11111111", offset: 6 },
      }), (session) => {
        resolveAll("accept")(session);
        session.insertText({ story: "body", paraId: "33333333", offset: 1 }, "Z");
      }, (saved) => `${view(saved, "11111111")}|${view(saved, "33333333")}`);
      expect(next).toBe(direct);
      expect(next).toContain("|1Z2]");
    });
  }
}

test("a continued field's cached paragraphs do not restore a deleted nested field's markers", async () => {
  let bytes = docx(p("11111111", char("begin") + instr(" IF ") + char("separate")) + p("33333333", char("begin") + instr(" PAGE ")) + p("44444444", char("separate") + run("1") + char("end") + char("end")) + tail);
  const session = await open(bytes);
  session.deleteRange({ story: "body", start: { paraId: "33333333", offset: 0 }, end: { paraId: "33333333", offset: 1 } });
  bytes = await publish(bytes, session.encodeState());
  session.destroy();
  for (let publication = 0; publication < 3; publication++) {
    expect(view(bytes, "44444444")).toBe("1]");
    expect([...documentXml(bytes).matchAll(/w:fldCharType="begin"/g)]).toHaveLength(1);
    const reopened = await open(bytes);
    edit(reopened, "22222222", "x");
    bytes = await publish(bytes, reopened.encodeState());
    reopened.destroy();
  }
});

test("a text-box field does not shift a host paragraph's continued field", async () => {
  const drawing = `<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="41" name="Text Box 41"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent><w:p>${field(run("1"), " PAGE ")}</w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  const parts = unzipContainer(docx(p("11111111", drawing + char("begin") + instr(" IF ")) + p("33333333", char("separate") + run("yes") + char("end")) + tail));
  parts["word/document.xml"] = new TextEncoder().encode(new TextDecoder().decode(parts["word/document.xml"]).replace("<w:document ", '<w:document xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '));
  let bytes = rezipContainer(parts);
  for (let publication = 0; publication < 3; publication++) {
    const session = await open(bytes);
    if (publication) edit(session, "22222222", "x");
    bytes = await publish(bytes, session.encodeState());
    session.destroy();
    const xml = documentXml(bytes);
    const first = xml.slice(xml.indexOf('w14:paraId="11111111"'), xml.indexOf('w14:paraId="33333333"'));
    const second = xml.slice(xml.indexOf('w14:paraId="33333333"'), xml.indexOf('w14:paraId="22222222"'));
    const markers = (value: string) => [...value.matchAll(/w:fldCharType="([^"]+)"/g)].map((match) => match[1]);
    expect(markers(first)).toEqual(["begin", "separate", "end", "begin"]);
    expect(markers(second)).toEqual(["separate", "end"]);
  }
});

for (const [name, wrap] of [["link", link], ["insertion", ins], ["control", sdt]] as const) {
  test(`a continued field inside a ${name} keeps its markers when its captured owner is renumbered`, async () => {
    const bytes = docx(p("11111111", run("prefix") + wrap(char("begin") + instr(" TOC "))) + p("33333333", ins(char("separate") + run("result") + char("end"))) + tail);
    const { next, direct } = await landed(bytes, (session) => session.deleteRange({
      story: "body", start: { paraId: "11111111", offset: 0 }, end: { paraId: "11111111", offset: 6 },
    }), (session) => {
      resolveAll("accept")(session);
      session.insertText({ story: "body", paraId: "33333333", offset: 1 }, "Z");
    },
    (saved) => `${view(saved, "11111111")}|${view(saved, "33333333")}`);
    expect(next).toBe(direct);
    expect(next).toContain("|rZesult]");
  });
  test.each([false, true])(`a continued field with nested ${name} characters stays open across paragraphs, nested opener: %s`, async (nestedOpener) => {
    const opening = char("begin") + instr(" TOC ");
    const closing = run("prefix") + char("separate") + run("result") + char("end") + run("suffix");
    let bytes = docx(p("11111111", nestedOpener ? wrap(opening) : opening) + p("33333333", nestedOpener ? closing : wrap(closing)) + tail);
    for (let publication = 0; publication < 3; publication++) {
      const session = await open(bytes);
      edit(session, "22222222", "x");
      bytes = await publish(bytes, session.encodeState());
      session.destroy();
      const paragraphs = [...documentXml(bytes).matchAll(/<w:p\b[^>]*>([\s\S]*?)<\/w:p>/g)];
      expect([...paragraphs[0][1].matchAll(/w:fldCharType="([^"]+)"/g)].map((match) => match[1])).toEqual(["begin"]);
      expect([...paragraphs[1][1].matchAll(/w:fldCharType="([^"]+)"/g)].map((match) => match[1])).toEqual(["separate", "end"]);
      expect(paragraphs[1][1]).toContain("result");
      const shown = view(bytes, "33333333");
      expect(shown.replace(/H\(|\)|S\{|\+\{|\}/g, "")).toBe("prefix|result]suffix");
      if (!nestedOpener) {
        const tag = { link: "hyperlink", insertion: "ins", control: "sdtContent" }[name];
        const containers = [...paragraphs[1][1].matchAll(new RegExp(`<w:${tag}\\b[^>]*>([\\s\\S]*?)</w:${tag}>`, "g"))];
        for (const type of ["separate", "end"]) {
          expect(containers.some((container) => container[1].includes(`w:fldCharType="${type}"`))).toBe(true);
        }
      }
    }
    const resolved = await open(bytes);
    resolveAll("accept")(resolved);
    bytes = await publish(bytes, resolved.encodeState());
    resolved.destroy();
    expect(documentXml(bytes)).not.toMatch(/<w:ins\b/);
    expect(view(bytes, "33333333").replace(/H\(|\)|S\{|\}/g, "")).toBe("prefix|result]suffix");
  });
}

test("continued fields without paragraph ids have distinct anchors across stories", async () => {
  const content = (text: string) => `<w:p>${char("begin")}${instr(" TOC ")}${char("separate")}${run(text)}</w:p><w:p>${run("last")}${char("end")}</w:p>`;
  let bytes = withStories(content("body") + tail, content("header"), p("55555555", run("note")));
  for (let publication = 0; publication < 3; publication++) {
    const session = await open(bytes);
    const ids = ["body", "hf:rId10"].flatMap((story) => session.storySegments(story).flatMap((segment) =>
      segment.kind === "embed" && typeof segment.payload.continuationId === "string" ? [segment.payload.continuationId] : []
    ));
    expect(new Set(ids).size).toBe(2);
    edit(session, "22222222", "x");
    bytes = await publish(bytes, session.encodeState());
    session.destroy();
    for (const part of ["word/document.xml", "word/header1.xml"]) {
      const paragraphs = [...documentXml(bytes, part).matchAll(/<w:p\b[^>]*>([\s\S]*?)<\/w:p>/g)];
      expect(paragraphs[0][1]).not.toContain('w:fldCharType="end"');
      expect(paragraphs[1][1]).toContain('w:fldCharType="end"');
    }
  }
});

test.each(["accept", "reject"] as const)("%s all resolves a formatting-only revision", async (mode) => {
  const properties = '<w:rPr><w:b/><w:rPrChange w:id="4" w:author="A"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr>';
  let bytes = paragraph(`<w:r>${properties}<w:t>changed</w:t></w:r>`);
  const session = await open(bytes);
  resolveAll(mode)(session);
  bytes = await publish(bytes, session.encodeState());
  session.destroy();
  for (let publication = 0; publication < 3; publication += 1) {
    const xml = documentXml(bytes);
    expect(xml).not.toContain("w:rPrChange");
    expect(xml).toContain(mode === "accept" ? "<w:b/>" : "<w:i/>");
    expect(xml).not.toContain(mode === "accept" ? "<w:i/>" : "<w:b/>");
    const reopened = await open(bytes);
    bytes = await publish(bytes, reopened.encodeState());
    reopened.destroy();
  }
});

test.each(["accept", "reject"] as const)("%s all resolves an empty run's formatting revision", async (mode) => {
  const bytes = paragraph('<w:r><w:rPr><w:b/><w:rPrChange w:id="4" w:author="A"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t/></w:r>');
  const session = await open(bytes);
  expect(session.listRevisions().some((change) => change.kind === "rPrChange")).toBe(true);
  resolveAll(mode)(session);
  expect(session.listRevisions()).toHaveLength(0);
  const saved = await publish(bytes, session.encodeState());
  session.destroy();
  expect(documentXml(saved)).not.toContain("w:rPrChange");
});

test.each(["accept", "reject"] as const)("%s all resolves formatting revisions inside nested containers", async (mode) => {
  const changed = '<w:r><w:rPr><w:b/><w:rPrChange w:id="4" w:author="A"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>changed</w:t></w:r>';
  for (const wrap of [link, sdt, fs, field, (xml: string) => field(link(sdt(xml)))]) {
    const bytes = paragraph(wrap(changed));
    const session = await open(bytes);
    resolveAll(mode)(session);
    const saved = await publish(bytes, session.encodeState());
    session.destroy();
    const xml = documentXml(saved);
    expect(xml).not.toContain("w:rPrChange");
    expect(xml).toContain(mode === "accept" ? "<w:b/>" : "<w:i/>");
    expect(xml).not.toContain(mode === "accept" ? "<w:i/>" : "<w:b/>");
  }
});

test.each(["accept", "reject"] as const)("%s all resolves move wrappers and their range markers", async (mode) => {
  const moved = '<w:moveFromRangeStart w:id="80" w:name="m"/>' + moveFrom(deleted("old")) + '<w:moveFromRangeEnd w:id="80"/><w:moveToRangeStart w:id="81" w:name="m"/>' + moveTo(run("new")) + '<w:moveToRangeEnd w:id="81"/>';
  for (const wrap of [(xml: string) => xml, field, (xml: string) => field(sdt(xml))]) {
    const bytes = paragraph(wrap(moved));
    const session = await open(bytes);
    resolveAll(mode)(session);
    const saved = await publish(bytes, session.encodeState());
    session.destroy();
    const xml = documentXml(saved);
    expect(xml).not.toMatch(/<w:(?:moveFrom|moveTo|ins|del)\b|<w:move(?:From|To)Range(?:Start|End)\b/);
    expect(xml).toContain(mode === "accept" ? "new" : "old");
    expect(xml).not.toContain(mode === "accept" ? "old" : "new");
  }
});

test("text beside field markers stays on its own side across publications", async () => {
  let bytes = paragraph('<w:r><w:t>before</w:t><w:fldChar w:fldCharType="begin"/><w:instrText> DATE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/><w:t>2026</w:t></w:r><w:r><w:t> result</w:t><w:fldChar w:fldCharType="end"/><w:t>after</w:t></w:r>');
  for (let publication = 0; publication < 3; publication += 1) {
    const session = await open(bytes);
    edit(session, "22222222", "x");
    bytes = await publish(bytes, session.encodeState());
    session.destroy();
    expect(view(bytes)).toBe("a before[«DATE»|2026 result]after b");
  }
});

test("a block content control in a table cell keeps its content through edits and publication", async () => {
  const control = sdt(p("33333333", run("kept")));
  let bytes = docx(`<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>${control}</w:tc></w:tr></w:tbl>${tail}`);
  const story = "body:t0:r0c0:sdt0";
  for (let publication = 0; publication < 3; publication += 1) {
    const session = await open(bytes);
    expect(session.paragraphs(story)[0].text).toBe(`${"x".repeat(publication)}kept`);
    edit(session, "22222222", "z");
    const captured = session.encodeState();
    const exported = await publish(bytes, captured);
    session.insertText({ story, paraId: "33333333", offset: 0 }, "x");
    const latest = session.encodeState();
    const result = await rebaseOffice(bytes, checkpoint(bytes, captured), checkpoint(bytes, latest), exported);
    bytes = await publish(exported, result.state);
    session.destroy();
    const xml = documentXml(bytes);
    expect(xml).toMatch(/<w:tc>(?:<w:tcPr>[\s\S]*?<\/w:tcPr>)?<w:sdt>/);
    expect(xml).toContain(`${"x".repeat(publication + 1)}kept`);
  }
});

test("a dirty complex field keeps Word's update flag through publications", async () => {
  let bytes = paragraph(field(run("1"), " PAGE ").replace('w:fldCharType="begin"', 'w:fldCharType="begin" w:dirty="true"'));
  for (let publication = 0; publication < 3; publication += 1) {
    const session = await open(bytes);
    edit(session, "22222222", "x");
    bytes = await publish(bytes, session.encodeState());
    session.destroy();
    expect(documentXml(bytes)).toContain('w:fldCharType="begin" w:dirty="true"');
  }
});

test.each([" DATE ", " REF target \\h "])("typing at a %s field link's end keeps its Word link properties", async (instruction) => {
  const original = '<w:hyperlink w:anchor="target" w:history="1" w:tgtFrame="_blank" w:docLocation="section" w:tooltip="tip">' + run("BB") + '</w:hyperlink>';
  let bytes = paragraph(field(original, instruction));
  for (let publication = 0; publication < 3; publication += 1) {
    const session = await open(bytes);
    if (publication === 0) session.insertText({ story: "body", paraId: "11111111", offset: 4 }, "Z");
    else edit(session, "22222222", "x");
    bytes = await publish(bytes, session.encodeState());
    session.destroy();
    const saved = documentXml(bytes).match(/<w:hyperlink[^>]*>/)![0];
    expect(saved).toContain('w:history="1"');
    expect(saved).toContain('w:tgtFrame="_blank"');
    expect(saved).toContain('w:docLocation="section"');
    expect(saved).toContain('w:anchor="target"');
    expect(saved).toContain('w:tooltip="tip"');
  }
});

test("unbolding a field's first child after capture saves the same field formatting directly and after a rebase", async () => {
  const bold = '<w:hyperlink w:anchor="target"><w:r><w:rPr><w:b/></w:rPr><w:t>AA</w:t></w:r></w:hyperlink>';
  const bytes = paragraph(field(bold + link(run("BB")), " REF target \\h "));
  const session = await open(bytes);
  edit(session, "22222222", "x");
  const captured = session.encodeState();
  const exported = await publish(bytes, captured);
  session.formatRange({ story: "body", start: { paraId: "11111111", offset: 2 }, end: { paraId: "11111111", offset: 4 } }, { bold: false });
  const latest = session.encodeState();
  const direct = await publish(bytes, latest);
  const result = await rebaseOffice(bytes, checkpoint(bytes, captured), checkpoint(bytes, latest), exported);
  let saved = await publish(exported, result.state);
  session.destroy();
  expect(documentXml(saved)).toBe(documentXml(direct));
  for (let publication = 0; publication < 2; publication += 1) {
    const reopened = await open(saved);
    const previous = documentXml(saved);
    saved = await publish(saved, reopened.encodeState());
    reopened.destroy();
    expect(documentXml(saved)).toBe(previous);
  }
});

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
    const result = (saved: Uint8Array) => view(saved).match(/\|(.*)$/)?.[1];
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
        expect(view(saved, "33333333")).toBe("Entry2 2]");
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

test("typing at a field child's end after a captured sibling deletion lands exactly", async () => {
  const bytes = paragraph(field(link(run("AA")).replace('w:anchor="target"', 'w:anchor="a"') + link(run("BB")).replace('w:anchor="target"', 'w:anchor="b"')));
  const { next, direct } = await landed(bytes, (session) => {
    const offset = childAt(session, "AA");
    session.deleteRange({ story: "body", start: { paraId: "11111111", offset }, end: { paraId: "11111111", offset: offset + 2 } });
  }, (session) => session.insertText({ story: "body", paraId: "11111111", offset: childAt(session, "BB") + 2 }, "Z"));
  expect(next).toBe(direct);
  expect(next).toContain("Z");
  expect(next).not.toContain("AA");
});

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
  ["Accept All after a bookmark", `<w:bookmarkStart w:id="7" w:name="m"/>${run("x")}<w:bookmarkEnd w:id="7"/>${uncovered}`, resolveAll("accept"), "26"],
  ["Accept All beside a mid-paragraph break", `${pageBreak}${uncovered}${refField("30", "b")}`, resolveAll("accept"), "30"],
] as const)("text typed in a projected child after the capture lands in its field: %s", async (_, xml, before, text) => {
  const { next, direct } = await landed(paragraph(xml), before, typeIn(text));
  expect(next).toBe(direct);
  expect(next).toContain("Z");
  expect(next).toContain(`H(${text[0]}Z${text.slice(1)})`);
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
  expect(view(out)).toBe("[«TOC \\o \\h»|H(InZtro 1)X");
  expect(view(out, "33333333")).toBe("Entry2 2]");
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
/**
 * Deletes a selection covering just the embed of the field whose instruction holds `instruction`: its embed goes.
 * Backspace beside a field that shows nothing steps over it instead (decided 2026-10-02).
 */
const deleteField = (instruction: string): Edit => (session) => {
  const { story, paraId, offset } = fieldAt(session, instruction);
  session.deleteRange({ story, start: { paraId, offset }, end: { paraId, offset: offset + 1 } });
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
  ["a selection deletes its embed", paragraph(`${refField("20", "a")}${run(" mid ")}${refField("30", "b")}`), () => {}, deleteField("REF a"), orphaned],
  ["a delete across its end", paragraph(`${refField("20", "a")}${run(" mid ")}${refField("30", "b")}`), () => {}, deleteAcrossEnd("20", "REF a"), "a H(2)mid [«REF b \\h»|H(30)] b"],
  ["text is typed in it, its embed gone before the capture", paragraph(`${refField("20", "a")}${run(" mid ")}${refField("30", "b")}`), deleteField("REF a"), typeInText("20"), "a H(2Z0) mid [«REF b \\h»|H(30)] b"],
  ["the next field follows at once", paragraph(`${refField("20", "a")}${refField("30", "b")}`), () => {}, deleteField("REF a"), "a H(20)[«REF b \\h»|H(30)] b"],
  ["the next field has no child at its index", paragraph(`${field(`${run("q")}${link(run("20"))}`, " REF a \\h ")}${run(" mid ")}${refField("30", "b")}`), () => {}, deleteField("REF a"), orphaned],
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
  const { next, direct } = await landed(bytes, () => {}, deleteField("REF a"), (out) => view(out, paraId, path));
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
  deleteField("REF a")(b);
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
    "[«TOC \\o \\h»|H(IZntro 1)Old",
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

// Text typed at a link's end stays in that link (item 8), so the export's seed holds it as the capture does.
test.each([
  ["one link", field(link(run("BB")), " REF b \\h "), "BB"],
  ["two links", field(`${linkTo("AA")}${linkTo("BB")}`, " REF b \\h "), "AA"],
])("text typed at a link's end before the capture lands exactly when a rebase deletes its field (%s)", async (_, xml, text) => {
  const { next, direct } = await landed(paragraph(xml), typeInText(text, text.length), deleteField("REF b"));
  expect(direct).toContain(`H(${text}Z)`);
  expect(next).toBe(direct);
});

test("text typed at a link's end before the capture lands when the field is left alone", async () => {
  const bytes = paragraph(field(link(run("BB")), " REF b \\h "));
  const { next, direct } = await landed(bytes, typeInText("BB", 2), (session) => edit(session, "22222222", "y"));
  expect(direct).toBe("a [«REF b \\h»|H(BBZ)] b");
  expect(next).toBe(direct);
});

test.each(["accept", "reject"] as const)(
  "%s all keeps where it stands the link of a field deleted after a join, which shares the next field's number",
  async (mode) => {
    const out = await directly(
      twoParagraphs(`${run("x")}${link(run("BB"))}${ins(run("26"))}`),
      joinNext("11111111"),
      deleteField("REF a"),
      resolveAll(mode)
    );
    expect(out).toBe(mode === "accept" ? "H(AA)[«REF b \\h»|xH(BB)26]" : "H(AA)[«REF b \\h»|xH(BB)]");
  }
);

test.each([
  [
    "Accept All renumbers the next field",
    paragraph(`${refField("AA", "a")}${field(`${link(run("BB"))}${ins(run("26"))}`, " REF b \\h ")}`),
    [deleteField("REF a"), resolveAll("accept")],
    "a H(AA)[«REF b \\h»|H(BB)26] b",
  ],
  [
    "Accept All renumbers the next field, its links kept in order",
    paragraph(`${field(`${linkTo("AA")}${linkTo("CC")}`, " REF a \\h ")}${field(`${link(run("BB"))}${ins(run("26"))}`, " REF b \\h ")}`),
    [deleteField("REF a"), resolveAll("accept")],
    "a H(AA)H(CC)[«REF b \\h»|H(BB)26] b",
  ],
  ["a join gives both fields one number", twoParagraphs(link(run("BB"))), [joinNext("11111111"), deleteField("REF a")], "H(AA)[«REF b \\h»|H(BB)]"],
] as const)("a deleted field's link joins no other field when %s", async (_, bytes, edits, expected) => {
  expect(await directly(bytes, ...edits)).toBe(expected);
});

/** Types or pastes `text` over the units from `from` into the text holding `start` through the embed of `instruction`. */
const typeOver = (start: string, from: number, instruction: string, text = "Q"): Edit => (session) => {
  const at = textAt(session, start);
  const field = fieldAt(session, instruction);
  session.replaceRange({ story: "body", start: { paraId: at.paraId, offset: at.offset + from }, end: { paraId: field.paraId, offset: field.offset + 1 } }, text);
};
test.each([
  [
    "part of its link and its end, then a join",
    docx(p("11111111", field(linkTo("BB"), " REF b \\h ")) + p("33333333", field(linkTo("CC"), " REF c \\h ")) + tail),
    [typeOver("BB", 1, "REF b"), joinNext("11111111")],
    "H(BQ)[«REF c \\h»|H(CC)]",
  ],
  [
    "the whole field, pasting, then Accept All renumbering the next",
    paragraph(`${field(linkTo("AA"), " REF a \\h ")}${field(`${linkTo("BB")}${ins(run("26"))}`, " REF b \\h ")}`),
    [typeOver("AA", 0, "REF a", "pasted"), resolveAll("accept")],
    "a H(pasted)[«REF b \\h»|H(BB)26] b",
  ],
] as const)("text typed over a field it removes joins no other field: %s", async (_, bytes, edits, expected) => {
  expect(await directly(bytes, ...edits)).toBe(expected);
});

test.each([
  ["a REF field's link", paragraph(field(linkTo("BB"), " REF b \\h ")), "BB", 1, "a [«REF b \\h»|H(BB)] b"],
  [
    "a TOC's first entry",
    docx(p("11111111", `${char("begin")}${instr(" TOC \\o \\h ")}${char("separate")}${link(run("Introduction 1"))}`) + p("33333333", `${link(run("Entry2 2"))}${char("end")}`) + tail),
    "Introduction",
    2,
    "[«TOC \\o \\h»|H(Introduction 1)",
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
  deleteField("REF a")(session);
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
// Enter between a field's links splits the field across the paragraphs (round 2 D), so the join lands as saved directly.
test("a split between two of a field's links before the capture and a join after land exactly", async () => {
  const split = (session: YrsSession) => session.splitParagraph({ story: "body", paraId: "11111111", offset: textAt(session, "BB").offset });
  const { next, direct } = await landed(threeLinks, split, joinNext("11111111"));
  expect(direct).toBe("a [«DATE»|H(AA)H(BB)H(CC)] b");
  expect(next).toBe(direct);
});
test.each([
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

// A field showing a kept insertion of its own keeps Enter as it was: its links stop projecting, which refuses.
test("a rebase fails as a RebaseError when a child's field stops projecting in the export", async () => {
  const bytes = paragraph(field(`${link(run("AA"))}${link(run("BB"))}${ins(run("26"))}`));
  const split: Edit = (session) =>
    void session.splitParagraph({ story: "body", paraId: "11111111", offset: fieldAt(session, "DATE").offset });
  await expect(landed(bytes, split, joinNext("11111111"))).rejects.toBeInstanceOf(RebaseError);
});

// Enter before a field's embed now splits the field across the paragraphs (round 2 D): its links keep projecting.
test("a split before a field's embed before the capture and a join after land exactly", async () => {
  const bytes = paragraph(field(`${link(run("AA"))}${link(run("BB"))}`));
  const split: Edit = (session) =>
    void session.splitParagraph({ story: "body", paraId: "11111111", offset: fieldAt(session, "DATE").offset });
  const { next, direct } = await landed(bytes, split, joinNext("11111111"));
  expect(direct).toBe("a [«DATE»|H(AA)H(BB)] b");
  expect(next).toBe(direct);
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

// A delete right before a field's embed used to give the embed the deleted
// text's link and field marker (yrs scanned past the embed for formatting to
// clean up), and the export then dropped the field with its text.
/** A Word table of contents: entries are links to headings, each with a tab and a PAGEREF field. */
const tocEntry = (anchor: string, text: string, page: string) =>
  `<w:hyperlink w:anchor="${anchor}" w:history="1">${run(text)}<w:r><w:tab/></w:r>${char("begin")}${instr(` PAGEREF ${anchor} \\h `)}${char("separate")}${run(page)}${char("end")}</w:hyperlink>`;
const wordToc = docx(
  p("11111111", `${char("begin")}${instr(" TOC \\o \\h ")}${char("separate")}${tocEntry("_Toc1", "Introduction", "1")}`) +
    p("33333333", `${tocEntry("_Toc2", "Details", "2")}${char("end")}`) +
    tail
);
/** The offset of the TOC field's embed: the end of its first entry. */
const tocEnd = (session: YrsSession) => {
  const { paraId, offset } = fieldAt(session, "TOC");
  return { paraId, offset };
};
const atTocEnd: Record<string, [Edit, string]> = {
  "Backspace at the end of the first entry (its page number field)": [
    (session) => void session.deleteAt({ story: "body", ...tocEnd(session) }, "backward"),
    "[«TOC \\o \\h»|H(Introduction)",
  ],
  "retyping the first entry's page number": [
    (session) => {
      const { paraId, offset } = tocEnd(session);
      session.replaceRange({ story: "body", start: { paraId, offset: offset - 1 }, end: { paraId, offset } }, "7");
    },
    "[«TOC \\o \\h»|H(Introduction7)",
  ],
  "typing at the end of the first entry and deleting it": [
    (session) => {
      const { paraId, offset } = tocEnd(session);
      session.insertText({ story: "body", paraId, offset }, "Z");
      session.deleteAt({ story: "body", paraId, offset: offset + 1 }, "backward");
    },
    "[«TOC \\o \\h»|H(Introduction[«PAGEREF _Toc1 \\h»|1])",
  ],
};
test.each(Object.keys(atTocEnd))("%s keeps the TOC field across three publications", async (name) => {
  const [change, expected] = atTocEnd[name]!;
  let bytes = wordToc;
  for (let publication = 0; publication < 3; publication += 1) {
    const session = await open(bytes);
    if (publication === 0) change(session);
    edit(session, "22222222", "x");
    bytes = await publish(bytes, session.encodeState());
    session.destroy();
    expect(view(bytes)).toBe(expected);
  }
});
test("Backspace at the end of a TOC's first entry after the capture lands with the field", async () => {
  const [change, expected] = atTocEnd["Backspace at the end of the first entry (its page number field)"]!;
  const { next, direct } = await landed(wordToc, () => {}, change);
  expect(direct).toBe(expected);
  expect(next).toBe(direct);
});
test.each([
  [
    "Backspace at the end of a field's link right before another field",
    paragraph(`${field(linkTo("BBB"), " REF b \\h ")}${field(linkTo("CC"), " REF c \\h ")}`),
    (session: YrsSession) => {
      const { paraId, offset } = textAt(session, "BBB");
      session.deleteAt({ story: "body", paraId, offset: offset + 3 }, "backward");
    },
    "a [«REF b \\h»|H(BB)][«REF c \\h»|H(CC)] b",
  ],
  [
    "retyping the last character of a field's link right before another link",
    paragraph(`${field(linkTo("BB"), " REF b \\h ")}${linkTo("next")}`),
    (session: YrsSession) => {
      const { paraId, offset } = textAt(session, "BB");
      session.replaceRange({ story: "body", start: { paraId, offset: offset + 1 }, end: { paraId, offset: offset + 2 } }, "Q");
    },
    "a [«REF b \\h»|H(BQ)]H(next) b",
  ],
] as const)("%s keeps the field", async (_, bytes, change, expected) => {
  expect(await directly(bytes, change)).toBe(expected);
});

test.each([
  ["typing", "Q", "a H(AA)[«REF b \\h»|H(QB)] b"],
  ["pasting", "XYZ", "a H(AA)[«REF b \\h»|H(XYZB)] b"],
] as const)("%s over a removed field's embed into the next field's link keeps the text in that field", async (_, text, expected) => {
  const bytes = paragraph(`${field(linkTo("AA"), " REF a \\h ")}${field(linkTo("BB"), " REF b \\h ")}`);
  const over: Edit = (session) => {
    const from = fieldAt(session, "REF a");
    const to = textAt(session, "BB");
    session.replaceRange({ story: "body", start: from, end: { paraId: to.paraId, offset: to.offset + 1 } }, text);
  };
  expect(await directly(bytes, over)).toBe(expected);
});

test("deleting a field with the tracked break after it, after the capture, lands when the export's seed numbers the next field otherwise", async () => {
  // The seed reads the deleted break as the next paragraph's, so the seed of
  // the export numbers the second field one lower than the capture does.
  const pageBreak = '<w:r><w:br w:type="page"/></w:r>';
  const bytes = docx(
    p("33333333", run("prev")) + p("11111111", `${run("a ")}${field(link(run("26")))}${del(pageBreak)}${field(link(run("30")))}`) + tail
  );
  const removeField: Edit = (session) => {
    const { story, paraId, offset } = textAt(session, "26");
    session.deleteRange({ story, start: { paraId, offset }, end: { paraId, offset: offset + 3 } });
  };
  const { next, direct } = await landed(bytes, () => {}, removeField);
  expect(next).toBe(direct);
});

const bookmarkStart = (id: number) => `<w:bookmarkStart w:id="${id}" w:name="m${id}"/>`;
const bookmarkEnd = (id: number) => `<w:bookmarkEnd w:id="${id}"/>`;
/** The body's bookmark and field characters in order, one string per paragraph: B5/E5, [ | ]. */
const markers = (bytes: Uint8Array) =>
  [...documentXml(bytes).matchAll(/<w:p [\s\S]*?<\/w:p>/g)].map(([paragraph]) =>
    [...paragraph.matchAll(/<w:(bookmarkStart|bookmarkEnd|fldChar) w:(?:id|fldCharType)="(\w+)"/g)]
      .map(([, tag, value]) => (tag === "fldChar" ? { begin: "[", separate: "|", end: "]" }[value!] : `${tag === "bookmarkStart" ? "B" : "E"}${value}`))
      .join(" ")
  ).filter(Boolean);
/** Every bookmark opens before it closes. */
const balanced = (sequence: string) =>
  sequence.split(" ").every((marker, index, all) => !marker.startsWith("E") || all.slice(0, index).includes(`B${marker.slice(1)}`));
const nestedContinued =
  p("11111111", run("a") + char("begin") + instr(" IF ") + char("begin") + instr(" PAGE ")) +
  p("33333333", char("separate") + run("one") + char("end") + bookmarkStart(5) + char("separate") + bookmarkEnd(5) + run("yes") + char("end") + run("z"));

test.each([
  [
    "bookmarks a join collapses to one point",
    p("11111111", run("ab") + bookmarkStart(3) + bookmarkStart(4) + run("c")) +
      p("33333333", run("d") + bookmarkStart(5) + bookmarkEnd(3) + run("e") + bookmarkEnd(4) + bookmarkEnd(5) + run("f")),
    (session: YrsSession) => session.deleteRange({ story: "body", start: { paraId: "11111111", offset: 2 }, end: { paraId: "33333333", offset: 2 } }),
  ],
  [
    "bookmarks around a continued field a join removes",
    p("11111111", run("a") + char("begin") + instr(" IF ") + char("begin") + instr(" PAGE ")) +
      p("33333333", char("separate") + run("one") + bookmarkStart(3) + bookmarkStart(4) + char("end") + bookmarkStart(5) + bookmarkEnd(3) + char("separate") + bookmarkEnd(4) + bookmarkEnd(5) + run("yes") + char("end") + run("z")),
    (session: YrsSession) => session.deleteRange({ story: "body", start: { paraId: "11111111", offset: 1 }, end: { paraId: "33333333", offset: 0 } }),
  ],
])("%s save each start before its end", async (_, body, change) => {
  let bytes = docx(body + tail);
  for (let publication = 0; publication < 3; publication += 1) {
    const session = await open(bytes);
    if (publication === 0) change(session);
    else edit(session, "22222222", "z");
    bytes = await publish(bytes, session.encodeState());
    session.destroy();
    const [first] = markers(bytes);
    expect(first!.split(" ").filter((marker) => marker.startsWith("B"))).toHaveLength(3);
    expect(balanced(first!)).toBe(true);
  }
});

test("a bookmark between a nested continued field's inner end and outer separate stays there", async () => {
  let bytes = docx(nestedContinued + tail);
  const source = markers(bytes);
  expect(source).toEqual(["[ [", "| ] B5 | E5 ]"]);
  for (let publication = 0; publication < 3; publication += 1) {
    const session = await open(bytes);
    if (publication > 0) edit(session, "22222222", "z");
    bytes = await publish(bytes, session.encodeState());
    session.destroy();
    expect(markers(bytes)).toEqual(source);
  }
});

test("coincident continued-field markers keep their source order in the editor", async () => {
  const bytes = docx(nestedContinued + tail);
  for (let attempt = 0; attempt < 6; attempt += 1) {
    const session = await open(bytes);
    const bookmarks = session.paragraphs("body")[1]!.properties.bookmarks as Array<{ kind: string; offset: number }>;
    expect(bookmarks.map(({ kind, offset }) => `${kind}@${offset}`)).toEqual([
      "fieldseparate@0", "fieldend@3", "start@3", "fieldseparate@3", "end@3", "fieldend@6",
    ]);
    session.destroy();
  }
});

/** Types Q inside the projected child holding `text`, in whichever paragraph it is. */
function typeInChild(session: YrsSession, text: string) {
  let offset = 0;
  let found: number | undefined;
  for (const segment of session.storySegments("body")) {
    if (segment.kind === "pilcrow") {
      if (found !== undefined) return void session.insertText({ story: "body", paraId: segment.paraId, offset: found }, "Q");
      offset = 0;
      continue;
    }
    if (found === undefined && segment.kind === "text" && segment.attributes.fieldResult && segment.text.includes(text))
      found = offset + segment.text.indexOf(text) + 1;
    offset += segment.kind === "text" ? segment.text.length : 1;
  }
  throw new Error(`no projected child holds ${text}`);
}
const holder = p("55555555", bookmarkStart(5) + bookmarkEnd(5));
const tocFields = field(link(run("26"))) + field(link(run("30")));
const handOff = (session: YrsSession) => session.deleteAt({ story: "body", paraId: "55555555", offset: 0 }, "forward");
// R6-N1: the hand-off writes the bookmarks ahead of the survivor's fields, so it renumbers them as the export's seed does.
test.each([
  ["a paragraph before", p("33333333", run("prev")) + holder + p("44444444", pageBreak + tocFields)],
  ["the story start", holder + p("44444444", pageBreak + tocFields)],
  ["a column break after the page break", p("33333333", run("prev")) + holder + p("44444444", pageBreak + `<w:r><w:br w:type="column"/></w:r>` + tocFields)],
  ["text before the fields", p("33333333", run("prev")) + holder + p("44444444", pageBreak + run("a") + tocFields)],
  ["a table instead of the break", p("33333333", run("prev")) + holder + `<w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc>${p("12121212", run("cell"))}</w:tc></w:tr></w:tbl>` + p("44444444", tocFields)],
])("typing in a field's link after a captured bookmark hand-off with %s lands exactly", async (_, body) => {
  const bytes = docx(body + tail);
  const session = await open(bytes);
  handOff(session);
  const captured = session.encodeState();
  const exported = await publish(bytes, captured);
  typeInChild(session, "30");
  const latest = session.encodeState();
  session.destroy();
  const direct = documentXml(await publish(bytes, latest));
  const { state } = await rebaseOffice(bytes, checkpoint(bytes, captured), checkpoint(bytes, latest), exported);
  const rebased = await open(exported, state);
  const next = await publish(exported, rebased.encodeState());
  rebased.destroy();
  expect(documentXml(next)).toBe(direct);
  expect(view(next, "44444444")).toContain("H(3Q0)");
});

/** The first paragraph's segments as the editor holds them: text with its field result and link, or an embed. */
const segments = (session: YrsSession) =>
  session.storySegments("body").map((segment) =>
    segment.kind === "text"
      ? `${segment.text}${segment.attributes.fieldResult ? "<" : ""}${segment.attributes.hyperlink ? "@" : ""}`
      : segment.kind === "pilcrow" ? "¶" : `[${segment.embedKind}]`
  ).join("").split("¶")[0];
// Item 8 (N18 follow-up): text typed at the end of a field result's projected link stays in that link and field.
test.each([
  ["a REF field's link", field(link(run("AA")), " REF a \\h ") + field(link(run("BB")), " REF b \\h "), "AA", "[«REF a \\h»|H(AAZ)][«REF b \\h»|H(BB)]"],
  ["a field's only link", field(link(run("2026"))), "2026", "[«DATE»|H(2026Z)]"],
  ["a link before a kept insertion", field(link(run("20")) + ins(run("26"))), "20", "[«DATE»|H(20Z)+{26}]"],
])("text typed at the end of %s stays in it across publications", async (_, xml, text, saved) => {
  let bytes = paragraph(xml);
  const session = await open(bytes);
  session.insertText({ story: "body", paraId: "11111111", offset: childAt(session, text) + text.length }, "Z");
  const editor = segments(session);
  bytes = await publish(bytes, session.encodeState());
  session.destroy();
  expect(view(bytes)).toContain(saved);
  for (let publication = 0; publication < 2; publication += 1) {
    const reopened = await open(bytes);
    expect(segments(reopened)).toBe(editor);
    edit(reopened, "22222222", "z");
    bytes = await publish(bytes, reopened.encodeState());
    reopened.destroy();
  }
});

// Round 2 C: text typed after a field result's projected simple field stays a unit inside that field across publications.
test("text typed after a projected simple field stays in its field across publications", async () => {
  let bytes = paragraph(field(fs(run("20"), " PAGE ")));
  const session = await open(bytes);
  session.insertText({ story: "body", paraId: "11111111", offset: 3 }, "Y");
  const editor = segments(session);
  expect(editor).toBe("a [field]Y<[field] b");
  bytes = await publish(bytes, session.encodeState());
  session.destroy();
  expect(view(bytes)).toContain("[«DATE»|F<20>Y]");
  for (let publication = 0; publication < 2; publication += 1) {
    const reopened = await open(bytes);
    expect(segments(reopened)).toBe(editor);
    edit(reopened, "22222222", "z");
    bytes = await publish(bytes, reopened.encodeState());
    reopened.destroy();
  }
});

// Round 2 D: Enter inside a projected link splits its field across the two paragraphs (the field ends in the second).
const tocParagraphs =
  p("11111111", `${run("C")}${char("begin")}${instr(" TOC ")}${char("separate")}${link(run("Introduction"))}`) +
  p("33333333", `${link(run("Details"))}${char("end")}`);
test.each([
  ["a field's only link", paragraph(field(link(run("2026")))), "2026", 4, ["a [«DATE»|H(2026)", "]"]],
  ["a table of contents' first entry", docx(tocParagraphs + tail), "Introduction", 5, ["C[«TOC»|H(Intro)", "H(duction)", "H(Details)]"]],
])("Enter inside %s splits its field across the paragraphs", async (_, source, text, at, saved) => {
  let bytes = source;
  const session = await open(bytes);
  session.splitParagraph({ story: "body", paraId: "11111111", offset: childAt(session, text) + at });
  const editor = matrixUnits(session, "body").replace(/¶[^¶]*tail¶$/, "");
  bytes = await publish(bytes, session.encodeState());
  session.destroy();
  const paragraphs = [...documentXml(bytes).matchAll(/<w:p [^>]*w14:paraId="([0-9A-F]+)"/g)].map(([, id]) => id!);
  expect(paragraphs.slice(0, saved.length).map((id) => view(bytes, id).replace(/ b$/, ""))).toEqual(saved);
  for (let publication = 0; publication < 2; publication += 1) {
    const reopened = await open(bytes);
    expect(matrixUnits(reopened, "body").replace(/¶[^¶]*tail¶$/, "")).toBe(editor);
    edit(reopened, "22222222", "z");
    bytes = await publish(bytes, reopened.encodeState());
    reopened.destroy();
  }
});


// Round 2 E (zz-n4): Accept all uncovers a link a comment around its field does not cover until a publication; a
// later delete of that field's embed must not leave the rebased save covering the link where the direct save does not.
test("deleting a field after Accept all uncovered its link under a comment lands exactly or refuses", async () => {
  const bytes = docx(
    p("33333333", run("prev")) +
      p("11111111", `${ins(run("xy"))}<w:commentRangeStart w:id="5"/>${field(run("20") + ins(link(run("26"))))}${field(link(run("30")))}<w:commentRangeEnd w:id="5"/>${ref(5)}`) +
      tail
  );
  const dropFirst: Edit = (session) => {
    const { story, paraId, offset } = fieldAt(session, "DATE");
    session.deleteRange({ story, start: { paraId, offset }, end: { paraId, offset: offset + 1 } });
  };
  const show = (saved: Uint8Array) => sig(saved, "word/document.xml");
  const result = await landed(bytes, (session) => void session.acceptChange({ all: true }), dropFirst, show).catch(
    (error: unknown) => error
  );
  if (result instanceof RebaseError) return;
  const { next, direct } = result as { next: string; direct: string };
  expect(direct).toContain("L(26)<c1");
  expect(next).toBe(direct);
});

// Round 2 G: Undo of a join restores bookmarks where the save writes them.
test("Undo of a join collapsing bookmarks leaves them where the save writes them", async () => {
  const bytes = docx(
    p("11111111", run("ab") + bookmarkStart(3) + bookmarkStart(4) + run("c")) +
      p("33333333", run("d") + bookmarkStart(5) + bookmarkEnd(3) + run("e") + bookmarkEnd(4) + bookmarkEnd(5) + run("f")) +
      tail
  );
  const placed = (session: YrsSession) =>
    session.paragraphs("body").map(({ properties }) =>
      ((properties.bookmarks ?? []) as Array<{ kind: string; id: unknown; offset: number }>).map(({ kind, id, offset }) => `${kind}${id}@${offset}`).join(" ")
    );
  const session = await open(bytes);
  session.deleteRange({ story: "body", start: { paraId: "11111111", offset: 2 }, end: { paraId: "33333333", offset: 2 } });
  expect(session.undo()).toBe(true);
  const editor = placed(session);
  const saved = await publish(bytes, session.encodeState());
  session.destroy();
  const reopened = await open(saved);
  expect(placed(reopened)).toEqual(editor);
  reopened.destroy();
});

// Round 2 F: a join removing a nested continued field's paragraph, then Undo after a capture: the rebased state
// holds the restored fields only with their continued separate and end, or the rebase refuses.
test("Undo of a join removing a nested continued field lands exactly or refuses, in every story", async () => {
  await prime();
  for (const where of ["body", "cell", "header"] as const) {
    const row = await runRow({
      id: `nested continued join, undo in ${where}`,
      bytes: matrixDocx(where, nestedContinued),
      where,
      before: (session, story) =>
        void session.deleteRange({ story, start: { paraId: "11111111", offset: 1 }, end: { paraId: "33333333", offset: 0 } }),
      after: (session) => void session.undo(),
    });
    expect(row.cls).not.toStartWith("silent");
  }
});

// Round 2 review B, I: Enter inside a projected link, then Delete or Backspace back, restores the link and its field.
const twoLinks = paragraph(field(link(run("AA")) + linkTo("BB"), " REF a \\h "));
test.each([
  ["Delete at the first half's end", twoLinks, "AA", (session: YrsSession, first: string) => {
    const { length } = session.paragraphSpans("body").find((span) => span.paraId === first)!;
    session.deleteAt({ story: "body", paraId: first, offset: length }, "forward");
  }],
  ["Backspace at the second half's start", twoLinks, "AA", (session: YrsSession, _: string, second: string) =>
    void session.deleteAt({ story: "body", paraId: second, offset: 0 }, "backward")],
  ["Delete in a table of contents", docx(tocParagraphs + tail), "Introduction", (session: YrsSession, first: string) => {
    const { length } = session.paragraphSpans("body").find((span) => span.paraId === first)!;
    session.deleteAt({ story: "body", paraId: first, offset: length }, "forward");
  }],
])("Enter inside a projected link, then %s, restores the original", async (_, bytes, text, join) => {
  const before = await open(bytes);
  const original = matrixUnits(before, "body");
  before.destroy();
  const session = await open(bytes);
  const { firstParaId, secondParaId } = session.splitParagraph({ story: "body", paraId: "11111111", offset: childAt(session, text) + 1 });
  join(session, firstParaId, secondParaId);
  expect(matrixUnits(session, "body")).toBe(original);
  const saved = await publish(bytes, session.encodeState());
  session.destroy();
  expect(documentXml(saved)).toBe(documentXml(await publish(bytes, (await open(bytes)).encodeState())));
});

// Round 2 review B, O: text typed at the end of a paragraph whose field code continues into the next stays out of the code.
test("text typed after a field whose code continues into the next paragraph lands before the field", async () => {
  let bytes = docx(
    p("11111111", run("a") + char("begin") + instr(" DATE ")) + p("33333333", char("separate") + run("20") + char("end") + run("z")) + tail
  );
  const session = await open(bytes);
  const { length } = session.paragraphSpans("body").find((span) => span.paraId === "11111111")!;
  session.insertText({ story: "body", paraId: "11111111", offset: length }, "Q");
  const editor = matrixUnits(session, "body");
  bytes = await publish(bytes, session.encodeState());
  session.destroy();
  expect(view(bytes)).toBe("aQ[«DATE»");
  const reopened = await open(bytes);
  expect(matrixUnits(reopened, "body")).toBe(editor);
  reopened.destroy();
});


// Round 2 review B: Enter after a projected simple field splits its field (K, M; a kept insertion keeps the old
// Enter), and Enter inside a projected link then a join back restores it (I), with a comment around it too.
const holder44 = (xml: string) => p("44444444", `${run("a ")}${xml}${run(" b")}`);
const enterAfterPage: MatrixEdit = (s, st) => {
  const at = matrixFieldAt(s, st, "PAGE");
  s.splitParagraph({ ...at, story: st, offset: at.offset + 1 });
};
const enterInAA = (join: "back" | "Backspace"): MatrixEdit => (s, st) => {
  const at = matrixLocate(s, st, "AA");
  const { firstParaId, secondParaId } = s.splitParagraph({ story: st, paraId: at.paraId, offset: at.offset + 1 });
  if (join === "back") s.deleteAt({ story: st, paraId: firstParaId, offset: matrixLen(s, st, firstParaId) }, "forward");
  else s.deleteAt({ story: st, paraId: secondParaId, offset: 0 }, "backward");
};
test.each([
  ["Enter after a projected simple field with result text before it", field(run("x") + fs(run("20"), " PAGE ") + run("y")), enterAfterPage],
  ["Enter after a projected simple field with a link and text after it", field(fs(run("20"), " PAGE ") + link(run("AA")) + run("x")), enterAfterPage],
  ["Enter after a projected simple field before a kept insertion", field(fs(run("20"), " PAGE ") + ins(run("y")) + run("x")), enterAfterPage],
  ["Enter in a link of a commented field, Delete back", `<w:commentRangeStart w:id="5"/>${field(link(run("AA")) + linkTo("BB"), " REF a \\h ")}<w:commentRangeEnd w:id="5"/>${ref(5)}`, enterInAA("back")],
  ["Enter in a link before a simple field, Backspace back", field(link(run("AA")) + fs(run("20"), " PAGE ") + linkTo("BB"), " REF a \\h "), enterInAA("Backspace")],
])("%s lands exactly in every capture order", async (_, xml, edit) => {
  await prime();
  for (const where of ["body", "cell", "header"] as const)
    for (const order of orders(`${where}`, null, edit)) {
      const row = await runRow({ id: order.id, bytes: matrixDocx(where, holder44(xml)), where, before: order.before, after: order.after });
      expect(`${order.id}: ${row.cls}`).toEndWith(": exact");
    }
});

// Round 2 review C: Enter in a projected link of a field showing its own result text, then Undo (P) or a join back
// (Q), restores the field without duplicating its links; a deletion or bookmark after the split splits (R).
const ownText: Record<string, string> = {
  "[REF|L(AA)yL(BB)]": field(link(run("AA")) + run("y") + linkTo("BB"), " REF a \\\\h "),
  "[REF|L(AA)y]": field(link(run("AA")) + run("y"), " REF a \\\\h "),
  "[REF|xL(AA)y]": field(run("x") + link(run("AA")) + run("y"), " REF a \\\\h "),
  "[REF|L(AA)L(BB)y]": field(link(run("AA")) + linkTo("BB") + run("y"), " REF a \\\\h "),
};
const splitAA = (session: YrsSession) => session.splitParagraph({ story: "body", paraId: "11111111", offset: childAt(session, "AA") + 1 });
const restoring: Record<string, (session: YrsSession) => void> = {
  "Enter, Undo": (session) => {
    splitAA(session);
    session.undo();
  },
  "Enter, Backspace": (session) => {
    const { secondParaId } = splitAA(session);
    session.deleteAt({ story: "body", paraId: secondParaId, offset: 0 }, "backward");
  },
  "Enter, Delete": (session) => {
    const { firstParaId } = splitAA(session);
    const { length } = session.paragraphSpans("body").find((span) => span.paraId === firstParaId)!;
    session.deleteAt({ story: "body", paraId: firstParaId, offset: length }, "forward");
  },
  "Enter, a range delete across the paragraph mark": (session) => {
    const { firstParaId, secondParaId } = splitAA(session);
    const { length } = session.paragraphSpans("body").find((span) => span.paraId === firstParaId)!;
    session.deleteRange({ story: "body", start: { paraId: firstParaId, offset: length }, end: { paraId: secondParaId, offset: 0 } });
  },
  "Enter, Backspace, Undo": (session) => {
    const { secondParaId } = splitAA(session);
    session.deleteAt({ story: "body", paraId: secondParaId, offset: 0 }, "backward");
    session.undo();
  },
};
test.each(Object.keys(ownText).flatMap((file) => Object.keys(restoring).map((how) => [file, how] as const)))(
  "%s: %s restores the original without duplicating its links",
  async (file, how) => {
    const bytes = paragraph(ownText[file]!);
    const untouched = await open(bytes);
    const original = documentXml(await publish(bytes, untouched.encodeState()));
    untouched.destroy();
    const session = await open(bytes);
    restoring[how]!(session);
    const editor = matrixUnits(session, "body");
    const saved = await publish(bytes, session.encodeState());
    session.destroy();
    const xml = view(saved);
    expect(xml.split("AA").length - 1).toBe(1);
    expect(documentXml(saved)).toBe(original);
    const reopened = await open(saved);
    expect(matrixUnits(reopened, "body")).toBe(editor);
    reopened.destroy();
  }
);
test.each([
  ["a tracked deletion", field(fs(run("20"), " PAGE ") + del(deleted("w")) + run("y"))],
  ["a bookmark", field(fs(run("20"), " PAGE ") + `<w:bookmarkStart w:id="7" w:name="m7"/>` + run("y") + `<w:bookmarkEnd w:id="7"/>`)],
])("Enter after a projected simple field with %s after it splits the field exactly", async (_, xml) => {
  await prime();
  for (const where of ["body", "cell", "header"] as const)
    for (const order of orders(`${where}`, null, enterAfterPage)) {
      const row = await runRow({ id: order.id, bytes: matrixDocx(where, holder44(xml)), where, before: order.before, after: order.after });
      expect(`${order.id}: ${row.cls}`).toEndWith(": exact");
    }
});


// Round 2 review D, W: a join after an Enter split keeps every result run and the text typed between the halves.
const bodyText = (bytes: Uint8Array) =>
  [...documentXml(bytes).matchAll(/<w:p [^>]*>([\s\S]*?)<\/w:p>/g)]
    .map(([, paragraph]) => [...paragraph!.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(([, text]) => text).join(""))
    .filter((text) => text !== "tail")
    .join("¶");
const runsField = (xml: string) => paragraph(field(xml, " REF a \\\\h "));
test.each([
  ["[REF|L(AA)yz], Backspace", runsField(link(run("AA")) + run("y") + run("z")), "join", "a AAyz b"],
  ["[REF|L(AA)yzL(BB)w], Backspace", runsField(link(run("AA")) + run("y") + run("z") + linkTo("BB") + run("w")), "join", "a AAyzBBw b"],
  ["[REF|L(AA)yz], Enter at the 2nd half's start, type X, Backspace", runsField(link(run("AA")) + run("y") + run("z")), "X between", "a AX¶Ayz b"],
  ["[REF|L(AA)yz], type X at the 1st half's end, Backspace", runsField(link(run("AA")) + run("y") + run("z")), "X first", "a AXAyz b"],
])("Enter in a projected link of %s keeps every run and typed text", async (_, bytes, flow, text) => {
  const session = await open(bytes);
  const { firstParaId, secondParaId } = splitAA(session);
  if (flow === "X between") {
    const inner = session.splitParagraph({ story: "body", paraId: secondParaId, offset: 0 });
    session.insertText({ story: "body", paraId: inner.firstParaId, offset: 0 }, "X");
    session.deleteAt({ story: "body", paraId: inner.firstParaId, offset: 0 }, "backward");
  } else {
    if (flow === "X first") {
      const { length } = session.paragraphSpans("body").find((span) => span.paraId === firstParaId)!;
      session.insertText({ story: "body", paraId: firstParaId, offset: length }, "X");
    }
    session.deleteAt({ story: "body", paraId: secondParaId, offset: 0 }, "backward");
  }
  const saved = await publish(bytes, session.encodeState());
  session.destroy();
  expect(bodyText(saved)).toBe(text);
});

test("Enter after a projected simple field followed by two runs, then Backspace, restores the field", async () => {
  const bytes = paragraph(field(fs(run("20"), " PAGE ") + run("y") + run("z")));
  const before = await open(bytes);
  const original = matrixUnits(before, "body");
  before.destroy();
  const session = await open(bytes);
  const at = fieldAt(session, "PAGE");
  const { secondParaId } = session.splitParagraph({ story: "body", paraId: at.paraId, offset: at.offset + 1 });
  session.deleteAt({ story: "body", paraId: secondParaId, offset: 0 }, "backward");
  expect(matrixUnits(session, "body")).toBe(original);
  const saved = await publish(bytes, session.encodeState());
  session.destroy();
  expect(bodyText(saved)).toBe("a 20yz b");
});

// Follow-up 3: the paragraphs Enter makes of the first one, and Backspace at the second's start.
const paraIds = (session: YrsSession) => session.paragraphSpans("body").map((span) => span.paraId);
const nthParagraph = (session: YrsSession, k: number) => paraIds(session)[paraIds(session).indexOf("11111111") + k]!;
const startOf = (session: YrsSession, k: number) => ({ story: "body", paraId: nthParagraph(session, k), offset: 0 });
const backspace = (session: YrsSession) => void session.deleteAt(startOf(session, 1), "backward");

// Follow-up 3 (rv5-ur2): Enter, Undo, Redo in a field whose result after the link is one multi-character run puts the
// field's end where Enter put it.
test("Enter, Undo, Redo keeps a split field's end after its moved run", async () => {
  const bytes = paragraph(field(link(run("AA")) + run("yy"), " REF a \\h "));
  const enter = async (redo: boolean) => {
    const session = await open(bytes);
    session.splitParagraph({ story: "body", paraId: "11111111", offset: childAt(session, "AA") + 1 });
    if (redo) {
      session.undo();
      session.redo();
    }
    const editor = matrixUnits(session, "body");
    const saved = await publish(bytes, session.encodeState());
    session.destroy();
    const reopened = await open(saved);
    const units = matrixUnits(reopened, "body");
    reopened.destroy();
    return { editor, saved: sig(saved, "word/document.xml"), reopened: units };
  };
  const [plain, redone] = [await enter(false), await enter(true)];
  expect(redone).toEqual(plain);
  expect(plain.saved).toContain("yy]");
});

// Round 3 review, finding 1 (rv6-peers): peer A presses Undo and Redo after Enter split a field while peer B edits
// the moved run, concurrently or first. Both peers and the reopened save end the field at the same place, since the
// Undo or Redo re-anchors the field's end (decided 2026-10-02); each typed character lands once. Redo re-inserting
// the moved run can put B's W after it or revive B's deletion (Yjs Undo semantics, as for item 1).
const fieldEndOf = (session: YrsSession) => {
  for (const { paraId, properties } of session.paragraphs("body")) {
    const end = ((properties as { bookmarks?: Array<{ kind: string; offset: number }> }).bookmarks ?? []).find((mark) => mark.kind === "fieldend");
    if (end) return { story: "body", paraId, offset: end.offset };
  }
  throw new Error("no field end");
};
const peerSteps: Record<string, (session: YrsSession) => void> = {
  "W at the field's end": (session) => session.insertText(fieldEndOf(session), "W"),
  "W between the y's": (session) => session.insertText({ ...startOf(session, 1), offset: 2 }, "W"),
  "W at the 2nd paragraph's start": (session) => session.insertText(startOf(session, 1), "W"),
  "deleting the 1st y": (session) => void session.deleteAt({ ...startOf(session, 1), offset: 2 }, "backward"),
};
test.each(
  ["yy", "yyyy"].flatMap((moved) =>
    [
      ...Object.keys(peerSteps).map((step) => ["U,R", step] as const),
      ["U,R,U,R", "W at the field's end"] as const,
    ].flatMap(([steps, step]) => (["concurrently", "after B"] as const).map((mode) => [moved, steps, step, mode] as const))
  )
)("L(AA)%s split, A presses %s beside B %s, %s: the peers and the save end the field alike", async (moved, steps, step, mode) => {
  const bytes = paragraph(field(link(run("AA")) + run(moved), " REF a \\h "));
  const A = await open(bytes);
  const B = await open(bytes);
  A.splitParagraph({ story: "body", paraId: "11111111", offset: childAt(A, "AA") + 1 });
  const sync = () => {
    const [toB, toA] = [A.encodeStateAsUpdate(B.encodeStateVector()), B.encodeStateAsUpdate(A.encodeStateVector())];
    A.applyUpdate(toA);
    B.applyUpdate(toB);
  };
  sync();
  const press = () => {
    for (const key of steps.split(",")) key === "U" ? A.undo() : A.redo();
  };
  if (mode === "concurrently") {
    press();
    peerSteps[step]!(B);
  } else {
    peerSteps[step]!(B);
    sync();
    press();
  }
  sync();
  const saved = await publish(bytes, A.encodeState());
  const reopened = await open(saved);
  expect(matrixUnits(A, "body")).toBe(matrixUnits(B, "body"));
  expect(matrixUnits(reopened, "body")).toBe(matrixUnits(A, "body"));
  if (step.startsWith("W")) expect([...bodyText(saved)].sort().join("")).toBe([...`a A¶AW${moved} b`].sort().join(""));
  for (const session of [A, B, reopened]) session.destroy();
});

// Round 3 review, finding 2: Undo after a redone deletion removes only its own step (the vendored yrs followed a
// redone item to its start, so this Undo also took the text typed in the step before).
test("Undo after deleting, undoing and redoing removes only its own step", async () => {
  const bytes = docx(p("11111111", run("x")) + tail);
  const session = await open(bytes);
  const at = (offset: number) => ({ story: "body", paraId: "11111111", offset });
  const text = () => session.paragraphs("body")[0]!.text;
  session.insertText(at(1), "abc");
  session.addUndoBoundary();
  session.insertText(at(4), "def");
  session.addUndoBoundary();
  session.deleteRange({ story: "body", start: at(0), end: at(7) });
  session.addUndoBoundary();
  session.undo();
  session.redo();
  session.undo();
  expect(text()).toBe("xabcdef");
  session.undo();
  expect(text()).toBe("xabc");
  session.undo();
  expect(text()).toBe("x");
  session.redo();
  expect(text()).toBe("xabc");
  const saved = await publish(bytes, session.encodeState());
  session.destroy();
  expect(view(saved)).toBe("xabc");
});

// Follow-up 3 (rv5-tab): a join gives the field the text its seed shows, without the tab of its moved run.
test("Enter then Backspace in a field whose moved run holds a tab restores the field's shown text", async () => {
  const bytes = paragraph(field(link(run("AA")) + `<w:r><w:t>y</w:t><w:tab/><w:t>z</w:t></w:r>`, " REF a \\h "));
  const before = await open(bytes);
  const original = matrixUnits(before, "body");
  before.destroy();
  const session = await open(bytes);
  session.splitParagraph({ story: "body", paraId: "11111111", offset: childAt(session, "AA") + 1 });
  backspace(session);
  expect(matrixUnits(session, "body")).toBe(original);
  session.destroy();
});

// Follow-ups after round 3: Enter in a projected link, alone or joined back, leaves a field showing what the seed of
// the saved state shows, in every story. A nested complex field before the link shows nothing in the seed, so it shows
// nothing after the split either; a moved run holding only a w:ptab (which seeds nothing) rejoins with the runs around
// it, so the join saves one link. Oracle: the original seed, the untouched save and the saved text.
const ptabRun = `<w:r><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="dot"/></w:r>`;
const splitShapes: Record<string, [string, string]> = {
  "[REF|[PAGE|7]L(AA)yy]": [field(field(run("7"), " PAGE ") + link(run("AA")) + run("yy"), " REF a \\h "), "a 7AAyy b"],
  "[REF|L(AA)y<ptab>z]": [field(link(run("AA")) + run("y") + ptabRun + run("z"), " REF a \\h "), "a AAyz b"],
  "[REF|L(AA)<ptab>yz]": [field(link(run("AA")) + ptabRun + run("yz"), " REF a \\h "), "a AAyz b"],
  // Review of ee62d514: a moved run holding only a line break, and an empty run after a projected simple field.
  "[REF|L(AA)y<br>z]": [field(link(run("AA")) + run("y") + "<w:r><w:br/></w:r>" + run("z"), " REF a \\h "), "a AAyz b"],
  "[REF|L(AA)F{7}<b/>y]": [field(link(run("AA")) + fs(run("7"), " PAGE ") + "<w:r><w:rPr><w:b/></w:rPr></w:r>" + run("y"), " REF a \\h "), "a AA7y b"],
};
const splitFlows: Record<string, { joined: boolean; act: (session: YrsSession, story: string, first: string, second: string) => void }> = {
  Enter: { joined: false, act: () => {} },
  "Enter, Backspace": { joined: true, act: (session, story, _, second) => void session.deleteAt({ story, paraId: second, offset: 0 }, "backward") },
  "Enter, Delete": {
    joined: true,
    act: (session, story, first) => void session.deleteAt({ story, paraId: first, offset: session.paragraphSpans(story).find((span) => span.paraId === first)!.length }, "forward"),
  },
  "Enter, Undo, Redo": { joined: false, act: (session) => (session.undo(), session.redo()) },
  "Enter, Backspace, Undo, Redo": {
    joined: true,
    act: (session, story, _, second) => {
      session.deleteAt({ story, paraId: second, offset: 0 }, "backward");
      session.undo();
      session.redo();
    },
  },
};
/** Each field's code and shown text. */
const fieldsShown = (session: YrsSession, story: string) =>
  session
    .storySegments(story)
    .filter((g) => g.kind === "embed" && g.embedKind === "field")
    .map((g) => `${String(g.payload.instruction).trim()}=${String(g.payload.displayText)}`)
    .join(" | ");
test.each(
  (["body", "cell", "header"] as const).flatMap((where) =>
    Object.keys(splitShapes).flatMap((shape) => Object.keys(splitFlows).map((flow) => [where, shape, flow] as const))
  )
)("%s | %s: %s shows what the seed of the save shows", async (where, shape, flow) => {
  const [xml, text] = splitShapes[shape]!;
  const [story, part] = STORY[where];
  const bytes = matrixDocx(where, holder44(xml));
  const untouched = await open(bytes);
  const original = { units: matrixUnits(untouched, story), shown: fieldsShown(untouched, story) };
  const originalSave = await publish(bytes, untouched.encodeState());
  untouched.destroy();
  const session = await open(bytes);
  const at = matrixLocate(session, story, "AA");
  const { firstParaId, secondParaId } = session.splitParagraph({ story, paraId: at.paraId, offset: at.offset + 1 });
  splitFlows[flow]!.act(session, story, firstParaId, secondParaId);
  const saved = await publish(bytes, session.encodeState());
  const reopened = await open(saved);
  expect(fieldsShown(session, story)).toBe(fieldsShown(reopened, story));
  expect(storyText(saved, where).replaceAll("¶", "")).toBe(text);
  const ptabs = (file: Uint8Array) => partXml(file, part).split("<w:ptab ").length - 1;
  expect(ptabs(saved)).toBe(ptabs(bytes));
  if (splitFlows[flow]!.joined) {
    expect(fieldsShown(session, story)).toBe(original.shown);
    expect(matrixUnits(session, story)).toBe(original.units);
    expect(sig(saved, part)).toBe(sig(originalSave, part));
  } else expect(matrixUnits(reopened, story)).toBe(matrixUnits(session, story));
  session.destroy();
  reopened.destroy();
});

// Review of ee62d514, finding 1: once a split field's first paragraph loses its last projected link, the seed of the
// save no longer projects the field and shows its whole result, a nested field's included; the editor shows the same,
// and Undo brings back what it showed before.
test.each(
  (["body", "cell", "header"] as const).flatMap((where) => (["Backspace", "range delete"] as const).map((how) => [where, how] as const))
)("%s | [REF|[PAGE|7]L(AA)yy]: Enter, then a %s of the link left in the first paragraph, shows the whole result", async (where, how) => {
  const [story] = STORY[where];
  const bytes = matrixDocx(where, holder44(splitShapes["[REF|[PAGE|7]L(AA)yy]"]![0]));
  const session = await open(bytes);
  const at = matrixLocate(session, story, "AA");
  const { firstParaId } = session.splitParagraph({ story, paraId: at.paraId, offset: at.offset + 1 });
  const split = fieldsShown(session, story);
  session.addUndoBoundary();
  if (how === "Backspace") session.deleteAt({ story, paraId: firstParaId, offset: at.offset + 1 }, "backward");
  else session.deleteRange({ story, start: { paraId: firstParaId, offset: at.offset }, end: { paraId: firstParaId, offset: at.offset + 1 } });
  const saved = await publish(bytes, session.encodeState());
  const reopened = await open(saved);
  expect(fieldsShown(session, story)).toBe("REF a \\h=7");
  expect(fieldsShown(reopened, story)).toBe(fieldsShown(session, story));
  expect(storyText(saved, where)).toBe("a 7¶Ayy b¶");
  session.undo();
  expect(fieldsShown(session, story)).toBe(split);
  session.destroy();
  reopened.destroy();
});

// Review of ee62d514, finding 3: an empty run moved out by Enter goes back only with its neighbours, so a join that
// leaves the run after it in a later paragraph does not save it ahead of that run; joining all the paragraphs back
// restores the untouched save.
test.each(["body", "cell", "header"] as const)("%s | [REF|L(AA)y,z,<b/>]: Enter twice, then joins, keep the runs in order", async (where) => {
  const [story, part] = STORY[where];
  const bold = "<w:r><w:rPr><w:b/></w:rPr></w:r>";
  const bytes = matrixDocx(where, holder44(field(link(run("AA")) + run("y") + run("z") + bold, " REF a \\h ")));
  const untouched = await open(bytes);
  const original = await publish(bytes, untouched.encodeState());
  untouched.destroy();
  const order = (file: Uint8Array) => {
    const xml = partXml(file, part);
    return [xml.indexOf(">y</w:t>"), xml.indexOf(">z</w:t>"), xml.indexOf("<w:b/></w:rPr></w:r>")];
  };
  const session = await open(bytes);
  const at = matrixLocate(session, story, "AA");
  const { secondParaId } = session.splitParagraph({ story, paraId: at.paraId, offset: at.offset + 1 });
  session.splitParagraph({ story, paraId: secondParaId, offset: 2 });
  session.deleteAt({ story, paraId: secondParaId, offset: 0 }, "backward");
  const [y, z, b] = order(await publish(bytes, session.encodeState()));
  expect(y).toBeLessThan(z);
  if (b >= 0) expect(b).toBeGreaterThan(z);
  const third = session.paragraphSpans(story).map((span) => span.paraId)[session.paragraphSpans(story).findIndex((span) => span.paraId === at.paraId) + 1]!;
  session.deleteAt({ story, paraId: third, offset: 0 }, "backward");
  const joined = await publish(bytes, session.encodeState());
  expect(sig(joined, part)).toBe(sig(original, part));
  const [, joinedZ, joinedB] = order(joined);
  expect(joinedB).toBeGreaterThan(joinedZ!);
  session.destroy();
});

// Decided 2026-10-04: w:ptab round-trips, in ordinary text and in a field's result, drawn as a tab.
test.each(["body", "cell", "header"] as const)("%s | an untouched file keeps its positional tabs", async (where) => {
  const [story, part] = STORY[where];
  const bytes = matrixDocx(where, p("44444444", `${run("a")}${ptabRun}${run("b")}`) + holder44(field(link(run("AA")) + ptabRun + run("z"), " REF a \\h ")));
  const session = await open(bytes);
  const saved = await publish(bytes, session.encodeState());
  const tabs = session.storySegments(story).filter((g) => g.kind === "embed" && g.embedKind === "tab");
  session.destroy();
  expect(tabs.length).toBe(1);
  expect(partXml(saved, part).match(/<w:ptab [^>]*\/>/g)).toEqual([
    '<w:ptab w:relativeTo="margin" w:alignment="right" w:leader="dot"/>',
    '<w:ptab w:relativeTo="margin" w:alignment="right" w:leader="dot"/>',
  ]);
});

// Decision 2026-10-02: Backspace and Delete beside a field marker that shows nothing step over it and delete the
// visible unit past it; only a selection covering it removes the field. Each file has one such marker: a table of
// contents' own embed after its first entry, a REF field over links only, a field whose own result text Enter moved
// to the next paragraph, and a field over a nested simple field and the text after it.
const tocEntryXml = (anchor: string, text: string, page: string) =>
  `<w:hyperlink w:anchor="${anchor}">${run(text)}<w:r><w:tab/></w:r>${char("begin")}${instr(` PAGEREF ${anchor} \\h `)}${char("separate")}${run(page)}${char("end")}</w:hyperlink>`;
const markerFiles: Record<string, { xml: string; code: string; split?: boolean }> = {
  "TOC entry": {
    xml:
      p("44444444", `${char("begin")}${instr(" TOC \\o \\h ")}${char("separate")}${tocEntryXml("_Toc1", "Intro", "1")}`) +
      p("45454545", `${tocEntryXml("_Toc2", "Body", "2")}${char("end")}`),
    code: "TOC",
  },
  "REF over links": { xml: holder44(field(linkTo("AA") + linkTo("BB"), " REF a \\h ")), code: "REF" },
  "own text moved by Enter": { xml: holder44(field(link(run("AA")) + run("yz"), " REF a \\h ")), code: "REF", split: true },
  "nested simple field": { xml: holder44(field(fs(run("20"), " PAGE ") + run("y"))), code: "DATE" },
};
/** The story's saved text: characters, tabs as \t, ¶ per paragraph, the fixture's own paragraphs left out. */
const storyText = (bytes: Uint8Array, where: Where) =>
  [...partXml(bytes, STORY[where][1]).matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:tab\/>|<\/w:p>/g)]
    .map(([match, text]) => (text !== undefined ? text : match === "<w:tab/>" ? "\t" : "¶"))
    .join("")
    .replace(/(^|¶)(tail|head|foot|note|x)¶/g, "$1")
    .replace(/(^|¶)(tail|head|foot|note|x)¶/g, "$1");
/** Field codes and simple fields the story saves. */
const fieldsSaved = (bytes: Uint8Array, where: Where) => (sig(bytes, STORY[where][1]).match(/«[^»]*»|F\{/g) ?? []).sort().join(" | ");
/** The story offset of the marker: the embed of the field whose code holds `code`, showing nothing. */
function markerAt(session: YrsSession, story: string, code: string) {
  const ids = session.paragraphs(story).map((x) => x.paraId);
  let index = 0;
  let offset = 0;
  for (const g of session.storySegments(story)) {
    if (g.kind === "pilcrow") {
      index += 1;
      offset = 0;
      continue;
    }
    if (g.kind === "embed" && g.embedKind === "field" && String(g.payload.instruction ?? "").includes(code)) {
      expect(String(g.payload.displayText ?? "")).toBe("");
      return { story, paraId: ids[index]!, offset };
    }
    offset += g.kind === "text" ? g.text.length : 1;
  }
  throw new Error(`no ${code} marker in ${story}`);
}
type MarkerKey = "Backspace after" | "Delete before" | "Backspace before" | "Delete after";
/** What each key at the marker leaves: the visible unit past it goes, the field stays. */
const markerExpect: Record<string, Record<MarkerKey, { text: string; gone?: string }>> = {
  "TOC entry": {
    "Backspace after": { text: "Intro\t¶Body\t2¶", gone: "«PAGEREF _Toc1 \\h»" },
    "Delete before": { text: "Intro\t1Body\t2¶" },
    "Backspace before": { text: "Intro\t¶Body\t2¶", gone: "«PAGEREF _Toc1 \\h»" },
    "Delete after": { text: "Intro\t1Body\t2¶" },
  },
  "REF over links": {
    "Backspace after": { text: "a AAB b¶" },
    "Delete before": { text: "a AABBb¶" },
    "Backspace before": { text: "a AAB b¶" },
    "Delete after": { text: "a AABBb¶" },
  },
  "own text moved by Enter": {
    "Backspace after": { text: "a ¶Ayz b¶" },
    "Delete before": { text: "a AAyz b¶" },
    "Backspace before": { text: "a ¶Ayz b¶" },
    "Delete after": { text: "a AAyz b¶" },
  },
  "nested simple field": {
    "Backspace after": { text: "a 20 b¶" },
    "Delete before": { text: "a 20yb¶" },
    "Backspace before": { text: "a 20 b¶" },
    "Delete after": { text: "a 20yb¶" },
  },
};
const pressAtMarker = (session: YrsSession, story: string, file: string, key: MarkerKey) => {
  if (markerFiles[file]!.split)
    session.splitParagraph({ story, paraId: "44444444", offset: textAt(session, "AA").offset + 1 });
  const at = markerAt(session, story, markerFiles[file]!.code);
  const after = key.endsWith("after");
  session.deleteAt({ ...at, offset: at.offset + (after ? 1 : 0) }, key.startsWith("Backspace") ? "backward" : "forward");
};
test.each(
  (["body", "cell", "header"] as const).flatMap((where) =>
    Object.keys(markerFiles).flatMap((file) =>
      (Object.keys(markerExpect[file]!) as MarkerKey[]).map((key) => [where, file, key] as const)
    )
  )
)("%s | %s: %s an invisible field marker deletes the visible unit past it and keeps the field", async (where, file, key) => {
  const bytes = matrixDocx(where, markerFiles[file]!.xml);
  const story = STORY[where][0];
  const untouched = await open(bytes);
  const before = fieldsSaved(await publish(bytes, untouched.encodeState()), where);
  untouched.destroy();
  const session = await open(bytes);
  pressAtMarker(session, story, file, key);
  const editor = matrixUnits(session, story);
  const saved = await publish(bytes, session.encodeState());
  session.destroy();
  const { text, gone } = markerExpect[file]![key];
  expect(storyText(saved, where)).toBe(text);
  expect(fieldsSaved(saved, where)).toBe(before.split(" | ").filter((code) => code !== gone).join(" | "));
  const reopened = await open(saved);
  // Joining two table of contents entries leaves the editor's TOC continued where the reopened file holds it in one
  // paragraph (as Delete at the first entry's end always has); the save is the same.
  if (!(file === "TOC entry" && key.startsWith("Delete"))) expect(matrixUnits(reopened, story)).toBe(editor);
  reopened.destroy();
});

// Two peers: one presses Backspace after the marker while the other types at the end of the link before it; the
// field stays once, with the typed character and without the deleted one.
test("Backspace after an invisible field marker beside a peer's typing keeps the field once", async () => {
  const bytes = matrixDocx("body", markerFiles["REF over links"]!.xml);
  const A = await open(bytes);
  const B = await open(bytes);
  const at = markerAt(A, "body", "REF");
  A.deleteAt({ ...at, offset: at.offset + 1 }, "backward");
  B.insertText({ ...at, offset: at.offset - 1 }, "W");
  const [toB, toA] = [A.encodeStateAsUpdate(B.encodeStateVector()), B.encodeStateAsUpdate(A.encodeStateVector())];
  A.applyUpdate(toA);
  B.applyUpdate(toB);
  expect(matrixUnits(A, "body")).toBe(matrixUnits(B, "body"));
  const saved = await publish(bytes, A.encodeState());
  A.destroy();
  B.destroy();
  expect(storyText(saved, "body")).toBe("a AABW b¶");
  expect(fieldsSaved(saved, "body")).toBe("«REF a \\h»");
});

// Follow-up item 4 (decided 2026-10-02, line breaks and comment references 2026-10-04): a field whose result continues
// into the next paragraph seeds the runs that end its result in its first paragraph as text after its embed, so text
// typed at that paragraph's end is result text in the editor, where the save puts it, and the reopened file reads the
// same. Shapes: a DATE result split over two paragraphs, a table of contents without links (entries as plain runs), a
// tail run holding a line break, and a tail ending in a comment's reference.
// [name, body, field code, a text of the untouched save, the saved text after Backspace at the first paragraph's end]
const continuedResults = [
  ["a DATE result over two paragraphs", p("11111111", `${run("a")}${char("begin")}${instr(" DATE ")}${char("separate")}${run("2")}`) + p("33333333", `${run("0")}${char("end")}${run("z")}`), " DATE ", ">2</w:t>", "a0ztail"],
  [
    "a table of contents without links",
    p("11111111", `${char("begin")}${instr(" TOC \\o ")}${char("separate")}${run("Intro")}<w:r><w:tab/></w:r>${run("1")}`) + p("33333333", `${run("Body")}<w:r><w:tab/></w:r>${run("2")}${char("end")}`),
    " TOC \\o ",
    ">Intro</w:t>",
    "IntroBody2tail",
  ],
  [
    "a result ending in a line break and text",
    p("11111111", `${run("x")}${char("begin")}${instr(" INCLUDETEXT x ")}${char("separate")}${run("A")}<w:r><w:t>B</w:t><w:br/><w:t>C</w:t></w:r>`) + p("33333333", `${run("D")}${char("end")}`),
    " INCLUDETEXT x ",
    'w:type="textWrapping"/>',
    "xABDtail",
  ],
  [
    "a result ending in a comment's reference",
    p("11111111", `${commentStart(1)}${run("x")}${char("begin")}${instr(" DOCVARIABLE v ")}${char("separate")}${run("AB")}${commentEnd(1)}${commentRef(1)}`) + p("33333333", `${run("EF")}${char("end")}${run("y")}`),
    " DOCVARIABLE v ",
    ">AB</w:t>",
    "xAEFytail",
  ],
] as const;
/** The saved document's text runs, joined. */
const savedText = (bytes: Uint8Array) => [...documentXml(bytes).matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map(([, text]) => text).join("");
test.each(continuedResults)("Backspace at the end of %s's first paragraph deletes one character and keeps the field", async (_, body, code, __, text) => {
  const bytes = matrixDocx("body", body);
  const session = await open(bytes);
  const { length } = session.paragraphSpans("body").find((span) => span.paraId === "11111111")!;
  session.deleteAt({ story: "body", paraId: "11111111", offset: length }, "backward");
  const editor = matrixUnits(session, "body");
  const saved = await publish(bytes, session.encodeState());
  session.destroy();
  const reopened = await open(saved);
  expect(matrixUnits(reopened, "body")).toBe(editor);
  reopened.destroy();
  const xml = documentXml(saved);
  expect(xml).toContain(code);
  expect(xml).toContain('w:fldCharType="end"');
  expect(savedText(saved)).toBe(text);
});
test.each(continuedResults)("text typed at the end of %s's first paragraph shows where the save puts it", async (_, body, __, kept) => {
  const bytes = matrixDocx("body", body);
  const untouched = await open(bytes);
  const original = documentXml(await publish(bytes, untouched.encodeState()));
  untouched.destroy();
  // Untouched, the save keeps the source's text.
  expect(original).toContain(kept);
  const session = await open(bytes);
  const { length } = session.paragraphSpans("body").find((span) => span.paraId === "11111111")!;
  session.insertText({ story: "body", paraId: "11111111", offset: length }, "Q");
  const editor = matrixUnits(session, "body");
  const saved = await publish(bytes, session.encodeState());
  session.destroy();
  const reopened = await open(saved);
  expect(matrixUnits(reopened, "body")).toBe(editor);
  reopened.destroy();
  // Inside the result: after the separate, before the end in the next paragraph.
  const xml = documentXml(saved);
  expect(xml.indexOf("Q</w:t>")).toBeGreaterThan(xml.indexOf('w:fldCharType="separate"'));
  expect(xml.indexOf("Q</w:t>")).toBeLessThan(xml.indexOf('w:fldCharType="end"'));
});

// Follow-up item 4: Enter inside the link of a field whose result continues into the next paragraph, then the
// rewriting join, restores the field with its tail text after the embed (the join stops at the tail instead of
// leaving the link split).
test.each(["Backspace", "Delete"] as const)("Enter in the link of a field whose result continues, then %s, restores it", async (key) => {
  const bytes = docx(
    p("11111111", `${run("a ")}${char("begin")}${instr(" REF a \\h ")}${char("separate")}${link(run("AA"))}${run("yy")}`) +
      p("33333333", `${run("zz")}${char("end")}${run(" b")}`) +
      tail
  );
  const before = await open(bytes);
  const original = matrixUnits(before, "body");
  const untouched = await publish(bytes, before.encodeState());
  before.destroy();
  const session = await open(bytes);
  const { firstParaId, secondParaId } = session.splitParagraph({ story: "body", paraId: "11111111", offset: childAt(session, "AA") + 1 });
  if (key === "Backspace") session.deleteAt({ story: "body", paraId: secondParaId, offset: 0 }, "backward");
  else session.deleteAt({ story: "body", paraId: firstParaId, offset: session.paragraphSpans("body").find((span) => span.paraId === firstParaId)!.length }, "forward");
  expect(matrixUnits(session, "body")).toBe(original);
  const saved = await publish(bytes, session.encodeState());
  session.destroy();
  expect(view(saved)).toBe("a [«REF a \\h»|H(AA)yy");
  expect(documentXml(saved)).toBe(documentXml(untouched));
});

// Follow-up item 4, two peers: both type at the end of the first paragraph of a field whose result continues; both
// texts stay, in the editors and in the reopened save, as ordinary typing.
test("two peers typing at the end of a continued field result's first paragraph both keep their text", async () => {
  const bytes = docx(p("11111111", `${run("a")}${char("begin")}${instr(" DATE ")}${char("separate")}${run("2")}`) + p("33333333", `${run("0")}${char("end")}${run("z")}`) + tail);
  const A = await open(bytes);
  const B = await open(bytes);
  const end = (session: YrsSession) => ({ story: "body", paraId: "11111111", offset: session.paragraphSpans("body").find((span) => span.paraId === "11111111")!.length });
  A.insertText(end(A), "Q");
  B.insertText(end(B), "R");
  const [toB, toA] = [A.encodeStateAsUpdate(B.encodeStateVector()), B.encodeStateAsUpdate(A.encodeStateVector())];
  A.applyUpdate(toA);
  B.applyUpdate(toB);
  expect(matrixUnits(A, "body")).toBe(matrixUnits(B, "body"));
  const saved = await publish(bytes, A.encodeState());
  const reopened = await open(saved);
  expect(matrixUnits(reopened, "body")).toBe(matrixUnits(A, "body"));
  expect(view(saved)).toMatch(/^a\[«DATE»\|2(QR|RQ)$/);
  for (const session of [A, B, reopened]) session.destroy();
});

// Round 4 review, finding 1: keys pressed slowly (each its own Undo step) that delete a continued field's end, then
// as many Undos, put the end back where it was. The editor, a fresh replica and the save equal the untouched file: an
// Undo re-anchors only a marker whose own text it restored, not one next to text it restored.
const slowEndShapes: Record<string, [string, string, number | "end"]> = {
  "a table of contents with links, its end after the last entry": [
    p("11111111", `${char("begin")}${instr(' TOC \\o "1-3" \\h ')}${char("separate")}${tocEntryXml("_Toc1", "Intro", "1")}`) +
      p("33333333", `${tocEntryXml("_Toc2", "Body", "2")}${char("end")}`),
    "33333333",
    "end",
  ],
  "an IF result, its end after text": [
    p("11111111", `${run("Dear ")}${char("begin")}${instr(" IF 1 = 1 ")}${char("separate")}${run("Sir,")}`) + p("33333333", `${run("Thanks")}${char("end")}${run(" ok")}`),
    "33333333",
    6,
  ],
  "a DATE result, its end after text": [
    p("11111111", `${run("a")}${char("begin")}${instr(" DATE ")}${char("separate")}${run("2")}`) + p("33333333", `${run("lo wo")}${char("end")}${run("zz")}`),
    "33333333",
    5,
  ],
};
test.each(
  (["body", "cell"] as const).flatMap((where) =>
    Object.keys(slowEndShapes).flatMap((shape) =>
      (["Backspace", "Delete"] as const).flatMap((key) => [2, 3].map((times) => [where, shape, key, times] as const))
    )
  )
)("%s | %s: %s ×%i slowly, then as many Undos, restores the field's end", async (where, shape, key, times) => {
  const [xml, paraId, at] = slowEndShapes[shape]!;
  const [story, part] = STORY[where];
  const bytes = matrixDocx(where, xml);
  const untouched = await open(bytes);
  const seeded = matrixUnits(untouched, story);
  const original = sig(await publish(bytes, untouched.encodeState()), part);
  untouched.destroy();
  const session = await open(bytes);
  session.beginUndoCapture();
  const end = () => session.paragraphSpans(story).find((span) => span.paraId === paraId)!.length;
  let caret = at === "end" ? end() : at;
  for (let i = 0; i < times; i += 1) {
    if (i > 0) session.addUndoBoundary();
    if (key === "Backspace") caret = session.deleteAt({ story, paraId, offset: caret }, "backward").caret.offset;
    else session.deleteAt({ story, paraId, offset: (at === "end" ? end() : at) - 1 }, "forward");
  }
  for (let i = 0; i < times; i += 1) session.undo();
  const replica = await open(bytes, session.encodeState());
  const saved = await publish(bytes, session.encodeState());
  expect(matrixUnits(session, story)).toBe(seeded);
  expect(matrixUnits(replica, story)).toBe(seeded);
  expect(sig(saved, part)).toBe(original);
  session.destroy();
  replica.destroy();
});

// Each delete step its own Undo step or one, then all undone: the end stays put (round 4 review, fe4's flows; the
// single-step ones are what re-anchoring the end fixed in this round).
const fe4Flows: Record<string, (session: YrsSession, story: string) => void> = {
  "del [1,2), del [0,1), Undo ×2": (s, st) => (cut(s, st, 1, 2), s.addUndoBoundary(), cut(s, st, 0, 1), s.undo(), s.undo()),
  "del [0,2), del [0,1), Undo ×2": (s, st) => (cut(s, st, 0, 2), s.addUndoBoundary(), cut(s, st, 0, 1), s.undo(), s.undo()),
  "del [1,2) twice, Undo ×2": (s, st) => (cut(s, st, 1, 2), s.addUndoBoundary(), cut(s, st, 1, 2), s.undo(), s.undo()),
  "del [0,3), Undo": (s, st) => (cut(s, st, 0, 3), s.undo()),
  "del [0,3), Undo, Redo, Undo": (s, st) => (cut(s, st, 0, 3), s.undo(), s.redo(), s.undo()),
  "del [1,3), del [0,1), Undo ×2": (s, st) => (cut(s, st, 1, 3), s.addUndoBoundary(), cut(s, st, 0, 1), s.undo(), s.undo()),
  "del [0,1), del [0,2), Undo ×2": (s, st) => (cut(s, st, 0, 1), s.addUndoBoundary(), cut(s, st, 0, 2), s.undo(), s.undo()),
};
/** Deletes `start..end` of the second paragraph. */
function cut(session: YrsSession, story: string, start: number, end: number) {
  session.deleteRange({ story, start: { paraId: "33333333", offset: start }, end: { paraId: "33333333", offset: end } });
}
test.each(
  (["body", "cell"] as const).flatMap((where) =>
    ["date", "link"].flatMap((shape) => Object.keys(fe4Flows).map((flow) => [where, shape, flow] as const))
  )
)("%s | a continued %s field: %s keeps its end", async (where, shape, flow) => {
  const xml =
    shape === "date"
      ? p("11111111", `${run("a")}${char("begin")}${instr(" DATE ")}${char("separate")}${run("2")}`) + p("33333333", `${run("10")}${char("end")}${run("zz")}`)
      : p("11111111", `${run("a ")}${char("begin")}${instr(" REF a \\h ")}${char("separate")}${link(run("AA"))}`) + p("33333333", `${run("yy")}${char("end")}${run(" b")}`);
  const [story, part] = STORY[where];
  const bytes = matrixDocx(where, xml);
  const untouched = await open(bytes);
  const seeded = matrixUnits(untouched, story);
  const original = sig(await publish(bytes, untouched.encodeState()), part);
  untouched.destroy();
  const session = await open(bytes);
  session.beginUndoCapture();
  fe4Flows[flow]!(session, story);
  const replica = await open(bytes, session.encodeState());
  const saved = await publish(bytes, session.encodeState());
  expect(matrixUnits(session, story)).toBe(seeded);
  expect(matrixUnits(replica, story)).toBe(seeded);
  expect(sig(saved, part)).toBe(original);
  session.destroy();
  replica.destroy();
});

// Round 4 review, finding 3: Delete in front of an invisible marker that ends a story's last paragraph steps over it
// and finds nothing to delete, as Delete at a story's end does: nothing changes and nothing throws (a throw left the
// editor's input queue failed). Backspace after one that opens a story's first paragraph likewise.
const tcField = `${char("begin")}${instr(" TC x ")}${char("end")}`;
test.each([
  ["body", "a comment's reference ending the document", `${commentStart(5)}${run("last word")}${commentEnd(5)}${commentRef(5)}`],
  ["cell", "a TC field ending a cell", `${tcField}${run("ab")}${tcField}`],
  ["header", "a TC field ending a header", `${tcField}${run("ab")}${tcField}`],
  ["footnote", "a TC field ending a footnote", `${tcField}${run("ab")}${tcField}`],
] as const)("%s | %s: Delete before it and Backspace at the start change nothing", async (where, _, xml) => {
  const story = STORY[where][0];
  // In the body the marker ends the document, so the content replaces its last paragraph.
  const bytes = where === "body" ? docx(p("22222222", xml)) : matrixDocx(where, p("44444444", xml));
  const paraId = where === "body" ? "22222222" : "44444444";
  const session = await open(bytes);
  const before = matrixUnits(session, story);
  const length = session.paragraphSpans(story).find((span) => span.paraId === paraId)!.length;
  expect(() => session.deleteAt({ story, paraId, offset: length - 1 }, "forward")).not.toThrow();
  if (where === "cell" || where === "header") expect(() => session.deleteAt({ story, paraId, offset: 1 }, "backward")).not.toThrow();
  expect(matrixUnits(session, story)).toBe(before);
  // The session takes the next key as usual.
  session.insertText({ story, paraId, offset: length - 1 }, "Z");
  expect(matrixUnits(session, story)).toBe(before.replace(where === "body" ? "word" : "ab", where === "body" ? "wordZ" : "abZ"));
  session.destroy();
});

// Round 4 review, finding 4: a field whose result continues keeps in its own data any run that is more than text, tabs,
// line breaks and comment references (a page break), and its characters keep the formatting its whole result lent
// them, so an untouched file saves as it did before the result's tail became text. Since 2026-10-04 a comment's
// reference in the tail is tail text and loses its CommentReference style on save, as references elsewhere do.
const untouchedTails: Record<string, [string, string]> = {
  "a page break in the tail": [
    p("11111111", `${run("x")}${char("begin")}${instr(" INCLUDETEXT x ")}${char("separate")}${run("A")}<w:r><w:br w:type="page"/><w:t>B</w:t></w:r>`) + p("33333333", `${run("C")}${char("end")}`),
    '<w:p w14:paraId="11111111"><w:r><w:t>x</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> INCLUDETEXT x </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>A</w:t></w:r><w:r><w:br w:type="page"/><w:t>B</w:t></w:r></w:p>',
  ],
  "a page break before plain text in the tail": [
    p("11111111", `${run("x")}${char("begin")}${instr(" INCLUDETEXT x ")}${char("separate")}${run("A")}<w:r><w:br w:type="page"/><w:t>B</w:t></w:r>${run("D")}`) + p("33333333", `${run("C")}${char("end")}`),
    '<w:p w14:paraId="11111111"><w:r><w:t>x</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> INCLUDETEXT x </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>A</w:t></w:r><w:r><w:br w:type="page"/><w:t>B</w:t></w:r><w:r><w:t>D</w:t></w:r></w:p>',
  ],
  "a bibliography entry with an italic title": [
    `<w:p w14:paraId="11111111"><w:pPr><w:pStyle w:val="Bibliography"/></w:pPr>${char("begin")}${instr(" BIBLIOGRAPHY ")}${char("separate")}${run("Smith, J. (2020). ")}<w:r><w:rPr><w:i/><w:iCs/></w:rPr><w:t>Title.</w:t></w:r>${run(" Pub.")}</w:p>` +
      p("33333333", `${run("Doe, A. (2021). ")}${char("end")}`),
    '<w:p w14:paraId="11111111"><w:pPr><w:pStyle w:val="Bibliography"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> BIBLIOGRAPHY </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Smith, J. (2020). </w:t></w:r><w:r><w:rPr><w:i/><w:iCs/></w:rPr><w:t>Title.</w:t></w:r><w:r><w:t> Pub.</w:t></w:r></w:p>',
  ],
  "a comment ending with its reference in the tail": [
    p("11111111", `${commentStart(1)}${run("x")}${char("begin")}${instr(" DOCVARIABLE v ")}${char("separate")}${run("AB")}${commentEnd(1)}${commentRef(1)}${run("CD")}`) + p("33333333", `${run("EF")}${char("end")}${run("y")}`),
    '<w:p w14:paraId="11111111"><w:commentRangeStart w:id="1"/><w:r><w:t>x</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> DOCVARIABLE v </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>AB</w:t></w:r><w:r><w:commentReference w:id="1"/></w:r><w:r><w:t>CD</w:t></w:r><w:commentRangeEnd w:id="1"/></w:p>',
  ],
};
test.each(Object.keys(untouchedTails))("an untouched file with %s saves as before", async (shape) => {
  const [xml, expected] = untouchedTails[shape]!;
  const bytes = matrixDocx("body", xml);
  const session = await open(bytes);
  const saved = documentXml(await publish(bytes, session.encodeState()));
  session.destroy();
  expect(saved.replaceAll(' xml:space="preserve"', "").match(/<w:p w14:paraId="11111111">[\s\S]*?<\/w:p>/)![0]).toBe(expected);
});

// Decided 2026-10-04: a line break (w:br, w:cr) in the runs ending a continued result's first paragraph is tail text
// too. An untouched save may regroup those runs and write the break's w:type="textWrapping": it equals the source once
// runs are merged and breaks written alike.
const lineBreakTails: Record<string, [string, string]> = {
  "a line break inside the tail's last run": [
    p("11111111", `${run("x")}${char("begin")}${instr(" INCLUDETEXT x ")}${char("separate")}${run("A")}<w:r><w:t>B</w:t><w:br/><w:t>C</w:t></w:r>`) + p("33333333", `${run("D")}${char("end")}`),
    '<w:p w14:paraId="11111111"><w:r><w:t>x</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> INCLUDETEXT x </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>A</w:t></w:r><w:r><w:t>B</w:t><w:br/><w:t>C</w:t></w:r></w:p>',
  ],
  "a carriage return ending the tail": [
    p("11111111", `${run("x")}${char("begin")}${instr(" INCLUDETEXT x ")}${char("separate")}${run("A")}<w:r><w:cr/></w:r>`) + p("33333333", `${run("D")}${char("end")}`),
    '<w:p w14:paraId="11111111"><w:r><w:t>x</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> INCLUDETEXT x </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>A</w:t></w:r><w:r><w:br/></w:r></w:p>',
  ],
};
/** A paragraph's XML with runs merged and breaks written alike. */
const runsMerged = (xml: string) =>
  xml
    .replaceAll(' xml:space="preserve"', "")
    .replaceAll('<w:br w:type="textWrapping"/>', "<w:br/>")
    .replaceAll("<w:cr/>", "<w:br/>")
    .replaceAll("</w:r><w:r>", "")
    .replaceAll("</w:t><w:t>", "");
test.each(Object.keys(lineBreakTails))("an untouched file with %s saves its runs as before", async (shape) => {
  const [xml, expected] = lineBreakTails[shape]!;
  const bytes = matrixDocx("body", xml);
  const session = await open(bytes);
  const saved = documentXml(await publish(bytes, session.encodeState()));
  session.destroy();
  expect(runsMerged(saved.match(/<w:p w14:paraId="11111111">[\s\S]*?<\/w:p>/)![0])).toBe(runsMerged(expected));
});

// Round 4 review, finding 5: the projector (documentToYrs) reads a continued field's result tail as the engine seed
// does, its field data included, in every story.
test.each(
  (["body", "cell", "header", "footnote"] as const).flatMap((where) =>
    [
      ["a DATE result", p("11111111", `${run("a")}${char("begin")}${instr(" DATE ")}${char("separate")}${run("2")}`) + p("33333333", `${run("0")}${char("end")}${run("z")}`)],
      ["a link then own text", p("11111111", `${run("a ")}${char("begin")}${instr(" REF a \\h ")}${char("separate")}${link(run("AA"))}${run("yy")}`) + p("33333333", `${run("zz")}${char("end")}${run(" b")}`)],
      ...Object.entries(untouchedTails).map(([name, [xml]]) => [name, xml] as const),
      ["a bold tab first in the tail", p("11111111", `${char("begin")}${instr(" DOCVARIABLE v ")}${char("separate")}<w:r><w:rPr><w:b/></w:rPr><w:tab/></w:r>${run("A")}`) + p("33333333", `${run("D")}${char("end")}`)],
      ...Object.entries(lineBreakTails).map(([name, [xml]]) => [name, xml] as const),
    ].map(([name, xml]) => [where, name, xml] as const)
  )
)("%s | %s: the projector and the engine seed agree", async (where, _, xml) => {
  expect(await parity(matrixDocx(where, xml), STORY[where][0])).toBe(true);
});

// Round 4 recheck (decided 2026-10-03): a tab run with its own formatting ending a continued result stays in the
// field, as do the runs before it, so an untouched save keeps the tab's formatting (the save drops a tab unit's own
// formatting). Expected: the save before the result's tail became text (dcf7d9b5 through 86744da3), byte for byte.
const plainRun = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
test.each(
  ["<w:b/>", `<w:u w:val="single"/>`, `<w:highlight w:val="yellow"/>`, `<w:color w:val="FF0000"/>`].flatMap((rpr) =>
    ["first", "last", "between"].map((place) => [rpr, place] as const)
  )
)("an untouched continued result with a %s tab %s in its tail saves as before", async (rpr, place) => {
  const tab = `<w:r><w:rPr>${rpr}</w:rPr><w:tab/></w:r>`;
  const tailXml = place === "first" ? `${tab}${plainRun("A")}` : place === "last" ? `${plainRun("A")}${tab}` : `${plainRun("A")}${tab}${plainRun("B")}`;
  const bytes = matrixDocx("body", p("11111111", `${char("begin")}${instr(" DOCVARIABLE v ")}${char("separate")}${tailXml}`) + p("33333333", `${plainRun("D")}${char("end")}`));
  const session = await open(bytes);
  const saved = documentXml(await publish(bytes, session.encodeState()));
  session.destroy();
  // The field's characters take its result's first run's formatting.
  const own = place === "first" ? `<w:rPr>${rpr}</w:rPr>` : "";
  expect(saved.match(/<w:p w14:paraId="11111111">[\s\S]*?<\/w:p>/)![0]).toBe(
    `<w:p w14:paraId="11111111"><w:r>${own}<w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DOCVARIABLE v </w:instrText></w:r><w:r>${own}<w:fldChar w:fldCharType="separate"/></w:r>${tailXml}</w:p>`
  );
});
