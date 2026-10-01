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
