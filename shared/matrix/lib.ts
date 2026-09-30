// Shared machinery for the DOCX edit/publish/rebase matrices. See README.md.
import { createHash } from "node:crypto";
import { parseDocx } from "../../packages/docx/src/docx";
import { rezipContainer, unzipContainer } from "../../packages/docx/src/wasm/opc";
import { createYrsSession, type YrsSession } from "../../packages/docx/src/yrs";
import { documentToYrs } from "../../packages/docx/src/yrs/documentToYrs";
import { exportOffice, rebaseOffice, seedOffice } from "../office-checkpoint";
import { RebaseError } from "../office-rebase";

// ---------------------------------------------------------------- XML builders

export const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
export const deleted = (text: string) => `<w:r><w:delText xml:space="preserve">${text}</w:delText></w:r>`;
export const p = (id: string | null, xml: string, ppr = "") =>
  `<w:p${id ? ` w14:paraId="${id}"` : ""}>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ""}${xml}</w:p>`;
export const PB = `<w:r><w:br w:type="page"/></w:r>`;
export const CB = `<w:r><w:br w:type="column"/></w:r>`;
const change = (tag: string, xml: string, id: number) =>
  `<w:${tag} w:id="${id}" w:author="A" w:date="2026-09-01T00:00:00Z">${xml}</w:${tag}>`;
export const ins = (xml: string, id = 90) => change("ins", xml, id);
export const del = (xml: string, id = 91) => change("del", xml, id);
export const moveTo = (xml: string, id = 92) => change("moveTo", xml, id);
export const moveFrom = (xml: string, id = 93) => change("moveFrom", xml, id);
export const table = (xml: string) =>
  `<w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc>${xml}</w:tc></w:tr></w:tbl>`;
export const S = (id = 1) => `<w:commentRangeStart w:id="${id}"/>`;
export const E = (id = 1) => `<w:commentRangeEnd w:id="${id}"/>`;
export const ref = (id = 1) =>
  `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`;
export const bm = (xml: string, id = 5, name = "mark") =>
  `<w:bookmarkStart w:id="${id}" w:name="${name}"/>${xml}<w:bookmarkEnd w:id="${id}"/>`;
export const link = (xml: string, anchor = "target") => `<w:hyperlink w:anchor="${anchor}">${xml}</w:hyperlink>`;
/** A link to an external URL through the part's own relationship rId50 (each part names its own URL). */
export const webLink = (xml: string) => `<w:hyperlink r:id="rId50">${xml}</w:hyperlink>`;
export const sdt = (xml: string, id = 8) =>
  `<w:sdt><w:sdtPr><w:id w:val="${id}"/></w:sdtPr><w:sdtContent>${xml}</w:sdtContent></w:sdt>`;
/** A simple field (`w:fldSimple`). */
export const fs = (xml: string, instr = " DATE ") => `<w:fldSimple w:instr="${instr}">${xml}</w:fldSimple>`;
const fldChar = (type: string) => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
export const instr = (text: string) => `<w:r><w:instrText xml:space="preserve">${text}</w:instrText></w:r>`;
/** A complex field; `code` is the markup between begin and separate. */
export const cf = (code: string, result: string) =>
  `${fldChar("begin")}${code}${fldChar("separate")}${result}${fldChar("end")}`;
export const field = (result: string, code = " DATE ") => cf(instr(code), result);
export const spaceBefore = `<w:spacing w:before="480"/>`;

// ---------------------------------------------------------------- packages

const W =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
const OFFICE = "application/vnd.openxmlformats-officedocument.wordprocessingml";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** Every story kind: its id in the session and the part it saves to. */
export const STORY = {
  body: ["body", "word/document.xml"],
  control: ["body:sdt0", "word/document.xml"],
  cell: ["body:t0:r0c0", "word/document.xml"],
  nested: ["body:t0:r0c0:t0:r0c0", "word/document.xml"],
  header: ["hf:rId20", "word/header1.xml"],
  headerCell: ["hf:rId20:t0:r0c0", "word/header1.xml"],
  footer: ["hf:rId21", "word/footer1.xml"],
  footnote: ["fn:1", "word/footnotes.xml"],
  endnote: ["en:1", "word/endnotes.xml"],
} as const;
export type Where = keyof typeof STORY;
export const ALL = Object.keys(STORY) as Where[];

/** The body's last paragraph; `typeTail` edits it. */
export const TAIL = "22222222";

/** A package holding `content` in story `where`, then the tail paragraph in the body. */
export function docx(where: Where, content: string): Uint8Array {
  const body =
    where === "body"
      ? content
      : where === "control"
      ? sdt(content, 7)
      : where === "cell"
      ? table(content)
      : where === "nested"
      ? table(`${table(content)}${p("99999999", "")}`)
      : where === "footnote"
      ? p("11111111", `${run("x")}<w:r><w:footnoteReference w:id="1"/></w:r>`)
      : where === "endnote"
      ? p("11111111", `${run("x")}<w:r><w:endnoteReference w:id="1"/></w:r>`)
      : "";
  const header =
    where === "header" ? content : where === "headerCell" ? `${table(content)}${p("66666665", "")}` : p("66666666", run("head"));
  const note = (kind: "footnote" | "endnote", mine: boolean, id: string) =>
    `<w:${kind}s ${W}><w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}><w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}><w:${kind} w:id="1">${
      mine ? content : p(id, run("note"))
    }</w:${kind}></w:${kind}s>`;
  const web = (name: string) =>
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId50" Type="${REL}/hyperlink" Target="https://example.com/${name}" TargetMode="External"/></Relationships>`;
  const parts: Record<string, string> = {
    "word/_rels/header1.xml.rels": web("header"),
    "word/_rels/footer1.xml.rels": web("footer"),
    "word/_rels/footnotes.xml.rels": web("footnotes"),
    "word/_rels/endnotes.xml.rels": web("endnotes"),
    "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="${OFFICE}.comments+xml"/><Override PartName="/word/header1.xml" ContentType="${OFFICE}.header+xml"/><Override PartName="/word/footer1.xml" ContentType="${OFFICE}.footer+xml"/><Override PartName="/word/footnotes.xml" ContentType="${OFFICE}.footnotes+xml"/><Override PartName="/word/endnotes.xml" ContentType="${OFFICE}.endnotes+xml"/></Types>`,
    "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
    "word/_rels/document.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="${REL}/comments" Target="comments.xml"/><Relationship Id="rId20" Type="${REL}/header" Target="header1.xml"/><Relationship Id="rId21" Type="${REL}/footer" Target="footer1.xml"/><Relationship Id="rId22" Type="${REL}/footnotes" Target="footnotes.xml"/><Relationship Id="rId23" Type="${REL}/endnotes" Target="endnotes.xml"/><Relationship Id="rId50" Type="${REL}/hyperlink" Target="https://example.com/document" TargetMode="External"/></Relationships>`,
    "word/document.xml": `<w:document ${W}><w:body>${body}${p(TAIL, run("tail"))}<w:sectPr><w:headerReference w:type="default" r:id="rId20"/><w:footerReference w:type="default" r:id="rId21"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
    // Source comment 1 keeps the comments part, which a rebase needs the export to keep.
    "word/comments.xml": `<w:comments ${W}><w:comment w:id="1" w:author="Src" w:date="2026-09-01T00:00:00Z"><w:p><w:r><w:t>c1</w:t></w:r></w:p></w:comment></w:comments>`,
    "word/header1.xml": `<w:hdr ${W}>${header}</w:hdr>`,
    "word/footer1.xml": `<w:ftr ${W}>${where === "footer" ? content : p("77777777", run("foot"))}</w:ftr>`,
    "word/footnotes.xml": note("footnote", where === "footnote", "88888888"),
    "word/endnotes.xml": note("endnote", where === "endnote", "88888889"),
  };
  return rezipContainer(Object.fromEntries(Object.entries(parts).map(([k, v]) => [k, new TextEncoder().encode(v)])));
}

// ---------------------------------------------------------------- sessions

const fixed = { seed: "0".repeat(64), now: "2026-09-29T00:00:00.000Z" };
// Client ids are reset for every row: editor comment ids and concurrent-insert order derive from them,
// so a row's result must not depend on the rows before it.
let clientId = 0;
const nextClient = () => (clientId += 1);

export const prime = () => seedOffice("docx", docx("body", p("11111111", run("x"))));

export async function open(bytes: Uint8Array, state?: Uint8Array): Promise<YrsSession> {
  const session = await createYrsSession({ clientId: nextClient() });
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

export const publish = (base: Uint8Array, state: Uint8Array) => exportOffice(base, checkpoint(base, state), fixed);

export const partXml = (bytes: Uint8Array, part: string) => {
  const data = unzipContainer(bytes)[part];
  return data ? new TextDecoder().decode(data) : "";
};

// ---------------------------------------------------------------- reading a saved story

const SIG =
  /<(\/?)w:(p|br|t|delText|instrText|ins|del|moveTo|moveFrom|tbl|commentRangeStart|commentRangeEnd|commentReference|bookmarkStart|bookmarkEnd|fldSimple|fldChar|hyperlink|sdt|oMath|smartTag|customXml|sectPr)\b([^>]*?)(\/?)>([^<]*)/g;

/**
 * The saved story's part in order: ¶ paragraph end, [PB]/[CB], T[…] table, text, «code», F{…} simple field,
 * [ | ] complex field, L(…) link (L@name(…) to an external URL), S{…} content control, +{…}/-{…} tracked, <cN cN> comment range (numbered by
 * first appearance, as editor comment ids are hashes), R reference, BN/EN bookmark, § section break.
 */
export function sig(bytes: Uint8Array, part: string): string {
  const xml = partXml(bytes, part);
  const from = Math.max(xml.indexOf("<w:body>"), xml.indexOf('<w:footnote w:id="1"'), xml.indexOf('<w:endnote w:id="1"'), 0);
  const to = part === "word/document.xml" ? xml.lastIndexOf("<w:sectPr") : xml.length;
  const rels = partXml(bytes, part.replace(/([^/]+)$/, "_rels/$1.rels"));
  const target = (id: string) => new RegExp(`Id="${id}"[^>]*Target="[^"]*/([^"/]*)"`).exec(rels)?.[1] ?? `${id}?`;
  const comments = new Map<string, number>();
  const comment = (attrs: string) => {
    const id = /w:id="(-?\d+)"/.exec(attrs)![1];
    if (!comments.has(id)) comments.set(id, comments.size + 1);
    return comments.get(id);
  };
  return [...xml.slice(from, to).matchAll(SIG)]
    .map(([, close, tag, attrs, selfClose, text]) => {
      const id = /w:id="(-?\d+)"/.exec(attrs)?.[1];
      switch (tag) {
        case "t":
        case "delText":
          return close ? "" : text;
        case "instrText":
          return close ? "" : `«${text.trim()}»`;
        case "br":
          return /"page"/.test(attrs) ? "[PB]" : /"column"/.test(attrs) ? "[CB]" : "";
        case "p":
          return close || selfClose ? "¶" : "";
        case "tbl":
          return close ? "]" : "T[";
        case "commentRangeStart":
          return `<c${comment(attrs)}`;
        case "commentRangeEnd":
          return `c${comment(attrs)}>`;
        case "commentReference":
          return "R";
        case "bookmarkStart":
          return `B${id}`;
        case "bookmarkEnd":
          return `E${id}`;
        case "fldChar":
          return { begin: "[", separate: "|", end: "]" }[/fldCharType="(\w+)"/.exec(attrs)![1]] ?? "";
        case "fldSimple":
          return close ? "}" : selfClose ? "F{}" : "F{";
        case "hyperlink": {
          const rid = /r:id="([^"]+)"/.exec(attrs)?.[1];
          return close ? ")" : rid ? `L@${target(rid)}(` : "L(";
        }
        case "sdt":
          return close ? "}" : "S{";
        case "sectPr":
          return close ? "" : "§";
        default:
          if (selfClose) return "";
          return close ? "}" : `${{ ins: "+", del: "-", moveTo: "M+", moveFrom: "M-" }[tag] ?? tag}{`;
      }
    })
    .join("");
}

/** Comment markers and references next to each other, sorted: the accepted marker-order class compares these. */
export function markerOrder(s: string): string {
  return s.replace(/(?:<c\d+|c\d+>|R)+/g, (cluster) => (cluster.match(/<c\d+|c\d+>|R/g) ?? []).sort().join(""));
}

// ---------------------------------------------------------------- reading an editing state

type Props = { pPrIns?: unknown; pPrDel?: unknown; sectPr?: unknown; bookmarks?: Array<{ kind: string; id: unknown; offset?: number }> };

/**
 * A story's units: text (+(…)/-(…) tracked, a trailing < when it sits in a field result), ¶ (¶+/¶- tracked mark,
 * then § section break and {B5,E5} bookmarks), [PB]/[CB]/[kind] embeds (+/- tracked, ^ leading flag; [field:code|text]), {ref}
 * comment references.
 */
export function units(s: YrsSession, story: string): string {
  const marks = new Map(
    s.paragraphs(story).map(({ paraId, properties }) => {
      const props = properties as Props;
      const mark = props.pPrIns ? "+" : props.pPrDel ? "-" : "";
      const marks = (props.bookmarks ?? []).map((b) =>
        b.kind === "fieldend" || b.kind === "fieldseparate"
          ? `${b.kind === "fieldend" ? "]" : "|"}@${b.offset}`
          : `${b.kind === "start" ? "B" : "E"}${b.id}`
      ).join(",");
      return [paraId, `${mark}${props.sectPr ? "§" : ""}${marks ? `{${marks}}` : ""}`];
    })
  );
  return s
    .storySegments(story)
    .map((g) => {
      if (g.kind === "text") {
        const a = g.attributes as Record<string, unknown>;
        const t = a.fieldResult ? `${g.text}<` : g.text;
        return a.ins ? `+(${t})` : a.del ? `-(${t})` : t;
      }
      if (g.kind === "pilcrow") return `¶${marks.get(g.paraId) ?? ""}`;
      if (g.payload.modelKind === "commentReference") return "{ref}";
      const mark = g.attributes.ins ? "+" : g.attributes.del ? "-" : "";
      const kind = g.embedKind === "pageBreak" ? "PB" : g.embedKind === "columnBreak" ? "CB" : g.embedKind;
      const text = g.embedKind === "field" ? `:${String(g.payload.instruction ?? "").trim()}|${String(g.payload.displayText ?? "")}` : "";
      return `${mark}[${kind}${g.payload.leading ? "^" : ""}${text}]`;
    })
    .join("");
}

/** The bridge's blocks for a story, page breaks with their space-before flag: what the editor renders. */
export const blocks = (s: YrsSession, story: string) =>
  (s.yrsBlocksForStory(story) as Array<{ kind: string; keepsLeadingSpacing?: boolean }>)
    .map((b) => (b.kind === "pageBreak" ? (b.keepsLeadingSpacing ? "PB^" : "PB") : b.kind === "columnBreak" ? "CB" : b.kind[0]))
    .join(" ");

/** The text each comment covers in a story (# for an embed). */
export function covered(s: YrsSession, story: string): string[] {
  const flat = s
    .storySegments(story)
    .flatMap((g) =>
      g.kind === "text" ? [...g.text] : g.kind === "pilcrow" ? ["¶"] : [g.payload.modelKind === "commentReference" ? "" : "#"]
    );
  return s
    .listComments()
    .flatMap(({ id }) => s.resolveComment(id).filter((a) => a.story === story).map(({ start, end }) => flat.slice(start, end).join("")))
    .sort();
}

/** Projector (documentToYrs over the parsed file) and engine seed read a story the same way. */
export async function parity(bytes: Uint8Array, story: string): Promise<boolean> {
  const parsed = await parseDocx(bytes.slice().buffer, { preloadFonts: false });
  const projected = await createYrsSession({ clientId: nextClient() });
  const engine = await createYrsSession({ clientId: nextClient() });
  try {
    documentToYrs(projected, parsed);
    engine.seedFromDocx(bytes);
    const read = (s: YrsSession) =>
      JSON.stringify([s.storySegments(story), covered(s, story), s.paragraphs(story).map((x) => x.properties)]);
    return read(projected) === read(engine);
  } finally {
    projected.destroy();
    engine.destroy();
  }
}

// ---------------------------------------------------------------- rows

export type Edit = (s: YrsSession, story: string) => void;
export const none: Edit = () => {};
export const both = (...edits: Edit[]): Edit => (s, story) => edits.forEach((edit) => edit(s, story));
/** Typing in the body's tail paragraph: an edit away from the story under test. */
export const typeTail: Edit = (s) => void s.insertText({ story: "body", paraId: TAIL, offset: 0 }, "z");

/** A paragraph's length in units (its breaks, text and embeds). */
export const len = (s: YrsSession, story: string, paraId: string) => {
  const span = s.paragraphSpans(story).find((x) => x.paraId === paraId);
  if (!span) throw new Error(`no paragraph ${paraId} in ${story}`);
  return span.length;
};
export const has = (s: YrsSession, story: string, paraId: string) => s.paragraphSpans(story).some((x) => x.paraId === paraId);
/** The offset after the breaks and references that open a paragraph: where its text starts. */
export function textStart(s: YrsSession, story: string, paraId: string): number {
  const ids = s.paragraphs(story).map((x) => x.paraId);
  let index = 0;
  let offset = 0;
  for (const g of s.storySegments(story)) {
    if (g.kind === "pilcrow") {
      if (ids[index] === paraId) return offset;
      index += 1;
      offset = 0;
      continue;
    }
    if (ids[index] === paraId && (g.kind === "text" || !["pageBreak", "columnBreak"].includes(g.embedKind) && g.payload.modelKind !== "commentReference"))
      return offset;
    offset += g.kind === "text" ? g.text.length : 1;
  }
  throw new Error(`no paragraph ${paraId} in ${story}`);
}
/**
 * Where `needle` starts in a story: the paragraph and unit offset of its `nth` occurrence, text split across
 * formatting and comment boundaries included (an embed or reference counts as one unit that matches nothing).
 */
export function locate(s: YrsSession, story: string, needle: string, nth = 0): { paraId: string; offset: number } {
  const ids = s.paragraphs(story).map((x) => x.paraId);
  const lines: string[] = [""];
  for (const g of s.storySegments(story))
    if (g.kind === "pilcrow") lines.push("");
    else lines[lines.length - 1] += g.kind === "text" ? g.text : "\u0000";
  let seen = 0;
  for (const [index, line] of lines.entries())
    for (let at = line.indexOf(needle); at >= 0; at = line.indexOf(needle, at + 1)) if (seen++ === nth) return { paraId: ids[index], offset: at };
  throw new Error(`no ${JSON.stringify(needle)} in ${story}`);
}
/** Offset of the `nth` field embed (not a comment reference) whose instruction contains `code`. */
export function fieldAt(s: YrsSession, story: string, code: string, nth = 0): { paraId: string; offset: number } {
  const ids = s.paragraphs(story).map((x) => x.paraId);
  let index = 0;
  let offset = 0;
  let seen = 0;
  for (const g of s.storySegments(story)) {
    if (g.kind === "pilcrow") {
      index += 1;
      offset = 0;
      continue;
    }
    if (g.kind === "embed" && g.embedKind === "field" && g.payload.modelKind !== "commentReference" && String(g.payload.instruction ?? "").includes(code) && seen++ === nth)
      return { paraId: ids[index], offset };
    offset += g.kind === "text" ? g.text.length : 1;
  }
  throw new Error(`no field ${code} in ${story}`);
}
export const typeAt =
  (paraId: string, text: string, at: number | "end" | "text" = "end"): Edit =>
  (s, story) =>
    void s.insertText(
      { story, paraId, offset: at === "end" ? len(s, story, paraId) : at === "text" ? textStart(s, story, paraId) : at },
      text
    );
const note = [{ type: "paragraph", content: [{ type: "run", content: [{ type: "text", text: "note" }] }] }];
/** An editor comment from one paragraph offset to another; returns its id. */
export const addComment = (s: YrsSession, story: string, from: [string, number], to: [string, number], author = "Ed") =>
  s.addComment(
    [{ story, start: { paraId: from[0], offset: from[1] }, end: { paraId: to[0], offset: to[1] } }],
    author,
    "2026-09-29T00:00:00Z",
    note
  ).commentId;
export const reply = (s: YrsSession, parent: string) =>
  s.applyRawOps("body", [
    { op: "patchComment", id: `reply-${parent}`, fields: { author: "R2", date: "2026-09-29T00:00:00Z", body: note, parentId: parent } },
  ]);
export const removeComment = (s: YrsSession, id: string) => s.applyRawOps("body", [{ op: "removeComment", id }]);

export interface Row {
  /** Stable, human-readable key: the baseline's row id. */
  id: string;
  bytes: Uint8Array;
  where: Where;
  /** Applied before the capture. */
  before: Edit;
  /** Applied after the capture, to the live state (rebase) or to the reopened publication (publish-then-edit). */
  after: Edit;
  /** Also compare projector and engine on the direct save. */
  parity?: boolean;
}

/**
 * The capture orders of one setup and edit: the edit before the capture then typing elsewhere, the setup
 * before and the edit after it, and both after it. Without a setup the last two are the same row.
 */
export function orders(id: string, setup: Edit | null, edit: Edit) {
  const out: Array<{ id: string; before: Edit; after: Edit }> = [
    { id: `${id} | edit<cap<tail`, before: both(setup ?? none, edit), after: typeTail },
  ];
  if (setup) out.push({ id: `${id} | setup<cap<edit`, before: setup, after: edit });
  out.push({ id: `${id} | cap<${setup ? "setup+" : ""}edit`, before: none, after: both(setup ?? none, edit) });
  return out;
}

/**
 * Rebase outcome, best first. timing and marker-order are the accepted silent classes; a refusal ranks above them,
 * as it costs one publication while they save a file that depends on when the capture ran (a class the decisions
 * move to refusals, as for the break-flag difference, then shows as an improvement, and the reverse as a regression).
 */
const RANK = { exact: 0, refused: 1, timing: 2, "marker-order": 2, silent: 3, error: 4 };
type Rebase = keyof typeof RANK;

export interface Result {
  id: string;
  cls: string;
  detail?: string;
}

const breaksOf = (u: string) => (u.match(/\[[+-]?(PB|CB)/g) ?? []).length;
/** The text, field codes and field text a story holds, as a multiset: where they sit is `moved`, not `lost`. */
const textOf = (u: string) => [...u.replace(/\[[^\]:]*(?::([^\]]*))?\]|\{[^}]*\}|[¶§<+\-()]/g, "$1")].sort().join("");
const marksOf = (u: string) => (u.match(/[BE]\d+|§/g) ?? []).sort().join();
/** Units compared across a save: the leading flag and reference marks are the save's to place. */
const placed = (u: string) => u.replaceAll("{ref}", "").replaceAll("^", "");
/**
 * Decided 2026-09-29: a break with no paragraph before it and no text to lead (story start) saves as its own
 * break-only paragraph, so the reopened story gains that paragraph.
 */
const breakParagraph = (editor: string, saved: string) =>
  /^\[[+-]?(PB|CB)/.test(placed(editor)) && placed(saved) === `¶${placed(editor)}`;

async function stable(bytes: Uint8Array, part: string, story: string): Promise<string | null> {
  const seen: string[] = [];
  for (let i = 0; i < 3; i += 1) {
    const s = await open(bytes);
    seen.push(`${sig(bytes, part).replace(/z+tail/g, "tail")} ${JSON.stringify(covered(s, story))}`);
    typeTail(s, story);
    bytes = await publish(bytes, s.encodeState());
    s.destroy();
  }
  return new Set(seen).size === 1 ? null : seen.join(" => ");
}

const message = (e: unknown) => String(e instanceof Error ? e.message : e).slice(0, 160);

/**
 * One row: the direct save of before+after (checked over three publications), publish-then-edit, and the rebase
 * of `after` over a capture taken after `before`, classified against each other.
 */
export async function runRow(row: Row): Promise<Result> {
  clientId = 0;
  const [story, part] = STORY[row.where];
  const s = await open(row.bytes);
  let captured: Uint8Array, exported: Uint8Array, latest: Uint8Array, editor: string, bridge: string;
  try {
    row.before(s, story);
    captured = s.encodeState();
    exported = await publish(row.bytes, captured);
    row.after(s, story);
    latest = s.encodeState();
    editor = units(s, story);
    bridge = blocks(s, story);
  } catch (e) {
    return { id: row.id, cls: "error", detail: `edit: ${message(e)}` };
  } finally {
    s.destroy();
  }
  const directBytes = await publish(row.bytes, latest);
  const direct = sig(directBytes, part);

  let seq: string;
  try {
    const s2 = await open(exported);
    row.after(s2, story);
    seq = sig(await publish(exported, s2.encodeState()), part);
    s2.destroy();
  } catch (e) {
    seq = `error ${message(e)}`;
  }

  let next: string;
  let rebase: Rebase;
  try {
    const { state } = await rebaseOffice(row.bytes, checkpoint(row.bytes, captured), checkpoint(row.bytes, latest), exported);
    const r = await open(exported, state);
    next = sig(await publish(exported, r.encodeState()), part);
    r.destroy();
    rebase = next === direct ? "exact" : next === seq ? "timing" : markerOrder(next) === markerOrder(direct) ? "marker-order" : "silent";
  } catch (e) {
    next = `${e instanceof RebaseError ? "refused" : "error"} ${message(e)}`;
    rebase = e instanceof RebaseError ? "refused" : "error";
  }

  // Flags on the direct save: lost (a break, text, field, bookmark or section break is gone), moved (the reopened
  // units differ; break-paragraph is its decided form), unstable (no fixed point over three publications), render
  // (the editor's blocks differ from the reopened file's), parity (projector and engine read the save differently).
  const flags: string[] = [];
  const reopened = await open(directBytes);
  const saved = units(reopened, story);
  const render = blocks(reopened, story);
  reopened.destroy();
  if (breaksOf(saved) !== breaksOf(editor) || textOf(saved) !== textOf(editor) || marksOf(saved) !== marksOf(editor))
    flags.push("lost");
  else if (breakParagraph(editor, saved)) flags.push("break-paragraph");
  else if (placed(saved) !== placed(editor)) flags.push("moved");
  const unstable = await stable(directBytes, part, story);
  if (unstable) flags.push("unstable");
  if (render !== bridge && !(breakParagraph(editor, saved) && render === `p ${bridge}`)) flags.push("render");
  if (row.parity && !(await parity(directBytes, story))) flags.push("parity");

  const cls = [rebase, ...flags].join("+");
  if (cls === "exact" && !process.env.MATRIX_DETAIL) return { id: row.id, cls };
  const detail = [
    `editor ${editor}`,
    `saved  ${saved}`,
    `direct ${direct}`,
    `next   ${next}`,
    `seq    ${seq}`,
    ...(unstable ? [`pubs   ${unstable}`] : []),
    ...(render !== bridge ? [`blocks ${bridge} | reopened ${render}`] : []),
  ].join("\n");
  return { id: row.id, cls, detail };
}

/** Whether `now` is worse than `was`: a worse rebase outcome, a different one of the same rank, or a new flag. */
export function worse(was: string, now: string): boolean {
  const [a, ...af] = was.split("+");
  const [b, ...bf] = now.split("+");
  const rank = (cls: string) => RANK[cls as Rebase];
  return rank(b) > rank(a) || (rank(b) === rank(a) && a !== b) || bf.some((flag) => !af.includes(flag));
}

/**
 * A generator's entry point: runs its rows (every n-th from `--shard i/n`, those whose id contains `--only`) and
 * writes one JSON result per line to `--out` (stdout by default). run.ts spawns the shards and compares them with
 * the baseline. MATRIX_DETAIL=1 keeps the detail of exact rows too.
 */
export async function main(rows: () => Row[]): Promise<void> {
  const arg = (name: string) => {
    const at = process.argv.indexOf(name);
    return at < 0 ? undefined : process.argv[at + 1];
  };
  const [shard, shards] = (arg("--shard") ?? "0/1").split("/").map(Number);
  await prime();
  const only = arg("--only");
  const all = rows().filter((row) => !only || row.id.includes(only));
  const ids = new Set<string>();
  for (const row of all) {
    if (ids.has(row.id)) throw new Error(`duplicate row id: ${row.id}`);
    ids.add(row.id);
  }
  const lines: string[] = [];
  for (let i = shard; i < all.length; i += shards) lines.push(JSON.stringify(await runRow(all[i])));
  const out = arg("--out");
  if (out) await Bun.write(out, lines.join("\n"));
  else console.log(lines.join("\n"));
}
