// Insert and Update table of contents (docx-toc track, 2026-10-06): inserting one in an empty paragraph and mid
// paragraph, updating Insert's and Word's own, editing an entry or a heading after it, and two peers; review 1:
// a Word TOC whose end paragraph breaks its section (H1), two peers updating at once (M1), a Table of Figures (M2),
// Insert inside a TOC (L1); review 2: the old table's own paragraphs carrying an outline level (R2-1) and a peer
// typing at the first entry's end during an Update (R2-2).
import type { YrsSession, YrsTocLayout } from "../../packages/docx/src/yrs";
import { both, docx, field, instr, link, locate, main, orders, p, peers, run, bm, type Edit, type Row } from "./lib";

/** Headings by their own outline level (`\u`): the matrix package has no styles. */
const heading = (id: string, level: number, text: string) => p(id, run(text), `<w:outlineLvl w:val="${level - 1}"/>`);
const HEADINGS = [
  heading("10000002", 1, "Introduction"),
  p("10000003", run("Some text.")),
  heading("10000004", 2, "Background"),
  heading("10000005", 4, "Deep"),
  heading("10000006", 3, "Details"),
].join("");
const EMPTY = `${p("10000001", "")}${HEADINGS}`;
const MID = `${p("10000001", run("Hello world"))}${HEADINGS}`;

// Word's own TOC: the end opening the paragraph after the entries, `\o "1-2"`, a bookmark Word left on a heading.
const fc = (type: string) => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
const entry = (text: string, n: number) =>
  link(`${run(text)}<w:r><w:tab/></w:r>${field(run("9"), ` PAGEREF _Toc${n} \\h `)}`, `_Toc${n}`);
const WORD = [
  p("30000001", `${fc("begin")}${instr(' TOC \\o "1-2" \\h \\z \\u ')}${fc("separate")}${entry("Old one", 1)}`, `<w:pStyle w:val="TOC1"/>`),
  p("30000002", entry("Old two", 2), `<w:pStyle w:val="TOC2"/>`),
  p("30000003", `${fc("end")}${run("after")}`),
  p("10000002", bm(run("Introduction"), 1, "_Toc1"), `<w:outlineLvl w:val="0"/>`),
  p("10000003", run("Some text.")),
  heading("10000004", 2, "Background"),
  heading("10000006", 3, "Details"),
].join("");

// The UAT journey's shape: text before the field in its first paragraph, the end inside the last entry.
const CONTENTS = [
  p("30000001", `<w:r><w:t>Contents</w:t><w:tab/></w:r>${fc("begin")}${instr(' TOC \\o "1-3" \\h \\z \\u ')}${fc("separate")}${entry("Old one", 1)}`),
  p("30000002", `${entry("Old two", 2)}${fc("end")}`),
  HEADINGS,
].join("");

// Word's TOC on its own roman-numbered page: the end alone in the paragraph that breaks the section (review H1).
const SECTION = [
  p("30000001", `${fc("begin")}${instr(' TOC \\o "1-3" \\h \\z \\u ')}${fc("separate")}${entry("Old one", 1)}`, `<w:pStyle w:val="TOC1"/>`),
  p("30000002", entry("Old two", 2), `<w:pStyle w:val="TOC2"/>`),
  p("30000003", fc("end"), `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgNumType w:fmt="lowerRoman"/></w:sectPr>`),
  HEADINGS,
].join("");

// Old tables whose own paragraphs carry an outline level (review 2, R2-1): the field opening in a "Contents" title
// paragraph, and an entry paragraph with its own level. Update lists neither.
const TITLED = [
  p("30000001", `${run("Contents")}${fc("begin")}${instr(' TOC \\o "1-3" \\h \\z \\u ')}${fc("separate")}${entry("Old one", 1)}`, `<w:outlineLvl w:val="0"/>`),
  p("30000002", `${entry("Old two", 2)}${fc("end")}`, `<w:pStyle w:val="TOC2"/>`),
  HEADINGS,
].join("");
const LEVELLED = [
  p("30000001", `${fc("begin")}${instr(' TOC \\o "1-3" \\h \\z \\u ')}${fc("separate")}${entry("Old one", 1)}`, `<w:pStyle w:val="TOC1"/>`),
  p("30000002", `${entry("Old two", 2)}${fc("end")}`, `<w:pStyle w:val="TOC2"/><w:outlineLvl w:val="1"/>`),
  HEADINGS,
].join("");

// A Table of Figures (`\c`) before the headings (review M2): Update leaves it alone.
const FIGURES = `${p("40000001", `${fc("begin")}${instr(' TOC \\h \\z \\c "Figure" ')}${fc("separate")}${run("Figure 1")}${fc("end")}`)}${EMPTY}`;

const layout: YrsTocLayout = {
  pages: { "10000002": "1", "10000004": "2", "10000006": "ii" },
  tabTwips: 9350,
  emptyText: "No table of contents entries found.",
};
const at = (s: YrsSession, st: string, needle: string, plus = 0) => {
  const { paraId, offset } = locate(s, st, needle);
  return { story: st, paraId, offset: offset + plus };
};
const insert: Edit = (s, st) => void s.insertTableOfContents({ story: st, paraId: "10000001", offset: 0 }, layout);
const insertMid: Edit = (s, st) => void s.insertTableOfContents(at(s, st, "world"), layout);
const update: Edit = (s) => void s.updateTableOfContents(null, layout);
/** The heading paragraph holding exactly `text` (an entry also holds a tab and its page). */
const headingAt = (s: YrsSession, st: string, text: string) => s.paragraphs(st).find((x) => x.text === text)!.paraId;
const rename: Edit = (s, st) => {
  const paraId = headingAt(s, st, "Background");
  void s.replaceRange({ story: st, start: { paraId, offset: 0 }, end: { paraId, offset: 10 } }, "Context");
};
const addHeading: Edit = (s, st) => {
  const { secondParaId } = s.splitParagraph(at(s, st, "Some text.", 10));
  s.insertText({ story: st, paraId: secondParaId, offset: 0 }, "Method");
  s.setParagraphAttr(secondParaId, "outlineLevel", 1);
};

const EDITS: Record<string, Edit> = {
  "type in 1st entry": (s, st) => void s.insertText(at(s, st, "Introduction", 5), "Z"),
  "Backspace at 1st entry end": (s, st) => void s.deleteAt(at(s, st, "Introduction", 12), "backward"),
  "Enter in 2nd entry": (s, st) => void s.splitParagraph(at(s, st, "Background", 4)),
  "type in a heading": (s, st) => void s.insertText({ story: st, paraId: headingAt(s, st, "Details"), offset: 7 }, "Z"),
};

/** Typing at the first entry's end (review 2, R2-2). */
const typeEntryEnd: Edit = (s, st) => void s.insertText(at(s, st, "Introduction", 12), "YY");

/** Insert with the caret inside the table's second entry: Word replaces the table (review L1). */
const insertInside: Edit = (s, st) => void s.insertTableOfContents(at(s, st, "Background", 4), layout);

function rows(): Row[] {
  const out: Row[] = [];
  const push = (xml: string, id: string, setup: Edit | null, edit: Edit, same?: Edit) => {
    const bytes = docx("body", xml);
    const oracles = same ? orders(id, setup, same) : undefined;
    for (const [index, o] of orders(id, setup, edit).entries())
      out.push({ id: o.id, bytes, where: "body", before: o.before, after: o.after, ...(oracles ? { same: oracles[index]! } : {}) });
  };
  push(EMPTY, "insert in an empty paragraph", null, insert);
  push(MID, "insert mid paragraph", null, insertMid);
  push(EMPTY, "insert, update after a rename", both(insert, rename), update);
  push(WORD, "update Word's own", null, update);
  push(WORD, "update Word's own after a rename", rename, update);
  push(CONTENTS, "update one with text before it", null, update);
  for (const [name, edit] of Object.entries(EDITS)) push(EMPTY, `insert | ${name}`, insert, edit);
  push(EMPTY, "peers | update while renaming", insert, peers(update, rename));
  push(EMPTY, "peers | update while adding a heading", insert, peers(update, addHeading));
  push(EMPTY, "peers | insert while renaming", null, peers(insert, rename));
  push(SECTION, "update Word's own ending in a section break", null, update);
  // Accepted with the concurrent-join class: each peer's table stays whole, so there are two (`text` against one).
  push(EMPTY, "peers | update while updating", insert, peers(update, update), update);
  push(FIGURES, "update with a table of figures only", null, update);
  push(FIGURES, "update with a table of figures and one of headings", insert, update);
  push(EMPTY, "insert inside a table", insert, insertInside);
  push(TITLED, "update one opening in a heading-level title", null, update);
  push(LEVELLED, "update one with an entry at an outline level", null, update);
  // Accepted with the concurrent-join class: the typing sits among the old table's tombstones, after the new table.
  push(EMPTY, "peers | update while typing at the 1st entry's end", insert, peers(update, typeEntryEnd));
  push(EMPTY, "peers | typing at the 1st entry's end while updating", insert, peers(typeEntryEnd, update));
  return out;
}

await main(rows);
