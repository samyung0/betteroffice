// Field containers by edit (fields reviews 1–6 and review 7: zz-matrix, zz-locations, zz-edits, zz-rebase,
// the projected-children probes zz5-join/zz6-adv/zz7-s15/zz7-leak, and the UAT journey's table of contents).
import type { YrsSession } from "../../packages/docx/src/yrs";
import {
  addComment,
  bm,
  both,
  cf,
  del,
  deleted,
  docx,
  field,
  fieldAt,
  fs,
  ins,
  instr,
  link,
  locate,
  main,
  moveFrom,
  moveTo,
  none,
  orders,
  p,
  peers,
  ref,
  run,
  S,
  E,
  sdt,
  type Edit,
  type Row,
  type Where,
} from "./lib";

const P = "44444444";

/** Containers around and inside fields (zz-cases): w:fldSimple (fs), complex fields (cf), links, tracked changes, content controls, bookmarks. */
const CASES: Record<string, string> = {
  "fs>bm": fs(bm(run("2026"))),
  "link>fs": link(fs(run("2026"))),
  "ins>fs": ins(fs(run("2026"))),
  "fs>ins": fs(ins(run("2026"))),
  "fs>sdt": fs(sdt(run("2026"))),
  "cf>ins": field(`${run("20")}${ins(run("26"))}`),
  "cf>del": field(`${run("20")}${del(deleted("26"))}`),
  "cf>sdt": field(`${run("20")}${sdt(run("26"))}`),
  "fs>del": fs(`${run("20")}${del(deleted("26"))}`),
  "fs>moveTo": fs(moveTo(run("2026"))),
  "fs>moveFrom": fs(`${run("20")}${moveFrom(deleted("26"))}`),
  "fs>link": fs(link(run("2026"))),
  "fs>fs": fs(`${run("20")}${fs(run("26"), " PAGE ")}`),
  "fs>cf": fs(`${run("20")}${field(run("26"), " PAGE ")}`),
  "fs>ins+del": fs(`${ins(run("20"), 70)}${del(deleted("19"), 71)}${run("26")}`),
  "fs>bm-partial": `${fs(`<w:bookmarkStart w:id="5" w:name="mark"/>${run("20")}`)}${run("26")}<w:bookmarkEnd w:id="5"/>`,
  "bm-open>fs-close": `<w:bookmarkStart w:id="5" w:name="mark"/>${run("20")}${fs(`${run("26")}<w:bookmarkEnd w:id="5"/>`)}`,
  "link>fs>bm": link(fs(bm(run("2026")))),
  "link>fs>ins": link(fs(ins(run("2026")))),
  "ins>fs>bm": ins(fs(bm(run("2026")))),
  "ins>fs>ins": ins(fs(ins(run("2026"), 71)), 70),
  "ins>link>fs": ins(link(fs(run("2026")))),
  "ins>fs+run": ins(`${run("x")}${fs(run("2026"))}${run("y")}`),
  "del>fs": del(fs(deleted("2026"))),
  "moveTo>fs": moveTo(fs(run("2026"))),
  "moveFrom>fs": moveFrom(fs(deleted("2026"))),
  "sdt>fs>ins": sdt(fs(ins(run("2026")))),
  "fs>sdt>ins": fs(sdt(ins(run("2026")))),
  "cf>ins>fs": field(`${run("20")}${ins(fs(run("26"), " PAGE "))}`),
  "cf>link>ins": field(`${run("20")}${link(ins(run("26")))}`),
  "cf>sdt>ins": field(`${run("20")}${sdt(ins(run("26")))}`),
  "cf>cf>ins": field(`${run("20")}${field(`${run("2")}${ins(run("6"))}`, " PAGE ")}`),
  "cf>bm+ins": field(`${bm(run("20"))}${ins(run("26"))}`),
  "cf>ins>del": field(`${ins(run("20"))}${del(deleted("19"))}${run("26")}`),
  "cf>moveTo": field(`${run("20")}${moveTo(run("26"))}`),
  "cf-code>ins": cf(ins(instr(" DATE ")), run("2026")),
  "cf-all-ins": [
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r>`,
    instr(" DATE "),
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>`,
    run("2026"),
    `<w:r><w:fldChar w:fldCharType="end"/></w:r>`,
  ]
    .map((xml, i) => ins(xml, 70 + i))
    .join(""),
  "link>cf>ins": link(field(`${run("20")}${ins(run("26"))}`)),
  "fs>unknown": fs(`<w:proofErr w:type="spellStart"/>${run("20")}<w:permStart w:id="1" w:edGrp="everyone"/>${run("26")}<w:permEnd w:id="1"/>`),
  "fs>smartTag": fs(`<w:smartTag w:uri="u" w:element="e">${run("2026")}</w:smartTag>`),
  "fs>customXml": fs(`<w:customXml w:element="e">${run("2026")}</w:customXml>`),
  "fs>math": fs(`${run("20")}<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>`),
  "fs>tab-br": fs(`<w:r><w:t>20</w:t><w:tab/><w:t>26</w:t><w:br/></w:r>`),
  "fs>empty": fs(""),
  "fs>bm-only": fs(bm("")),
  "ctl:ins>link": ins(link(run("2026"))),
  "ctl:link>ins": link(ins(run("2026"))),
  "ctl:sdt>ins": sdt(ins(run("2026"))),
  "link>cf(toc)": link(`${run("Intro")}${field(run("2"), " PAGEREF _Toc1 \\h ")}`),
  "fs>foreign": fs(`${run("20")}<x:ext xmlns:x="urn:x" x:v="1"/>${run("26")}`),
  "cf>foreign": field(`${run("20")}<x:ext xmlns:x="urn:x" x:v="1"/>${run("26")}`),
  "fs>ins>fs": fs(ins(fs(run("2026"), " PAGE "))),
  "fs>del>fs": fs(`${run("20")}${del(fs(deleted("26"), " PAGE "))}`),
  "deep:link>fs>fs>ins": link(fs(fs(ins(run("2026")), " PAGE "))),
  "deep:cf>cf>sdt": field(`${run("20")}${field(sdt(run("26")), " PAGE ")}`),
  "link>x+cf+y": link(`${run("x")}${field(run("2026"), " PAGE ")}${run("y")}`),
  "pj:cf>link+ins": field(`${link(run("20"))}${ins(run("26"))}`),
  "pj:cf>fs+sdt": field(`${fs(run("20"), " PAGE ")}${sdt(run("26"))}`),
  "pj:cf>fs": field(fs(run("20"), " PAGE ")),
  "pj:cf>link": field(link(run("2026"))),
};
const holder = (xml: string) => p(P, `${run("a ")}${xml}${run(" b")}`);

/** The containers whose edits found defects in rounds 1–7; the rest are checked untouched. */
const FOCUS = [
  "fs>bm", "link>fs", "ins>fs", "fs>ins", "fs>sdt", "cf>ins", "cf>del", "cf>sdt", "fs>del", "fs>fs", "cf>bm+ins",
  "link>fs>ins", "ins>fs>bm", "fs>bm-partial", "bm-open>fs-close", "pj:cf>link", "pj:cf>link+ins", "pj:cf>fs",
];

/** The first field embed in the holder paragraph (a projected child's field comes before its owner). */
const f0 = (s: YrsSession, st: string) => fieldAt(s, st, "").offset;
const pos = (st: string, offset: number) => ({ story: st, paraId: P, offset });
const span = (st: string, a: number, b: number) => ({ story: st, start: { paraId: P, offset: a }, end: { paraId: P, offset: b } });
const OPS: Record<string, Edit> = {
  "type before field": (s, st) => void s.insertText(pos(st, f0(s, st)), "X"),
  "type after field": (s, st) => void s.insertText(pos(st, f0(s, st) + 1), "Y"),
  "Backspace before field": (s, st) => void s.deleteAt(pos(st, f0(s, st)), "backward"),
  "delete field": (s, st) => s.deleteRange(span(st, f0(s, st), f0(s, st) + 1)),
  "delete across field": (s, st) => s.deleteRange(span(st, 1, f0(s, st) + 2)),
  "Enter before field": (s, st) => void s.splitParagraph(pos(st, f0(s, st))),
  "Enter after field": (s, st) => void s.splitParagraph(pos(st, f0(s, st) + 1)),
  "comment over field": (s, st) => void addComment(s, st, [P, f0(s, st)], [P, f0(s, st) + 1]),
  "comment across field": (s, st) => void addComment(s, st, [P, 1], [P, f0(s, st) + 2]),
};
const EDIT_STORIES: Where[] = ["body", "header"];

// A table of contents as Word writes it (the UAT journey's fixture): the field spans two paragraphs, each entry a
// link holding a PAGEREF field (R7-P1, N20 and the journey's refusal).
const tocChars = (code: string) =>
  [`<w:r><w:fldChar w:fldCharType="begin"/></w:r>${instr(code)}<w:r><w:fldChar w:fldCharType="separate"/></w:r>`, `<w:r><w:fldChar w:fldCharType="end"/></w:r>`];
const [tocBegin, tocEnd] = tocChars(' TOC \\o "1-3" \\h \\z \\u ');
const entry = (text: string, n: number) =>
  link(`${run(text)}<w:r><w:tab/></w:r>${field(run("1"), ` PAGEREF _Toc${n} \\h `)}`, `_Toc${n}`);
const TOC = [
  p("10000001", run("Title")),
  p("10000002", `<w:r><w:t>Contents</w:t><w:tab/></w:r>${tocBegin}${entry("Introduction", 1)}`),
  p("10000003", `${entry("Details", 2)}${tocEnd}`),
  p("10000004", bm(run("Introduction"), 1, "_Toc1")),
  p("10000005", run("Some text.")),
  p("10000006", bm(run("Details"), 2, "_Toc2")),
].join("");
const intro = (s: YrsSession, st: string, plus = 0) => {
  const { paraId, offset } = locate(s, st, "Introduction");
  return { story: st, paraId, offset: offset + plus };
};
const TOC_OPS: Record<string, Edit> = {
  "type in 1st entry": (s, st) => void s.insertText(intro(s, st, 5), "Z"),
  "type at 1st entry start": (s, st) => void s.insertText(intro(s, st), "Z"),
  "type at 1st entry end": (s, st) => void s.insertText(intro(s, st, 12), "Z"),
  "Backspace at 1st entry end": (s, st) => void s.deleteAt(intro(s, st, 12), "backward"),
  "retype 1st entry's last letter": (s, st) => void s.replaceRange({ story: st, start: intro(s, st, 11), end: intro(s, st, 12) }, "n"),
  "Enter inside 1st entry": (s, st) => void s.splitParagraph(intro(s, st, 5)),
  "Enter at 1st entry start": (s, st) => void s.splitParagraph(intro(s, st)),
  "type in body text": (s, st) => void s.insertText({ story: st, paraId: "10000005", offset: 4 }, "Z"),
};

// Two REF fields whose results are links (zz7-s15): type-over across a field's end (S15, R7-S1), Clear formatting
// on a link (S16), a delete before a field (R7-P1), a field deleted after typing in its link (B6), joins (S12).
const lk = (n: number, text: string) => `<w:hyperlink w:anchor="t${n}" w:history="1" w:tooltip="tip${n}">${run(text)}</w:hyperlink>`;
const REF = (n: string) => ` REF ${n} \\h `;
const REFS: Record<string, string> = {
  "a [REF a|L(AA)][REF b|L(BB)] b": p(P, `${run("a ")}${field(lk(0, "AA"), REF("a"))}${field(lk(1, "BB"), REF("b"))}${run(" b")}`),
  "x [REF a|L(AA)]¶[REF b|L(BB)] y": `${p(P, `${run("x ")}${field(lk(0, "AA"), REF("a"))}`)}${p("45454545", `${field(lk(1, "BB"), REF("b"))}${run(" y")}`)}`,
};
const aa = (s: YrsSession, st: string, plus = 0) => {
  const { paraId, offset } = locate(s, st, "AA");
  return { story: st, paraId, offset: offset + plus };
};
const refA = (s: YrsSession, st: string, plus = 0) => {
  const { paraId, offset } = fieldAt(s, st, "REF a");
  return { story: st, paraId, offset: offset + plus };
};
const bb = (s: YrsSession, st: string, plus = 0) => {
  const { paraId, offset } = locate(s, st, "BB");
  return { story: st, paraId, offset: offset + plus };
};
const REF_OPS: Record<string, Edit> = {
  "type in AA": (s, st) => void s.insertText(aa(s, st, 1), "Z"),
  "type at AA end": (s, st) => void s.insertText(aa(s, st, 2), "Z"),
  "Backspace at AA end": (s, st) => void s.deleteAt(aa(s, st, 2), "backward"),
  "delete REF a": (s, st) => s.deleteRange({ story: st, start: refA(s, st), end: refA(s, st, 1) }),
  "type over REF a..first B": (s, st) => void s.replaceRange({ story: st, start: refA(s, st), end: bb(s, st, 1) }, "Q"),
  "type over AA and REF a": (s, st) => void s.replaceRange({ story: st, start: aa(s, st), end: refA(s, st, 1) }, "Q"),
  "type over A|A..B|B": (s, st) => void s.replaceRange({ story: st, start: aa(s, st, 1), end: bb(s, st, 1) }, "Q"),
  "clear formatting on the first A": (s, st) => s.clearFormatting({ story: st, start: aa(s, st), end: aa(s, st, 1) }),
};

// Joins collapsing bookmarks and continued fields, then Undo and Redo (round 2 F and G, the review's C rows).
const fc = (type: string) => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
const bs = (id: number) => `<w:bookmarkStart w:id="${id}" w:name="m${id}"/>`;
const be = (id: number) => `<w:bookmarkEnd w:id="${id}"/>`;
const JOINS: Record<string, [string, Edit]> = {
  "nested continued field": [
    p("11111111", run("a") + fc("begin") + instr(" IF ") + fc("begin") + instr(" PAGE ")) +
      p(P, fc("separate") + run("one") + fc("end") + bs(5) + fc("separate") + be(5) + run("yes") + fc("end") + run("z")),
    (s, st) => void s.deleteRange({ story: st, start: { paraId: "11111111", offset: 1 }, end: { paraId: P, offset: 0 } }),
  ],
  "bookmarks collapsing": [
    p("11111111", run("ab") + bs(3) + bs(4) + run("c")) + p(P, run("d") + bs(5) + be(3) + run("e") + be(4) + be(5) + run("f")),
    (s, st) => void s.deleteRange({ story: st, start: { paraId: "11111111", offset: 2 }, end: { paraId: P, offset: 2 } }),
  ],
};
// Round 2 review B (rv2): Enter inside projected links and after projected simple fields, joins back, and typing
// after a field whose code continues into the next paragraph.
const len = (s: YrsSession, st: string, paraId: string) => s.paragraphSpans(st).find((x) => x.paraId === paraId)!.length;
const SPLITS: Record<string, string> = {
  "[REF|L(AA)L(BB)]": holder(field(`${link(run("AA"))}${link(run("BB"), "other")}`, " REF a \\h ")),
  "<c>[REF|L(AA)L(BB)]</c>": holder(`<w:commentRangeStart w:id="1"/>${field(`${link(run("AA"))}${link(run("BB"), "other")}`, " REF a \\h ")}<w:commentRangeEnd w:id="1"/>${ref()}`),
  "[REF|L(AA)F{20}L(BB)]": holder(field(`${link(run("AA"))}${fs(run("20"), " PAGE ")}${link(run("BB"), "other")}`, " REF a \\h ")),
  "[TOC|L(Intro)L(Body)]": p(P, field(`${link(run("Intro"))}${link(run("Body"), "t2")}`, ' TOC \\o "1-3" \\h ')),
  // Round 2 review C: a field showing its own result text.
  "[REF|L(AA)yL(BB)]": holder(field(`${link(run("AA"))}${run("y")}${link(run("BB"), "other")}`, " REF a \\h ")),
  "[REF|xL(AA)y]": holder(field(`${run("x")}${link(run("AA"))}${run("y")}`, " REF a \\h ")),
  // Follow-ups after round 3: a nested complex field before the link, and a moved run holding only a w:ptab.
  "[REF|[PAGE|7]L(AA)yy]": holder(field(`${field(run("7"), " PAGE ")}${link(run("AA"))}${run("yy")}`, " REF a \\h ")),
  "[REF|L(AA)y<ptab>z]": holder(
    field(`${link(run("AA"))}${run("y")}<w:r><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="dot"/></w:r>${run("z")}`, " REF a \\h ")
  ),
  // Review of ee62d514: a moved run holding only a line break, and an empty run after a projected simple field.
  "[REF|L(AA)y<br>z]": holder(field(`${link(run("AA"))}${run("y")}<w:r><w:br/></w:r>${run("z")}`, " REF a \\h ")),
  "[REF|L(AA)F{7}<b/>y]": holder(field(`${link(run("AA"))}${fs(run("7"), " PAGE ")}<w:r><w:rPr><w:b/></w:rPr></w:r>${run("y")}`, " REF a \\h ")),
};
const firstLink = (s: YrsSession, st: string) => {
  const at = locate(s, st, s.storySegments(st).some((g) => g.kind === "text" && g.text.includes("AA")) ? "AA" : "Intro");
  return { story: st, paraId: at.paraId, offset: at.offset + 1 };
};
const SPLIT_OPS: Record<string, Edit> = {
  "Enter in 1st, join back": (s, st) => {
    const { firstParaId } = s.splitParagraph(firstLink(s, st));
    s.deleteAt({ story: st, paraId: firstParaId, offset: len(s, st, firstParaId) }, "forward");
  },
  "Enter in 1st, Backspace at 2nd start": (s, st) => {
    const { secondParaId } = s.splitParagraph(firstLink(s, st));
    s.deleteAt({ story: st, paraId: secondParaId, offset: 0 }, "backward");
  },
  "Enter in 1st, Undo": (s, st) => {
    s.splitParagraph(firstLink(s, st));
    s.undo();
  },
  "Enter in 1st, range delete back": (s, st) => {
    const { firstParaId, secondParaId } = s.splitParagraph(firstLink(s, st));
    s.deleteRange({ story: st, start: { paraId: firstParaId, offset: len(s, st, firstParaId) }, end: { paraId: secondParaId, offset: 0 } });
  },
};
const AFTER_SIMPLE: Record<string, string> = {
  "[DATE|xF{20}y]": holder(field(`${run("x")}${fs(run("20"), " PAGE ")}${run("y")}`)),
  "[DATE|F{20}L(AA)x]": holder(field(`${fs(run("20"), " PAGE ")}${link(run("AA"))}${run("x")}`)),
  "[DATE|F{20}+{y}x]": holder(field(`${fs(run("20"), " PAGE ")}${ins(run("y"))}${run("x")}`)),
  "[DATE|F{20}-{w}y]": holder(field(`${fs(run("20"), " PAGE ")}${del(deleted("w"))}${run("y")}`)),
  // Round 2 review D: a tracked move after the split point refuses a rebase of Enter after a capture (accepted).
  "[DATE|F{20}M-{w}y]": holder(
    field(
      `${fs(run("20"), " PAGE ")}<w:moveFromRangeStart w:id="94" w:name="mv1" w:author="A" w:date="2026-09-01T00:00:00Z"/>${moveFrom(run("w"))}<w:moveFromRangeEnd w:id="94"/>${run("y")}`
    )
  ),
  "[DATE|F{20}<bm>y]": holder(field(`${fs(run("20"), " PAGE ")}<w:bookmarkStart w:id="7" w:name="m7"/>${run("y")}<w:bookmarkEnd w:id="7"/>`)),
};
const enterAfterPage: Edit = (s, st) => {
  const at = fieldAt(s, st, "PAGE");
  s.splitParagraph({ story: st, paraId: at.paraId, offset: at.offset + 1 });
};
const CODE_CONTINUES =
  p(P, `${run("a")}<w:r><w:fldChar w:fldCharType="begin"/></w:r>${instr(" DATE ")}`) +
  p("45454545", `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${run("20")}<w:r><w:fldChar w:fldCharType="end"/></w:r>${run("z")}`);
const RESULT_CONTINUES =
  p(P, `${run("a")}<w:r><w:fldChar w:fldCharType="begin"/></w:r>${instr(" DATE ")}<w:r><w:fldChar w:fldCharType="separate"/></w:r>${run("2")}`) +
  p("45454545", `${run("0")}<w:r><w:fldChar w:fldCharType="end"/></w:r>${run("z")}`);
// Decided 2026-10-04: a line break or a comment's reference ending the first paragraph's result is tail text too.
const RESULT_CONTINUES_BREAK =
  p(P, `${run("a")}<w:r><w:fldChar w:fldCharType="begin"/></w:r>${instr(" INCLUDETEXT x ")}<w:r><w:fldChar w:fldCharType="separate"/></w:r>${run("2")}<w:r><w:t>B</w:t><w:br/><w:t>C</w:t></w:r>`) +
  p("45454545", `${run("0")}<w:r><w:fldChar w:fldCharType="end"/></w:r>${run("z")}`);
const RESULT_CONTINUES_REF =
  p(P, `${S(1)}${run("a")}<w:r><w:fldChar w:fldCharType="begin"/></w:r>${instr(" DOCVARIABLE v ")}<w:r><w:fldChar w:fldCharType="separate"/></w:r>${run("AB")}${E(1)}${ref(1)}`) +
  p("45454545", `${run("0")}<w:r><w:fldChar w:fldCharType="end"/></w:r>${run("z")}`);
// Review of ee62d514, finding 1: the link left in the first paragraph deleted after Enter.
const deleteLeftLink: Edit = (s, st) => {
  const at = locate(s, st, "AA");
  const { firstParaId } = s.splitParagraph({ story: st, paraId: at.paraId, offset: at.offset + 1 });
  s.deleteAt({ story: st, paraId: firstParaId, offset: at.offset + 1 }, "backward");
};
// Recheck of e7bf01fc: a story holding a w:ptab (seeded as a tab embed) renders, untouched and edited, and after
// Enter leaves one in the second paragraph.
const PTAB = `<w:r><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="dot"/></w:r>`;
const PTABS: Record<string, string> = {
  "a<ptab>b": p(P, `${run("a")}${PTAB}${run("b")}`),
  "a<b ptab>b": p(P, `${run("a")}<w:r><w:rPr><w:b/></w:rPr><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="dot"/></w:r>${run("b")}`),
  "L(A<ptab>B)": holder(link(`${run("A")}${PTAB}${run("B")}`)),
  "[REF|L(A<ptab>B)]": holder(field(link(`${run("A")}${PTAB}${run("B")}`), " REF a \\h ")),
};
const PTAB_SPLITS: Record<string, string> = {
  "[REF|L(AA)<ptab>yz]": holder(field(`${link(run("AA"))}${PTAB}${run("yz")}`, " REF a \\h ")),
  "[REF|L(AA)yz<ptab>]": holder(field(`${link(run("AA"))}${run("yz")}${PTAB}`, " REF a \\h ")),
  "[REF|L(AA)<ptab>L(BB)]": holder(field(`${link(run("AA"))}${PTAB}${link(run("BB"), "other")}`, " REF a \\h ")),
};
// Recheck of 7d3ef664: d959a718's formatted and clearing line breaks ending a continued result's tail (they stay in
// the field), and 14fa634d's refresh paths after Enter in a field's link: Backspace or a range delete of the link
// left in the first paragraph, Accept All of its suggested deletion, typing over it, and a plain run left ending the
// continued result (tail text, the recorded children after it renumbered); also a native continued field losing its
// last link. Two-peer rows: both peers drop the same link, and each drops half of it.
const continuedWith = (tail: string) =>
  p(P, `${run("a")}<w:r><w:fldChar w:fldCharType="begin"/></w:r>${instr(" INCLUDETEXT x ")}<w:r><w:fldChar w:fldCharType="separate"/></w:r>${run("2")}${tail}`) +
  p("45454545", `${run("0")}<w:r><w:fldChar w:fldCharType="end"/></w:r>${run("z")}`);
const BREAK_TAILS: Record<string, string> = {
  "sized line break": continuedWith(`<w:r><w:rPr><w:sz w:val="40"/></w:rPr><w:br/></w:r>`),
  "clearing line break": continuedWith(`<w:r><w:br w:clear="all"/></w:r>`),
  "bold text and line break": continuedWith(`<w:r><w:rPr><w:b/></w:rPr><w:t>B</w:t><w:br/></w:r>`),
};
const REFRESHED: Record<string, string> = {
  "[REF|[PAGE|7]L(AA)yy]": SPLITS["[REF|[PAGE|7]L(AA)yy]"]!,
  "[REF|[PAGE|7]L(AA)L(BB)]": holder(field(`${field(run("7"), " PAGE ")}${link(run("AA"))}${link(run("BB"), "other")}`, " REF a \\h ")),
  "[REF|xL(AA)yy]": holder(field(`${run("x")}${link(run("AA"))}${run("yy")}`, " REF a \\h ")),
};
// After Enter in "AA", the link left in the first paragraph is its unit 2 ("a " before it).
const enterInAA: Edit = (s, st) => void s.splitParagraph(firstLink(s, st));
const LEFT_LINK_OPS: Record<string, Edit> = {
  "Backspace the link left": (s, st) => void s.deleteAt({ story: st, paraId: P, offset: 3 }, "backward"),
  "range delete the link left": (s, st) => void s.deleteRange({ story: st, start: { paraId: P, offset: 2 }, end: { paraId: P, offset: 3 } }),
  "Accept All of its suggested deletion": (s, st) => {
    s.deleteRange({ story: st, start: { paraId: P, offset: 2 }, end: { paraId: P, offset: 3 } }, { name: "Bo", date: "2026-10-04T00:00:00Z" });
    s.acceptChange({ all: true });
  },
  "type over the link left": (s, st) => void s.replaceRange({ story: st, start: { paraId: P, offset: 2 }, end: { paraId: P, offset: 3 } }, "Q"),
};
const NATIVE_TAIL =
  p(P, `${run("a")}<w:r><w:fldChar w:fldCharType="begin"/></w:r>${instr(" REF a \\h ")}<w:r><w:fldChar w:fldCharType="separate"/></w:r>${run("x")}${link(run("AA"))}`) +
  p("45454545", `${run("0")}<w:r><w:fldChar w:fldCharType="end"/></w:r>${run("z")}`);
// The native field's link is units 1..3 of its paragraph ("a" before it).
const dropUnits = (from: number, to: number): Edit => (s, st) =>
  void s.deleteRange({ story: st, start: { paraId: P, offset: from }, end: { paraId: P, offset: to } });
const undo: Edit = (s) => void s.undo();
const redo: Edit = (s) => void s.redo();

function rows(): Row[] {
  const out: Row[] = [];
  const push = (where: Where, xml: string, id: string, setup: Edit | null, edit: Edit, same?: Edit) => {
    const bytes = docx(where, xml);
    const oracles = same && orders(id, setup, same);
    for (const [index, o] of orders(`${where} | ${id}`, setup, edit).entries())
      out.push({ id: o.id, bytes, where, before: o.before, after: o.after, ...(oracles ? { same: oracles[index]! } : {}) });
  };
  const untouched = (where: Where, name: string, xml: string) =>
    out.push({ id: `${where} | ${name} | untouched`, bytes: docx(where, xml), where, before: none, after: none, parity: true });
  for (const [name, xml] of Object.entries(CASES)) untouched("body", name, holder(xml));
  for (const where of ["control", "cell", "header", "footnote"] as Where[]) for (const name of FOCUS) untouched(where, name, holder(CASES[name]));
  for (const where of EDIT_STORIES)
    for (const name of FOCUS) for (const [op, edit] of Object.entries(OPS)) push(where, holder(CASES[name]), `${name} | ${op}`, null, edit);
  untouched("body", "toc", TOC);
  for (const [op, edit] of Object.entries(TOC_OPS)) push("body", TOC, `toc | ${op}`, null, edit);
  for (const where of EDIT_STORIES)
    for (const [name, xml] of Object.entries(REFS)) {
      untouched(where, name, xml);
      const ops = name.includes("¶")
        ? ["type over A|A..B|B", "delete REF a", "Backspace at AA end"]
        : Object.keys(REF_OPS);
      for (const op of ops) push(where, xml, `${name} | ${op}`, null, REF_OPS[op]);
      push(where, xml, `${name} | type at AA end, then delete REF a`, REF_OPS["type at AA end"], REF_OPS["delete REF a"]);
      if (name.includes("¶"))
        push(where, xml, `${name} | join the paragraphs`, null, (s, st) =>
          void s.deleteAt({ story: st, paraId: P, offset: s.paragraphSpans(st).find((x) => x.paraId === P)!.length }, "forward"));
    }
  for (const where of ["body", "cell", "header"] as Where[]) {
    for (const [name, xml] of Object.entries(SPLITS)) {
      if (name.startsWith("[TOC") && where !== "body") continue;
      for (const [op, edit] of Object.entries(SPLIT_OPS)) push(where, xml, `${name} | ${op}`, null, edit);
    }
    for (const [name, xml] of Object.entries(AFTER_SIMPLE)) push(where, xml, `${name} | Enter after the simple field`, null, enterAfterPage);
    push(where, CODE_CONTINUES, "field code continued | type at its paragraph's end", null, (s, st) =>
      void s.insertText({ story: st, paraId: P, offset: len(s, st, P) }, "Q"));
    push(where, RESULT_CONTINUES, "field result continued | type at its paragraph's end", null, (s, st) =>
      void s.insertText({ story: st, paraId: P, offset: len(s, st, P) }, "Q"));
    for (const [name, xml] of [["line break", RESULT_CONTINUES_BREAK], ["comment reference", RESULT_CONTINUES_REF]] as const) {
      push(where, xml, `field result continued, ${name} in its tail | type at its paragraph's end`, null, (s, st) =>
        void s.insertText({ story: st, paraId: P, offset: len(s, st, P) }, "Q"));
      push(where, xml, `field result continued, ${name} in its tail | Backspace at its paragraph's end`, null, (s, st) =>
        void s.deleteAt({ story: st, paraId: P, offset: len(s, st, P) }, "backward"));
    }
    push(where, SPLITS["[REF|[PAGE|7]L(AA)yy]"]!, "[REF|[PAGE|7]L(AA)yy] | Enter in 1st, Backspace the link left", null, deleteLeftLink);
  }
  for (const where of ["body", "cell", "header", "footnote"] as Where[])
    for (const [name, xml] of Object.entries(PTABS)) {
      untouched(where, name, xml);
      push(where, xml, `${name} | type at its paragraph's start`, null, (s, st) => void s.insertText({ story: st, paraId: P, offset: 0 }, "Q"));
    }
  for (const where of ["body", "cell", "header"] as Where[])
    for (const [name, xml] of Object.entries(PTAB_SPLITS)) {
      push(where, xml, `${name} | Enter in 1st`, null, (s, st) => void s.splitParagraph(firstLink(s, st)));
      push(where, xml, `${name} | Enter in 1st, join back`, null, SPLIT_OPS["Enter in 1st, join back"]!);
    }
  for (const where of ["body", "cell", "header"] as Where[]) {
    for (const [name, xml] of Object.entries(BREAK_TAILS)) {
      untouched(where, `field result continued, ${name} in its tail`, xml);
      push(where, xml, `field result continued, ${name} in its tail | type at its paragraph's end`, null, (s, st) =>
        void s.insertText({ story: st, paraId: P, offset: len(s, st, P) }, "Q"));
      push(where, xml, `field result continued, ${name} in its tail | Backspace at its paragraph's end`, null, (s, st) =>
        void s.deleteAt({ story: st, paraId: P, offset: len(s, st, P) }, "backward"));
    }
    for (const [name, xml] of Object.entries(REFRESHED))
      for (const [op, edit] of Object.entries(LEFT_LINK_OPS))
        // The nested field's Backspace row is deleteLeftLink's.
        if (!(name === "[REF|[PAGE|7]L(AA)yy]" && op === "Backspace the link left"))
          push(where, xml, `${name} | Enter in 1st, ${op}`, null, both(enterInAA, edit));
    push(where, NATIVE_TAIL, "native [REF|xL(AA)¶0] | delete the link", null, dropUnits(1, 3));
    push(where, NATIVE_TAIL, "native [REF|xL(AA)¶0] | type over the link", null, (s, st) =>
      void s.replaceRange({ story: st, start: { paraId: P, offset: 1 }, end: { paraId: P, offset: 3 } }, "Q"));
    // Two peers.
    push(where, REFRESHED["[REF|xL(AA)yy]"]!, "[REF|xL(AA)yy] | Enter in 1st, both peers Backspace the link left", null,
      both(enterInAA, peers(LEFT_LINK_OPS["Backspace the link left"]!, LEFT_LINK_OPS["Backspace the link left"]!)),
      both(enterInAA, LEFT_LINK_OPS["Backspace the link left"]!));
    push(where, REFRESHED["[REF|[PAGE|7]L(AA)yy]"]!, "[REF|[PAGE|7]L(AA)yy] | Enter in 1st, both peers Backspace the link left", null,
      both(enterInAA, peers(LEFT_LINK_OPS["Backspace the link left"]!, LEFT_LINK_OPS["Backspace the link left"]!)),
      both(enterInAA, LEFT_LINK_OPS["Backspace the link left"]!));
    push(where, NATIVE_TAIL, "native [REF|xL(AA)¶0] | both peers delete the link", null, peers(dropUnits(1, 3), dropUnits(1, 3)), dropUnits(1, 3));
    push(where, NATIVE_TAIL, "native [REF|xL(AA)¶0] | each peer deletes half the link", null, peers(dropUnits(1, 2), dropUnits(2, 3)), dropUnits(1, 3));
  }
  for (const where of ["body", "cell", "header"] as Where[])
    for (const [name, [xml, join]] of Object.entries(JOINS)) {
      push(where, xml, `join ${name}`, null, join);
      push(where, xml, `join ${name}, undo`, join, undo);
      push(where, xml, `join ${name}, undo, redo`, join, (s, st) => {
        undo(s, st);
        redo(s, st);
      });
    }
  return out;
}

await main(rows);
