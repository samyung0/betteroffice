// Accept all and Reject all (fields reviews 2–6: zz4-matrix, zz3-fixpoint, zz2-undo, zz4-stories), with tracked
// changes inside field results, in field code and outside fields.
import type { YrsSession } from "../../packages/docx/src/yrs";
import {
  E,
  S,
  bm,
  both,
  cf,
  del,
  deleted,
  docx,
  field,
  fs,
  ins,
  instr,
  link,
  main,
  moveFrom,
  moveTo,
  none,
  orders,
  p,
  run,
  sdt,
  table,
  webLink,
  type Edit,
  type Row,
  type Where,
} from "./lib";

const P = "44444444";
const holder = (xml: string) => p(P, `${run("a ")}${xml}${run(" b")}`);
const accept: Edit = (s) => void s.acceptChange({ all: true });
const reject: Edit = (s) => void s.rejectChange({ all: true });
const undo: Edit = (s) => void s.undo();

const delInstr = (t: string) => `<w:r><w:delInstrText xml:space="preserve">${t}</w:delInstrText></w:r>`;
const rPrChange = (text: string) =>
  `<w:r><w:rPr><w:b/><w:rPrChange w:id="95" w:author="A" w:date="2026-09-01T00:00:00Z"><w:rPr/></w:rPrChange></w:rPr><w:t>${text}</w:t></w:r>`;
const moveRanges = (from: string, to: string) =>
  `<w:moveFromRangeStart w:id="60" w:author="A" w:date="2026-09-01T00:00:00Z" w:name="move1"/>${moveFrom(from)}<w:moveFromRangeEnd w:id="60"/>` +
  `<w:moveToRangeStart w:id="61" w:author="A" w:date="2026-09-01T00:00:00Z" w:name="move1"/>${moveTo(to)}<w:moveToRangeEnd w:id="61"/>`;

/** Fields holding changes Accept and Reject all resolve (zz2-cases ACCEPT and zz3-fixpoint EXTRA). */
const SHAPES: Record<string, string> = {
  "cf>ins": field(`${run("20")}${ins(run("26"))}`),
  "cf>del": field(`${run("20")}${del(deleted("26"))}`),
  "cf>moveTo": field(`${run("20")}${moveTo(run("26"))}`),
  "cf>moveFrom": field(`${run("20")}${moveFrom(deleted("26"))}`),
  "cf>moveRanges": field(`${run("20")}${moveRanges(deleted("x"), run("26"))}`),
  "cf>ins>del": field(`${run("20")}${ins(`${run("2")}${del(deleted("5"), 71)}${run("6")}`)}`),
  "cf>del+ins": field(`${del(deleted("19"))}${ins(run("20"), 72)}${run("26")}`),
  "cf>ins-rPrChange": field(`${run("20")}${ins(rPrChange("26"))}`),
  "cf>run-rPrChange": field(`${run("20")}${rPrChange("26")}`),
  "cf>ins>link": field(`${run("20")}${ins(link(run("26")))}`),
  "cf>ins>fs": field(`${run("20")}${ins(fs(run("26"), " PAGE "))}`),
  "cf>ins>cf": field(`${run("20")}${ins(field(run("26"), " PAGE "))}`),
  "cf>cf>ins": field(`${run("20")}${field(`${run("2")}${ins(run("6"))}`, " PAGE ")}`),
  "cf>ins+cf": field(`${ins(run("20"))}${field(run("26"), " PAGE ")}`),
  "cf>sdt>ins": field(`${run("20")}${sdt(ins(run("26")))}`),
  "cf>ins>bm": field(`${run("20")}${ins(bm(run("26")))}`),
  "fs>ins": fs(ins(run("2026"))),
  "fs>del": fs(`${run("20")}${del(deleted("26"))}`),
  "fs>moveTo": fs(moveTo(run("2026"))),
  "fs>ins+del": fs(`${ins(run("20"), 70)}${del(deleted("19"), 71)}${run("26")}`),
  "fs>cf>ins": fs(`${run("20")}${field(`${run("2")}${ins(run("6"))}`, " PAGE ")}`),
  "ins>fs>ins": ins(fs(ins(run("2026"), 71)), 70),
  "link>fs>ins": link(fs(ins(run("2026")))),
  "sdt>cf>ins": sdt(field(`${run("20")}${ins(run("26"))}`)),
  "sdt>sdt>cf>del": sdt(sdt(field(`${run("20")}${del(deleted("26"))}`), 9)),
  "sdt>fs>ins": sdt(fs(ins(run("2026")))),
  "pj:cf>link+ins": field(`${link(run("20"))}${ins(run("26"))}`),
  "pj:cf>fs>ins+ins": field(`${fs(ins(run("20"), 70), " PAGE ")}${ins(run("26"), 71)}`),
  "code:ins": cf(ins(instr(" DATE ")), run("2026")),
  "code:del": cf(`${instr(" DATE ")}${del(delInstr("\\@ yyyy "))}`, run("2026")),
  "code:del+ins": cf(`${instr(" DATE ")}${del(delInstr("\\@ M "), 70)}${ins(instr("\\@ yyyy "), 71)}`, run("2026")),
  "code:ins-type": cf(`${del(delInstr(" PAGE "), 70)}${ins(instr(" DATE "), 71)}`, run("3")),
  "code:nested-ins": cf(`${instr(" IF ")}${field(run("1"), " PAGE ")}${ins(instr(' = 1 "a" "b" '))}`, run("a")),
  "code:in-nested": field(`${run("20")}${cf(ins(instr(" PAGE ")), run("26"))}`),
  "code:foreign": cf(`${instr(" DATE ")}<x:ext xmlns:x="urn:x" x:v="1"/>`, run("2026")),
  "fs>ins>link": fs(ins(link(run("2026")))),
  "fs>del>link": fs(`${run("20")}${del(link(deleted("26")))}`),
  "fs>ins>fs": fs(ins(fs(run("2026"), " PAGE "))),
  "sdt>cf>ins>link": sdt(field(`${run("20")}${ins(link(run("26")))}`)),
  "comment before:cf>ins>link": `${S()}${run("c")}${E()}${field(`${run("20")}${ins(link(run("26")))}`)}`,
  "bm before:cf>ins>link": `<w:bookmarkStart w:id="3" w:name="m"/>${run("c")}<w:bookmarkEnd w:id="3"/>${field(`${run("20")}${ins(link(run("26")))}`)}`,
  "2 fields:pj+ins>link": `${field(link(run("x")))}${field(`${run("20")}${ins(link(run("26")))}`)}`,
};

// Edits before and after the capture around Accept or Reject all, on fields whose results hold links (zz4-matrix:
// S9, S10, B5, B7, N20). A tab before the field shifts the numbering the export gives it (the journey's case).
const CHILD_SHAPES: Record<string, [string, string]> = {
  "link+ins": [field(`${link(run("20"))}${ins(run("26"))}`), "20"],
  "link+del": [field(`${link(run("20"))}${del(deleted("26"))}`), "20"],
  "ins>link": [field(`${run("20")}${ins(link(run("26")))}`), "26"],
  "fs child+ins": [field(`${link(run("20"))}${fs(run("1"), " PAGE ")}${ins(run("26"))}`), "20"],
  "nested fs>ins in link": [field(`${link(`${run("20")}${fs(ins(run("1")), " PAGE ")}`)}${run("x")}`), "20"],
  "two links, ins between": [field(`${link(run("20"))}${ins(run("-"))}${link(run("26"))}`), "26"],
};
/** Offset of `text` among a paragraph's field result children (-1 while it is not a child). */
function childAt(s: YrsSession, story: string, text: string): number {
  let offset = 0;
  for (const g of s.storySegments(story)) {
    if (g.kind === "pilcrow") break;
    if (g.kind === "text" && g.text.includes(text) && (g.attributes as Record<string, unknown>).fieldResult) return offset + g.text.indexOf(text);
    offset += g.kind === "text" ? g.text.length : 1;
  }
  return -1;
}
const typeIn = (text: string): Edit => (s, st) => {
  const at = childAt(s, st, text);
  if (at >= 0) s.insertText({ story: st, paraId: P, offset: at + 1 }, "Z");
};
const deleteChild = (text: string): Edit => (s, st) => {
  const at = childAt(s, st, text);
  if (at >= 0) s.deleteRange({ story: st, start: { paraId: P, offset: at }, end: { paraId: P, offset: at + text.length } });
};

// Accept all in other stories: an uncovered link resolves with its own part's relationships (R4-N2, S11).
const STORY_SHAPES: Record<string, [string, string]> = {
  "cf>ins>webLink": [field(`${run("20")}${ins(webLink(run("26")))}`), "26"],
  "pj:cf>webLink+ins": [field(`${webLink(run("20"))}${ins(run("26"))}`), "20"],
  "cf>ins": [SHAPES["cf>ins"], "26"],
  "fs>ins": [SHAPES["fs>ins"], "2026"],
};

// Tracked changes outside fields, paragraph marks included.
const PLAIN: Record<string, string> = {
  "ins run": holder(ins(run("new"))),
  "del run": holder(del(deleted("old"))),
  "moveFrom+moveTo": holder(moveRanges(deleted("m"), run("m"))),
  "rPrChange run": holder(rPrChange("bold")),
  "inserted paragraph mark": `${p(P, run("one"), `<w:rPr><w:ins w:id="80" w:author="A" w:date="2026-09-01T00:00:00Z"/></w:rPr>`)}${p("45454545", run("two"))}`,
  "deleted paragraph mark": `${p(P, run("one"), `<w:rPr><w:del w:id="81" w:author="A" w:date="2026-09-01T00:00:00Z"/></w:rPr>`)}${p("45454545", run("two"))}`,
  "inserted mark before a table": `${p(P, run("one"), `<w:rPr><w:ins w:id="82" w:author="A" w:date="2026-09-01T00:00:00Z"/></w:rPr>`)}${table(p("46464646", run("cell")))}`,
};

function rows(): Row[] {
  const out: Row[] = [];
  const push = (where: Where, xml: string, id: string, setup: Edit | null, edit: Edit) => {
    const bytes = docx(where, xml);
    for (const o of orders(`${where} | ${id}`, setup, edit)) out.push({ id: o.id, bytes, where, before: o.before, after: o.after });
  };
  for (const [name, xml] of Object.entries(SHAPES)) {
    push("body", holder(xml), `${name} | accept all`, null, accept);
    push("body", holder(xml), `${name} | reject all`, null, reject);
    push("body", holder(xml), `${name} | accept all, undo`, null, both(accept, undo));
  }
  for (const [shape, [xml, child]] of Object.entries(CHILD_SHAPES))
    for (const pre of ["", `<w:r><w:t>x</w:t><w:tab/></w:r>`]) {
      const bytes = docx("body", holder(`${pre}${xml}`));
      const type = typeIn(child);
      const befores: Record<string, Edit> = { none, type, delchild: deleteChild(child), accept, reject, "type+accept": both(type, accept) };
      const afters: Record<string, Edit> = { accept, reject, type, "type+accept": both(type, accept), delchild: deleteChild(child) };
      for (const [bn, before] of Object.entries(befores))
        for (const [an, after] of Object.entries(afters))
          out.push({ id: `body | ${shape}${pre ? " +tab" : ""} | before=${bn} | after=${an}`, bytes, where: "body", before, after });
    }
  for (const where of ["cell", "header", "footnote"] as Where[])
    for (const [name, [xml, child]] of Object.entries(STORY_SHAPES)) {
      push(where, holder(xml), `${name} | accept all`, null, accept);
      push(where, holder(xml), `${name} | reject all`, null, reject);
      push(where, holder(xml), `${name} | accept all, then type in ${child}`, accept, typeIn(child));
    }
  for (const where of ["body", "cell"] as Where[])
    for (const [name, xml] of Object.entries(PLAIN)) {
      push(where, xml, `${name} | accept all`, null, accept);
      push(where, xml, `${name} | reject all`, null, reject);
    }
  return out;
}

await main(rows);
