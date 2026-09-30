// Comments next to page and column breaks (reviews 7–11: zz8-p2, zz9-refmove, zz9-textstart, zz11-stale,
// zz11-endpara). Fixtures and names follow those probes, so rows compare with their logs.
import type { YrsSession } from "../../packages/docx/src/yrs";
import {
  CB,
  E,
  PB,
  S,
  addComment,
  both,
  del,
  docx,
  ins,
  len,
  main,
  none,
  orders,
  p,
  ref,
  removeComment,
  reply,
  run,
  textStart,
  typeAt,
  type Edit,
  type Row,
  type Where,
} from "./lib";

const PREV = "33333333";
const H = "44444444";
const NEXT = "45454545";
const p33 = (xml: string) => p(PREV, xml);
const p44 = (xml: string) => p(H, xml);
const p45 = (xml: string) => p(NEXT, xml);
const heading = (breaks: string) => `${p33(run("prev"))}${p44(`${breaks}${run("Heading")}`)}`;
const breakOnly = (breaks: string) => `${p33(run("prev"))}${p44(breaks)}${p45(run("next"))}`;

/** Document part, table cell, header part: stories behaved alike in rounds 8–11 once client ids were fixed. */
const STORIES: Where[] = ["body", "cell", "header"];

const ts = (s: YrsSession, st: string) => textStart(s, st, H);
const ed = (from: [string, number] | ((s: YrsSession, st: string) => [string, number]), to: (s: YrsSession, st: string) => [string, number]) =>
  ((s, st) => void addComment(s, st, typeof from === "function" ? from(s, st) : from, to(s, st))) as Edit;
const at = (offset: number) => () => [H, offset] as [string, number];
const atText = (plus = 0) => (s: YrsSession, st: string) => [H, ts(s, st) + plus] as [string, number];

// ---- A: editor comments ending before, between or right after a slot's breaks (R8-B1, R9-B1, R9-N1, R10-N1).
const A_FILES: Record<string, string> = {
  "prev¶[PB]Heading": heading(PB),
  "prev¶[CB]Heading": heading(CB),
  "prev¶[PB][CB]Heading": heading(PB + CB),
  "prev¶[CB]¶next": breakOnly(CB),
  "prev¶[PB][CB]¶next": breakOnly(PB + CB),
};
const A_COVERS: Record<string, Edit> = {
  "ed prev..44@0": ed([PREV, 0], at(0)),
  "two ed ending at 44@0": both(ed([PREV, 2], at(0)), ed([PREV, 0], at(0))),
  "ed prev..44@0 + reply": (s, st) => reply(s, addComment(s, st, [PREV, 0], [H, 0])),
  "ed over 1st break": ed([H, 0], at(1)),
  "ed prev..44@1": ed([PREV, 0], at(1)),
  "ed prev@0..text start": ed([PREV, 0], atText()),
  "two ed ..text start": both(ed([PREV, 0], atText()), ed([PREV, 2], atText())),
  "ed text start..+3": ed((s, st) => [H, ts(s, st)], atText(3)),
};

// ---- B: source comments whose range or reference sits next to the break (R8-N1, R9-N3).
const B_FILES: Record<string, string> = {
  "src <prev¶ E[PB]R Heading": `${p33(`${S()}${run("prev")}`)}${p44(`${E()}${PB}${ref()}${run("Heading")}`)}`,
  "src <prev¶ E[PB]Heading (no ref)": `${p33(`${S()}${run("prev")}`)}${p44(`${E()}${PB}${run("Heading")}`)}`,
  "src <prev¶ ER[PB]Heading (Word shape)": `${p33(`${S()}${run("prev")}`)}${p44(`${E()}${ref()}${PB}${run("Heading")}`)}`,
  "src <prev¶[PB]E Heading (no ref)": `${p33(`${S()}${run("prev")}`)}${p44(`${PB}${E()}${run("Heading")}`)}`,
  "src comment over [PB]Heading": `${p33(run("prev"))}${p44(`${S()}${PB}${run("Heading")}${E()}${ref()}`)}`,
  "src comment over Heading, ref in next para": `${p33(run("prev"))}${p44(`${PB}${S()}${run("Heading")}${E()}`)}${p45(`${ref()}${run("next")}`)}`,
};
const B_COVERS: Record<string, Edit> = {
  none,
  "reply to src 1": (s) => reply(s, "1"),
  "ed prev..44@0": A_COVERS["ed prev..44@0"],
  "ed over Heading": ed((s, st) => [H, ts(s, st)], (s, st) => [H, len(s, st, H)]),
};

// ---- C: comments over a break-led heading whose text goes (R7-P2, R8), at the story's start too.
const C_FILES: Record<string, string> = {
  "prev¶[PB]Heading": heading(PB),
  "prev¶[PB][CB]Heading": heading(PB + CB),
  "[PB]Heading¶next (story start)": `${p44(`${PB}${run("Heading")}`)}${p45(run("next"))}`,
};
const C_COVERS: Record<string, Edit> = {
  "ed over [PB]Heading": ed([H, 0], (s, st) => [H, len(s, st, H)]),
  "ed over [PB]": ed([H, 0], at(1)),
  "ed over Heading": B_COVERS["ed over Heading"],
  "two ed: [PB] and [PB]Heading": both(ed([H, 0], at(1)), ed([H, 0], (s, st) => [H, len(s, st, H)])),
  "ed over [PB]Heading + reply": (s, st) => reply(s, addComment(s, st, [H, 0], [H, len(s, st, H)])),
};

// ---- E: Word's tracked breaks with a comment at them, then Accept or Reject all.
const E_FILES: Record<string, string> = {
  "prev¶ins[PB]Heading": heading(ins(PB)),
  "prev¶del[PB]Heading": heading(del(PB)),
};
const E_COVERS: Record<string, Edit> = { none, "ed prev..44@0": A_COVERS["ed prev..44@0"], "ed over 1st break": A_COVERS["ed over 1st break"] };

const EDITS: Record<string, Edit> = {
  "type Q at 44 end": typeAt(H, "Q"),
  "delete 44 text": (s, st) => {
    const from = ts(s, st);
    if (from < len(s, st, H)) s.deleteRange({ story: st, start: { paraId: H, offset: from }, end: { paraId: H, offset: len(s, st, H) } });
  },
  "retype 44 text as X": (s, st) =>
    void s.replaceRange({ story: st, start: { paraId: H, offset: ts(s, st) }, end: { paraId: H, offset: len(s, st, H) } }, "X"),
  "remove ed comments": (s) => {
    for (const { id, author } of s.listComments()) if (author === "Ed") removeComment(s, id);
  },
  "Backspace after 1st break": (s, st) => void s.deleteAt({ story: st, paraId: H, offset: 1 }, "backward"),
  "accept all": (s) => void s.acceptChange({ all: true }),
  "reject all": (s) => void s.rejectChange({ all: true }),
};
const pick = (...names: string[]) => names.map((name) => [name, EDITS[name]] as const);

// ---- D: round 11. A comment removed and another added after the capture ahead of a break slot (R11-S1), and a
// comment added after the capture ending at a paragraph that closes with a trailing break (R11-N2).
const by = (author: string, from: [string, number], to: [string, number]): Edit => (s, st) => void addComment(s, st, from, to, author);
const drop = (author: string): Edit => (s) => {
  for (const c of s.listComments()) if (c.author === author) removeComment(s, c.id);
};
const cB = by("B", [PREV, 0], [PREV, 2]);
const cA = by("A", [PREV, 0], [H, 0]);
const cC = by("C", [PREV, 2], [PREV, 4]);
const Q = EDITS["type Q at 44 end"];
const STALE: Array<[string, string, Edit, Record<string, Edit>]> = [];
for (const [name, xml] of [
  ["prev¶[CB]¶next, A ends at 44@0", breakOnly(CB)],
  ["prev¶[PB][CB]¶next, A ends at 44@0", breakOnly(PB + CB)],
])
  STALE.push([name, xml, both(cB, cA), { "remove B + add C + type Q": both(drop("B"), cC, Q), "add C + type Q": both(cC, Q), "remove B + type Q": both(drop("B"), Q), "type Q": Q }]);
for (const [name, xml] of [
  ["prev[CB]¶[PB]Heading", `${p33(run("prev") + CB)}${p44(PB + run("Heading"))}`],
  ["prev[PB]¶[CB]Heading", `${p33(run("prev") + PB)}${p44(CB + run("Heading"))}`],
])
  STALE.push([name, xml, cB, { "remove B + add C": both(drop("B"), cC), "remove B + add C + type Q": both(drop("B"), cC, Q), "add C": cC, "remove B": drop("B") }]);
const ENDPARA: Record<string, string> = {
  "prev[PB]¶Heading": `${p33(run("prev") + PB)}${p44(run("Heading"))}${p45(run("next"))}`,
  "prev[CB]¶Heading": `${p33(run("prev") + CB)}${p44(run("Heading"))}${p45(run("next"))}`,
  "prev[CB]¶[PB]Heading": `${p33(run("prev") + CB)}${p44(PB + run("Heading"))}${p45(run("next"))}`,
  "prev[PB]¶¶next": `${p33(run("prev") + PB)}${p44("")}${p45(run("next"))}`,
};
const ENDPARA_ADDS: Record<string, Edit> = { "add C over ev": by("C", [PREV, 2], [PREV, 4]), "add C over prev": by("C", [PREV, 0], [PREV, 4]) };

function rows(): Row[] {
  const out: Row[] = [];
  const untouched = new Set<string>();
  const grid = (where: Where, files: Record<string, string>, covers: Record<string, Edit>, edits: ReadonlyArray<readonly [string, Edit]>) => {
    for (const [fname, xml] of Object.entries(files))
      for (const [cname, cover] of Object.entries(covers)) {
        const bytes = docx(where, xml);
        // The file untouched, and the cover alone: projector and engine read the saves alike.
        if (!untouched.has(`${where} | ${fname}`)) {
          untouched.add(`${where} | ${fname}`);
          out.push({ id: `${where} | ${fname} | untouched`, bytes, where, before: none, after: none, parity: true });
        }
        if (cover !== none) out.push({ id: `${where} | ${fname} | ${cname} | none | setup<cap`, bytes, where, before: cover, after: none, parity: true });
        for (const [ename, edit] of edits)
          for (const o of orders(`${where} | ${fname} | ${cname} | ${ename}`, cover === none ? null : cover, edit))
            out.push({ id: o.id, bytes, where, before: o.before, after: o.after });
      }
  };
  for (const where of STORIES) {
    const text = (f: string) => !f.endsWith("¶next");
    for (const [fname, xml] of Object.entries(A_FILES))
      grid(
        where,
        { [fname]: xml },
        Object.fromEntries(Object.entries(A_COVERS).filter(([c]) => text(fname) || c !== "ed text start..+3")),
        pick("type Q at 44 end", "remove ed comments", "Backspace after 1st break", ...(text(fname) ? ["delete 44 text"] : []))
      );
    for (const [cname, cover] of Object.entries(B_COVERS))
      grid(where, B_FILES, { [cname]: cover }, pick("type Q at 44 end", "delete 44 text", "Backspace after 1st break", ...(cname.startsWith("ed") ? ["remove ed comments"] : [])));
    grid(where, C_FILES, C_COVERS, pick("delete 44 text", "retype 44 text as X"));
    grid(where, E_FILES, E_COVERS, pick("type Q at 44 end", "accept all", "reject all"));
    for (const [fname, xml, before, afters] of STALE)
      for (const [aname, after] of Object.entries(afters))
        out.push({ id: `${where} | ${fname} | ${aname} | setup<cap<edit`, bytes: docx(where, xml), where, before, after });
    for (const [fname, xml] of Object.entries(ENDPARA))
      for (const [aname, after] of Object.entries(ENDPARA_ADDS))
        for (const [bname, before] of Object.entries({ "no earlier comment": none, "B on pr before": cB }))
          out.push({ id: `${where} | ${fname} | ${aname} | ${bname} | setup<cap<edit`, bytes: docx(where, xml), where, before, after });
  }
  return out;
}

await main(rows);
