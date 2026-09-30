// Page and column breaks by story, file shape and edit (reviews 1–7 and 11: zz-a-stories, zz-b-edits,
// zz-c-tracked, zz-d-rebase, zz-m5-rebase, zz11-plain, the hand-off and no-paraId probes).
import {
  ALL,
  CB,
  PB,
  bm,
  del,
  docx,
  fs,
  ins,
  len,
  link,
  main,
  none,
  orders,
  p,
  run,
  spaceBefore,
  textStart,
  typeAt,
  type Edit,
  type Row,
  type Where,
} from "./lib";

const PREV = "33333333";
const H = "44444444";
const NEXT = "45454545";
const EMPTY = "43434343";
const prev = p(PREV, run("prev"));
const next = p(NEXT, run("next"));
const heading = (breaks: string, ppr = "") => p(H, `${breaks}${run("Heading")}`, ppr);

/** Five stories stand for the nine: document part, block content control, table cell, header part, notes part. */
const FIVE: Where[] = ["body", "control", "cell", "header", "footnote"];

interface File {
  xml: string;
  /** The file has a paragraph before the break's paragraph. */
  prev: boolean;
  /** The break's paragraph holds text after its breaks. */
  text: boolean;
  /** The break's paragraph opens with a break. */
  leads: boolean;
}
const file = (xml: string, prev: boolean, text: boolean, leads = true): File => ({ xml, prev, text, leads });

const FILES: Record<string, File> = {
  "prev¶[PB]Heading": file(`${prev}${heading(PB)}${next}`, true, true),
  "prev¶[CB]Heading": file(`${prev}${heading(CB)}${next}`, true, true),
  "prev¶[PB][CB]Heading": file(`${prev}${heading(PB + CB)}${next}`, true, true),
  "prev[CB]¶[PB]Heading": file(`${p(PREV, run("prev") + CB)}${heading(PB)}${next}`, true, true),
  "prev[PB]¶Heading": file(`${p(PREV, run("prev") + PB)}${heading("")}${next}`, true, true, false),
  "prev¶[PB]Heading space-before": file(`${prev}${heading(PB, spaceBefore)}${next}`, true, true),
  // Word's page-break-only paragraph seeds its break before the paragraph's own slot: nothing to Backspace after.
  "prev¶[PB]¶next": file(`${prev}${p(H, PB)}${next}`, true, false, false),
  "prev¶[CB]¶next": file(`${prev}${p(H, CB)}${next}`, true, false),
  "prev¶[PB][CB]¶next": file(`${prev}${p(H, PB + CB)}${next}`, true, false),
  "[PB]Heading (story start)": file(`${heading(PB)}${next}`, false, true),
  "[PB][CB]¶next (story start)": file(`${p(H, PB + CB)}${next}`, false, false),
  "prev¶ab[PB]cd (mid)": file(`${prev}${p(H, `${run("ab")}${PB}${run("cd")}`)}${next}`, true, true, false),
};

type Needs = { prev?: boolean; text?: boolean; leads?: boolean };
const EDITS: Record<string, [Edit, Needs]> = {
  "type Q at 44 end": [typeAt(H, "Q"), {}],
  "type Q at 44@0": [typeAt(H, "Q", 0), {}],
  "Enter at 44@0": [(s, st) => void s.splitParagraph({ story: st, paraId: H, offset: 0 }), {}],
  "Enter at 44 text start": [(s, st) => void s.splitParagraph({ story: st, paraId: H, offset: textStart(s, st, H) }), {}],
  "Enter at 44 end, type N": [
    (s, st) => {
      const { secondParaId } = s.splitParagraph({ story: st, paraId: H, offset: len(s, st, H) });
      s.insertText({ story: st, paraId: secondParaId, offset: 0 }, "N");
    },
    {},
  ],
  "Enter at 44@0, Delete": [
    (s, st) => {
      const before = new Set(s.paragraphs(st).map((x) => x.paraId));
      s.splitParagraph({ story: st, paraId: H, offset: 0 });
      const empty = s.paragraphs(st).find((x) => !before.has(x.paraId))!.paraId;
      s.deleteAt({ story: st, paraId: empty, offset: len(s, st, empty) }, "forward");
    },
    { leads: true },
  ],
  "Backspace after 1st break": [(s, st) => void s.deleteAt({ story: st, paraId: H, offset: 1 }, "backward"), { leads: true }],
  "Delete at prev end": [(s, st) => void s.deleteAt({ story: st, paraId: PREV, offset: len(s, st, PREV) }, "forward"), { prev: true }],
  "delete 44 text": [
    (s, st) => s.deleteRange({ story: st, start: { paraId: H, offset: textStart(s, st, H) }, end: { paraId: H, offset: len(s, st, H) } }),
    { text: true },
  ],
  "retype 44 text as X": [
    (s, st) =>
      void s.replaceRange({ story: st, start: { paraId: H, offset: textStart(s, st, H) }, end: { paraId: H, offset: len(s, st, H) } }, "X"),
    { text: true },
  ],
  "join next into 44": [(s, st) => void s.deleteAt({ story: st, paraId: NEXT, offset: 0 }, "backward"), {}],
  "select prev@2..44 text+2, delete": [
    (s, st) => s.deleteRange({ story: st, start: { paraId: PREV, offset: 2 }, end: { paraId: H, offset: textStart(s, st, H) + 2 } }),
    { prev: true, text: true },
  ],
  "toolbar break at 44 text+2": [
    (s, st) => {
      const { secondParaId } = s.splitParagraph({ story: st, paraId: H, offset: textStart(s, st, H) + 2 });
      s.insertPageBreak({ story: st, paraId: secondParaId, offset: 0 });
    },
    { text: true },
  ],
  "type Z at prev end": [typeAt(PREV, "Z"), { prev: true }],
};
const fits = (f: File, needs: Needs) =>
  (!needs.prev || f.prev) && (!needs.text || f.text) && (!needs.leads || f.leads);

// Tracked breaks from Word files (the editor no longer suggests): kept wrappers, Accept and Reject all.
const TRACKED: Record<string, File> = {
  "prev¶ins[PB]Heading": file(`${prev}${heading(ins(PB))}${next}`, true, true),
  "prev¶del[PB]Heading": file(`${prev}${heading(del(PB))}${next}`, true, true),
  "prev ins[PB]¶Heading": file(`${p(PREV, run("prev") + ins(PB))}${heading("")}${next}`, true, true, false),
  "prev del[PB]¶Heading": file(`${p(PREV, run("prev") + del(PB))}${heading("")}${next}`, true, true, false),
};
const TRACKED_EDITS: Record<string, [Edit, Needs]> = {
  "type Q at 44 end": EDITS["type Q at 44 end"],
  "delete 44 text": EDITS["delete 44 text"],
  "Enter at 44 text start": EDITS["Enter at 44 text start"],
  "Backspace after 1st break": EDITS["Backspace after 1st break"],
  "accept all": [(s) => void s.acceptChange({ all: true }), {}],
  "reject all": [(s) => void s.rejectChange({ all: true }), {}],
};

// A capture after a state the export writes differently (the break-only paragraph, the leading flag), then an edit.
const REWRITTEN = ["prev¶[PB]Heading", "prev¶[PB][CB]Heading", "[PB]Heading (story start)"];
const SETUPS = ["delete 44 text", "Enter at 44 text start"];
const LATER = ["type Q at 44 end", "Backspace after 1st break", "Enter at 44 text start"];

// An empty bookmarked paragraph before a break's paragraph: the hand-off on Delete or Backspace (R3-S1, R5-B1, R6).
const HANDOFF: Record<string, string> = {
  "prev¶{bm}¶[PB][CB]¶next": `${prev}${p(EMPTY, bm(""))}${p(H, PB + CB)}${next}`,
  "{bm}¶[PB][CB]¶next (story start)": `${p(EMPTY, bm(""))}${p(H, PB + CB)}${next}`,
  "prev¶{bm}¶[CB]¶next": `${prev}${p(EMPTY, bm(""))}${p(H, CB)}${next}`,
  "prev¶{bm}¶[PB]Heading": `${prev}${p(EMPTY, bm(""))}${heading(PB)}${next}`,
};
const handOff: Edit = (s, st) => void s.deleteAt({ story: st, paraId: EMPTY, offset: 0 }, "forward");
const HANDOFF_EDITS: Record<string, Edit> = {
  "Backspace at 44@0": (s, st) => void s.deleteAt({ story: st, paraId: H, offset: 0 }, "backward"),
  "type Q at 44 end": typeAt(H, "Q"),
  "Delete at 44 end": (s, st) => void s.deleteAt({ story: st, paraId: H, offset: len(s, st, H) }, "forward"),
};

// An empty paragraph ending a section before a break (R2-S1); body only.
const sectPr = `<w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`;
const SECTION: Record<string, string> = {
  "prev¶§¶[PB]Heading": `${prev}${p(EMPTY, "", sectPr)}${heading(PB)}${next}`,
  "prev¶§¶[CB]Heading": `${prev}${p(EMPTY, "", sectPr)}${heading(CB)}${next}`,
};

// A mid-paragraph break right after a field, link or tracked insertion, with and without w14:paraId (R2-S2, R3-N4).
const beside = (inline: string, ids: boolean) =>
  `${p(ids ? PREV : null, run("prev"))}${p(ids ? H : null, `${run("a ")}${inline}${PB}${run("x")}`)}${p(ids ? NEXT : null, run("next"))}`;
const BESIDE: Record<string, string> = {};
for (const ids of [true, false])
  for (const [name, inline] of [
    ["F", fs(run("2026"))],
    ["L", link(run("2026"))],
    ["+", ins(run("2026"))],
  ])
    BESIDE[`prev¶a ${name}[PB]x${ids ? "" : " (no paraId)"}`] = beside(inline, ids);
/** Paragraph ids by position: files without w14:paraId get generated ones. */
const nth = (s: Parameters<Edit>[0], st: string, i: number) => s.paragraphs(st)[i].paraId;
const BESIDE_EDITS: Record<string, Edit> = {
  "Enter at prev end": (s, st) => void s.splitParagraph({ story: st, paraId: nth(s, st, 0), offset: len(s, st, nth(s, st, 0)) }),
  "type Q at 44 end": (s, st) => void s.insertText({ story: st, paraId: nth(s, st, 1), offset: len(s, st, nth(s, st, 1)) }, "Q"),
  "Enter at 44@0": (s, st) => void s.splitParagraph({ story: st, paraId: nth(s, st, 1), offset: 0 }),
};
const BESIDE_STORIES: Where[] = ["body", "cell", "header"];

function rows(): Row[] {
  const out: Row[] = [];
  const push = (where: Where, xml: string, id: string, setup: Edit | null, edit: Edit) => {
    for (const o of orders(`${where} | ${id}`, setup, edit)) out.push({ id: o.id, bytes: docx(where, xml), where, before: o.before, after: o.after });
  };
  // Every file untouched in every story, three publications, projector and engine compared.
  const untouched: Array<[string, string, Where[]]> = [
    ...Object.entries(FILES).map(([n, f]): [string, string, Where[]] => [n, f.xml, ALL]),
    ...Object.entries(TRACKED).map(([n, f]): [string, string, Where[]] => [n, f.xml, ALL]),
    ...Object.entries(HANDOFF).map(([n, x]): [string, string, Where[]] => [n, x, ALL]),
    ...Object.entries(SECTION).map(([n, x]): [string, string, Where[]] => [n, x, ["body"]]),
    ...Object.entries(BESIDE).map(([n, x]): [string, string, Where[]] => [n, x, ALL]),
  ];
  for (const [name, xml, stories] of untouched)
    for (const where of stories)
      out.push({ id: `${where} | ${name} | untouched`, bytes: docx(where, xml), where, before: none, after: none, parity: true });
  for (const where of FIVE) {
    for (const [fname, f] of Object.entries(FILES))
      for (const [ename, [edit, needs]] of Object.entries(EDITS)) if (fits(f, needs)) push(where, f.xml, `${fname} | ${ename}`, null, edit);
    for (const [fname, f] of Object.entries(TRACKED))
      for (const [ename, [edit, needs]] of Object.entries(TRACKED_EDITS)) if (fits(f, needs)) push(where, f.xml, `${fname} | ${ename}`, null, edit);
    for (const fname of REWRITTEN)
      for (const sname of SETUPS)
        for (const ename of LATER) push(where, FILES[fname].xml, `${fname} | ${sname}, then ${ename}`, EDITS[sname][0], EDITS[ename][0]);
    for (const [fname, xml] of Object.entries(HANDOFF)) {
      push(where, xml, `${fname} | Delete at bm end`, null, handOff);
      push(where, xml, `${fname} | Backspace at 44@0`, null, HANDOFF_EDITS["Backspace at 44@0"]);
      for (const ename of ["type Q at 44 end", "Delete at 44 end"])
        push(where, xml, `${fname} | Delete at bm end, then ${ename}`, handOff, HANDOFF_EDITS[ename]);
    }
  }
  for (const [fname, xml] of Object.entries(SECTION)) {
    push("body", xml, `${fname} | Delete at § end`, null, handOff);
    push("body", xml, `${fname} | Backspace at 44@0`, null, HANDOFF_EDITS["Backspace at 44@0"]);
  }
  for (const where of BESIDE_STORIES)
    for (const [fname, xml] of Object.entries(BESIDE))
      for (const [ename, edit] of Object.entries(BESIDE_EDITS)) push(where, xml, `${fname} | ${ename}`, null, edit);
  return out;
}

await main(rows);
