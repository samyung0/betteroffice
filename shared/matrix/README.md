# DOCX edit/publish/rebase matrices

A pre-landing regression suite for DOCX breaks, comments and fields. Run it before landing a change to the DOCX
editor, seed, export or rebase, and from the manual `DOCX matrices` workflow. It is not part of `bun run test`.
Capy's `office_matrix` job runs it and `shared/docx-*.test.ts` whenever the BetterOffice pin moves.

```sh
bun run build:wasm            # the docx, xlsx and pptx cores the checkpoint bundle imports
bun run test:matrix           # every generator, compared with baseline/
bun run test:matrix -- --update-baseline   # accept the current classes
bun run test:matrix -- docx-fields --jobs 4
bun shared/matrix/docx-breaks.matrix.ts --only "body | prev¶[PB]Heading | type Q"   # rows as JSON lines
MATRIX_DETAIL=1 …             # keep the detail of exact rows too
```

The runner fails when a row's class got worse, or when rows were added or removed; it prints improvements
separately. A fix shows up as `better …` lines and a baseline update; a regression as `WORSE …` with the row's
editor units, saved units and the direct, rebased and publish-then-edit saves. The full detail of every row that is
not exact goes to the `report.txt` it names.

## A row

Each row is a file, a story, an edit before the capture and an edit after it. The runner makes three saves:

- **direct**: both edits in one session, published, then published twice more with typing elsewhere;
- **publish-then-edit**: the capture's export reopened, the later edit applied, published;
- **rebase**: the later edit rebased over the capture onto its export, published.

`orders()` gives an edit its capture orders: `edit<cap<tail` (the edit, then typing in the body's tail after the
capture), `setup<cap<edit`, and `cap<setup+edit`.

Client ids are reset for every row. Editor comment ids and the order of concurrent inserts derive from them, and the
reviewers' logs showed two-comment rows changing class with the row's position in the process.

## Classes

The rebase outcome, best first:

| Class | Meaning |
|---|---|
| `exact` | The rebased save equals the direct save. |
| `refused` | The rebase threw a `RebaseError`; the next publication captures the latest state. |
| `timing` | Accepted: it equals publish-then-edit instead (publication timing; includes decision 48's retyped-heading rows). |
| `marker-order` | Accepted (2026-09-30, rounds 9–10): only the order of adjacent comment markers and references differs. |
| `silent` | Any other difference. |
| `error` | The edit or the rebase threw something else. |

Then flags on the direct save, joined with `+`:

| Flag | Meaning |
|---|---|
| `lost` | The reopened save lacks a break, text, field code or text, a bookmark or a section break the editor held. |
| `moved` | Nothing lost, but the reopened units differ from the editor's (leading flags and reference marks aside). |
| `break-paragraph` | Accepted (2026-09-29): a break with no paragraph before it and no text to lead saves as its own paragraph. |
| `unstable` | The saved story or its comment coverage changes over three publications. |
| `render` | The editor's blocks (page breaks with their space-before flag) differ from the reopened file's. |
| `parity` | The projector (`documentToYrs`) and the engine seed read the save differently (untouched rows only). |

A row is worse when its rebase class ranks lower (`timing` and `marker-order` rank together), changes within a rank,
or gains a flag. A refusal ranks above the accepted classes: it costs one publication, while they save a file that
depends on when the capture ran. So a class the decisions move to refusals (the break-flag difference, 2026-09-30)
shows as an improvement, and a refusal that stops firing (R11-S1's misses) as a regression.

## Generators

| File | Rows | Covers |
|---|---|---|
| `docx-breaks.matrix.ts` | breaks by story, file shape and edit | reviews 1–7 and 11: every file untouched in all nine stories; 12 shapes × 14 edits in five stories; tracked breaks with Accept/Reject all; captures after a state the export rewrites; the bookmark hand-off; a section-ending paragraph; breaks beside fields, links and insertions with and without `w14:paraId` |
| `docx-comments-breaks.matrix.ts` | comments next to breaks | rounds 7–11 (`zz8-p2`, `zz9-refmove`, `zz9-textstart`, `zz11-stale`, `zz11-endpara`), fixtures and names as in those probes, in body, cell and header |
| `docx-fields.matrix.ts` | field containers by edit | fields rounds 1–6 and round 7: 62 containers untouched, 18 edited by 9 edits, the journey's table of contents, REF fields whose results are links |
| `docx-review-all.matrix.ts` | Accept and Reject all | fields rounds 2–6: 41 fields with kept changes (accept, reject, accept then undo), `zz4-matrix`'s before/after grid, other stories' link relationships, tracked changes outside fields |

## Known non-exact rows in the baseline

Accepted differences and unresolved cases stay in the baseline, so fixes show as improvements:

- `timing+unstable` rows with two comments ending beside a break-only paragraph retain the accepted marker order
  and empty-range differences. An empty range becomes a reference-only comment after publication in every story.
- `unstable` rows with two editor comments ending at one point: their markers reorder on the second publication.
- `moved` rows in `docx-fields`: text typed after a projected field saves as its owner's result text, and Enter
  between a projected child and its field, or inside a table of contents' first entry, saves the child outside it.
- `refused` rows follow decisions: the break-flag refusal (R10-N1, 2026-09-30), field children the rebase cannot
  match one-to-one (round 5, N20), a table of contents entry typed in after a capture (the UAT journey).

The follow-up fixes leave no `silent` or `lost` rows. Field end/separate markers compare by kind and offset, since
their internal owner id can change on export; bookmark identities still compare exactly.
