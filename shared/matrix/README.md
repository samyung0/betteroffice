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
| `text` | The editor's story text, in document order, differs from what an equivalent edit leaves (`Row.same`, e.g. one peer doing what two did). |

A row is worse when its rebase class ranks lower (`timing` and `marker-order` rank together), changes within a rank,
or gains a flag. A refusal ranks above the accepted classes: it costs one publication, while they save a file that
depends on when the capture ran. So a class the decisions move to refusals (the break-flag difference, 2026-09-30)
shows as an improvement, and a refusal that stops firing (R11-S1's misses) as a regression.

## Generators

| File | Rows | Covers |
|---|---|---|
| `docx-breaks.matrix.ts` | breaks by story, file shape and edit | reviews 1–7 and 11: every file untouched in all nine stories; 12 shapes × 14 edits in five stories; tracked breaks with Accept/Reject all; captures after a state the export rewrites; the bookmark hand-off; a section-ending paragraph; breaks beside fields, links and insertions with and without `w14:paraId` |
| `docx-comments-breaks.matrix.ts` | comments next to breaks | rounds 7–11 (`zz8-p2`, `zz9-refmove`, `zz9-textstart`, `zz11-stale`, `zz11-endpara`), fixtures and names as in those probes, in body, cell and header; round 2's column break a comment boundary precedes and emptied comment over a source reference; round 2 review B's breaks a comment ends at or a bookmark sits at, with joins and typing |
| `docx-fields.matrix.ts` | field containers by edit | fields rounds 1–6 and round 7: 62 containers untouched, 18 edited by 9 edits, the journey's table of contents, REF fields whose results are links; joins collapsing bookmarks and continued fields with Undo and Redo (round 2); round 2 review B's Enter in projected links joined back (also beside a nested complex field, and with moved `w:ptab`, line-break and empty runs), Enter after projected simple fields, and typing after a field whose code continues; typing and Backspace after a continued result whose tail ends in a line break or a comment's reference; stories holding a `w:ptab`, untouched, typed in and left holding one after Enter, so a story that fails to render fails its rows; formatted and clearing line breaks ending a continued result; the shown-text refresh after Enter (Backspace, range delete, Accept All, type-over, a tail left) and two peers dropping the same link |
| `docx-review-all.matrix.ts` | Accept and Reject all | fields rounds 2–6: 41 fields with kept changes (accept, reject, accept then undo), `zz4-matrix`'s before/after grid, other stories' link relationships, tracked changes outside fields; a field deleted after Accept all under a comment (`zz-n4`) |

## Known non-exact rows in the baseline

Accepted differences and unresolved cases stay in the baseline, so fixes show as improvements:

- `timing+unstable` rows with two comments ending beside a break-only paragraph retain the accepted marker order
  and empty-range differences. An empty range becomes a reference-only comment after publication in every story.
  Since a text-less paragraph's breaks lead the text typed after them (round 2), the 6 such rows with text typed
  after a capture land with that marker order instead of refusing.
- `unstable` rows with two editor comments ending at one point: their markers reorder on the second publication.
- `moved` rows in `docx-fields` (accepted, round 2): Enter between a projected link and its field when the field's
  result holds a kept tracked insertion stays as it was, saving the link outside the field. (Text typed at the end of
  a paragraph whose field result continues into the next lands exactly since follow-up item 4: the seed reads that
  paragraph's result tail as text after the field.)
- `unstable` rows with a continued result whose first-paragraph tail ends in a comment's reference and text typed
  after it (same classes on dd87c75e): the parser hoists the comment's end out of the field, so the reopened file
  counts the typed text inside the comment and the next publication writes its end after it.
- `exact+moved` rows for text left in a field after its last link is deleted (accepted 2026-10-04): a plain run left
  ending a continued result's first paragraph stays in the field, which shows it, until the next publication reads
  it as text after the field (no tail move), so concurrent deletes converge without duplicates. One peer, two peers
  deleting the same link, and two peers each deleting half of it all land here; the `text` oracle checks the two-peer
  rows against one peer's delete.
- `exact+lost+text` rows (queued): a half-link delete by two peers leaves a stale shown text in the editor until
  publication. Neither partial delete removes the field's last link, so no refresh runs; with a nested field before
  the link the editor shows `REF=` where the save, which equals the single-peer save, and its reopened file show
  `REF=7`.
- `refused` rows follow decisions:
  - field children the rebase cannot match one-to-one (round 5, N20), including text typed next to a projected
    simple field before a capture, which shares that field's result slot the export's seed splits (accepted like
    N20, round 2 review B), and Enter inside a table of contents' first entry after a capture;
  - a table of contents entry typed in after a capture (the UAT journey);
  - a field deleted after Accept all uncovered a link a comment around it covers only from the next publication
    (zz-n4, round 2), when an edit after the capture touches the comment's range or what it covers; other edits
    land with the accepted timing difference (round 2 review B);
  - Undo restoring a continued field whose separate and end the export dropped, and a comment the rebase would lose
    with its reference, or whose range typing left ending before it starts (round 2 review);
  - the capture's save writing a new comment's reference ahead of breaks that open the latest paragraph, or ahead
    of bookmarks at a paragraph's start (round 2 review B, N);
  - an Enter taken after a capture that splits a field holding a tracked move after the split point (accepted,
    round 2 review D).

  The break-flag refusal (R10-N1, 2026-09-30) no longer fires for a text-less paragraph's own breaks, which lead text
  typed after them in the editor too, nor for a column break a comment boundary precedes (round 2).

The follow-up fixes leave no `silent` or `lost` rows. Field end/separate markers compare by kind and offset, since
their internal owner id can change on export; bookmark identities still compare exactly.
