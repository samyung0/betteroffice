# Evo Office browser proof

This directory answers two questions before Evo Notes commits to a deep
BetterOffice fork: can the framework-free browser cores open, view, edit, save,
and reopen modern Office files without silently discarding unrelated OOXML, and
can view-only DOCX/XLSX/PPTX sessions avoid loading their editor exports?

The answer for this fixture tier is **yes, with one benign XLSX package cleanup
and a remaining DOCX lowering caveat**.

## Run it

BetterOffice CI pins Bun 1.3.14. Use that version; older Bun releases fail
upstream tests because their runtime APIs differ.

```sh
bun install --frozen-lockfile
bun run build:packages
bun run typecheck:packages
bun run test
bun run test:poc
```

`test:poc` writes disposable edited/no-op files and a detailed JSON report to
`poc/output/`. It performs this sequence for DOCX, XLSX, and PPTX:

1. Open the source fixture and build the format's view model.
2. Save without an edit and reopen the result.
3. Make one marker edit, save, and reopen again.
4. Compare the source/output ZIP part names and SHA-256 hashes.
5. Assert format-specific logical invariants rather than accepting a successful
   ZIP write as proof of fidelity.

It then checks that the DOCX/XLSX/PPTX viewer artifacts remain below their
payload budgets and do not export save, edit, or collaboration methods.

The generated fixtures are documented in `fixtures/README.md`. Their generation
scripts stay beside the harness so a failing input can be reproduced exactly.

## Baseline result

Measured on macOS arm64 with Bun 1.3.14. Timings are local process timings, not
browser cold-start or network budgets.

| Format | View proof                                                         | No-op package result                                                                                       | Edited package result                                                                                 |
| ------ | ------------------------------------------------------------------ | ---------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------- |
| DOCX   | 20 stories, 17 body paragraphs, 2 sections, header/footer retained | 19/19 parts retained; 4 canonicalized                                                                      | 19/19 retained; independent render preserves the targeted run style                                   |
| XLSX   | 2 sheets, 181 draw commands, formula and chart present             | logical snapshot identical; 4/13 parts rewritten and an empty, unreferenced `xl/sharedStrings.xml` removed | marker edit reopens; formula, merges, freeze panes, representative styles, and chart geometry survive |
| PPTX   | 2 slides, 20 shapes, 36 render primitives                          | byte-identical payloads for all 26 parts                                                                   | only `ppt/slides/slide1.xml` changes                                                                  |

The edited outputs also pass independent rendering: LibreOffice for DOCX and
the artifact-tool import/render path for XLSX/PPTX. The PPTX overflow validator
reports no slide overflow.

## Payload finding

The generated, uncompressed WASM assets are currently:

| Runtime     | Raw bytes | Brotli bytes | Share of editor |
| ----------- | --------: | -----------: | --------------: |
| DOCX viewer | 6,767,841 |    1,716,992 |           76.2% |
| DOCX editor | 8,883,671 |    2,090,860 |            100% |
| XLSX viewer |   728,948 |      267,493 |           17.6% |
| XLSX editor | 4,140,006 |    1,208,100 |            100% |
| PPTX viewer | 1,407,997 |      509,359 |           50.0% |
| PPTX editor | 2,816,263 |      857,304 |            100% |

On the feature-rich fixtures, the WASM linear memory after the first render is
1,376,256 bytes for the XLSX viewer versus 2,424,832 for its editor, and
2,818,048 bytes for the PPTX viewer versus 3,342,336 for its editor. These
numbers exclude file buffers, JavaScript objects, canvases, and browser process
overhead; the check is useful as an engine regression gate, not a total-tab RAM
claim.

XLSX now has a dedicated parser/renderer crate with no Yrs, calculation, undo,
save, collaboration, or raster dependencies. PPTX builds its initial render
snapshot directly from parsed OOXML, so it does not construct a Yrs document.
Its renderer still shares the snapshot type definitions with the edit crate;
the linker removes the unused edit implementation from the viewer artifact.
DOCX still uses the shared OOXML-to-Yrs lowering bridge while opening, but the
viewer retains only the immutable display list. Its separate size-optimized
Cargo profile saves about 12.6% raw and 5.6% Brotli compared with the normal
release profile. A host should run this one-time lowering in a disposable
worker so the temporary editing projection and WASM memory are reclaimed.
The XLSX viewer displays formula results cached in the file. Opening the editor
can recalculate those values, so the host should treat that transition as a new
document revision when the results differ.

Use `@betteroffice/docx/viewer`, `@betteroffice/xlsx/viewer`, and
`@betteroffice/pptx/viewer` in a view iframe. The existing root imports remain
editor-compatible, and explicit editor entry points are available for the
replacement edit iframe.

## Decision

Keep the fork as a separate repository and keep the Rust crates: they are the
browser WASM engines, not optional multi-platform baggage. The Python bindings
do not need to ship in Evo Notes and can be excluded from our release workflow
without deleting them during the proof phase.

The next implementation milestone is an isolated browser host with one
versioned `postMessage` protocol. The editor iframe should replace the viewer
iframe on demand so the browser can reclaim viewer memory. SharedArrayBuffer
depends on cross-origin isolation; the iframe is useful as a containment and
lifecycle boundary, but it does not grant SharedArrayBuffer access by itself.

This fixture tier is sufficient to continue engineering on the fork. It is not
the final production-fidelity gate. Before integrating with Evo Notes, add a
small corpus authored by desktop Word, Excel, and PowerPoint, especially tracked
changes/comments, large formula workbooks, pivot tables, grouped/animated
slides, embedded media, and files with unusual fonts.
