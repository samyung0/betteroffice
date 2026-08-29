# Evo Office browser proof

This directory answers one narrow question before Evo Notes commits to a deep
BetterOffice fork: can the framework-free browser cores open, view, edit, save,
and reopen modern Office files without silently discarding unrelated OOXML?

The answer for this first fixture tier is **yes, with one benign XLSX package
cleanup and a significant DOCX payload warning**.

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

| Runtime                   |      Bytes | Consequence                                      |
| ------------------------- | ---------: | ------------------------------------------------ |
| DOCX parse + layout + OPC |  8,210,524 | plausible viewer-only payload before compression |
| DOCX edit                 | 11,417,142 | should remain out of the initial viewer path     |
| XLSX combined engine      |  4,171,799 | view and edit are not physically split yet       |
| PPTX combined engine      |  2,852,483 | view and edit are not physically split yet       |

DOCX already has separate WASM artifacts, but its current React viewer still
creates an editing/Yrs session. A real lightweight viewer entry point is
therefore achievable in the fork, but not available merely by setting a prop.
XLSX and PPTX each ship one combined WASM module; a read-only UI mode alone will
not materially reduce the engine payload. A true split there requires separate
Rust feature builds or dedicated read-only cores.

## Decision

Keep the fork as a separate repository and keep the Rust crates: they are the
browser WASM engines, not optional multi-platform baggage. The Python bindings
do not need to ship in Evo Notes and can be excluded from our release workflow
without deleting them during the proof phase.

The next implementation milestone should be an isolated browser host with two
entry points (`viewer` and `editor`) and one versioned `postMessage` protocol.
The editor iframe should replace the viewer iframe on demand so all viewer
memory can be reclaimed. SharedArrayBuffer depends on cross-origin isolation;
the iframe is useful as a containment and lifecycle boundary, but it does not by
itself grant SharedArrayBuffer access.

This fixture tier is sufficient to continue engineering on the fork. It is not
the final production-fidelity gate. Before integrating with Evo Notes, add a
small corpus authored by desktop Word, Excel, and PowerPoint—especially tracked
changes/comments, large formula workbooks, pivot tables, grouped/animated
slides, embedded media, and files with unusual fonts.
