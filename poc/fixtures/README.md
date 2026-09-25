# Browser round-trip fixtures

These files are deterministic inputs for the Capy Notebook office proof of concept. Each
contains a unique `CAPY_EDIT_MARKER_*` string so the harness can make one narrow
edit, save, reopen, and distinguish the intended change from collateral OOXML
rewrites.

- `feature-rich.docx` covers runs, a hyperlink relationship, lists, a table,
  merged cells, headers/footers, hard pagination, Unicode, and sections.
- `feature-rich.xlsx` covers multiple worksheets, formulas, merged cells,
  formatting, frozen panes, Unicode-safe strings, and a native chart.
- `feature-rich.pptx` covers multiple slides, styled text boxes, shapes, theme
  colors, Unicode, and a native chart.

- `exchange-plan.docx` (Capy `e2e/fixtures/files/rich-content`) carries two native
  charts with embedded workbooks.
- `opaque-objects.docx` (Capy storage probe `gen_files.py`, `opaque_docx`) carries two
  charts, two text boxes as `mc:AlternateContent` with VML fallbacks, a VML `w:pict`
  rectangle and an OLE `w:object`.
- `lecture.pptx` (Capy `e2e/fixtures/files/rich-content`) carries 20 slides, two masters,
  22 layouts, 22 media parts, a table, comments and notes.
- `book-30p.docx`, `images-10.docx` and `deck-50.pptx` come from a full run of the Capy
  storage probe `gen_files.py` (seeded RNG): a text-only book, ten large pictures, and 50
  text-heavy slides. The golden seed tests pin their seeds and state sizes.

Regenerate the `feature-rich` fixtures with the scripts in `../scripts`. Preview outputs live
under `../preview` and are intentionally not part of the fixture contract.
