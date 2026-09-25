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

Regenerate the `feature-rich` fixtures with the scripts in `../scripts`. Preview outputs live
under `../preview` and are intentionally not part of the fixture contract.
