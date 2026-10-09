# Third-party notices

This file records third-party software included in, or from which substantial
portions of, this repository derive, together with the applicable licenses.
Entries are appended as new third-party material is incorporated.

---

## eigenpal docx editor

The API design and portions of the TypeScript packages in this repository
derive from the eigenpal docx editor (upstream repository:
https://github.com/eigenpal/docx-editor, mirrored at
https://github.com/sorenlouv/docx-editor), published on npm as the following
packages (last release: 1.9.0):

- `@eigenpal/docx-editor-core`
- `@eigenpal/docx-editor-react`
- `@eigenpal/docx-editor-vue`
- `@eigenpal/docx-editor-i18n`
- `@eigenpal/docx-editor-agents`
- `@eigenpal/nuxt-docx-editor`

License: Apache License, Version 2.0 (per the `license` field of the published
package metadata and the `LICENSE` file shipped in the packages).
Copyright 2026 EigenPal Inc.

The license terms are identical to this repository's root `LICENSE`; that
file serves as the copy of the license for this derivation.

---

## Carlito (metric-compatible with Calibri)

`crates/xlsx-raster/assets/Carlito-Regular.ttf` is vendored unmodified from
the google/fonts repository, `ofl/carlito` (https://github.com/google/fonts);
upstream project https://github.com/googlefonts/carlito. The bytes are
compiled into the `betteroffice-xlsx-raster` crate via `include_bytes!`, and
shared unmodified by the `betteroffice-vsdx-raster` crate through the same path.

Copyright 2013 The Carlito Project Authors, with Reserved Font Name "Carlito".

License: SIL Open Font License, Version 1.1
(`crates/xlsx-raster/assets/OFL.txt`).

---

## Bundled font binaries (`packages/fonts`, `packages/fonts-cjk`)

Thirty font binaries are vendored unmodified and redistributed in the published
`@betteroffice/fonts` and `@betteroffice/fonts-cjk` packages. All are licensed
under the SIL Open Font License, Version 1.1. Each package ships the applicable
license texts, with the copyright notices they cover, under its `LICENSES/`
directory; those files are the authoritative copies for redistribution.

`packages/fonts/assets` — 25 faces:

- **Carlito** (4 faces) — Copyright 2013 The Carlito Project Authors, with
  Reserved Font Name "Carlito". From google/fonts `ofl/carlito`; upstream
  https://github.com/googlefonts/carlito. `LICENSES/OFL-Carlito.txt`.
- **Caladea** (4 faces) — Copyright 2012 The Caladea Project Authors. From
  google/fonts `ofl/caladea`; upstream
  https://github.com/huertatipografica/Caladea. `LICENSES/OFL-Caladea.txt`.
- **Liberation Sans / Serif / Mono** (12 faces) — Digitized data copyright (c)
  2010 Google Corporation; Copyright (c) 2012 Red Hat, Inc., with Reserved Font
  Name "Liberation". From the Liberation Fonts 2.1.5 release,
  https://github.com/liberationfonts/liberation-fonts.
  `LICENSES/OFL-Liberation.txt`.
- **Noto Sans Arabic, Noto Naskh Arabic** (3 faces) — Copyright 2022 The Noto
  Project Authors. From https://github.com/notofonts/notofonts.github.io;
  upstream https://github.com/notofonts/arabic. `LICENSES/OFL-NotoArabic.txt`.
- **Noto Sans Hebrew** (2 faces) — Copyright 2022 The Noto Project Authors.
  From https://github.com/notofonts/notofonts.github.io; upstream
  https://github.com/notofonts/hebrew. `LICENSES/OFL-NotoSansHebrew.txt`. The
  regular face is also compiled into the `betteroffice-vsdx-raster` tests via
  `include_bytes!`.

`packages/fonts-cjk/assets` — 5 faces:

- **Noto Sans SC / TC / JP / KR** (4 faces) — Copyright 2014-2021 Adobe, with
  Reserved Font Name "Source".
- **Noto Serif SC** (1 face) — Copyright 2017-2024 Adobe, with Reserved Font
  Name "Source".

Both are static `SubsetOTF` Regulars from
https://github.com/notofonts/noto-cjk. `LICENSES/OFL-NotoCJK.txt`.

---

## yrs (vendored copy)

`third_party/yrs` is the `yrs` 0.27.3 crate from crates.io (upstream
https://github.com/y-crdt/y-crdt). Its sources are byte-identical apart from:

- one match arm in `clean_format_gap` (`src/types/text.rs`) that counts a
  shared type as content;
- checked UTF-8 decoding instead of `from_utf8_unchecked` at three call sites
  (`src/encoding/read.rs`, `src/updates/decoder.rs`, `src/lib.rs`), so both
  update decoders reject invalid strings (see issue #224), with a unit test;
- `follow_redone` (`src/store.rs`) keeps the offset into each redone item, as
  Yjs's `followRedone` does, so a sticky index inside a multi-unit item that
  Undo and Redo restored resolves to the same unit, and Undo after a delete,
  Undo and Redo removes only its own step instead of the text typed before it,
  with unit tests;
- `UndoManager::try_process` (`src/undo.rs`) redoes a step's items in the
  order its deletions list them, as Yjs does, not in hash order, so restored
  copies get the same clocks every time and a later Undo no longer deletes a
  neighbouring copy that merged with its own (tests in
  `crates/pptx-edit/tests/yrs_undo_order.rs`, which CI runs);
- `Store::next_live_item` (`src/store.rs`), a read-only lookup of the first
  live item right of a (deleted) item, so the DOCX editor finds the field a
  peer's delete ends at without computing a position (its unit test lives in
  `crates/docx-edit`, which CI runs);
- `Item::gc` (`src/block.rs`) and `GCCollector::collect_marked` (`src/gc.rs`)
  collect a collected type's children even when Undo keeps them, as Yjs does
  (upstream PR #682), so none is left pointing into the type's freed branch,
  and `UndoManager::try_process` (`src/undo.rs`) skips an inserted item whose
  redone copy was collected instead of dropping the whole step (CI runs
  their tests in `crates/pptx-edit/tests/yrs_undo.rs` and `yrs_memory.rs`);
- `Item::redo` (`src/block.rs`) lets Undo restore a map value past a value
  set concurrently that lost (deleted, its origin the value being restored,
  the winner to its right), so Undo of a concurrent set no longer leaves the
  key without a value; every other value blocks as in Yjs, a live one and a
  removal of the current value alike, also when that was the concurrent
  winner (Epo 2026-10-08 narrow Undo rule); a value is
  restored under its own parent, as Yjs 13.6.31 does (#757); and a restored
  value's parents are kept, as Yjs's `keepItem` does (CI runs the tests in
  `crates/pptx-edit/tests/yrs_undo.rs` and `convergence.rs`,
  `crates/docx-edit/tests/yrs_memory.rs` and `concurrent_undo.rs`, and
  `crates/betteroffice-xlsx/tests/stable_collaboration.rs`);
- `UndoManager::clear_undo` and `clear_redo` (`src/undo.rs`) clear only their
  own stack (tested in `crates/pptx-edit/tests/yrs_undo.rs`);
- `Update::decode` (`src/update.rs`) refuses an update in which a client's
  section starts inside or before an earlier one for that client, which yrs
  placed twice (CI runs its test in `crates/pptx-edit/tests/yrs_memory.rs`);
- a `[lints]` block in its `Cargo.toml` that allows its upstream warnings.

The workspace, `apps/native-viewer`, `bindings` and `fuzz` substitute it for
the registry crate via `[patch.crates-io]`, so it is compiled into the crates
that use `yrs`, their WASM builds and the macOS app; crates published to
crates.io depend on the registry crate.

Copyright (c) 2020 Bartosz Sypytkowski, Kevin Jahns.

License: MIT (`third_party/yrs/LICENSE`).
