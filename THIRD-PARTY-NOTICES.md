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
https://github.com/y-crdt/y-crdt). Its sources are byte-identical apart from
the changes below, each marked "Patched for BetterOffice" in the code.

Text:

- one match arm in `clean_format_gap` (`src/types/text.rs`) that counts a
  shared type as content;
- `Text::apply_delta_as_yjs` (`src/types/text.rs`, with
  `yjs_clean_format_gap`): a delta whose deletions clean the format items
  they leave behind as Yjs 13.6.31's `deleteText` does, for the native Office
  rebase (tested in `crates/office-service`, which CI runs);
- `insert_attributes` and `insert_negated_attributes` (`src/types/text.rs`)
  write an edit's format items in key order, not in hash order, so one edit
  writes the same items in every run and layout (tests in
  `crates/docx-edit/tests/yrs_format_order.rs`, which CI runs);
- a split at a UTF-16 offset inside a surrogate pair, at a slice boundary or
  in an `ItemContent` split, gives two replacement characters as Yjs does,
  instead of panicking (`split_str_utf16_lossy` in `src/block.rs`, with a unit
  test).

Transactions and updates:

- a transaction reads its insert set instead of building whole state vectors
  to tell what it inserted (`emit_update_v1`/`v2` and `write_blocks_since` in
  `src/store.rs`, `encode_update` and `add_changed_type` in
  `src/transaction.rs`, `event_keys` in `src/types/mod.rs`), so its cost no
  longer grows with every client the document saw. One result differs from a
  state-vector diff. A struct that arrives before one its client wrote earlier
  is integrated behind a skip, where the state vector stops. Its
  transaction's update now holds that struct (upstream sent none), and later
  transactions count it as made before them, so changes to it raise events
  (upstream counted it as new and raised none). Tests are in
  `crates/docx-edit/tests/yrs_transaction_work.rs`, which pins the skip case;
- a pending update waits on each client's held state (`BlockStore::get_state`
  in `src/block_store.rs`, used by `src/update.rs` and `src/transaction.rs`)
  instead of its clock past skips, and on every client its stacked structs
  wait on, so structs that arrive before one they depend on are applied once
  it arrives (tests in `crates/docx-edit/tests/yrs_pending.rs`);
- `Update::decode` (`src/update.rs`) refuses an update in which a client's
  section starts inside or before an earlier one for that client, which yrs
  placed twice (CI runs its test in `crates/pptx-edit/tests/yrs_memory.rs`);
- limits for hostile input. `Any` decoding stops at `MAX_DECODE_DEPTH` (128)
  nesting levels (`src/any.rs`). No decoder reserves memory from a length the
  input declares: `src/any.rs`, the client and struct lists in
  `src/update.rs`, `src/state_vector.rs`, `src/id_set.rs`,
  `src/sync/awareness.rs` and two `ItemContent` decoders in `src/block.rs`.
  An update listing a client with no structs applies as nothing
  (`src/update.rs`). Tests are in `crates/pptx-edit/tests/yrs_hostile.rs` and
  `yrs_decode_alloc.rs`, and in `tests/decode_memory.rs` with its `[[test]]`
  entry in the crate's `Cargo.toml`;
- checked UTF-8 decoding instead of `from_utf8_unchecked` at three call sites
  (`src/encoding/read.rs`, `src/updates/decoder.rs`, `src/lib.rs`), so both
  update decoders reject invalid strings (see issue #224), with a unit test;
- a work counter with an optional budget (`Work` in `src/block_store.rs`,
  `TransactionMut::limit_work` and `work` in `src/transaction.rs`,
  `UpdateError::WorkBudgetExceeded` in `src/error.rs`).
  - Integration counts each item a conflict scan passes, each step of a map
    key's walk, the blocks a split or merge moves, picker switches and pending
    merges (`src/block.rs`, `src/update.rs`, `src/transaction.rs`,
    `src/id_set.rs`, `src/store.rs`).
  - Past the budget, `apply_update` stops before the next struct, a scan
    stops and the commit merges no more; the document must then be dropped.
  - A budget ends with its transaction (tests in
    `crates/pptx-edit/tests/yrs_placement.rs`);
- `BlockPicker::next` (`src/update.rs`) takes the next struct in a loop, not
  a call per drained client, so a chain of 20,000 clients no longer
  overflows the stack (tests in `crates/pptx-edit/tests/yrs_placement.rs`);
- `BlockStore::push` (`src/block_store.rs`) replaces only a skip; a block
  already placed stays, so it is never freed while its sequence links it
  (tests in `crates/pptx-edit/tests/yrs_placement.rs`);
- `TransactionMut::drop` (`src/transaction.rs`) skips the commit while a
  panic unwinds through it, so a second panic cannot abort the process; the
  document must then be dropped (tests in
  `crates/pptx-edit/tests/yrs_placement.rs`).

Undo and garbage collection:

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
- the same function deletes no more units of a redone copy than its step
  inserted (as y-crdt PR #674), so Undo no longer takes a peer's restored
  text along with the copy it merged into;
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
  own stack (tested in `crates/pptx-edit/tests/yrs_undo.rs`).

Read-only accessors, for structure walks and guards that compute no index:

- `Item::origin`, `right_origin`, `left`, `right`, `parent_sub`, `content`,
  `update_parent` (with `UpdateParent`), `parent_branch` and `redone` (the
  local link from a deleted item to the copy Undo or Redo made of it)
  (`src/block.rs`);
- `Branch::start`, `item`, `entry_item` and `map_items` (`src/branch.rs`);
- `Update::blocks` and `UpdateBlock` (`src/update.rs`, re-exported in
  `src/lib.rs`);
- `Store::get_item`, `get_block_range`, `blocks_in` (a client's blocks over a
  clock range, found with one search) and `next_live_item` (the first live
  item right of a deleted one, so the DOCX editor finds the field an edit
  ends at) (`src/store.rs`, with unit tests and tests in `crates/docx-edit`);
- `StickyIndex::get_offsets` (`src/sticky_index.rs`), which resolves many
  indexes in one walk as `get_offset` resolves each, with unit tests.

Its `Cargo.toml` also has a `[lints]` block that allows its upstream warnings.

The workspace, `apps/native-viewer`, `bindings` and `fuzz` substitute it for
the registry crate via `[patch.crates-io]`, so it is compiled into the crates
that use `yrs`, their WASM builds and the macOS app; crates published to
crates.io depend on the registry crate.

Copyright (c) 2020 Bartosz Sypytkowski, Kevin Jahns.

License: MIT (`third_party/yrs/LICENSE`).
