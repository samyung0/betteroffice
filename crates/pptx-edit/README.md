# betteroffice-pptx-edit

The collaborative PPTX deck model, backed by Yrs.

`DeckSession` owns the shared document. Slide order, slides, shapes, and text
stories are separate shared types, so reordering slides does not conflict with
editing text on one of them. Text stories reuse the same one-story-is-one-CRDT-
text rule as the DOCX editing core.

`DeckUndoManager` tracks local user-origin transactions only; remote updates
stay out of local history.

State vectors, diffs, and updates are standard Yrs v1, so any transport that
speaks Yjs sync-v1 works.

Deck schema v3 keeps parsed metadata in `pptx:meta.packageJson` with an empty
`media` list. `pptx:meta.media` is one immutable Yrs array of
`[partPath, contentType, Buffer]` entries, preserving package order and exact
media bytes without JSON decimal-array expansion. This is private checkpoint
encoding; the parser's public JSON format is unchanged. The original bootstrap
metadata IDs are retired, and v3 metadata is written after the existing roots
at new IDs. State-vector synchronization therefore sends the binary overrides
to legacy parents while preserving slide, shape and story identities.

Opening v1/v2 updates migrates their JSON media before Undo starts. Live update
application also migrates if a concurrent legacy migration wins a map conflict.
Yrs garbage collection removes the replaced JSON payload from subsequently
encoded states. Migration preserves edited roots and the source fingerprint;
reopening without source bytes still works, and saving still requires the exact
original source package. Older engine versions reject schema v3 and must reload
with the new engine before they can receive its updates.

Run `cargo test -p betteroffice-pptx-edit` for schema migration, binary media,
late/concurrent updates, Undo, and source-preserving export checks.

Used by [betteroffice-pptx](https://crates.io/crates/betteroffice-pptx). The
`wasm` feature exposes the JavaScript surface consumed by
[@betteroffice/pptx](https://www.npmjs.com/package/@betteroffice/pptx).

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
