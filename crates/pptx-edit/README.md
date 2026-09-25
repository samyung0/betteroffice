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

Comments use a separate shared map. Saving patches existing XML and preserves
untouched comment parts byte for byte. Deleting a thread removes its known
replies; a reply added concurrently becomes a root when its parent is absent,
so both clients and saved files retain it.

The engine reads exactly one deck schema (5.0) and rejects every other version;
there are no migrations. The document stores the slide graph, stories and
comments only. Parsed package data (layouts, masters, themes, relationships and
media) never enters Yjs: `open_from_update_with_source` derives it from the
fingerprinted source package, so every open needs that package. Collaborators
must run the same engine.

`DeckSession::rebase_checkpoint(A, captured, latest, B, client_id)` promotes a
captured export B while preserving later saved edits. It exports the latest
state transiently, seeds a fresh graph, and retains only OPC parts differing
from B, media included, as a binary `sourceOverlay`. Reopening with B restores
these current source templates, including opaque XML and relationships needed
by a saved Undo. Unchanged parts and full prior source archives are not
retained. The next rebase replaces the overlay rather than accumulating
previous ones.

The rebase returns `state` plus a transient `indexed_state` whose object
identities match the new current graph while its content and positions describe
B. Project that indexed state to the compact retrieval baseline, then discard
it. Install the returned state with a new collaboration epoch; old updates and
Undo history must not cross that boundary. WASM exposes the same operation as
`PptxDocument.rebaseCheckpoint`, returning `state` and `indexedState` getters.

Used by [betteroffice-pptx](https://crates.io/crates/betteroffice-pptx). The
`wasm` feature exposes the JavaScript surface consumed by
[@betteroffice/pptx](https://www.npmjs.com/package/@betteroffice/pptx).

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.

`add_picture` inserts PNG, JPEG, GIF, BMP, TIFF, WebP, or SVG bytes (up to 8 MiB).
Pictures synchronize with the shared deck and resolve through `media_bytes` before save.
`bring_to_front`, `send_to_back`, `bring_forward`, and `send_backward` arrange slide objects.
