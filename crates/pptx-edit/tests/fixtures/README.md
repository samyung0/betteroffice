# Deck fixtures

The engine reads one deck schema and keeps no stored snapshots of older ones.

## Hidden shapes

`hidden-shapes.pptx` is the repository's demo deck with `hidden="1"` added to
these `p:cNvPr` elements; every other ZIP part payload is unchanged:

| Slide | Snapshot shape ID | Name |
| --- | --- | --- |
| 1 | `slide:0:256:shape:0` | Cobalt rail |
| 1 | `slide:0:256:shape:8` | BetterOffice editor preview |
| 1 | `slide:0:256:shape:8.13` | PPTX tab |
| 2 | `slide:1:257:shape:4` | Format connector |
| 2 | `slide:1:257:shape:16` | Panel divider three |

Slide 1 loses 21 primitives: the cobalt rail, 14 child shapes, and six child
text boxes. Thirteen children have no hidden flag of their own. Slide 2 loses
the two marked shapes; slide 3 is unchanged.

## Connectors and comments

`connectors.md` documents the connector decks; `modern-comments.md` documents
`modern-comments.pptx`.

## Composite source decks

`blip-shadow.pptx` (bitmap effects with shape shadows), `chart-text-overflow.pptx`
(chart fills, axis `noFill` and explicit overflow), `metafile-tracking.pptx` (an OLE
preview with 6-point character spacing) and `run-spacing-shadow.pptx` (a shadow with
character spacing) each combine two features in one deck. `baseline_matches_seeded_doc_snapshot` seeds
each of them.
