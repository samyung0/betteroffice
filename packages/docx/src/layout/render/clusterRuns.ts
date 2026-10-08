import type { DisplayPrimitive, GlyphRunPrimitive } from './displayList';

/**
 * Splits joined glyph runs (`clusterRuns`) back into one run per glyph
 * cluster, so each cluster keeps its own positioned mirror element and
 * selection and find highlights sit on the painted glyphs.
 */
export function splitClusterRuns(primitives: DisplayPrimitive[]): DisplayPrimitive[] {
  if (!primitives.some((p) => p.kind === 'glyphRun' && p.clusterRuns)) return primitives;
  return primitives.flatMap((p) => (p.kind === 'glyphRun' && p.clusterRuns ? clusterPieces(p) : [p]));
}

/** One run per glyph cluster of a joined run, as the line emitted them. */
export function clusterPieces(run: GlyphRunPrimitive): GlyphRunPrimitive[] {
  // glyph clusters are UTF-8 byte offsets into the Rust-side text
  const utf16At = new Map<number, number>();
  let bytes = 0;
  let utf16 = 0;
  for (const char of run.text) {
    utf16At.set(bytes, utf16);
    const code = char.codePointAt(0) ?? 0;
    bytes += code < 0x80 ? 1 : code < 0x800 ? 2 : code < 0x10000 ? 3 : 4;
    utf16 += char.length;
  }
  utf16At.set(bytes, utf16);
  const starts = [...new Set(run.glyphs.map((glyph) => glyph.cluster))];
  const { clusterRuns: _joined, ...base } = run;
  const docStart = run.docStart ?? 0;
  const order = run.logicalOrder ?? 0;
  return starts.map((start, index) => {
    const from = utf16At.get(start) ?? 0;
    const to = utf16At.get(starts[index + 1] ?? bytes) ?? utf16;
    const piece: GlyphRunPrimitive = {
      ...base,
      text: run.text.slice(from, to),
      glyphs: run.glyphs
        .filter((glyph) => glyph.cluster === start)
        .map((glyph) => ({ ...glyph, cluster: 0 })),
      docStart: docStart + from,
      docEnd: docStart + to,
      logicalOrder: order + index,
    };
    if (index > 0) {
      delete piece.tabsBefore;
      delete piece.breaksBefore;
    }
    if (index < starts.length - 1) {
      delete piece.tabsAfter;
      delete piece.breaksAfter;
    }
    return piece;
  });
}
