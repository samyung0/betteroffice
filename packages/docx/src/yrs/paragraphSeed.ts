import type { Paragraph, ParagraphFormatting } from '../types/document';

type Attrs = Record<string, unknown>;

/**
 * The paragraph properties a seed resolves: direct formatting, else the list
 * level, else the style (`stylePpr`, `null` when the package has no styles).
 * The save compares the editor's properties with these to find its changes.
 */
export function seededParagraphProperties(
  paragraph: Pick<Paragraph, 'formatting' | 'listRendering'>,
  stylePpr: ParagraphFormatting | undefined | null
): Attrs {
  const formatting = paragraph.formatting;
  const directFirst = formatting?.indentFirstLine === 0 ? undefined : formatting?.indentFirstLine;
  const firstLine = directFirst ?? paragraph.listRendering?.indentFirstLine ?? formatting?.indentFirstLine;
  const hanging = directFirst === undefined && paragraph.listRendering?.indentFirstLine !== undefined
    ? paragraph.listRendering.hangingIndent
    : formatting?.hangingIndent ?? paragraph.listRendering?.hangingIndent;
  const attrs: Attrs = {};
  if (stylePpr !== null) {
    attrs.alignment = formatting?.alignment ?? stylePpr?.alignment ?? null;
    attrs.spaceBefore = formatting?.spaceBefore ?? stylePpr?.spaceBefore ?? null;
    attrs.spaceAfter = formatting?.spaceAfter ?? stylePpr?.spaceAfter ?? null;
    attrs.spaceBeforeLines = formatting?.spaceBeforeLines ?? stylePpr?.spaceBeforeLines ?? null;
    attrs.spaceAfterLines = formatting?.spaceAfterLines ?? stylePpr?.spaceAfterLines ?? null;
    attrs.beforeAutospacing = formatting?.beforeAutospacing ?? stylePpr?.beforeAutospacing ?? null;
    attrs.afterAutospacing = formatting?.afterAutospacing ?? stylePpr?.afterAutospacing ?? null;
    attrs.lineSpacing = formatting?.lineSpacing ?? stylePpr?.lineSpacing ?? null;
    attrs.lineSpacingRule = formatting?.lineSpacingRule ?? stylePpr?.lineSpacingRule ?? null;
    attrs.spacingExplicit = formatting?.spacingExplicit || null;
    attrs.indentLeft =
      formatting?.indentLeft ??
      paragraph.listRendering?.indentLeft ??
      stylePpr?.indentLeft ??
      null;
    attrs.indentRight = formatting?.indentRight ?? stylePpr?.indentRight ?? null;
    const numberingRemoved =
      formatting?.numPr?.numId === 0 && stylePpr?.numPr && stylePpr.numPr.numId !== 0;
    const styleFirstLine = numberingRemoved ? undefined : stylePpr;
    attrs.indentFirstLine =
      firstLine ??
      styleFirstLine?.indentFirstLine ??
      null;
    attrs.hangingIndent =
      hanging ??
      styleFirstLine?.hangingIndent ??
      false;
    attrs.borders = formatting?.borders ?? stylePpr?.borders ?? null;
    attrs.shading = formatting?.shading ?? stylePpr?.shading ?? null;
    attrs.tabs = formatting?.tabs ?? stylePpr?.tabs ?? null;
    attrs.pageBreakBefore = formatting?.pageBreakBefore ?? stylePpr?.pageBreakBefore ?? null;
    attrs.keepNext = formatting?.keepNext ?? stylePpr?.keepNext ?? null;
    attrs.keepLines = formatting?.keepLines ?? stylePpr?.keepLines ?? null;
    attrs.widowControl = formatting?.widowControl ?? stylePpr?.widowControl ?? null;
    attrs.contextualSpacing = formatting?.contextualSpacing ?? stylePpr?.contextualSpacing ?? null;
    attrs.snapToGrid = formatting?.snapToGrid ?? stylePpr?.snapToGrid ?? null;
    attrs.autoSpaceDE = formatting?.autoSpaceDE ?? stylePpr?.autoSpaceDE ?? null;
    attrs.autoSpaceDN = formatting?.autoSpaceDN ?? stylePpr?.autoSpaceDN ?? null;
    attrs.outlineLevel = formatting?.outlineLevel ?? stylePpr?.outlineLevel ?? null;
    attrs.bidi = formatting?.bidi ?? stylePpr?.bidi ?? null;
    if (!formatting?.numPr && stylePpr?.numPr && stylePpr.numPr.numId !== 0) {
      attrs.numPr = stylePpr.numPr;
      attrs.numPrFromStyle = stylePpr.numPr;
    }
  } else {
    attrs.alignment = formatting?.alignment ?? null;
    attrs.spaceBefore = formatting?.spaceBefore ?? null;
    attrs.spaceAfter = formatting?.spaceAfter ?? null;
    attrs.spaceBeforeLines = formatting?.spaceBeforeLines ?? null;
    attrs.spaceAfterLines = formatting?.spaceAfterLines ?? null;
    attrs.beforeAutospacing = formatting?.beforeAutospacing ?? null;
    attrs.afterAutospacing = formatting?.afterAutospacing ?? null;
    attrs.lineSpacing = formatting?.lineSpacing ?? null;
    attrs.lineSpacingRule = formatting?.lineSpacingRule ?? null;
    attrs.spacingExplicit = formatting?.spacingExplicit || null;
    attrs.indentLeft = formatting?.indentLeft ?? paragraph.listRendering?.indentLeft ?? null;
    attrs.indentRight = formatting?.indentRight ?? null;
    attrs.indentFirstLine =
      firstLine ?? null;
    attrs.hangingIndent =
      hanging ?? false;
    attrs.borders = formatting?.borders ?? null;
    attrs.shading = formatting?.shading ?? null;
    attrs.tabs = formatting?.tabs ?? null;
    attrs.pageBreakBefore = formatting?.pageBreakBefore ?? null;
    attrs.keepNext = formatting?.keepNext ?? null;
    attrs.keepLines = formatting?.keepLines ?? null;
    attrs.widowControl = formatting?.widowControl ?? null;
    attrs.snapToGrid = formatting?.snapToGrid ?? null;
    attrs.autoSpaceDE = formatting?.autoSpaceDE ?? null;
    attrs.autoSpaceDN = formatting?.autoSpaceDN ?? null;
    attrs.outlineLevel = formatting?.outlineLevel ?? null;
    attrs.bidi = formatting?.bidi ?? null;
  }
  return attrs;
}
