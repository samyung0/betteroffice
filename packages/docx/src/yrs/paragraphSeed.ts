import { computeListRendering, getCachedNumberingMap } from '../docx/numberingParser';
import { withTableReads } from './tableParagraphFormatting';
import type {
  ListRendering,
  NumberingDefinitions,
  Paragraph,
  ParagraphFormatting,
} from '../types/document';
import type {
  YrsAuthor,
  YrsParagraphAttrs,
  YrsParagraphTabStop,
  YrsSession,
  YrsStoryRange,
} from './index';

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
  // A first line set on the paragraph, zero included, wins over its list level's, as in Word.
  const list = paragraph.listRendering;
  const [firstLine, hanging] =
    formatting?.indentFirstLine !== undefined
      ? [formatting.indentFirstLine, formatting.hangingIndent]
      : [list?.indentFirstLine, list?.indentFirstLine !== undefined ? list.hangingIndent : undefined];
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
      (firstLine !== undefined ? hanging : styleFirstLine?.hangingIndent) ?? false;
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

const LIST_INDENTS = new Set(['indentLeft', 'indentFirstLine', 'hangingIndent']);

/**
 * `attrs` as an editor operation stores them: a property cleared that the
 * paragraph's style sets (`style`, from `styleParagraphValues`), or an indent
 * of a `numbered` paragraph (its list level sets them), becomes 0 or false,
 * and a style tab stop left out becomes a `clear` stop, so the paragraph
 * shows, saves and reopens without the inherited value.
 */
export function explicitParagraphAttrs(
  attrs: YrsParagraphAttrs,
  style: Readonly<Record<string, unknown>>,
  numbered = false
): YrsParagraphAttrs {
  const result: Record<string, unknown> = { ...attrs };
  for (const [key, value] of Object.entries(attrs)) {
    if (value !== null || key === 'tabs') continue;
    const off =
      numbered && LIST_INDENTS.has(key)
        ? key === 'hangingIndent' ? false : 0
        : typeof style[key] === 'number' && style[key] !== 0
          ? 0
          : style[key] === true
            ? false
            : undefined;
    if (off !== undefined) result[key] = off;
  }
  if ('tabs' in attrs) {
    const kept = attrs.tabs ?? [];
    // A stop ops stored before they took the seed's shape says `pos`.
    const at = (tab: YrsParagraphTabStop & { pos?: number }) => tab.position ?? tab.pos;
    const styleTabs = Array.isArray(style.tabs) ? (style.tabs as YrsParagraphTabStop[]) : [];
    const cleared = styleTabs
      .filter((stop) => stop.alignment !== 'clear' && !kept.some((tab) => at(tab) === stop.position))
      .map((stop): YrsParagraphTabStop => ({ position: stop.position, alignment: 'clear' }));
    result.tabs = kept.length + cleared.length > 0 ? [...kept, ...cleared] : null;
  }
  return result as YrsParagraphAttrs;
}

/**
 * A paragraph style's values for a paragraph in `story`: `styleParagraphValues`
 * with that story's cell context (`cellParagraphFormatting`) and the package's numbering.
 */
/**
 * A style's paragraph values in `story`. `list` false leaves the style's
 * numbering out; a numbering (`numPr`) gives the values of that style
 * numbered on the paragraph itself, with the level's rendering and indents.
 */
export type ParagraphStyleValues = (
  styleId: string | null,
  story: string,
  list?: boolean | { numId?: number; ilvl?: number }
) => Readonly<Record<string, unknown>>;

/**
 * Applies a paragraph style over the range as the editor does: the style's
 * values, given with those of the styles its paragraphs carry now, so a
 * paragraph's direct formatting stays.
 */
export function applyStyleValues(
  session: YrsSession,
  range: YrsStoryRange,
  styleId: string,
  styleValues: ParagraphStyleValues,
  suggesting?: YrsAuthor
): void {
  const paragraphs = session.paragraphs(range.story);
  const first = paragraphs.findIndex((paragraph) => paragraph.paraId === range.start.paraId);
  const last = paragraphs.findIndex((paragraph) => paragraph.paraId === range.end.paraId);
  const previous: Record<string, Readonly<Record<string, unknown>>> = {};
  for (const { properties } of paragraphs.slice(first, last + 1)) {
    const current = typeof properties.pStyle === 'string' ? properties.pStyle : '';
    previous[current] ??= styleValues(current || null, range.story);
  }
  session.applyParagraphStyle(range, styleId, styleValues(styleId, range.story), previous, suggesting);
}

/**
 * Gives the paragraphs of cells a table op made (`storyIds`) their style's
 * values in their cell, table style included, as the seed gives a file's
 * cells, so the editor shows a new header-row cell centred as Word does.
 * Run it with the op in `inOneUndoStep`.
 */
export function styleNewCells(
  session: YrsSession,
  storyIds: readonly string[],
  styleValues: ParagraphStyleValues
): void {
  withTableReads(() => {
    for (const story of storyIds) {
      const paragraphs = session.paragraphs(story);
      const first = paragraphs[0];
      const last = paragraphs.at(-1);
      if (!first || !last) continue;
      const styleId = typeof first.properties.pStyle === 'string' ? first.properties.pStyle : 'Normal';
      const range = { story, start: { paraId: first.paraId, offset: 0 }, end: { paraId: last.paraId, offset: 0 } };
      // A new cell's paragraphs carry only `styleId`, so its values are also the previous ones.
      const values = styleValues(styleId, story);
      session.applyParagraphStyle(range, styleId, values, { [styleId]: values });
    }
  });
}

/** Runs `edit` as one Undo step, whatever stories it selects. */
export function inOneUndoStep<T>(session: YrsSession, edit: () => T): T {
  const mode = session.undoCaptureMode();
  session.setUndoCaptureMode('manual');
  try {
    return edit();
  } finally {
    session.setUndoCaptureMode(mode);
  }
}

const LIST_ATTRS = [
  'listNumFmt',
  'listIsBullet',
  'listMarker',
  'listMarkerHidden',
  'listMarkerFontFamily',
  'listMarkerFontSize',
  'listMarkerBold',
  'listMarkerItalic',
  'listMarkerColor',
  'listMarkerSuffix',
  'listLevelNumFmts',
  'listAbstractNumId',
  'listStartOverride',
] as const;

const NO_LIST: YrsParagraphAttrs = {
  numPr: null,
  numPrFromStyle: null,
  ...Object.fromEntries(LIST_ATTRS.map((key) => [key, null])),
  indentLeft: null,
  indentFirstLine: null,
  hangingIndent: null,
};

const LIST_VALUE_KEYS = ['numPr', ...LIST_ATTRS, 'indentLeft', 'indentFirstLine', 'hangingIndent'] as const;

/**
 * Enter in an empty list item, as in Word: an item below the first level
 * moves up one level, and a first-level item leaves the list: numbering set
 * on the paragraph goes, and where its style gives a list it is turned off
 * (`numId` 0); the indents become the style's without its list. Returns
 * false, changing nothing, for a paragraph that is no empty list item (one
 * holding a field, picture or break is not empty).
 */
export function endEmptyListItem(
  session: YrsSession,
  story: string,
  paraId: string,
  styleValues: ParagraphStyleValues,
  suggesting?: YrsAuthor
): boolean {
  const paragraph = session.paragraphs(story).find((candidate) => candidate.paraId === paraId);
  const properties = paragraph?.properties ?? {};
  const numPr = properties.numPr as { numId?: number; ilvl?: number } | null | undefined;
  const empty = session.paragraphSpans(story).find((span) => span.paraId === paraId)?.length === 0;
  if (!paragraph || !empty || !numPr?.numId) return false;
  const styleId = typeof properties.pStyle === 'string' ? properties.pStyle : null;
  const level = numPr.ilvl ?? 0;
  const attrs: Record<string, unknown> = {};
  if (level > 0) {
    const values = styleValues(styleId, story, { numId: numPr.numId, ilvl: level - 1 });
    for (const key of LIST_VALUE_KEYS) attrs[key] = values[key] ?? null;
  } else {
    const styleList = styleValues(styleId, story).listNumFmt != null;
    const values = styleValues(styleId, story, false);
    Object.assign(attrs, NO_LIST, {
      numPr: styleList ? { numId: 0, ilvl: 0 } : null,
      indentLeft: values.indentLeft ?? null,
      indentFirstLine: values.indentFirstLine ?? null,
      hangingIndent: values.hangingIndent ?? null,
    });
  }
  session.setParagraphAttrs(
    { story, start: { paraId, offset: 0 }, end: { paraId, offset: 0 } },
    attrs as YrsParagraphAttrs,
    suggesting
  );
  return true;
}

/**
 * Gives the paragraph Enter made after one whose style names another next
 * style that style, starting clean as Word does: without the list the
 * previous paragraph set on itself, which the split carries.
 */
export function applyNextStyle(
  session: YrsSession,
  range: YrsStoryRange,
  nextStyleId: string,
  currentStyleId: string | null,
  styleValues: ParagraphStyleValues
): void {
  session.setParagraphAttrs(range, NO_LIST);
  session.applyParagraphStyle(range, nextStyleId, styleValues(nextStyleId, range.story), {
    [currentStyleId ?? '']: styleValues(currentStyleId, range.story),
  });
}

/**
 * The list rendering a style's numbering gives a paragraph carrying only that
 * style (`stylePpr`, resolved): its level's indents where the style sets
 * none, as the parser renders a style's list.
 */
export function styleListRendering(
  stylePpr: ParagraphFormatting | undefined,
  numbering: NumberingDefinitions | undefined
): ListRendering | null {
  const numPr = stylePpr?.numPr;
  if (!numPr || !numbering) return null;
  const map = getCachedNumberingMap(numbering);
  const listRendering = computeListRendering(numPr, map);
  if (!listRendering) return null;
  const level = map.getLevel(numPr.numId ?? 0, numPr.ilvl ?? 0)?.pPr;
  if (stylePpr?.indentLeft == null) listRendering.indentLeft = level?.indentLeft;
  if (stylePpr?.indentFirstLine == null && stylePpr?.hangingIndent == null) {
    listRendering.indentFirstLine = level?.indentFirstLine;
    listRendering.hangingIndent = level?.hangingIndent;
  }
  return listRendering;
}
