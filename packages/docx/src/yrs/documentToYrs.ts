/** Direct Document-to-yrs load projection. */

/* eslint-disable max-lines -- the complete load projection is intentionally co-located */

import { isRawXml } from '../types/content/rawXml';
import { emuToPixels } from '../utils/units';
import { isWrapNone } from '../docx/wrapTypes';
import { sdtPropsToAttrs } from '../types/sdtAttributes';
import { createStyleResolver, type StyleResolver } from '../styles';
import type {
  BlockContent,
  Chart,
  ComplexField,
  Document,
  FieldInlineContent,
  Hyperlink,
  Image,
  InlineSdt,
  MathEquation,
  NumberingDefinitions,
  Paragraph,
  ParagraphFormatting,
  ParagraphContent,
  Run,
  RunContent,
  Shape,
  SimpleField,
  SdtProperties,
  Table,
  TableBorders,
  TableCell,
  TableCellFormatting,
  TableRow,
  TextFormatting,
  Theme,
  TrackedChangeInfo,
} from '../types/document';
import { ensureHexPrefix, resolveColorToHex } from '../utils/colorResolver';
import { mergeTextFormatting } from '../utils/textFormattingMerge';
import {
  calculateRowSpans,
  tableCellParagraphFormatting,
  tableColumnCount,
  type RowSpanInfo,
} from './tableParagraphFormatting';
import type { Style } from '../types/styles';
import type { YrsRawOp, YrsSession } from './index';
import { noteYrsStoriesDirty } from './yrsToDocument';
import { seededParagraphProperties, styleListRendering } from './paragraphSeed';
import {
  blockSdtAttrsToPayload,
  blockSdtStoryId,
  dropNulls,
  footnoteStoryId,
  endnoteStoryId,
  headerFooterStoryId,
  paraAttrsToPpr,
  tableAttrsToGrid,
  tableAttrsToTblPr,
  tableCellAttrsToTcPr,
  tableCellStoryId,
  tableRowAttrsToTrPr,
  type YrsAttrs,
} from './storyAttributes';

type Attrs = Record<string, unknown>;

interface MarkDescriptor {
  name: string;
  /** Complete schema attrs, including non-null defaults. */
  attrs: Attrs;
}

interface TextUnit {
  kind: 'text';
  text: string;
  attrs: YrsAttrs;
  /** Node width, used only for bookmark offsets. */
  pmSize: number;
  /** Comments whose range holds this unit, in the order they opened. */
  comments?: number[];
  marks: MarkDescriptor[];
}

interface EmbedUnit {
  kind: 'embed';
  embedKind: string;
  payload: Attrs;
  attrs: YrsAttrs;
  pmSize: number;
  comments?: number[];
  marks: MarkDescriptor[];
}

type InlineUnit = TextUnit | EmbedUnit;

interface StoryPlan {
  storyId: string;
  units: InlineUnit[];
  commentCoverage: Map<number, Array<[number, number]>>;
}

interface StoryOptions {
  appendBodyTail: boolean;
  extraRunFormatting?: TextFormatting;
  tableParagraphFormatting?: ParagraphFormatting;
}

interface ProjectedCell {
  paragraphFormatting?: ParagraphFormatting;
  attrs: Attrs;
  content: BlockContent[];
  extraRunFormatting?: TextFormatting;
}

interface ProjectedRow {
  attrs: Attrs;
  cells: ProjectedCell[];
}

interface ProjectedTable {
  attrs: Attrs;
  rows: ProjectedRow[];
}

interface LoweringContext {
  styleResolver: StyleResolver | null;
  theme: Theme | null;
  plans: StoryPlan[];
  compatibilityMode: number;
  /** Comment ranges open at the lowering point, each with the plan it opened in. */
  openComments: Array<[id: number, plan: number]>;
}

/** The comments a unit of `plan` falls inside: a range covers only the story it opened in. */
function covering(context: LoweringContext, plan: number): number[] {
  return context.openComments.filter(([, home]) => home === plan).map(([id]) => id);
}

function compatibilityModeFromDocument(document: Document): number {
  const mode = document.package.settings?.compatibilityFlags?.compatibilityMode;
  return typeof mode === 'number' && Number.isFinite(mode) ? Math.trunc(mode) : 12;
}

const BOOLEAN_MARKS = new Set([
  'bold',
  'italic',
  'superscript',
  'subscript',
  'allCaps',
  'smallCaps',
  'emboss',
  'imprint',
  'textShadow',
  'textOutline',
  'hidden',
  'rtl',
]);

function stableStringify(value: unknown): string {
  if (value === null || value === undefined) return 'null';
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(',')}]`;
  if (typeof value === 'object') {
    const object = value as Attrs;
    return `{${Object.keys(object)
      .sort()
      .map((key) => `${JSON.stringify(key)}:${stableStringify(object[key])}`)
      .join(',')}}`;
  }
  return JSON.stringify(value);
}

function resolveTextFormatting(
  formatting: TextFormatting | undefined,
  styleResolver: StyleResolver | null
): TextFormatting | undefined {
  if (!formatting || !styleResolver) return formatting;
  const styleFormatting = styleResolver.resolveRunStyle(formatting.styleId);
  return styleFormatting ? mergeTextFormatting(styleFormatting, formatting) : formatting;
}

function formattingToMarks(formatting: TextFormatting | undefined): MarkDescriptor[] {
  if (!formatting) return [];
  const marks: MarkDescriptor[] = [];
  const add = (name: string, attrs: Attrs = {}) => marks.push({ name, attrs });

  if (formatting.bold) add('bold');
  if (formatting.italic) add('italic');
  if (formatting.underline && formatting.underline.style !== 'none') {
    add('underline', {
      style: formatting.underline.style,
      color: formatting.underline.color ?? null,
    });
  }
  if (formatting.strike || formatting.doubleStrike) {
    add('strike', { double: formatting.doubleStrike || false });
  }
  if (formatting.color && !formatting.color.auto) {
    add('textColor', {
      rgb: formatting.color.rgb ?? null,
      themeColor: formatting.color.themeColor ?? null,
      themeTint: formatting.color.themeTint ?? null,
      themeShade: formatting.color.themeShade ?? null,
    });
  }

  const shadingFill = formatting.shading?.fill;
  const shadingHighlight =
    (!formatting.shading?.pattern || formatting.shading.pattern === 'clear') &&
    shadingFill?.rgb &&
    !shadingFill.auto
      ? ensureHexPrefix(shadingFill.rgb)
      : undefined;
  const highlight =
    formatting.highlight && formatting.highlight !== 'none'
      ? formatting.highlight
      : shadingHighlight;
  if (highlight) add('highlight', { color: highlight });

  if (formatting.fontSize != null || formatting.fontSizeCs != null) {
    add('fontSize', {
      size: formatting.fontSize ?? null,
      sizeCs: formatting.fontSizeCs ?? null,
    });
  }
  if (formatting.fontFamily) {
    add('fontFamily', {
      ascii: formatting.fontFamily.ascii ?? null,
      hAnsi: formatting.fontFamily.hAnsi ?? null,
      eastAsia: formatting.fontFamily.eastAsia ?? null,
      cs: formatting.fontFamily.cs ?? null,
      asciiTheme: formatting.fontFamily.asciiTheme ?? null,
      hAnsiTheme: formatting.fontFamily.hAnsiTheme ?? null,
      eastAsiaTheme: formatting.fontFamily.eastAsiaTheme ?? null,
      csTheme: formatting.fontFamily.csTheme ?? null,
    });
  }
  if (formatting.language) {
    add('language', {
      latin: formatting.language.latin ?? null,
      eastAsia: formatting.language.eastAsia ?? null,
      bidi: formatting.language.bidi ?? null,
    });
  }
  if (formatting.fontFamily?.hint) add('fontHint', { hint: formatting.fontFamily.hint });
  if (formatting.vertAlign === 'superscript') add('superscript');
  else if (formatting.vertAlign === 'subscript') add('subscript');
  if (formatting.allCaps) add('allCaps');
  if (formatting.smallCaps) add('smallCaps');
  if (
    formatting.spacing != null ||
    formatting.position != null ||
    formatting.scale != null ||
    formatting.kerning != null
  ) {
    add('characterSpacing', {
      spacing: formatting.spacing ?? null,
      position: formatting.position ?? null,
      scale: formatting.scale ?? null,
      kerning: formatting.kerning ?? null,
    });
  }
  if (formatting.emboss) add('emboss');
  if (formatting.imprint) add('imprint');
  if (formatting.shadow) add('textShadow');
  if (formatting.emphasisMark && formatting.emphasisMark !== 'none') {
    add('emphasisMark', { type: formatting.emphasisMark });
  }
  if (formatting.outline) add('textOutline');
  if (formatting.hidden) add('hidden');
  if (formatting.rtl) add('rtl');
  // Document-grid opt-out (w:snapToGrid, default on): only an authored off
  // becomes a mark, mirroring how the layout bridge reads it.
  if (formatting.snapToGrid === false) marks.push({ name: 'snapToGrid', attrs: {} });
  if (formatting.effect && formatting.effect !== 'none') {
    add('textEffect', { effect: formatting.effect });
  }
  if (formatting.modernEffects) {
    add('modernTextEffects', { effects: formatting.modernEffects });
  }
  if (formatting.styleId) add('runStyle', { styleId: formatting.styleId });
  return marks;
}

function marksToYrsAttrs(marks: readonly MarkDescriptor[]): YrsAttrs {
  const attrs: YrsAttrs = {};
  for (const mark of marks) {
    if (mark.name === 'comment' || mark.name === 'footnoteRef') continue;
    if (BOOLEAN_MARKS.has(mark.name)) {
      attrs[mark.name] = true;
    } else if (mark.name === 'snapToGrid') {
      attrs.snapToGrid = false;
    } else if (mark.name === 'highlight') {
      attrs.highlight = mark.attrs.color;
    } else if (mark.name === 'insertion' || mark.name === 'deletion') {
      attrs[mark.name === 'insertion' ? 'ins' : 'del'] = dropNulls({
        id: mark.attrs.revisionId,
        author: mark.attrs.author,
        date: mark.attrs.date,
      });
    } else {
      attrs[mark.name] = dropNulls(mark.attrs);
    }
  }
  return attrs;
}

function marksKey(marks: readonly MarkDescriptor[]): string {
  return marks
    .filter((mark) => mark.name !== 'hyperlink' && mark.name !== 'comment')
    .map((mark) => `${mark.name}:${mark.name === 'rPrChange' ? stableStringify(mark.attrs) : JSON.stringify(mark.attrs)}`)
    .sort()
    .join('|');
}

function withMark(marks: readonly MarkDescriptor[], mark: MarkDescriptor): MarkDescriptor[] {
  return [...marks.filter((candidate) => candidate.name !== mark.name), mark];
}

function textUnit(text: string, marks: readonly MarkDescriptor[]): TextUnit {
  return {
    kind: 'text',
    text,
    attrs: marksToYrsAttrs(marks),
    pmSize: text.length,
    marks: [...marks],
  };
}

function embedUnit(
  embedKind: string,
  payload: Attrs,
  marks: readonly MarkDescriptor[] = [],
  pmSize = 1
): EmbedUnit {
  return {
    kind: 'embed',
    embedKind,
    payload,
    attrs: marksToYrsAttrs(marks),
    pmSize,
    marks: [...marks],
  };
}

function runFormatMarks(
  run: Run,
  styleFormatting: TextFormatting | undefined,
  styleResolver: StyleResolver | null
): MarkDescriptor[] {
  const runStyleFormatting = run.formatting?.styleId
    ? styleResolver?.getRunStyleOwnProperties(run.formatting.styleId)
    : undefined;
  const marks = formattingToMarks(
    mergeTextFormatting(mergeTextFormatting(styleFormatting, runStyleFormatting), run.formatting)
  );
  const styleId = run.formatting?.styleId;
  const styleName = styleId ? styleResolver?.getStyle(styleId)?.name : undefined;
  if ([styleId, styleName].some((name) => /^(?:Followed)?Hyperlink$/i.test(name ?? ''))) {
    for (const [property, name] of [['color', 'textColor'], ['underline', 'underline']] as const) {
      if (run.formatting?.[property] === undefined && runStyleFormatting?.[property] !== undefined) {
        const mark = marks.find((mark) => mark.name === name);
        if (mark) mark.attrs.inheritedHyperlink = true;
      }
    }
  }
  return marks;
}

function runMarks(
  run: Run,
  styleFormatting: TextFormatting | undefined,
  styleResolver: StyleResolver | null
): MarkDescriptor[] {
  const marks = runFormatMarks(run, styleFormatting, styleResolver);
  if (run.propertyChanges?.length) marks.push({ name: 'rPrChange', attrs: {
    changes: run.propertyChanges.map((change) => ({
      ...change,
      previousAttributes: marksToYrsAttrs(runFormatMarks({ type: 'run', content: [], formatting: change.previousFormatting }, styleFormatting, styleResolver)),
      currentAttributes: marksToYrsAttrs(runFormatMarks({ type: 'run', content: [], formatting: change.currentFormatting }, styleFormatting, styleResolver)),
    })),
  } });
  return marks;
}

function imagePayload(image: Image): Attrs {
  const width = image.size?.width ? emuToPixels(image.size.width) : undefined;
  const height = image.size?.height ? emuToPixels(image.size.height) : undefined;
  const wrapType = image.wrap.type;
  const wrapText = image.wrap.wrapText;
  const hAlign = image.position?.horizontal?.alignment;
  let cssFloat: 'left' | 'right' | 'none' | undefined;
  if (wrapType === 'inline' || wrapType === 'topAndBottom') cssFloat = 'none';
  else if (wrapType === 'square' || wrapType === 'tight' || wrapType === 'through') {
    if (wrapText === 'left') cssFloat = 'right';
    else if (wrapText === 'right') cssFloat = 'left';
    else if (hAlign === 'left' || hAlign === 'right') cssFloat = hAlign;
    else cssFloat = 'none';
  } else cssFloat = 'none';

  const displayMode =
    wrapType === 'inline'
      ? 'inline'
      : wrapType === 'topAndBottom'
        ? 'block'
        : isWrapNone(wrapType) || (cssFloat && cssFloat !== 'none')
          ? 'float'
          : 'block';
  const transforms: string[] = [];
  if (image.transform?.rotation) transforms.push(`rotate(${image.transform.rotation}deg)`);
  if (image.transform?.flipH) transforms.push('scaleX(-1)');
  if (image.transform?.flipV) transforms.push('scaleY(-1)');

  let borderWidth: number | undefined;
  let borderColor: string | undefined;
  let borderStyle: string | undefined;
  if (image.outline?.width) {
    borderWidth = Math.round((image.outline.width / 914400) * 96 * 100) / 100;
    if (image.outline.color?.rgb) borderColor = `#${image.outline.color.rgb}`;
    const styles: Record<string, string> = {
      solid: 'solid',
      dot: 'dotted',
      dash: 'dashed',
      lgDash: 'dashed',
      dashDot: 'dashed',
      lgDashDot: 'dashed',
      lgDashDotDot: 'dashed',
      sysDot: 'dotted',
      sysDash: 'dashed',
      sysDashDot: 'dashed',
      sysDashDotDot: 'dashed',
    };
    borderStyle = image.outline.style ? styles[image.outline.style] || 'solid' : 'solid';
  }

  return dropNulls({
    src: image.src || '',
    alt: image.alt ?? null,
    title: image.title ?? null,
    width: width ?? null,
    height: height ?? null,
    rId: image.rId ?? null,
    wrapType,
    displayMode,
    cssFloat: cssFloat ?? null,
    transform: transforms.length > 0 ? transforms.join(' ') : null,
    distTop: image.wrap.distT != null ? emuToPixels(image.wrap.distT) : null,
    distBottom: image.wrap.distB != null ? emuToPixels(image.wrap.distB) : null,
    distLeft: image.wrap.distL != null ? emuToPixels(image.wrap.distL) : null,
    distRight: image.wrap.distR != null ? emuToPixels(image.wrap.distR) : null,
    position: image.position
      ? {
          relativeHeight: image.position.relativeHeight,
          horizontal: image.position.horizontal
            ? {
                relativeTo: image.position.horizontal.relativeTo,
                posOffset: image.position.horizontal.posOffset,
                align: image.position.horizontal.alignment,
              }
            : undefined,
          vertical: image.position.vertical
            ? {
                relativeTo: image.position.vertical.relativeTo,
                posOffset: image.position.vertical.posOffset,
                align: image.position.vertical.alignment,
              }
            : undefined,
        }
      : null,
    borderWidth: borderWidth ?? null,
    borderColor: borderColor ?? null,
    borderStyle: borderStyle ?? null,
    wrapText: wrapText ?? null,
    hlinkHref: image.hlinkHref ?? null,
    cropTop: image.crop?.top ?? null,
    cropRight: image.crop?.right ?? null,
    cropBottom: image.crop?.bottom ?? null,
    cropLeft: image.crop?.left ?? null,
    shapeType: image.shapeType ?? null,
    opacity: image.opacity ?? null,
    effectExtentTop: image.padding?.top ? emuToPixels(image.padding.top) : null,
    effectExtentBottom: image.padding?.bottom ? emuToPixels(image.padding.bottom) : null,
    effectExtentLeft: image.padding?.left ? emuToPixels(image.padding.left) : null,
    effectExtentRight: image.padding?.right ? emuToPixels(image.padding.right) : null,
    layoutInCell: image.layoutInCell ?? null,
    allowOverlap: image.allowOverlap ?? null,
  }) as Attrs;
}

/** Carries the authored XML a picture or shape replays on save while unedited. */
function withSourceXml(payload: Attrs, sourceXml: string | undefined): Attrs {
  return sourceXml === undefined ? payload : { ...payload, sourceXml };
}

function shapePayload(shape: Shape): Attrs {
  return { shapeJson: JSON.stringify(shape) };
}

function chartPayload(chart: Chart): Attrs {
  return dropNulls({
    chartJson: JSON.stringify(chart),
    chartType: chart.chartType ?? null,
    title: chart.title ?? null,
    width: chart.size?.width ? emuToPixels(chart.size.width) : 320,
    height: chart.size?.height ? emuToPixels(chart.size.height) : 220,
    rId: chart.rId ?? null,
    path: chart.path ?? null,
  }) as Attrs;
}

function fieldPayload(
  field: SimpleField | ComplexField,
  styleFormatting?: TextFormatting
): {
  payload: Attrs;
  marks: MarkDescriptor[];
} {
  let displayText = '';
  let fieldFormatting: TextFormatting | undefined;
  const runs = field.type === 'simpleField' ? field.content : field.fieldResult;
  for (const child of runs) {
    if (child.type !== 'run') continue;
    for (const content of child.content) {
      if (content.type === 'text') displayText += content.text;
    }
    if (!fieldFormatting && child.formatting) fieldFormatting = child.formatting;
  }
  const formatting =
    fieldFormatting ?? (field.type === 'complexField' ? field.formatting : undefined);
  const displayMode = field.fieldTree?.displayMode ?? 'result';
  return {
    payload: {
      fieldType: field.fieldType,
      instruction: field.instruction,
      displayText,
      fieldKind: field.type === 'simpleField' ? 'simple' : 'complex',
      fldLock: field.fldLock ?? false,
      dirty: field.dirty ?? false,
      displayMode,
      hasCachedResult: displayText.length > 0,
      fieldData: fieldJson(field),
      modelKind: 'field',
      ...(field.type === 'complexField' && field.continuation ? { continuationId: field.continuation.id } : {}),
    },
    marks: formattingToMarks(mergeTextFormatting(styleFormatting, formatting)),
  };
}

function mathPayload(math: MathEquation): Attrs {
  return {
    display: math.display,
    ommlXml: math.ommlXml,
    plainText: math.plainText || '',
  };
}

function hyperlinkMark(hyperlink: Hyperlink): MarkDescriptor {
  return {
    name: 'hyperlink',
    attrs: {
      href: hyperlink.href || (hyperlink.anchor ? `#${hyperlink.anchor}` : ''),
      tooltip: hyperlink.tooltip ?? null,
      rId: hyperlink.rId ?? null,
    },
  };
}

/** Fields `continuedResultTail` trimmed: seed.rs serializes them with sorted keys, as it does every changed value. */
const trimmedFields = new WeakSet<object>();

/** A field's data as seed.rs stores it (`source_json`, `js_json` for a field it changed). */
function fieldJson(field: SimpleField | ComplexField): string {
  if (!trimmedFields.has(field)) return JSON.stringify(field);
  return JSON.stringify(field, (_, value: unknown) =>
    value && typeof value === 'object' && !Array.isArray(value)
      ? Object.fromEntries(Object.entries(value).sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0)))
      : value
  );
}

/**
 * A field whose result continues into a later paragraph, and the plain runs (text, plain line breaks, comment references,
 * and tabs or positional tabs without formatting) that end its result in this one, after its projected children: ordinary
 * text after the field's embed (`continued_result_tail` in crates/docx-edit/src/seed.rs).
 */
function continuedResultTail(value: SimpleField | ComplexField): [SimpleField | ComplexField, Run[]] {
  const inline = value.type === 'complexField' && value.continuation?.end && !value.continuation.separate
    ? value.structuredResult?.inline ?? []
    : [];
  // The save drops a tab's or break's own formatting and a break's w:clear (`tail_run` in seed.rs).
  const bare = (content: RunContent) => content.type === 'tab' || content.type === 'ptab' || content.type === 'break';
  const plain = (node: FieldInlineContent) =>
    node.type === 'run' &&
    node.content.every(
      (content) =>
        content.type === 'text' ||
        content.type === 'tab' ||
        content.type === 'ptab' ||
        content.type === 'commentReference' ||
        (content.type === 'break' &&
          (content.breakType === undefined || content.breakType === 'textWrapping') &&
          (content.clear === undefined || content.clear === 'none'))
    ) &&
    !(Object.keys(node.formatting ?? {}).length && node.content.some(bare));
  let tail = 0;
  while (tail < inline.length && plain(inline[inline.length - 1 - tail]!)) tail += 1;
  if (value.type !== 'complexField' || tail === 0) return [value, []];
  const kept = inline.slice(0, inline.length - tail);
  const fieldResult = shownRuns(kept);
  // With no result run left, the field takes its first one's formatting, which its characters save with (seed.rs).
  const first = value.fieldResult[0];
  const field: ComplexField = {
    ...value,
    structuredResult: { ...value.structuredResult, inline: kept },
    fieldResult,
    ...(value.fieldTree?.result?.inline ? { fieldTree: { ...value.fieldTree, result: { ...value.fieldTree.result, inline: kept } } } : {}),
    ...(!fieldResult.length && first ? { formatting: first.formatting } : {}),
  };
  trimmedFields.add(field);
  return [field, inline.slice(inline.length - tail) as Run[]];
}

function fieldToUnits(
  value: SimpleField | ComplexField,
  styleFormatting: TextFormatting | undefined,
  styleResolver: StyleResolver | null,
  projectionId: number
): InlineUnit[] {
  const result = value.structuredResult?.inline ?? [];
  const code = value.type === 'complexField' ? value.structuredCode?.inline ?? [] : [];
  const projectedChildren = [
    ...code.map((child, index) => ({ child, index: -index - 1 })),
    ...result.map((child, index) => ({ child, index })),
  ];
  if (numericFieldInstruction(value.instruction) || value.type !== 'complexField' || !projectedChildren.some(({ child }) => child.type === 'hyperlink' || child.type === 'simpleField')) {
    const field = fieldPayload(value, styleFormatting);
    return [embedUnit('field', field.payload, field.marks)];
  }
  const units: InlineUnit[] = [];
  const children: Attrs[] = [];
  // Result runs after a projected simple field are its field's text too.
  let afterSimple = false;
  const projectedRuns = new Set<number>();
  projectedChildren.forEach(({ child, index }) => {
    let projected: InlineUnit[];
    if (child.type === 'run') {
      if (!afterSimple) return;
      projected = runToUnits(child, styleFormatting, styleResolver);
      if (!projected.every((unit) => unit.kind === 'text')) return;
      projectedRuns.add(index);
    } else if (child.type === 'hyperlink' || child.type === 'simpleField') {
      afterSimple = child.type === 'simpleField' && index >= 0;
      const nested = child.type === 'simpleField' ? fieldPayload(child, styleFormatting) : null;
      projected = child.type === 'hyperlink'
        ? hyperlinkToUnits(child, styleFormatting, styleResolver)
        : [embedUnit('field', nested!.payload, nested!.marks)];
    } else return;
    children.push({ index, items: projected.map((unit) => unit.kind === 'text'
      ? { kind: 'text', text: unit.text, attributes: unit.attrs }
      : { kind: 'embed', embedKind: unit.embedKind, payload: unit.payload, attributes: unit.attrs }) });
    for (const unit of projected) unit.attrs = { ...unit.attrs, fieldResult: { id: projectionId, index } };
    units.push(...projected);
  });
  // The field shows the result its projected children leave: runs, and a
  // kept change's or content control's text.
  const visible = {
    ...value,
    fieldResult: shownRuns(
      result.filter((child, index) => !projectedRuns.has(index) && (child.type === 'run' || child.type === 'rawXml' || child.type === 'inlineSdt'))
    ),
  };
  const field = fieldPayload(visible, styleFormatting);
  field.payload.fieldData = fieldJson(value);
  field.payload.resultProjection = { id: projectionId, children };
  units.push(embedUnit('field', field.payload, field.marks));
  return units;
}

/** The runs field result `nodes` show (seed.rs `shown_runs`). */
function shownRuns(nodes: readonly FieldInlineContent[]): Run[] {
  return nodes.flatMap((node): Run[] => {
    switch (node.type) {
      case 'run':
        return [node];
      case 'hyperlink':
        return shownRuns(node.children.filter((child) => child.type === 'run'));
      case 'simpleField':
        return node.content.filter((child): child is Run => child.type === 'run');
      case 'complexField':
        return node.fieldResult;
      case 'inlineSdt':
      case 'insertion':
      case 'moveTo':
        return shownRuns(node.content);
      case 'rawXml':
        return node.shown ?? [];
      default:
        return [];
    }
  });
}

function noteRefUnit(
  id: number,
  noteType: 'footnote' | 'endnote',
  marks: readonly MarkDescriptor[]
): EmbedUnit {
  const noteMark: MarkDescriptor = {
    name: 'footnoteRef',
    attrs: { id: String(id), noteType },
  };
  const allMarks = [...marks, noteMark];
  return embedUnit(
    'noteRef',
    noteType === 'endnote' ? { endnoteRefId: id } : { footnoteRefId: id },
    allMarks
  );
}

function runContentToUnits(content: RunContent, marks: readonly MarkDescriptor[]): InlineUnit[] {
  switch (content.type) {
    case 'text':
      return content.text ? [textUnit(content.text, marks)] : [];
    case 'tab':
      return [textUnit('\t', marks)];
    case 'ptab': {
      // Drawn as a tab; its own attributes save it back as `w:ptab` (seed.rs).
      const { type: _, ...ptab } = content;
      return [embedUnit('tab', { ptab }, marks)];
    }
    case 'break':
      return content.breakType === undefined || content.breakType === 'textWrapping'
        ? [embedUnit('break', {}, marks)]
        : [embedUnit(content.breakType === 'page' ? 'pageBreak' : 'columnBreak', {}, marks)];
    case 'softHyphen':
      return [textUnit('\u00ad', marks)];
    case 'noBreakHyphen':
      return [textUnit('\u2011', marks)];
    case 'symbol': {
      const codePoint = Number.parseInt(content.char, 16);
      if (!Number.isInteger(codePoint) || codePoint < 0 || codePoint > 0x10ffff) return [];
      const symbolMark: MarkDescriptor = {
        name: 'fontFamily',
        attrs: {
          ascii: content.font || null,
          hAnsi: content.font || null,
          eastAsia: content.font || null,
          cs: content.font || null,
          asciiTheme: null,
          hAnsiTheme: null,
          eastAsiaTheme: null,
          csTheme: null,
        },
      };
      return [textUnit(String.fromCodePoint(codePoint), withMark(marks, symbolMark))];
    }
    case 'commentReference':
      return [
        embedUnit('field', {
          fieldType: 'COMMENT',
          instruction: '',
          displayText: '',
          fieldKind: 'simple',
          fldLock: false,
          dirty: false,
          displayMode: 'result',
          hasCachedResult: false,
          modelKind: 'commentReference',
          ...(content.id !== undefined ? { commentId: content.id } : {}),
        }),
      ];
    case 'drawing':
      return [embedUnit('image', withSourceXml(imagePayload(content.image), content.sourceXml))];
    case 'horizontalRule':
      return [embedUnit('horizontalRule', { rule: content.rule }, marks)];
    case 'shape':
      return [embedUnit('shape', withSourceXml(shapePayload(content.shape), content.sourceXml))];
    case 'chart':
      return [embedUnit('chart', chartPayload(content.chart))];
    case 'opaqueDrawing':
      return [embedUnit('opaqueDrawing', { kind: content.kind, xml: content.xml })];
    case 'footnoteRef':
      return [noteRefUnit(content.id, 'footnote', marks)];
    case 'endnoteRef':
      return [noteRefUnit(content.id, 'endnote', marks)];
    default:
      return [];
  }
}

function runToUnits(
  run: Run,
  styleFormatting: TextFormatting | undefined,
  styleResolver: StyleResolver | null,
  extraMarks: readonly MarkDescriptor[] = []
): InlineUnit[] {
  const marks = [...runMarks(run, styleFormatting, styleResolver), ...extraMarks];
  return run.content.flatMap((content) => content.type === 'fieldChar' && content.continuationId
    ? [embedUnit('bookmark', {
      id: content.continuationId,
      kind: `field${content.charType}`,
      run: { type: 'run', content: [content], ...(run.formatting && { formatting: run.formatting }) },
    }, [], 0)]
    : runContentToUnits(content, marks));
}

function hyperlinkToUnits(
  hyperlink: Hyperlink,
  styleFormatting: TextFormatting | undefined,
  styleResolver: StyleResolver | null,
  extraMarks: readonly MarkDescriptor[] = []
): InlineUnit[] {
  const marks = [...extraMarks, hyperlinkMark(hyperlink)];
  return (hyperlink.structuredChildren ?? hyperlink.children).flatMap((child) =>
    inlineToUnits(child, styleFormatting, styleResolver, marks)
  );
}

function inlineToUnits(
  child: FieldInlineContent,
  styleFormatting: TextFormatting | undefined,
  styleResolver: StyleResolver | null,
  extraMarks: readonly MarkDescriptor[] = []
): InlineUnit[] {
  let units: InlineUnit[];
  switch (child.type) {
    case 'run': units = runToUnits(child, styleFormatting, styleResolver); break;
    case 'hyperlink': units = hyperlinkToUnits(child, styleFormatting, styleResolver); break;
    case 'simpleField': case 'complexField': {
      const field = fieldPayload(child, styleFormatting);
      units = [embedUnit('field', field.payload, field.marks)];
      break;
    }
    case 'inlineSdt': units = [embedUnit('sdt', sdtPayload(child, styleFormatting, styleResolver))]; break;
    case 'mathEquation': units = [embedUnit('math', mathPayload(child))]; break;
    case 'bookmarkStart': case 'bookmarkEnd': {
      const data: Record<string, unknown> = { ...child, kind: child.type === 'bookmarkStart' ? 'start' : 'end' };
      delete data.type;
      delete data.position;
      units = [embedUnit('bookmark', data, [], 0)];
      break;
    }
    case 'insertion': case 'deletion': case 'moveFrom': case 'moveTo':
      units = trackedToUnits(child, styleFormatting, styleResolver); break;
    default: return [];
  }
  for (const unit of units) {
    for (const mark of extraMarks) unit.marks = withMark(unit.marks, mark);
    unit.attrs = marksToYrsAttrs(unit.marks);
    if (unit.kind === 'embed' && unit.embedKind === 'bookmark' && unit.attrs.hyperlink)
      unit.payload.inHyperlink = true;
  }
  return units;
}

function trackedMark(
  info: TrackedChangeInfo,
  kind: 'insertion' | 'deletion',
  isMovePair: boolean
): MarkDescriptor {
  return {
    name: kind,
    attrs: {
      revisionId: info.id,
      author: info.author,
      date: info.date ?? null,
      isMovePair,
    },
  };
}

function trackedToUnits(
  content: Extract<ParagraphContent, { type: 'insertion' | 'deletion' | 'moveFrom' | 'moveTo' }>,
  styleFormatting: TextFormatting | undefined,
  styleResolver: StyleResolver | null
): InlineUnit[] {
  const kind = content.type === 'insertion' || content.type === 'moveTo' ? 'insertion' : 'deletion';
  const mark = trackedMark(
    content.info,
    kind,
    content.type === 'moveFrom' || content.type === 'moveTo'
  );
  return content.content.flatMap((child) => inlineToUnits(child, styleFormatting, styleResolver, [mark]));
}

function sdtPayload(
  sdt: InlineSdt,
  styleFormatting: TextFormatting | undefined,
  styleResolver: StyleResolver | null
): Attrs {
  const content: Array<
    | { kind: 'text'; text: string; attrs: YrsAttrs }
    | { kind: 'tab'; attrs: YrsAttrs }
    | { kind: string; payload: Attrs; attrs: YrsAttrs }
  > = [];
  const append = (unit: InlineUnit): void => {
    if (unit.kind === 'text') {
      const kind = unit.text === '\t' ? 'tab' : 'text';
      if (kind === 'tab') {
        content.push({ kind: 'tab', attrs: unit.attrs });
        return;
      }
      const previous = content[content.length - 1];
      if (
        previous &&
        'text' in previous &&
        stableStringify(previous.attrs) === stableStringify(unit.attrs)
      ) {
        previous.text += unit.text;
      } else {
        content.push({ kind: 'text', text: unit.text, attrs: unit.attrs });
      }
      return;
    }
    content.push({ kind: unit.embedKind, payload: unit.payload, attrs: unit.attrs });
  };

  for (const child of sdt.content) {
    inlineToUnits(child, styleFormatting, styleResolver).forEach(append);
  }
  return dropNulls({
    ...sdtPropsToAttrs(sdt.properties),
    propertiesJson: JSON.stringify(sdt.properties),
    content,
  }) as Attrs;
}

function paragraphStyleFormatting(
  paragraph: Paragraph,
  styleResolver: StyleResolver | null,
  extraRunFormatting?: TextFormatting
): TextFormatting | undefined {
  const styleFormatting = styleResolver
    ? styleResolver.resolveParagraphStyle(paragraph.formatting?.styleId).runFormatting
    : undefined;
  return mergeTextFormatting(styleFormatting, extraRunFormatting);
}

/**
 * Note number marks carry no story unit, so the run boundary cache is the only
 * place a saved paragraph can learn they were there.
 */
function noteRefMarkTypes(run: Run): string[] {
  const marks: string[] = [];
  for (const content of run.content) {
    if (content.type === 'footnoteRefMark') marks.push('footnote');
    else if (content.type === 'endnoteRefMark') marks.push('endnote');
  }
  return marks;
}

function unitsText(units: readonly InlineUnit[]): string {
  return units
    .map((unit) => {
      if (unit.kind === 'text') return unit.text;
      const id = unit.payload.footnoteRefId ?? unit.payload.endnoteRefId;
      return String(id ?? '');
    })
    .join('');
}

/**
 * `w:br w:type="page"|"column"`, which the story carries as a block embed
 * beside the paragraph instead of as an inline unit.
 */
function flowBreakType(content: RunContent): 'page' | 'column' | null {
  if (content.type !== 'break') return null;
  return content.breakType === 'page' || content.breakType === 'column' ? content.breakType : null;
}

/**
 * Where a run's flow breaks sit in its text. They occupy no story unit, so
 * the save projection rebuilds them from these offsets.
 */
function flowBreakOffsets(run: Run): Array<{ offset: number; type: 'page' | 'column' }> {
  if (!run.content.some((content) => flowBreakType(content) !== null)) return [];
  const breaks: Array<{ offset: number; type: 'page' | 'column' }> = [];
  let offset = 0;
  for (const content of run.content) {
    const type = flowBreakType(content);
    if (type !== null) {
      breaks.push({ offset, type });
      continue;
    }
    offset += unitsText(runContentToUnits(content, [])).length;
  }
  return breaks;
}

function runBoundary(
  run: Run,
  styleFormatting: TextFormatting | undefined,
  styleResolver: StyleResolver | null,
  units = runToUnits(run, styleFormatting, styleResolver)
): Attrs | null {
  if (units.some((unit) => unit.kind !== 'text' && unit.embedKind !== 'noteRef')) return null;
  const keys = units.map((unit) => marksKey(unit.marks));
  if (keys.some((key) => key !== keys[0])) return null;
  const marks = noteRefMarkTypes(run);
  const breaks = flowBreakOffsets(run);
  const key = keys[0];
  const text = unitsText(units);
  return {
    text,
    ...(marks.length > 0 ? { noteMarks: marks } : {}),
    ...(breaks.length > 0 ? { breaks } : {}),
    ...(key !== undefined ? { marksKey: key } : {}),
    // Only a run without text restores from cached formatting; others use live marks.
    ...(run.formatting && text === '' ? { formatting: run.formatting } : {}),
    ...(run.propertyChanges ? { propertyChanges: run.propertyChanges } : {}),
  };
}

/** Where a style is applied: the cell's table-style paragraph formatting, the package's numbering. */
export interface StyleValueContext {
  /** From `cellParagraphFormatting`, for a paragraph in a table cell. */
  cell?: ParagraphFormatting;
  /** The package's numbering, so a style's own list renders as the seed renders it. */
  numbering?: NumberingDefinitions;
  /** Numbering set on the paragraph itself, rendered from its own level. */
  numPr?: ParagraphFormatting['numPr'];
}

/**
 * The paragraph values applying `styleId` gives a paragraph: what the seed
 * gives one carrying only that style there.
 */
export function styleParagraphValues(
  styles: StyleResolver | null,
  styleId: string | null,
  { cell, numbering, numPr }: StyleValueContext = {}
): Attrs {
  const paragraph: Paragraph = { type: 'paragraph', formatting: { styleId: styleId ?? undefined, numPr }, content: [] };
  const listRendering = styleListRendering(
    numPr ? { numPr } : styles?.resolveParagraphStyle(styleId, cell).paragraphFormatting,
    numbering
  );
  if (listRendering) paragraph.listRendering = listRendering;
  return paragraphAttrs(paragraph, styles, [], undefined, cell);
}

function paragraphAttrs(
  paragraph: Paragraph,
  styleResolver: StyleResolver | null,
  units: readonly InlineUnit[],
  runBoundaries: Attrs[] | undefined,
  tableParagraphFormatting?: ParagraphFormatting
): Attrs {
  const formatting = paragraph.formatting;
  const styleId = formatting?.styleId;
  const attrs: Attrs = {
    paraId: paragraph.paraId ?? null,
    textId: paragraph.textId ?? null,
    styleId: styleId ?? null,
    numPr: formatting?.numPr ?? null,
    numPrFromStyle: formatting?.numPrFromStyle ?? null,
    listNumFmt: paragraph.listRendering?.numFmt ?? null,
    listIsBullet: paragraph.listRendering?.isBullet ?? null,
    listMarker: paragraph.listRendering?.marker ?? null,
    listMarkerHidden: paragraph.listRendering?.markerHidden || null,
    listMarkerFontFamily: paragraph.listRendering?.markerFontFamily || null,
    listMarkerFontSize: paragraph.listRendering?.markerFontSize || null,
    listMarkerBold: paragraph.listRendering?.markerBold ?? null,
    listMarkerItalic: paragraph.listRendering?.markerItalic ?? null,
    listMarkerColor: paragraph.listRendering?.markerColor ?? null,
    listMarkerSuffix: paragraph.listRendering?.markerSuffix || null,
    listLevelNumFmts: paragraph.listRendering?.levelNumFmts || null,
    listAbstractNumId: paragraph.listRendering?.abstractNumId ?? null,
    listStartOverride: paragraph.listRendering?.startOverride ?? null,
    _originalFormatting: formatting ?? null,
  };

  if (styleResolver) {
    const resolved = styleResolver.resolveParagraphStyle(styleId, tableParagraphFormatting);
    Object.assign(attrs, seededParagraphProperties(paragraph, resolved.paragraphFormatting));
    const defaultChar = styleResolver.getDefaultCharacterStyle()?.rPr;
    const styleRpr = defaultChar
      ? mergeTextFormatting(resolved.runFormatting, defaultChar)
      : resolved.runFormatting;
    attrs.defaultTextFormatting =
      mergeTextFormatting(
        styleRpr,
        resolveTextFormatting(formatting?.runProperties, styleResolver)
      ) ?? null;
  } else {
    Object.assign(attrs, seededParagraphProperties(paragraph, null));
    attrs.defaultTextFormatting = formatting?.runProperties ?? null;
  }

  if (paragraph.sectionProperties) {
    attrs._sectionProperties = paragraph.sectionProperties;
    const start = paragraph.sectionProperties.sectionStart;
    if (
      start === 'nextPage' ||
      start === 'continuous' ||
      start === 'oddPage' ||
      start === 'evenPage' ||
      start === 'nextColumn'
    ) {
      attrs.sectionBreakType = start;
    }
  }
  if (paragraph.renderedPageBreakBefore) attrs.renderedPageBreakBefore = true;
  if (paragraph.pPrIns) {
    attrs.pPrIns = {
      revisionId: paragraph.pPrIns.id,
      author: paragraph.pPrIns.author,
      date: paragraph.pPrIns.date ?? null,
    };
  }
  if (paragraph.pPrDel) {
    attrs.pPrDel = {
      revisionId: paragraph.pPrDel.id,
      author: paragraph.pPrDel.author,
      date: paragraph.pPrDel.date ?? null,
    };
  }
  if (paragraph.propertyChanges?.length) attrs.pPrChange = paragraph.propertyChanges;

  const bookmarks: Attrs[] = [];
  let contentIndex = 0;
  let pmOffset = 0;
  // A bookmark ahead of some of the paragraph's leading breaks records how
  // many follow it, so the save writes it back before them.
  const leadingItems = paragraphFlowBreaks(paragraph)[0].map(({ item }) => item);
  const breaksAfter = (index: number) => {
    const after = leadingItems.filter((item) => item > index).length;
    return after > 0 ? { breaksAfter: after } : {};
  };
  // Continued field characters share offsets with bookmarks; with any in the
  // paragraph, a bookmark records how many precede it.
  const fieldMarker = (unit: InlineUnit | undefined) =>
    unit?.kind === 'embed' && unit.embedKind === 'bookmark' &&
    (unit.payload.kind === 'fieldseparate' || unit.payload.kind === 'fieldend');
  const fieldsBefore = units.some(fieldMarker) ? (count: number) => ({ fieldsBefore: count }) : () => ({});
  let fields = 0;
  for (const [index, content] of paragraph.content.entries()) {
    if (content.type === 'bookmarkStart') {
      bookmarks.push({
        id: content.id,
        name: content.name,
        kind: 'start',
        offset: pmOffset,
        ...(content.colFirst !== undefined ? { colFirst: content.colFirst } : {}),
        ...(content.colLast !== undefined ? { colLast: content.colLast } : {}),
        ...breaksAfter(index),
        ...fieldsBefore(fields),
      });
    } else if (content.type === 'bookmarkEnd') {
      bookmarks.push({ id: content.id, kind: 'end', offset: pmOffset, ...breaksAfter(index), ...fieldsBefore(fields) });
    } else {
      const count = unitsForParagraphContent(content);
      for (let index = 0; index < count; index += 1) {
        pmOffset += units[contentIndex]?.pmSize ?? 0;
        if (fieldMarker(units[contentIndex])) fields += 1;
        contentIndex += 1;
      }
    }
  }
  if (bookmarks.length > 0) attrs.bookmarks = bookmarks;
  if (runBoundaries?.length) attrs._originalRunBoundaries = runBoundaries;
  return attrs;
}

/* Set by paragraphUnits while paragraphAttrs calculates bookmark offsets. */
const paragraphContentUnitCounts = new WeakMap<object, number>();

function unitsForParagraphContent(content: ParagraphContent): number {
  return paragraphContentUnitCounts.get(content as object) ?? 0;
}

/** Opens or closes the range a comment marker marks. */
function markComment(
  openComments: LoweringContext['openComments'],
  content: ParagraphContent,
  plan: number
): void {
  if (content.type === 'commentRangeStart') {
    if (!openComments.some(([id]) => id === content.id)) openComments.push([content.id, plan]);
  } else if (content.type === 'commentRangeEnd') {
    const open = openComments.findIndex(([id]) => id === content.id);
    if (open >= 0) openComments.splice(open, 1);
  }
}

function paragraphUnits(
  context: LoweringContext,
  plan: number,
  paragraph: Paragraph,
  extraRunFormatting?: TextFormatting,
  tableParagraphFormatting?: ParagraphFormatting
): { units: InlineUnit[]; ppr: Attrs } {
  const { styleResolver, openComments } = context;
  const units: InlineUnit[] = [];
  const [, , inlineBreaks] = paragraphFlowBreaks(paragraph);
  let boundaries: Attrs[] | undefined = [];
  const styleFormatting = paragraphStyleFormatting(paragraph, styleResolver, extraRunFormatting);

  for (const [contentIndex, content] of paragraph.content.entries()) {
    const start = units.length;
    let breakIndex = 0;
    const keepInline = (unit: InlineUnit) => unit.kind !== 'embed' || !['pageBreak', 'columnBreak'].includes(unit.embedKind)
      || (inlineBreaks[contentIndex][breakIndex++] ?? false);
    if (content.type === 'commentRangeStart' || content.type === 'commentRangeEnd') {
      markComment(openComments, content, plan);
    } else if (content.type === 'run') {
      const runUnits = runToUnits(content, styleFormatting, styleResolver).filter(keepInline);
      const boundary = runBoundary(content, styleFormatting, styleResolver, runUnits);
      if (boundary && boundaries) boundaries.push(boundary);
      else boundaries = undefined;
      units.push(...runUnits);
    } else if (content.type === 'hyperlink') {
      boundaries = undefined;
      units.push(...hyperlinkToUnits(content, styleFormatting, styleResolver));
    } else if (content.type === 'simpleField' || content.type === 'complexField') {
      boundaries = undefined;
      const [field, tail] = continuedResultTail(content);
      units.push(...fieldToUnits(field, styleFormatting, styleResolver, contentIndex));
      for (const run of tail) units.push(...runToUnits(run, styleFormatting, styleResolver));
    } else if (content.type === 'inlineSdt') {
      boundaries = undefined;
      units.push(embedUnit('sdt', sdtPayload(content, styleFormatting, styleResolver), [], 2));
    } else if (
      content.type === 'insertion' ||
      content.type === 'deletion' ||
      content.type === 'moveFrom' ||
      content.type === 'moveTo'
    ) {
      boundaries = undefined;
      units.push(...trackedToUnits(content, styleFormatting, styleResolver));
    } else if (content.type === 'mathEquation') {
      boundaries = undefined;
      units.push(embedUnit('math', mathPayload(content)));
    } else if (content.type !== 'bookmarkStart' && content.type !== 'bookmarkEnd' && content.type !== 'rawXml') {
      boundaries = undefined;
    }
    if (content.type !== 'run') units.push(...units.splice(start).filter(keepInline));
    for (const unit of units.slice(start)) unit.comments = covering(context, plan);
    paragraphContentUnitCounts.set(content as object, units.length - start);
  }
  // The cache is kept only for what merging equal-formatted runs would lose.
  if (
    !boundaries?.some(
      (boundary) =>
        boundary.propertyChanges !== undefined ||
        boundary.noteMarks !== undefined ||
        boundary.breaks !== undefined ||
        boundary.text === ''
    )
  )
    boundaries = undefined;
  const attrs = paragraphAttrs(paragraph, styleResolver, units, boundaries, tableParagraphFormatting);
  return { units, ppr: paraAttrsToPpr(attrs) };
}

/** A page or column break (with the tracked change around it), or visible content. */
type FlowToken = {
  kind: 'pageBreak' | 'columnBreak' | 'visible' | 'mark' | 'reference';
  marker?: MarkDescriptor;
  /** The paragraph content item holding it. */
  item?: number;
};

function runTokens(run: Run, tokens: FlowToken[], marker?: MarkDescriptor): void {
  for (const content of run.content) {
    const flow =
      content.type === 'break' && (content.breakType === 'page' || content.breakType === 'column');
    if (flow)
      tokens.push({ kind: content.breakType === 'page' ? 'pageBreak' : 'columnBreak', marker });
    else if (content.type === 'commentReference') tokens.push({ kind: 'reference' });
    else if (content.type !== 'text' || content.text.length > 0) tokens.push({ kind: 'visible' });
  }
}

function inlineTokens(content: readonly { type: string }[], tokens: FlowToken[], marker?: MarkDescriptor): void {
  for (const item of content as readonly ParagraphContent[]) {
    if (item.type === 'run') runTokens(item, tokens, marker);
    else if (item.type === 'hyperlink') inlineTokens(item.structuredChildren ?? item.children, tokens, marker);
    else if (item.type === 'simpleField') {
      tokens.push({ kind: 'visible' });
      inlineTokens(item.structuredResult?.inline ?? item.content, tokens, marker);
    } else if (item.type === 'bookmarkStart' || item.type === 'bookmarkEnd') tokens.push({ kind: 'mark' });
    else if (item.type === 'complexField') inlineTokens([
      ...(item.structuredCode?.inline ?? item.fieldCode),
      ...(item.structuredResult?.inline ?? item.fieldResult),
    ], tokens, marker);
    else if (item.type === 'inlineSdt') inlineTokens(item.content, tokens, marker);
    else if (item.type === 'insertion' || item.type === 'deletion' || item.type === 'moveFrom' || item.type === 'moveTo') {
      inlineTokens(item.content, tokens, trackedMark(item.info,
        item.type === 'insertion' || item.type === 'moveTo' ? 'insertion' : 'deletion',
        item.type === 'moveFrom' || item.type === 'moveTo'));
    } else if (item.type === 'mathEquation') tokens.push({ kind: 'visible' });
  }
}

function exposedFlowBreaks(node: { type: string }, projectField = false): number {
  const item = node as ParagraphContent;
  const sum = (items: readonly { type: string }[]) => items.reduce((count, child) => count + exposedFlowBreaks(child), 0);
  if (item.type === 'run') return item.content.filter((entry) => flowBreakType(entry) !== null).length;
  if (item.type === 'hyperlink') return sum(item.structuredChildren ?? item.children);
  if (item.type === 'insertion' || item.type === 'deletion' || item.type === 'moveFrom' || item.type === 'moveTo') return sum(item.content);
  if (item.type === 'complexField' && projectField && !numericFieldInstruction(item.instruction)) return sum([
    ...(item.structuredCode?.inline ?? []), ...(item.structuredResult?.inline ?? []),
  ].filter((child) => child.type === 'hyperlink'));
  return 0;
}

/**
 * The break units before a paragraph's text (page ones flagged `leading`) and
 * after its pilcrow, as `paragraph_flow_breaks` in crates/docx-edit/src/seed.rs.
 */
function paragraphFlowBreaks(
  paragraph: Paragraph
): [Array<{ unit: EmbedUnit; item: number }>, EmbedUnit[], boolean[][]] {
  const tokens: FlowToken[] = [];
  paragraph.content.forEach((content, item) => {
    const start = tokens.length;
    inlineTokens([content], tokens);
    for (const token of tokens.slice(start)) token.item = item;
  });
  // Breaks lead the content that follows them: the text, or without text
  // the last bookmark. Without either, those up to the last column break lead.
  const lastBreak = tokens.map(({ kind }) => kind === 'pageBreak' || kind === 'columnBreak').lastIndexOf(true);
  const text = tokens.findIndex(({ kind }, index) => kind === 'visible' || (kind === 'reference' && index > lastBreak));
  const lastMark = tokens.map(({ kind }) => kind).lastIndexOf('mark');
  const content = text >= 0 ? text : lastMark;
  let split = content;
  if (content < 0) {
    split = 0;
    tokens.forEach((token, index) => {
      if (token.kind === 'columnBreak') split = index + 1;
    });
  }
  const lastText = tokens.map(({ kind }) => kind).lastIndexOf('visible');
  const exposed = paragraph.content.map((item) => exposedFlowBreaks(item, true));
  const counts = paragraph.content.map(() => 0);
  const inlineBreaks: boolean[][] = paragraph.content.map(() => []);
  for (const token of tokens) if (token.kind === 'pageBreak' || token.kind === 'columnBreak') counts[token.item!]++;
  const inline = tokens.map((token, index) => counts[token.item!] === exposed[token.item!]
    && text >= 0 && text < index && index < lastText);
  tokens.forEach((token, index) => {
    if (token.kind === 'pageBreak' || token.kind === 'columnBreak') inlineBreaks[token.item!].push(inline[index]);
  });
  const units = (from: number, to: number) =>
    tokens.slice(from, to).flatMap(({ kind, marker, item }, offset) =>
      inline[from + offset] || (kind !== 'pageBreak' && kind !== 'columnBreak')
        ? []
        : [
            {
              unit: embedUnit(
                kind,
                { ...(from + offset < content && { leading: true }), ...(from + offset >= split && { trailing: true }) },
                marker ? [marker] : []
              ),
              item: item!,
            },
          ]
    );
  return [units(0, split), units(split, tokens.length).map(({ unit }) => unit), inlineBreaks];
}

function revisionAttrs(info: TrackedChangeInfo): Attrs {
  return { revisionId: info.id, author: info.author, date: info.date ?? null };
}

function cellBorders(
  formatting: TableCellFormatting | undefined,
  tableBorders: TableBorders | undefined,
  firstRow: boolean,
  lastRow: boolean,
  firstColumn: boolean,
  lastColumn: boolean
): TableBorders | undefined {
  const inherited = tableBorders
    ? {
        top: firstRow ? tableBorders.top : tableBorders.insideH,
        bottom: lastRow ? tableBorders.bottom : tableBorders.insideH,
        left: firstColumn ? tableBorders.left : tableBorders.insideV,
        right: lastColumn ? tableBorders.right : tableBorders.insideV,
      }
    : undefined;
  return inherited || formatting?.borders
    ? { ...(inherited ?? {}), ...(formatting?.borders ?? {}) }
    : undefined;
}

function projectCell(
  cell: TableCell,
  options: {
    isHeader: boolean;
    rowspan: number;
    gridWidth?: number;
    firstRow: boolean;
    lastRow: boolean;
    firstColumn: boolean;
    lastColumn: boolean;
    tableBorders?: TableBorders;
    defaultMargins?: {
      top?: number;
      bottom?: number;
      left?: number;
      right?: number;
    };
    theme: Theme | null;
    tableBidi: boolean;
  }
): ProjectedCell {
  const formatting = cell.formatting;
  const backgroundColor = resolveColorToHex(formatting?.shading?.fill, options.theme);
  const width = formatting?.width?.value ?? options.gridWidth;
  const widthType = formatting?.width?.type ?? (options.gridWidth !== undefined ? 'pct' : null);
  const attrs: Attrs = {
    colspan: formatting?.gridSpan ?? 1,
    rowspan: options.rowspan,
    width: width ?? null,
    widthType,
    verticalAlign: formatting?.verticalAlign ?? null,
    backgroundColor: backgroundColor ?? null,
    borders:
      cellBorders(
        formatting,
        options.tableBorders,
        options.firstRow,
        options.lastRow,
        options.firstColumn,
        options.lastColumn
      ) ?? null,
    margins: formatting?.margins
      ? {
          top: formatting.margins.top?.value,
          bottom: formatting.margins.bottom?.value,
          left:
            formatting.margins.left?.value ??
            (options.tableBidi ? formatting.margins.end?.value : formatting.margins.start?.value),
          right:
            formatting.margins.right?.value ??
            (options.tableBidi ? formatting.margins.start?.value : formatting.margins.end?.value),
        }
      : (options.defaultMargins ?? null),
    textDirection: formatting?.textDirection ?? null,
    noWrap: formatting?.noWrap ?? false,
    _originalFormatting: formatting ?? null,
    _originalResolvedFill: backgroundColor ?? null,
  };
  if (cell.structuralChange) {
    const change = cell.structuralChange;
    const info = revisionAttrs(change.info);
    if (change.type === 'tableCellInsertion') attrs.cellMarker = { kind: 'ins', info };
    else if (change.type === 'tableCellDeletion') attrs.cellMarker = { kind: 'del', info };
    else if (change.type === 'tableCellMerge') {
      attrs.cellMarker = {
        kind: 'merge',
        info,
        vMerge: change.vMerge ?? 'cont',
        ...(change.vMergeOrig ? { vMergeOrig: change.vMergeOrig } : {}),
      };
    }
  }
  if (cell.propertyChanges?.length) attrs.tcPrChange = cell.propertyChanges;
  return {
    attrs: tableCellAttrsToTcPr(attrs, options.isHeader),
    content:
      cell.content.length > 0
        ? cell.content
        : [{ type: 'paragraph', content: [] } satisfies Paragraph],
  };
}

function projectRow(
  row: TableRow,
  table: Table,
  rowIndex: number,
  rowSpans: Map<string, RowSpanInfo>,
  tableBorders: TableBorders | undefined,
  defaultMargins: { top?: number; bottom?: number; left?: number; right?: number } | undefined,
  theme: Theme | null,
  tableStyle: Style | undefined,
  styleColumns: number
): ProjectedRow {
  const attrs: Attrs = {
    height: row.formatting?.height?.value ?? null,
    heightRule: row.formatting?.heightRule ?? null,
    isHeader: !!row.formatting?.header,
    _originalFormatting: row.formatting ?? null,
  };
  if (row.structuralChange) {
    if (row.structuralChange.type === 'tableRowInsertion') {
      attrs.trIns = revisionAttrs(row.structuralChange.info);
    } else if (row.structuralChange.type === 'tableRowDeletion') {
      attrs.trDel = revisionAttrs(row.structuralChange.info);
    }
  }
  if (row.propertyChanges?.length) attrs.trPrChange = row.propertyChanges;

  const columnWidths = table.columnWidths;
  const totalWidth = columnWidths?.reduce((sum, width) => sum + width, 0) ?? 0;
  const totalColumns =
    columnWidths?.length ??
    Math.max(
      0,
      ...table.rows.map((candidate) =>
        candidate.cells.reduce((sum, cell) => sum + (cell.formatting?.gridSpan ?? 1), 0)
      )
    );
  let column = 0;
  const cells: ProjectedCell[] = [];
  row.cells.forEach((cell) => {
    const colspan = cell.formatting?.gridSpan ?? 1;
    const startColumn = column;
    const rowSpan = rowSpans.get(`${rowIndex}-${startColumn}`);
    let gridWidth: number | undefined;
    if (columnWidths && totalWidth > 0) {
      let cellWidth = 0;
      for (
        let index = 0;
        index < colspan && startColumn + index < columnWidths.length;
        index += 1
      ) {
        cellWidth += columnWidths[startColumn + index];
      }
      gridWidth = Math.round((cellWidth / totalWidth) * 100);
    }
    column += colspan;
    if (rowSpan?.skip) return;
    cells.push({
      paragraphFormatting: tableCellParagraphFormatting(
        table,
        tableStyle,
        rowIndex,
        startColumn,
        column,
        styleColumns
      ),
      ...projectCell(cell, {
        isHeader: rowIndex === 0 && !!table.formatting?.look?.firstRow,
        rowspan: rowSpan?.rowSpan ?? 1,
        gridWidth,
        firstRow: rowIndex === 0,
        lastRow: rowIndex === table.rows.length - 1,
        firstColumn: startColumn === 0,
        lastColumn: column === totalColumns,
        tableBorders,
        defaultMargins,
        theme,
        tableBidi: Boolean(table.formatting?.bidi),
      }),
    });
  });
  if (cells.length === 0) {
    cells.push({
      paragraphFormatting: tableCellParagraphFormatting(
        table,
        tableStyle,
        rowIndex,
        0,
        totalColumns,
        styleColumns
      ),
      ...projectCell(
        {
          type: 'tableCell',
          formatting: totalColumns > 1 ? { gridSpan: totalColumns } : undefined,
          content: [{ type: 'paragraph', content: [] }],
        },
        {
          isHeader: rowIndex === 0 && !!table.formatting?.look?.firstRow,
          rowspan: 1,
          gridWidth: totalWidth > 0 ? 100 : undefined,
          firstRow: rowIndex === 0,
          lastRow: rowIndex === table.rows.length - 1,
          firstColumn: true,
          lastColumn: true,
          tableBorders,
          defaultMargins,
          theme,
          tableBidi: Boolean(table.formatting?.bidi),
        }
      ),
    });
  }
  return { attrs: tableRowAttrsToTrPr(attrs), cells };
}

function projectTable(
  table: Table,
  styleResolver: StyleResolver | null,
  theme: Theme | null,
  compatibilityMode: number
): ProjectedTable {
  const defaultStyle = styleResolver?.getDefaultTableStyle();
  const styleId = table.formatting?.styleId;
  const effectiveStyleId = styleId ?? defaultStyle?.styleId;
  const tableStyle = effectiveStyleId ? styleResolver?.getStyle(effectiveStyleId) : undefined;
  const borders =
    table.formatting?.borders ?? tableStyle?.tblPr?.borders ?? defaultStyle?.tblPr?.borders;
  const margins =
    table.formatting?.cellMargins ??
    tableStyle?.tblPr?.cellMargins ??
    defaultStyle?.tblPr?.cellMargins;
  const logicalLeft = table.formatting?.bidi ? margins?.end : margins?.start;
  const logicalRight = table.formatting?.bidi ? margins?.start : margins?.end;
  const defaultMargins = margins
    ? {
        top: margins.top?.value,
        bottom: margins.bottom?.value,
        left: margins.left?.value ?? logicalLeft?.value,
        right: margins.right?.value ?? logicalRight?.value,
      }
    : undefined;
  const basedOnStyleIds: string[] = [];
  const visited = new Set<string>();
  let inherited = tableStyle;
  while (inherited?.basedOn && basedOnStyleIds.length < 32) {
    if (visited.has(inherited.basedOn)) break;
    visited.add(inherited.basedOn);
    basedOnStyleIds.unshift(inherited.basedOn);
    inherited = styleResolver?.getStyle(inherited.basedOn);
  }
  const originalFormatting = {
    ...(table.formatting ?? {}),
    styleCascade: {
      ...(styleId ? { selectedStyleId: styleId } : {}),
      ...(defaultStyle?.styleId ? { defaultStyleId: defaultStyle.styleId } : {}),
      ...(basedOnStyleIds.length > 0 ? { basedOnStyleIds } : {}),
    },
  };
  const attrs: Attrs = {
    styleId: styleId ?? null,
    width: table.formatting?.width?.value ?? null,
    widthType: table.formatting?.width?.type ?? null,
    justification: table.formatting?.justification ?? null,
    columnWidths: table.columnWidths ?? null,
    tableLayout: table.formatting?.layout ?? null,
    floating: table.formatting?.floating ?? null,
    cellMargins: defaultMargins ?? null,
    look: table.formatting?.look ?? null,
    bidi: table.formatting?.bidi || null,
    // Omit the default so pre-existing yrs snapshots (no field) keep matching;
    // the bridge treats an absent mode as 12.
    compatibilityMode: compatibilityMode === 12 ? null : compatibilityMode,
    _originalFormatting: originalFormatting,
  };
  if (table.propertyChanges?.length) attrs.tblPrChange = table.propertyChanges;
  const rowSpans = calculateRowSpans(table);
  const styleColumns = tableColumnCount(table);
  return {
    attrs,
    rows: table.rows.map((row, rowIndex) =>
      projectRow(
        row,
        table,
        rowIndex,
        rowSpans,
        borders,
        defaultMargins,
        theme,
        tableStyle ?? defaultStyle,
        styleColumns
      )
    ),
  };
}

function blockSdtAttrs(properties: SdtProperties): Attrs {
  return blockSdtAttrsToPayload(sdtPropsToAttrs(properties));
}

const isCommentReference = (unit: InlineUnit): boolean =>
  unit.kind === 'embed' && unit.payload.modelKind === 'commentReference';

function addCommentCoverage(plan: StoryPlan): void {
  let offset = 0;
  for (const [index, unit] of plan.units.entries()) {
    if (unit.kind === 'embed' && unit.embedKind === 'bookmark') continue;
    const width = unit.kind === 'text' ? unit.text.length : 1;
    // A reference mark follows its own range's end, so another range holds
    // it only when that range goes on past it: coincident ends stay equal.
    const next = isCommentReference(unit)
      ? plan.units.slice(index + 1).find((candidate) => !isCommentReference(candidate)
        && !(candidate.kind === 'embed' && candidate.embedKind === 'bookmark'))
      : unit;
    for (const id of unit.comments ?? []) {
      if (!next?.comments?.includes(id)) continue;
      const intervals = plan.commentCoverage.get(id);
      const previous = intervals?.[intervals.length - 1];
      if (previous && previous[1] === offset) previous[1] = offset + width;
      else if (intervals) intervals.push([offset, offset + width]);
      else plan.commentCoverage.set(id, [[offset, offset + width]]);
    }
    offset += width;
  }
}

interface BlockCursor {
  paragraph: number;
  table: number;
  sdt: number;
}

/** Hands each block of a story the identity `visitStory` seeds it with. */
function takeBlockId(cursor: BlockCursor, storyId: string, block: BlockContent): string | null {
  if (isRawXml(block)) return null;
  if (block.type === 'paragraph') {
    const index = cursor.paragraph++;
    return block.paraId || `${storyId}:p${index}`;
  }
  if (block.type === 'table') return `${storyId}:t${cursor.table++}`;
  return blockSdtStoryId(storyId, cursor.sdt++);
}

function numericFieldInstruction(instruction: string): boolean {
  return /^\d+$/.test(instruction.trim());
}

/** How many story blocks a suppressed field's cached result duplicates. */
function cachedResultBlockCount(field: SimpleField | ComplexField): number | null {
  const blocks = field.structuredResult?.blocks ?? [];
  const last = blocks[blocks.length - 1];
  if (!last || last.type !== 'paragraph' || !last.content.every((node) => node.type === 'run' && node.content.every((entry) =>
    entry.type === 'fieldChar' && entry.charType === 'end' && entry.continuationId))) return null;
  return blocks.length;
}

/** Binds each suppressed field to the story blocks its cached result duplicates. */
function bindFieldResultBlocks(
  units: readonly InlineUnit[],
  storyId: string,
  blocks: readonly BlockContent[],
  owner: number,
  afterOwner: BlockCursor,
  tableIds: Map<number, string>
): void {
  for (const unit of units) {
    if (unit.kind !== 'embed' || unit.embedKind !== 'field') continue;
    const instruction = unit.payload.instruction;
    if (typeof instruction !== 'string' || !numericFieldInstruction(instruction)) continue;
    const data = unit.payload.fieldData;
    if (typeof data !== 'string') continue;
    let parsed: SimpleField | ComplexField;
    try {
      parsed = JSON.parse(data) as SimpleField | ComplexField;
    } catch {
      continue;
    }
    const count = cachedResultBlockCount(parsed);
    if (count === null || owner + 1 + count > blocks.length) continue;
    const cursor = { ...afterOwner };
    const ids: string[] = [];
    for (let offset = 0; offset < count; offset += 1) {
      const block = blocks[owner + 1 + offset]!;
      const id = takeBlockId(cursor, storyId, block);
      if (id === null) continue;
      if (block.type === 'table') tableIds.set(owner + 1 + offset, id);
      ids.push(id);
    }
    if (ids.length > 0) unit.payload.fieldResultBlocks = ids;
  }
}

function visitStory(
  context: LoweringContext,
  storyId: string,
  sourceBlocks: readonly BlockContent[],
  options: StoryOptions
): void {
  const plan: StoryPlan = {
    storyId,
    units: [],
    commentCoverage: new Map(),
  };
  const planIndex = context.plans.push(plan) - 1;
  /** Appends `unit` inside the comments open in this story. */
  const push = (unit: InlineUnit): void => {
    unit.comments = covering(context, planIndex);
    plan.units.push(unit);
  };
  const blocks =
    sourceBlocks.length > 0 ? [...sourceBlocks] : [{ type: 'paragraph', content: [] } as Paragraph];
  const cursor: BlockCursor = { paragraph: 0, table: 0, sdt: 0 };
  const resultTableIds = new Map<number, string>();
  let lastKind: 'paragraph' | 'table' | 'blockSdt' | null = null;

  for (const [blockIndex, block] of blocks.entries()) {
    if (isRawXml(block)) continue;
    const currentTable = cursor.table;
    const blockId = takeBlockId(cursor, storyId, block);
    if (blockId === null) continue;
    if (block.type === 'paragraph') {
      const [leadingBreaks, trailingBreaks] = paragraphFlowBreaks(block);
      // Comment markers ahead of a leading break take it into their range.
      let marked = 0;
      for (const { unit, item } of leadingBreaks) {
        for (const marker of block.content.slice(marked, item))
          markComment(context.openComments, marker, planIndex);
        marked = item;
        push(unit);
      }
      const paragraph = paragraphUnits(
        context,
        planIndex,
        block,
        options.extraRunFormatting,
        options.tableParagraphFormatting
      );
      bindFieldResultBlocks(
        paragraph.units,
        storyId,
        blocks,
        blockIndex,
        cursor,
        resultTableIds
      );
      plan.units.push(...paragraph.units);
      push(embedUnit('pilcrow', { ...paragraph.ppr, paraId: blockId }));
      for (const unit of trailingBreaks) push(unit);
      lastKind = 'paragraph';
      continue;
    }
    if (block.type === 'table') {
      const table = projectTable(
        block,
        context.styleResolver,
        context.theme,
        context.compatibilityMode
      );
      const rows = table.rows.map((row, rowIndex) => ({
        trPr: row.attrs,
        cells: row.cells.map((cell, cellIndex) => ({
          tcPr: cell.attrs,
          story: tableCellStoryId(storyId, currentTable, rowIndex, cellIndex),
        })),
      }));
      const resultTableId = resultTableIds.get(blockIndex);
      push(
        embedUnit('table', {
          tblPr: tableAttrsToTblPr(table.attrs),
          grid: tableAttrsToGrid(table.attrs),
          rows,
          ...(resultTableId === undefined ? {} : { blockId: resultTableId }),
        })
      );
      table.rows.forEach((row, rowIndex) => {
        row.cells.forEach((cell, cellIndex) => {
          visitStory(
            context,
            tableCellStoryId(storyId, currentTable, rowIndex, cellIndex),
            cell.content,
            {
              appendBodyTail: false,
              extraRunFormatting: cell.extraRunFormatting,
              tableParagraphFormatting: cell.paragraphFormatting,
            }
          );
        });
      });
      lastKind = 'table';
      continue;
    }
    const childStory = blockId;
    push(
      embedUnit('blockSdt', {
        ...blockSdtAttrs(block.properties),
        story: childStory,
      })
    );
    visitStory(context, childStory, block.content, {
      appendBodyTail: false,
      tableParagraphFormatting: options.tableParagraphFormatting,
    });
    lastKind = 'blockSdt';
  }

  if (options.appendBodyTail && (lastKind === 'table' || lastKind === 'blockSdt')) {
    push(
      embedUnit('pilcrow', {
        hangingIndent: false,
        paraId: `${storyId}:p${cursor.paragraph}`,
      })
    );
  }
  addCommentCoverage(plan);
}

function unitsToRawOps(units: readonly InlineUnit[]): YrsRawOp[] {
  const ops: YrsRawOp[] = [{ op: 'delete', index: 0, len: 1 }];
  let index = 0;
  let text = '';
  let attrs: YrsAttrs = {};
  let attrsKey = stableStringify(attrs);
  const flush = () => {
    if (!text) return;
    ops.push({ op: 'insert', index, text, attrs });
    index += text.length;
    text = '';
  };
  for (const unit of units) {
    if (unit.kind === 'text') {
      const key = stableStringify(unit.attrs);
      if (text && key !== attrsKey) flush();
      attrs = unit.attrs;
      attrsKey = key;
      text += unit.text;
    } else {
      flush();
      ops.push({
        op: 'insertEmbed',
        index,
        kind: unit.embedKind,
        payload: unit.payload,
        attrs: unit.attrs,
      });
      index += 1;
    }
  }
  flush();
  return ops;
}

function seedPlan(session: YrsSession, plan: StoryPlan): void {
  const bookmarks = takeBookmarks(plan.units);
  const ops = unitsToRawOps(plan.units);
  ops.push(...bookmarks);
  for (const [id, ranges] of plan.commentCoverage) {
    ops.push({
      op: 'setComment',
      id: String(id),
      ranges,
      author: '',
      date: '',
      body: null,
    });
  }
  session.applySeedRawOps(plan.storyId, ops);
}

/** Same offset-to-anchor lowering as seed.rs `take_bookmarks`. */
function takeBookmarks(units: InlineUnit[]): YrsRawOp[] {
  const found: Array<{ index: number; data: Record<string, unknown> }> = [];
  let at = 0;
  let pm = 0;
  let head = true;
  let spans: Array<{ raw: number; width: number; start: number; length: number }> = [];
  let leading: number[] = [];
  let fields: number[] = [];
  for (const unit of units) {
    const width = unit.kind === 'text' ? unit.text.length : 1;
    if (unit.kind === 'embed') {
      if (unit.embedKind === 'bookmark') {
        const data = { ...unit.payload };
        if (data.kind === 'fieldseparate' || data.kind === 'fieldend') {
          fields.push(found.length);
          data.order = found.length;
          if (Object.keys(unit.attrs).length) data.attributes = unit.attrs;
        }
        found.push({ index: at, data });
        continue;
      }
      if (unit.embedKind === 'pilcrow') {
        const bookmarks = unit.payload.bookmarks;
        delete unit.payload.bookmarks;
        if (Array.isArray(bookmarks)) for (const bookmark of bookmarks) {
          const data = { ...bookmark } as Record<string, unknown>;
          const offset = typeof data.offset === 'number' ? data.offset : 0;
          const after = typeof data.breaksAfter === 'number' ? data.breaksAfter : 0;
          delete data.offset;
          delete data.breaksAfter;
          // Between the field characters it follows and the next.
          if (typeof data.fieldsBefore === 'number' && fields.length) {
            const before = Math.min(data.fieldsBefore, fields.length);
            data.order = before === 0 ? fields[0]! - 0.5 : fields[before - 1]! + 0.5;
          }
          delete data.fieldsBefore;
          const span = spans.find(({ start, length }) => offset < start + length);
          if (offset === 0 && after > leading.length) throw new Error('Bookmark break position does not resolve');
          const index = offset === 0 && after > 0
            ? leading[leading.length - after]!
            : span ? span.raw + Math.min(Math.max(0, offset - span.start), span.width) : at;
          found.push({ index, data });
        }
        at++;
        pm = 0;
        spans = [];
        leading = [];
        fields = [];
        head = true;
        continue;
      }
      if (head && ['table', 'blockSdt', 'pageBreak', 'columnBreak'].includes(unit.embedKind)) {
        if (unit.embedKind.endsWith('Break')) leading.push(at);
        at++;
        continue;
      }
    }
    if (width > 0) {
      head = false;
      spans.push({ raw: at, width, start: pm, length: unit.pmSize });
      at += width;
      pm += unit.pmSize;
    }
  }
  let kept = 0;
  for (const unit of units) if (unit.kind !== 'embed' || unit.embedKind !== 'bookmark') units[kept++] = unit;
  units.length = kept;
  found.sort((a, b) => a.index - b.index || Number(a.data.id) - Number(b.data.id) || Number(a.data.kind === 'end') - Number(b.data.kind === 'end'));
  return found.map(({ index, data }) => ({ op: 'setBookmark', index, data }));
}

/**
 * Seeds every yrs-owned editable story directly from a parsed Document.
 *
 * Stories are `body`, `hf:{rId}`, `fn:{id}`, `en:{id}`, and recursively generated table
 * cell / block-SDT stories. The target session must not already contain any of
 * those story ids.
 *
 * @public
 */
export function documentToYrs(session: YrsSession, document: Document): void {
  noteYrsStoriesDirty(session, 'all');
  const context: LoweringContext = {
    styleResolver: document.package.styles ? createStyleResolver(document.package.styles) : null,
    theme: document.package.theme ?? null,
    plans: [],
    compatibilityMode: compatibilityModeFromDocument(document),
    openComments: [],
  };
  visitStory(context, 'body', document.package.document.content, {
    appendBodyTail: true,
  });
  for (const [rId, part] of document.package.headers ?? []) {
    visitStory(context, headerFooterStoryId(rId), part.content, {
      appendBodyTail: false,
    });
  }
  for (const [rId, part] of document.package.footers ?? []) {
    if (context.plans.some((plan) => plan.storyId === headerFooterStoryId(rId))) continue;
    visitStory(context, headerFooterStoryId(rId), part.content, {
      appendBodyTail: false,
    });
  }
  for (const note of document.package.footnotes ?? []) {
    visitStory(context, footnoteStoryId(note.id), note.content, {
      appendBodyTail: false,
    });
  }
  for (const note of document.package.endnotes ?? []) {
    visitStory(context, endnoteStoryId(note.id), note.content, {
      appendBodyTail: false,
    });
  }

  for (const plan of context.plans) session.createStory(plan.storyId, '', 'Normal', 'left');
  for (const plan of context.plans) seedPlan(session, plan);
  session.applySeedRawOps(
    'body',
    (document.package.document.comments ?? []).map((comment) => ({
      op: 'patchComment',
      id: String(comment.id),
      fields: {
        author: comment.author,
        date: comment.date ?? '',
        body: comment.content,
        parentId: comment.parentId == null ? null : String(comment.parentId),
        done: comment.done ?? false,
      },
    }))
  );
}
