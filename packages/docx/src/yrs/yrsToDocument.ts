/** yrs-to-OOXML save projection. */

/* eslint-disable max-lines -- the inverse mapping stays co-located with its save orchestrator */

import { projectYrsComments, commentSharedId } from './comments';
import { isRawXml } from '../types/content/rawXml';
import { pixelsToEmu } from '../utils/units';
import {
  applyContentControlValue,
  type ContentControlValue,
} from './contentControlValues';
import { sdtAttrsToProps } from '../types/sdtAttributes';
import {
  paragraphAttrsToFormatting,
  tableAttrsToFormatting,
  tableCellAttrsToFormatting,
  tableRowAttrsToFormatting,
  type ParagraphSaveAttrs,
  type TableCellSaveAttrs,
  type TableRowSaveAttrs,
  type TableSaveAttrs,
} from './saveFormatting';
import type {
  Document,
  BlockContent,
  Paragraph,
  ParagraphContent,
  Run,
  RunContent,
  HorizontalRuleContent,
  TextFormatting,
  Hyperlink,
  TrackedChangeInfo,
  Table,
  TableRow,
  TableCell,
  TableBorders,
  SimpleField,
  ComplexField,
  FieldType,
  MathEquation,
  Image,
  Shape,
  Chart,
  InlineSdt,
  SdtProperties,
  Comment,
  Footnote,
  Endnote,
} from '../types/document';
import type { YrsSession } from './index';

type Attrs = Record<string, unknown>;

interface YrsImageAttrs {
  src?: string;
  alt?: string;
  title?: string;
  width?: number;
  height?: number;
  rId?: string;
  wrapType?: Image['wrap']['type'];
  transform?: string;
  distTop?: number;
  distBottom?: number;
  distLeft?: number;
  distRight?: number;
  position?: {
    relativeHeight?: number;
    horizontal?: { relativeTo?: string; posOffset?: number; align?: string };
    vertical?: { relativeTo?: string; posOffset?: number; align?: string };
  };
  borderWidth?: number;
  borderColor?: string;
  borderStyle?: string;
  wrapText?: string;
  hlinkHref?: string;
  cropTop?: number;
  cropRight?: number;
  cropBottom?: number;
  cropLeft?: number;
  shapeType?: string;
  opacity?: number;
  layoutInCell?: boolean;
  allowOverlap?: boolean;
  effectExtentTop?: number;
  effectExtentRight?: number;
  effectExtentBottom?: number;
  effectExtentLeft?: number;
}

interface TextItem {
  kind: 'text';
  text: string;
  attributes: Attrs;
}

interface EmbedItem {
  kind: 'embed';
  embedKind: string;
  payload: Attrs;
  attributes: Attrs;
}

type InlineItem = TextItem | EmbedItem;

const PARAGRAPH_ATTR_DEFAULTS: Attrs = {
  paraId: null,
  textId: null,
  alignment: null,
  spaceBefore: null,
  spaceAfter: null,
  lineSpacing: null,
  lineSpacingRule: null,
  spacingExplicit: null,
  indentLeft: null,
  indentRight: null,
  indentFirstLine: null,
  hangingIndent: false,
  numPr: null,
  numPrFromStyle: null,
  listNumFmt: null,
  listIsBullet: null,
  listMarker: null,
  listMarkerHidden: null,
  listMarkerFontFamily: null,
  listMarkerFontSize: null,
  listMarkerSuffix: null,
  listLevelNumFmts: null,
  listAbstractNumId: null,
  listStartOverride: null,
  styleId: null,
  borders: null,
  shading: null,
  tabs: null,
  pageBreakBefore: null,
  renderedPageBreakBefore: null,
  keepNext: null,
  keepLines: null,
  widowControl: null,
  contextualSpacing: null,
  snapToGrid: null,
  autoSpaceDE: null,
  autoSpaceDN: null,
  defaultTextFormatting: null,
  sectionBreakType: null,
  bidi: null,
  outlineLevel: null,
  bookmarks: null,
  _originalFormatting: null,
  _originalRunBoundaries: null,
  _sectionProperties: null,
  pPrIns: null,
  pPrDel: null,
  pPrChange: null,
};

const TABLE_ATTR_DEFAULTS: Attrs = {
  styleId: null,
  width: null,
  widthType: null,
  justification: null,
  columnWidths: null,
  tableLayout: null,
  floating: null,
  cellMargins: null,
  look: null,
  bidi: null,
  _originalFormatting: null,
  tblPrChange: null,
};

const TABLE_ROW_ATTR_DEFAULTS: Attrs = {
  height: null,
  heightRule: null,
  isHeader: false,
  _originalFormatting: null,
  trIns: null,
  trDel: null,
  trPrChange: null,
};

const TABLE_CELL_ATTR_DEFAULTS: Attrs = {
  colspan: 1,
  rowspan: 1,
  colwidth: null,
  width: null,
  widthType: null,
  verticalAlign: null,
  backgroundColor: null,
  borders: null,
  margins: null,
  textDirection: null,
  noWrap: false,
  _originalFormatting: null,
  _originalResolvedFill: null,
  cellMarker: null,
  tcPrChange: null,
};

interface CommentBoundary {
  id: number;
  kind: 'start' | 'end';
  offset: number;
}

interface BookmarkBoundary extends CommentBoundary {
  name?: string;
  colFirst?: number;
  colLast?: number;
}

interface OriginalRunBoundary {
  text: string;
  noteMarks?: string[];
  /** Flow breaks the run held, at their offsets into `text`. */
  breaks?: Array<{ offset: number; type: 'page' | 'column' }>;
  marksKey?: string;
  formatting?: TextFormatting;
  propertyChanges?: Run['propertyChanges'];
}

interface TableCellPayload {
  tcPr?: Attrs;
  story?: string;
}

interface TableRowPayload {
  trPr?: Attrs;
  cells?: TableCellPayload[];
}

interface TablePayload extends Attrs {
  tblPr?: Attrs;
  grid?: unknown[];
  rows?: TableRowPayload[];
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

function asObject(value: unknown): Attrs | undefined {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
    ? (value as Attrs)
    : undefined;
}

function asFiniteNumber(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isFinite(value) ? value : undefined;
}

function asString(value: unknown): string | undefined {
  return typeof value === 'string' ? value : undefined;
}

function dropNulls(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(dropNulls);
  const object = asObject(value);
  if (!object) return value;
  const result: Attrs = {};
  for (const [key, entry] of Object.entries(object)) {
    if (entry === null || entry === undefined) continue;
    result[key] = dropNulls(entry);
  }
  return result;
}

function stableStringify(value: unknown): string {
  if (value === null || value === undefined) return 'null';
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(',')}]`;
  const object = asObject(value);
  if (object) {
    return `{${Object.keys(object)
      .sort()
      .map((key) => `${JSON.stringify(key)}:${stableStringify(object[key])}`)
      .join(',')}}`;
  }
  return JSON.stringify(value);
}

function formattingAttrs(attributes: Attrs): Attrs {
  const result = { ...attributes };
  delete result.hyperlink;
  delete result.ins;
  delete result.del;
  return result;
}

/** Converts attributes without manufacturing mark objects. */
function attrsToTextFormatting(attributes: Attrs): TextFormatting {
  const formatting: TextFormatting = {};

  if (attributes.bold) {
    formatting.bold = true;
    formatting.boldCs = true;
  }
  if (attributes.italic) {
    formatting.italic = true;
    formatting.italicCs = true;
  }

  const underline = asObject(attributes.underline);
  if (underline && underline.inheritedHyperlink !== true) {
    formatting.underline = {
      style: (asString(underline.style) || 'single') as NonNullable<
        TextFormatting['underline']
      >['style'],
      color: (asObject(underline.color) ?? null) as TextFormatting['color'],
    };
  }

  const strike = asObject(attributes.strike);
  if (strike) {
    if (strike.double) formatting.doubleStrike = true;
    else formatting.strike = true;
  }

  const textColor = asObject(attributes.textColor);
  if (textColor && textColor.inheritedHyperlink !== true) {
    formatting.color = {
      rgb: (asString(textColor.rgb) ?? null) as string | undefined,
      themeColor: (textColor.themeColor ?? null) as NonNullable<
        TextFormatting['color']
      >['themeColor'],
      themeTint: (asString(textColor.themeTint) ?? null) as string | undefined,
      themeShade: (asString(textColor.themeShade) ?? null) as string | undefined,
    };
  }

  if (typeof attributes.highlight === 'string') {
    formatting.highlight = attributes.highlight as TextFormatting['highlight'];
  }

  const fontSize = asObject(attributes.fontSize);
  if (fontSize) {
    const size = asFiniteNumber(fontSize.size);
    const sizeCs = asFiniteNumber(fontSize.sizeCs) ?? size;
    if (size !== undefined) formatting.fontSize = size;
    if (sizeCs !== undefined) formatting.fontSizeCs = sizeCs;
  }

  const language = asObject(attributes.language);
  if (language) {
    formatting.language = {
      latin: asString(language.latin),
      eastAsia: asString(language.eastAsia),
      bidi: asString(language.bidi),
    };
  }

  const fontFamily = asObject(attributes.fontFamily);
  if (fontFamily) {
    const ascii = asString(fontFamily.ascii);
    formatting.fontFamily = {
      ascii: (ascii ?? null) as string | undefined,
      hAnsi: (asString(fontFamily.hAnsi) ?? null) as string | undefined,
      eastAsia: asString(fontFamily.eastAsia),
      cs: asString(fontFamily.cs) || ascii,
      asciiTheme: (fontFamily.asciiTheme ?? null) as NonNullable<
        TextFormatting['fontFamily']
      >['asciiTheme'],
      hAnsiTheme: asString(fontFamily.hAnsiTheme),
      eastAsiaTheme: asString(fontFamily.eastAsiaTheme),
      csTheme: asString(fontFamily.csTheme),
    };
  }

  if (attributes.superscript) formatting.vertAlign = 'superscript';
  if (attributes.subscript) formatting.vertAlign = 'subscript';
  if (attributes.allCaps) formatting.allCaps = true;
  if (attributes.smallCaps) formatting.smallCaps = true;

  const spacing = asObject(attributes.characterSpacing);
  if (spacing) {
    const charSpacing = asFiniteNumber(spacing.spacing);
    const position = asFiniteNumber(spacing.position);
    const scale = asFiniteNumber(spacing.scale);
    const kerning = asFiniteNumber(spacing.kerning);
    if (charSpacing !== undefined) formatting.spacing = charSpacing;
    if (position !== undefined) formatting.position = position;
    if (scale !== undefined) formatting.scale = scale;
    if (kerning !== undefined) formatting.kerning = kerning;
  }

  if (attributes.emboss) formatting.emboss = true;
  if (attributes.imprint) formatting.imprint = true;
  if (attributes.textShadow) formatting.shadow = true;

  const emphasis = asObject(attributes.emphasisMark);
  if (emphasis) {
    formatting.emphasisMark = (asString(emphasis.type) || 'dot') as NonNullable<
      TextFormatting['emphasisMark']
    >;
  }
  if (attributes.textOutline) formatting.outline = true;
  if (attributes.hidden) formatting.hidden = true;
  if (attributes.rtl) formatting.rtl = true;

  const effect = asObject(attributes.textEffect);
  if (effect) {
    formatting.effect = (asString(effect.effect) || 'blinkBackground') as NonNullable<
      TextFormatting['effect']
    >;
  }

  const modern = asObject(attributes.modernTextEffects);
  if (modern?.effects) formatting.modernEffects = modern.effects as TextFormatting['modernEffects'];

  const runStyle = asObject(attributes.runStyle);
  const styleId = asString(runStyle?.styleId);
  if (styleId) formatting.styleId = styleId;

  return formatting;
}

function runContentForText(text: string, formatting: TextFormatting): RunContent[] {
  const content: RunContent[] = [];
  let plainText = '';
  const flushText = (): void => {
    if (!plainText) return;
    content.push({ type: 'text', text: plainText });
    plainText = '';
  };

  const fonts = formatting.fontFamily
    ? [
        formatting.fontFamily.ascii,
        formatting.fontFamily.hAnsi,
        formatting.fontFamily.eastAsia,
        formatting.fontFamily.cs,
      ].filter((font): font is string => Boolean(font))
    : [];
  const symbolFont = fonts.length > 0 && fonts.every((font) => font === fonts[0]) ? fonts[0] : null;

  for (const char of text) {
    const codePoint = char.codePointAt(0) ?? 0;
    if (char === '\u00ad') {
      flushText();
      content.push({ type: 'softHyphen' });
    } else if (char === '\u2011') {
      flushText();
      content.push({ type: 'noBreakHyphen' });
    } else if (symbolFont && codePoint >= 0xf000 && codePoint <= 0xf8ff) {
      flushText();
      content.push({
        type: 'symbol',
        font: symbolFont,
        char: codePoint.toString(16).toUpperCase().padStart(4, '0'),
      });
    } else {
      plainText += char;
    }
  }
  flushText();
  return content;
}

function createTextRun(text: string, attributes: Attrs): Run {
  const formatting = attrsToTextFormatting(formattingAttrs(attributes));
  return {
    type: 'run',
    formatting: Object.keys(formatting).length > 0 ? formatting : undefined,
    content: runContentForText(text, formatting),
  };
}

function appendTextRun(target: Run, source: Run): void {
  for (const item of source.content) {
    const previous = target.content[target.content.length - 1];
    if (previous?.type === 'text' && item.type === 'text') previous.text += item.text;
    else target.content.push(item);
  }
}

function fnv53(value: string): number {
  let hash = 0xcbf29ce484222325n;
  for (const byte of new TextEncoder().encode(value)) {
    hash = (hash ^ BigInt(byte)) * 0x100000001b3n;
    hash &= 0xffffffffffffffffn;
  }
  return Number(hash & ((1n << 53n) - 1n));
}

function revisionId(value: unknown): number {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value === 'string') {
    const parsed = Number(value);
    if (Number.isFinite(parsed)) return parsed;
    return fnv53(value);
  }
  return 0;
}

function trackedInfo(raw: unknown, _pmShape = false): TrackedChangeInfo | null {
  const value = asObject(raw);
  if (!value) return null;
  const author = asString(value.author) || 'Unknown';
  const date = asString(value.date);
  return {
    id: revisionId(value.revisionId ?? value.id),
    author,
    ...(date ? { date } : {}),
  };
}

function createHyperlink(attributes: Attrs): Hyperlink | null {
  const link = asObject(attributes.hyperlink);
  if (!link) return null;
  const href = asString(link.href) || '';
  if (href.startsWith('#')) {
    return {
      type: 'hyperlink',
      anchor: href.slice(1),
      tooltip: asString(link.tooltip) || undefined,
      children: [],
    };
  }
  return {
    type: 'hyperlink',
    href,
    tooltip: asString(link.tooltip) || undefined,
    rId: asString(link.rId) || undefined,
    children: [],
  };
}

function hyperlinkKey(attributes: Attrs): string | null {
  const link = asObject(attributes.hyperlink);
  return link ? asString(link.href) || '' : null;
}

function fieldFromPayload(payload: Attrs, attributes: Attrs): SimpleField | ComplexField {
  const fieldData = asString(payload.fieldData);
  if (fieldData && fieldData.length <= 2_000_000) {
    try {
      const stored = JSON.parse(fieldData) as SimpleField | ComplexField;
      const children = stored.type === 'simpleField' ? stored.content : stored.fieldResult;
      if (
        (stored.type === 'simpleField' || stored.type === 'complexField') &&
        typeof stored.instruction === 'string' &&
        Array.isArray(children)
      ) {
        const displayMode = asString(payload.displayMode) as 'result' | 'code' | undefined;
        if (stored.fieldTree && displayMode) stored.fieldTree.displayMode = displayMode;
        return stored;
      }
    } catch {
      // Malformed editor cache: rebuild the same minimal inert field as the save path.
    }
  }

  const formatting = attrsToTextFormatting(formattingAttrs(attributes));
  const displayRun: Run = {
    type: 'run',
    content: [{ type: 'text', text: asString(payload.displayText) || '' }],
    ...(Object.keys(formatting).length > 0 ? { formatting } : {}),
  };
  const instruction = asString(payload.instruction) || '';
  const fieldType = (asString(payload.fieldType) || 'UNKNOWN') as FieldType;
  const fldLock = payload.fldLock === true || undefined;
  const dirty = payload.dirty === true || undefined;
  const displayMode = (asString(payload.displayMode) || 'result') as 'result' | 'code';
  if (payload.fieldKind === 'complex') {
    return {
      type: 'complexField',
      instruction,
      fieldType,
      fieldCode: [],
      fieldResult: [displayRun],
      fldLock,
      dirty,
      fieldTree: { version: 1, displayMode },
    };
  }
  return {
    type: 'simpleField',
    instruction,
    fieldType,
    content: [displayRun],
    fldLock,
    dirty,
    fieldTree: { version: 1, displayMode },
  };
}

function mathFromPayload(payload: Attrs): MathEquation {
  return {
    type: 'mathEquation',
    display: (asString(payload.display) as 'inline' | 'block') || 'inline',
    ommlXml: asString(payload.ommlXml) || '',
    plainText: asString(payload.plainText) || undefined,
  };
}

function horizontalRuleRun(payload: Attrs, attributes: Attrs): Run {
  const value = asObject(payload.rule);
  const height = asFiniteNumber(value?.height);
  if (
    !value || height === undefined ||
    (value.width != null && asFiniteNumber(value.width) === undefined) ||
    (value.widthPercent != null && asFiniteNumber(value.widthPercent) === undefined) ||
    typeof value.alignment !== 'string' || typeof value.noShade !== 'boolean' ||
    typeof value.color !== 'string' || typeof value.xml !== 'string'
  ) throw new Error('Malformed horizontalRule embed payload');
  const rule: HorizontalRuleContent['rule'] = {
    width: asFiniteNumber(value.width) ?? null,
    widthPercent: asFiniteNumber(value.widthPercent) ?? null,
    height,
    alignment: value.alignment,
    noShade: value.noShade,
    color: value.color,
    xml: value.xml,
  };
  const formatting = attrsToTextFormatting(formattingAttrs(attributes));
  return {
    type: 'run',
    content: [{ type: 'horizontalRule', rule }],
    ...(Object.keys(formatting).length > 0 ? { formatting } : {}),
  };
}

/** The authored XML an unedited picture or shape replays; editing the embed drops it. */
function sourceXml(payload: Attrs): { sourceXml?: string } {
  const xml = asString(payload.sourceXml);
  return xml ? { sourceXml: xml } : {};
}

function imageRunFromPayload(payload: Attrs): Run {
  const attrs = payload as YrsImageAttrs & Attrs;
  const wrap: Image['wrap'] = {
    type: (asString(attrs.wrapType) || 'inline') as Image['wrap']['type'],
  };
  if (attrs.distTop !== undefined) wrap.distT = pixelsToEmu(Number(attrs.distTop));
  if (attrs.distBottom !== undefined) wrap.distB = pixelsToEmu(Number(attrs.distBottom));
  if (attrs.distLeft !== undefined) wrap.distL = pixelsToEmu(Number(attrs.distLeft));
  if (attrs.distRight !== undefined) wrap.distR = pixelsToEmu(Number(attrs.distRight));
  if (attrs.wrapText) wrap.wrapText = attrs.wrapText as Image['wrap']['wrapText'];

  const image: Image = {
    type: 'image',
    rId: asString(attrs.rId) || '',
    src: asString(attrs.src) || '',
    alt: asString(attrs.alt) || undefined,
    title: asString(attrs.title) || undefined,
    shapeType: asString(attrs.shapeType) || undefined,
    size: {
      width: pixelsToEmu(Number(attrs.width) || 0),
      height: pixelsToEmu(Number(attrs.height) || 0),
    },
    wrap,
  };

  if (attrs.transform) {
    const transform: NonNullable<Image['transform']> = {};
    const rotation = attrs.transform.match(/rotate\(([-\d.]+)deg\)/)?.[1];
    if (rotation) transform.rotation = Number.parseFloat(rotation);
    if (attrs.transform.includes('scaleX(-1)')) transform.flipH = true;
    if (attrs.transform.includes('scaleY(-1)')) transform.flipV = true;
    if (transform.rotation || transform.flipH || transform.flipV) image.transform = transform;
  }

  if (attrs.position?.horizontal && attrs.position.vertical) {
    image.position = {
      relativeHeight: attrs.position.relativeHeight,
      horizontal: {
        relativeTo: (attrs.position.horizontal.relativeTo || 'column') as NonNullable<
          Image['position']
        >['horizontal']['relativeTo'],
        alignment: attrs.position.horizontal.align as NonNullable<
          Image['position']
        >['horizontal']['alignment'],
        posOffset: attrs.position.horizontal.posOffset,
      },
      vertical: {
        relativeTo: (attrs.position.vertical.relativeTo || 'paragraph') as NonNullable<
          Image['position']
        >['vertical']['relativeTo'],
        alignment: attrs.position.vertical.align as NonNullable<
          Image['position']
        >['vertical']['alignment'],
        posOffset: attrs.position.vertical.posOffset,
      },
    };
  }

  if (attrs.borderWidth && attrs.borderWidth > 0) {
    const styles: Record<string, NonNullable<Image['outline']>['style']> = {
      solid: 'solid',
      dotted: 'dot',
      dashed: 'dash',
      double: 'solid',
      groove: 'solid',
      ridge: 'solid',
      inset: 'solid',
      outset: 'solid',
    };
    image.outline = {
      width: pixelsToEmu(attrs.borderWidth),
      color: attrs.borderColor ? { rgb: attrs.borderColor.replace('#', '') } : undefined,
      style: (attrs.borderStyle && styles[attrs.borderStyle]) || 'solid',
    };
  }
  if (attrs.hlinkHref) image.hlinkHref = attrs.hlinkHref;

  const crop: NonNullable<Image['crop']> = {};
  if (attrs.cropTop != null) crop.top = attrs.cropTop;
  if (attrs.cropRight != null) crop.right = attrs.cropRight;
  if (attrs.cropBottom != null) crop.bottom = attrs.cropBottom;
  if (attrs.cropLeft != null) crop.left = attrs.cropLeft;
  if (Object.keys(crop).length > 0) image.crop = crop;
  if (attrs.opacity != null && attrs.opacity < 1) image.opacity = attrs.opacity;
  if (attrs.layoutInCell != null) image.layoutInCell = attrs.layoutInCell;
  if (attrs.allowOverlap != null) image.allowOverlap = attrs.allowOverlap;

  const padding: NonNullable<Image['padding']> = {};
  if (attrs.effectExtentTop) padding.top = pixelsToEmu(attrs.effectExtentTop);
  if (attrs.effectExtentBottom) padding.bottom = pixelsToEmu(attrs.effectExtentBottom);
  if (attrs.effectExtentLeft) padding.left = pixelsToEmu(attrs.effectExtentLeft);
  if (attrs.effectExtentRight) padding.right = pixelsToEmu(attrs.effectExtentRight);
  if (Object.keys(padding).length > 0) image.padding = padding;

  return { type: 'run', content: [{ type: 'drawing', image, ...sourceXml(payload) }] };
}

/** The seeded shape, carrying the text body and everything else no payload field describes. */
function storedShape(value: unknown): Shape | undefined {
  const json = asString(value);
  if (!json || json.length > 2_000_000) return undefined;
  try {
    const parsed = JSON.parse(json) as Shape;
    if (parsed?.type !== 'shape' || typeof parsed.shapeType !== 'string') return undefined;
    if (!asObject(parsed.size)) parsed.size = { width: 0, height: 0 };
    return parsed;
  } catch {
    return undefined;
  }
}

function chartRunFromPayload(payload: Attrs): Run | null {
  const json = asString(payload.chartJson);
  if (!json) return null;
  try {
    const chart = JSON.parse(json) as Chart;
    if (chart?.type !== 'chart' || typeof chart.chartType !== 'string') return null;
    return { type: 'run', content: [{ type: 'chart', chart }] };
  } catch {
    return null;
  }
}

function opaqueDrawingRun(payload: Attrs): Run {
  return {
    type: 'run',
    content: [
      {
        type: 'opaqueDrawing',
        kind: asString(payload.kind) ?? '',
        xml: asString(payload.xml) ?? '',
      },
    ],
  };
}

function shapeRunFromPayload(payload: Attrs): Run {
  const shape: Shape = storedShape(payload.shapeJson) ?? {
    type: 'shape',
    shapeType: 'rect',
    size: { width: 0, height: 0 },
  };
  const shapeType = asString(payload.shapeType);
  if (shapeType) shape.shapeType = shapeType as Shape['shapeType'];
  const shapeId = asString(payload.shapeId);
  if (shapeId) shape.id = shapeId;
  if (payload.width) shape.size = { ...shape.size, width: pixelsToEmu(Number(payload.width)) };
  if (payload.height) shape.size = { ...shape.size, height: pixelsToEmu(Number(payload.height)) };
  if (Array.isArray(payload.geometryPath) && payload.geometryPath.length > 0) {
    shape.geometryPath = payload.geometryPath as NonNullable<Shape['geometryPath']>;
  }

  if (payload.fillType === 'gradient' && typeof payload.gradientStops === 'string') {
    try {
      const stops = JSON.parse(payload.gradientStops) as Array<{
        position: number;
        color: string;
      }>;
      shape.fill = {
        type: 'gradient',
        gradient: {
          type: (asString(payload.gradientType) || 'linear') as NonNullable<
            NonNullable<Shape['fill']>['gradient']
          >['type'],
          angle: asFiniteNumber(payload.gradientAngle) || undefined,
          stops: stops.map((stop) => ({
            position: stop.position,
            color: { rgb: stop.color.replace('#', '') },
          })),
        },
      };
    } catch {
      shape.fill = {
        type: 'solid',
        color: { rgb: (asString(payload.fillColor) || '000000').replace('#', '') },
      };
    }
  } else if (typeof payload.fillColor === 'string') {
    shape.fill = {
      type: (asString(payload.fillType) || 'solid') as 'solid' | 'none',
      color: { rgb: payload.fillColor.replace('#', '') },
    };
  } else if (payload.fillType === 'none') {
    shape.fill = { type: 'none' };
  }

  if (typeof payload.outlineWidth === 'number' && payload.outlineWidth > 0) {
    const styles: Record<string, NonNullable<Shape['outline']>['style']> = {
      solid: 'solid',
      dotted: 'dot',
      dashed: 'dash',
    };
    const outlineStyle = asString(payload.outlineStyle);
    shape.outline = {
      width: pixelsToEmu(payload.outlineWidth),
      color:
        typeof payload.outlineColor === 'string'
          ? { rgb: payload.outlineColor.replace('#', '') }
          : undefined,
      style: (outlineStyle && styles[outlineStyle]) || 'solid',
    };
  }

  const transform: NonNullable<Shape['transform']> = {};
  if (typeof payload.rotation === 'number') transform.rotation = payload.rotation;
  else if (typeof payload.transform === 'string') {
    const rotation = payload.transform.match(/rotate\(([-\d.]+)deg\)/)?.[1];
    if (rotation) transform.rotation = Number.parseFloat(rotation);
  }
  if (payload.flipH || String(payload.transform || '').includes('scaleX(-1)'))
    transform.flipH = true;
  if (payload.flipV || String(payload.transform || '').includes('scaleY(-1)'))
    transform.flipV = true;
  if (transform.rotation || transform.flipH || transform.flipV) shape.transform = transform;

  return { type: 'run', content: [{ type: 'shape', shape, ...sourceXml(payload) }] };
}

function inlineSdtFromPayload(payload: Attrs): InlineSdt {
  let properties: SdtProperties = sdtAttrsToProps(payload);
  const propertiesJson = asString(payload.propertiesJson);
  if (propertiesJson && propertiesJson.length <= 1_000_000) {
    try {
      const parsed = JSON.parse(propertiesJson) as SdtProperties;
      if (parsed && typeof parsed === 'object' && typeof parsed.sdtType === 'string') {
        properties = parsed;
      }
    } catch {
      // Keep the individually projected properties.
    }
  }
  const items: InlineItem[] = [];
  for (const raw of Array.isArray(payload.content) ? payload.content : []) {
    const entry = asObject(raw);
    const kind = asString(entry?.kind);
    if (!entry || !kind) continue;
    const attributes = asObject(entry.attrs) ?? {};
    if (kind === 'text') {
      const text = asString(entry.text);
      if (text !== undefined) items.push({ kind: 'text', text, attributes });
      continue;
    }
    items.push({
      kind: 'embed',
      embedKind: kind,
      payload: asObject(entry.payload) ?? {},
      attributes,
    });
  }
  let content = inlineSdtContent(buildParagraphContent(items));
  const authoredValue = contentControlValue(payload.value);
  if (authoredValue) {
    try {
      const applied = applyContentControlValue(properties, authoredValue);
      properties = applied.properties;
      const display = applied.content[0];
      content = inlineSdtContent(display?.type === 'paragraph' ? display.content : []);
    } catch {
      // Preserve the embedded control when a malformed authored value cannot
      // be applied to its captured OOXML properties.
    }
  }
  return { type: 'inlineSdt', properties, content };
}

function inlineSdtContent(content: ParagraphContent[]): InlineSdt['content'] {
  return content.filter(
    (child): child is InlineSdt['content'][number] =>
      child.type === 'run' ||
      child.type === 'hyperlink' ||
      child.type === 'simpleField' ||
      child.type === 'complexField' ||
      child.type === 'inlineSdt' ||
      child.type === 'mathEquation'
  );
}

function contentControlValue(value: unknown): ContentControlValue | null {
  const authored = asObject(value);
  if (authored?.kind === 'checkbox' && typeof authored.checked === 'boolean') {
    return { kind: 'checkbox', checked: authored.checked };
  }
  if (authored?.kind === 'dropdown' && typeof authored.value === 'string') {
    return { kind: 'dropdown', value: authored.value };
  }
  if (authored?.kind === 'date' && typeof authored.date === 'string') {
    return { kind: 'date', date: authored.date };
  }
  return null;
}

function commentReferenceFromPayload(payload: Attrs): Run | null {
  if (payload.modelKind !== 'commentReference') return null;
  const id = asFiniteNumber(payload.commentId);
  return {
    type: 'run',
    content: [{ type: 'commentReference', ...(id !== undefined ? { id } : {}) }],
  };
}

function ordinaryContentForItem(item: InlineItem): ParagraphContent | null {
  if (item.kind === 'text') return createTextRun(item.text, item.attributes);
  switch (item.embedKind) {
    case 'break':
      return { type: 'run', content: [{ type: 'break', breakType: 'textWrapping' }] };
    case 'flowBreak':
      return unitBreakRun(item.payload.breakType as FlowBreak);
    case 'tab':
      return { type: 'run', content: [{ type: 'tab' }] };
    case 'image':
      return imageRunFromPayload(item.payload);
    case 'horizontalRule':
      return horizontalRuleRun(item.payload, item.attributes);
    case 'shape':
      return shapeRunFromPayload(item.payload);
    case 'chart':
      return chartRunFromPayload(item.payload);
    case 'opaqueDrawing':
      return opaqueDrawingRun(item.payload);
    case 'field':
      return (
        commentReferenceFromPayload(item.payload) ?? fieldFromPayload(item.payload, item.attributes)
      );
    case 'math':
      return mathFromPayload(item.payload);
    case 'sdt':
      return inlineSdtFromPayload(item.payload);
    case 'noteRef': {
      const footnote = item.payload.footnoteRefId;
      const endnote = item.payload.endnoteRefId;
      return {
        type: 'run',
        content: [
          footnote !== undefined
            ? { type: 'footnoteRef', id: revisionId(footnote) }
            : { type: 'endnoteRef', id: revisionId(endnote) },
        ],
      };
    }
    default:
      return null;
  }
}

function trackedContentForItem(item: InlineItem, info: TrackedChangeInfo): ParagraphContent {
  let run: Run;
  if (item.kind === 'embed' && item.embedKind === 'image') run = imageRunFromPayload(item.payload);
  else if (item.kind === 'embed' && item.embedKind === 'horizontalRule')
    run = horizontalRuleRun(item.payload, item.attributes);
  else if (item.kind === 'embed' && item.embedKind === 'shape')
    run = shapeRunFromPayload(item.payload);
  else if (item.kind === 'embed' && item.embedKind === 'chart')
    run = chartRunFromPayload(item.payload) ?? { type: 'run', content: [] };
  else if (item.kind === 'embed' && item.embedKind === 'opaqueDrawing')
    run = opaqueDrawingRun(item.payload);
  else if (item.kind === 'embed' && item.embedKind === 'flowBreak')
    run = unitBreakRun(item.payload.breakType as FlowBreak);
  else if (item.kind === 'text') {
    const formatting = attrsToTextFormatting(formattingAttrs(item.attributes));
    run = {
      type: 'run',
      content: [{ type: 'text', text: item.text }],
      ...(Object.keys(formatting).length > 0 ? { formatting } : {}),
    };
  } else run = { type: 'run', content: [] };

  const raw = asObject(item.attributes.ins) ?? asObject(item.attributes.del);
  const isMovePair = raw?.isMovePair === true;
  if (item.attributes.ins) {
    const field =
      item.kind === 'embed' && item.embedKind === 'field' && item.payload.modelKind !== 'commentReference'
        ? fieldFromPayload(item.payload, item.attributes)
        : undefined;
    const content = field?.type === 'simpleField' ? [field] : [run];
    return isMovePair ? { type: 'moveTo', info, content } : { type: 'insertion', info, content };
  }
  return isMovePair
    ? { type: 'moveFrom', info, content: [run] }
    : { type: 'deletion', info, content: [run] };
}

function addToHyperlink(hyperlink: Hyperlink, item: InlineItem): void {
  // Once a field or equation joins the link, its full child list is structuredChildren.
  const add = (child: Run | SimpleField | ComplexField | MathEquation): void => {
    if (child.type === 'run') hyperlink.children.push(child);
    else hyperlink.structuredChildren ??= [...hyperlink.children];
    hyperlink.structuredChildren?.push(child);
  };
  if (item.kind === 'text') {
    add(createTextRun(item.text, item.attributes));
    return;
  }
  if (item.embedKind === 'break') {
    add({ type: 'run', content: [{ type: 'break', breakType: 'textWrapping' }] });
  } else if (item.embedKind === 'tab') {
    add({ type: 'run', content: [{ type: 'tab' }] });
  } else if (item.embedKind === 'horizontalRule') {
    add(horizontalRuleRun(item.payload, item.attributes));
  } else if (item.embedKind === 'field') {
    add(commentReferenceFromPayload(item.payload) ?? fieldFromPayload(item.payload, item.attributes));
  } else if (item.embedKind === 'math') {
    add(mathFromPayload(item.payload));
  }
}

function projectionSignature(items: InlineItem[]): string {
  const normalized: InlineItem[] = [];
  for (const source of items) {
    const attributes = { ...source.attributes };
    delete attributes.fieldResult;
    const item: InlineItem = source.kind === 'embed' && source.embedKind === 'tab'
      ? { kind: 'text', text: '\t', attributes }
      : { ...source, attributes };
    const previous = normalized.at(-1);
    if (item.kind === 'text' && previous?.kind === 'text' && stableStringify(previous.attributes) === stableStringify(attributes)) {
      previous.text += item.text;
    } else normalized.push(item);
  }
  return stableStringify(normalized);
}

function restoreProjectedFieldResults(items: InlineItem[]): InlineItem[] {
  const owners: { id: number; owner: EmbedItem; position: number }[] = [];
  for (const [position, item] of items.entries()) {
    if (item.kind !== 'embed' || item.embedKind !== 'field') continue;
    const id = asFiniteNumber(asObject(item.payload.resultProjection)?.id);
    if (id !== undefined) owners.push({ id, owner: item, position });
  }
  if (owners.length === 0) return items;
  const groups = new Map<EmbedItem, Map<number, InlineItem[]>>();
  const remaining = items.filter((item, position) => {
    const marker = asObject(item.attributes.fieldResult);
    const id = asFiniteNumber(marker?.id);
    const index = asFiniteNumber(marker?.index);
    const owner = owners.find((entry) => entry.id === id && entry.position > position)?.owner;
    if (index === undefined || !owner) return true;
    const children = groups.get(owner) ?? new Map<number, InlineItem[]>();
    const group = children.get(index) ?? [];
    const attributes = { ...item.attributes };
    delete attributes.fieldResult;
    group.push({ ...item, attributes });
    children.set(index, group);
    groups.set(owner, children);
    return false;
  });
  for (const { owner } of owners) {
    const stored = fieldFromPayload(owner.payload, owner.attributes);
    if (stored.type !== 'complexField') continue;
    const projection = asObject(owner.payload.resultProjection);
    const originals = Array.isArray(projection?.children) ? projection.children : [];
    const replacements = new Map<number, ReturnType<typeof inlineSdtContent>>();
    for (const raw of originals) {
      const child = asObject(raw);
      const index = asFiniteNumber(child?.index);
      if (index === undefined || !Array.isArray(child?.items)) continue;
      const current = groups.get(owner)?.get(index) ?? [];
      if (projectionSignature(current) === projectionSignature(child.items as InlineItem[])) continue;
      const rebuilt = inlineSdtContent(buildParagraphContent(current));
      const original = index < 0 ? stored.structuredCode?.inline?.[-index - 1] : stored.structuredResult?.inline?.[index];
      if (original?.type === 'hyperlink' && rebuilt.length === 1 && rebuilt[0]?.type === 'hyperlink') {
        rebuilt[0] = { ...original, ...rebuilt[0], structuredChildren: rebuilt[0].structuredChildren };
      }
      replacements.set(index, rebuilt);
    }
    if (replacements.size === 0) continue;
    const inline = (stored.structuredResult?.inline ?? []).flatMap((child, index) => replacements.get(index) ?? [child]);
    const code = stored.structuredCode?.inline?.flatMap((child, index) => replacements.get(-index - 1) ?? [child]);
    stored.structuredResult = { ...stored.structuredResult, inline };
    if (code) stored.structuredCode = { ...stored.structuredCode, inline: code };
    if (stored.fieldTree) {
      stored.fieldTree.result = { ...stored.fieldTree.result, inline };
      if (code) stored.fieldTree.code = { ...stored.fieldTree.code, inline: code };
    }
    stored.fieldResult = inline.flatMap((child) => child.type === 'run' ? [child]
      : child.type === 'hyperlink' ? child.children.filter((entry): entry is Run => entry.type === 'run') : []);
    owner.payload = { ...owner.payload, fieldData: JSON.stringify(stored) };
  }
  return remaining;
}

function buildParagraphContent(items: InlineItem[]): ParagraphContent[] {
  items = restoreProjectedFieldResults(items);
  const content: ParagraphContent[] = [];
  let currentRun: Run | null = null;
  let currentFormattingKey: string | null = null;
  let currentHyperlink: Hyperlink | null = null;

  const flushRun = (): void => {
    if (currentRun) content.push(currentRun);
    currentRun = null;
    currentFormattingKey = null;
  };
  const flushHyperlink = (): void => {
    if (currentHyperlink) content.push(currentHyperlink);
    currentHyperlink = null;
  };

  for (const item of items) {
    // A note reference is handled before tracked/link marks.
    if (item.kind === 'embed' && item.embedKind === 'noteRef') {
      flushRun();
      flushHyperlink();
      const note = ordinaryContentForItem(item);
      if (note) content.push(note);
      continue;
    }

    const revision = trackedInfo(item.attributes.ins ?? item.attributes.del);
    if (revision) {
      flushRun();
      flushHyperlink();
      content.push(trackedContentForItem(item, revision));
      continue;
    }

    const linkKey = hyperlinkKey(item.attributes);
    if (linkKey !== null) {
      flushRun();
      const currentKey = currentHyperlink
        ? currentHyperlink.href || (currentHyperlink.anchor ? `#${currentHyperlink.anchor}` : '')
        : null;
      if (!currentHyperlink || currentKey !== linkKey) {
        flushHyperlink();
        currentHyperlink = createHyperlink(item.attributes);
      }
      if (currentHyperlink) addToHyperlink(currentHyperlink, item);
      continue;
    }

    flushHyperlink();
    if (item.kind === 'text') {
      const key = stableStringify(formattingAttrs(item.attributes));
      const nextRun = createTextRun(item.text, item.attributes);
      if (currentRun && currentFormattingKey === key) appendTextRun(currentRun, nextRun);
      else {
        flushRun();
        currentRun = nextRun;
        currentFormattingKey = key;
      }
      continue;
    }

    flushRun();
    const child = ordinaryContentForItem(item);
    if (child) content.push(child);
  }

  flushRun();
  flushHyperlink();
  return content;
}

function marksKeyToYrsAttrs(marksKey: string | undefined): Attrs | null {
  if (!marksKey) return {};
  const attrs: Attrs = {};
  // Mark attribute JSON does not normally contain `|`; if a custom string
  // does, declining restoration merely coalesces runs and never loses text.
  for (const part of marksKey.split('|')) {
    const colon = part.indexOf(':');
    if (colon <= 0) return null;
    const name = part.slice(0, colon);
    let value: Attrs;
    try {
      value = JSON.parse(part.slice(colon + 1)) as Attrs;
    } catch {
      return null;
    }
    if (name === 'comment' || name === 'footnoteRef') continue;
    if (BOOLEAN_MARKS.has(name)) attrs[name] = true;
    else if (name === 'highlight') attrs.highlight = value.color;
    else if (name === 'insertion' || name === 'deletion') {
      attrs[name === 'insertion' ? 'ins' : 'del'] = dropNulls({
        id: value.revisionId,
        author: value.author,
        date: value.date,
        isMovePair: value.isMovePair,
      });
    } else attrs[name] = dropNulls(value);
  }
  return attrs;
}

/** Rebuilds the note number marks a run held; they occupy no story unit. */
function noteMarkContent(boundary: OriginalRunBoundary): RunContent[] | null {
  if (!boundary.noteMarks?.length) return null;
  return boundary.noteMarks.map((mark) => ({
    type: mark === 'endnote' ? 'endnoteRefMark' : 'footnoteRefMark',
  }));
}

/**
 * Rebuilds a run from its recorded boundary. Flow breaks occupy no story unit,
 * so without their recorded offsets an authored page break would not be saved.
 */
function boundaryContent(boundary: OriginalRunBoundary, formatting: TextFormatting): RunContent[] {
  const notes = noteMarkContent(boundary);
  if (notes) return notes;
  if (!boundary.breaks?.length) return runContentForText(boundary.text, formatting);
  const content: RunContent[] = [];
  let cursor = 0;
  for (const entry of boundary.breaks) {
    const at = Math.min(Math.max(entry.offset, cursor), boundary.text.length);
    if (at > cursor) {
      content.push(...runContentForText(boundary.text.slice(cursor, at), formatting));
    }
    content.push({ type: 'break', breakType: entry.type });
    cursor = at;
  }
  if (cursor < boundary.text.length) {
    content.push(...runContentForText(boundary.text.slice(cursor), formatting));
  }
  return content;
}

function restoreOriginalRuns(
  content: ParagraphContent[],
  items: InlineItem[],
  boundaries: OriginalRunBoundary[] | undefined
): ParagraphContent[] {
  if (
    !boundaries?.length ||
    !content.every(
      (child) => child.type === 'run' && child.content.every((entry) => entry.type === 'text')
    ) ||
    items.some((item) => item.kind !== 'text' || item.attributes.hyperlink) ||
    boundaries.some((boundary) => boundary.noteMarks?.length && boundary.text.length > 0)
  ) {
    return content;
  }
  const fullText = items.map((item) => (item.kind === 'text' ? item.text : '')).join('');
  if (fullText !== boundaries.map((boundary) => boundary.text).join('')) return content;

  let itemIndex = 0;
  let itemOffset = 0;
  const restoredAttrs: Attrs[] = [];
  for (const boundary of boundaries) {
    const expected = marksKeyToYrsAttrs(boundary.marksKey);
    if (!expected) return content;
    restoredAttrs.push(expected);
    let remaining = boundary.text.length;
    while (remaining > 0) {
      const item = items[itemIndex];
      if (item?.kind !== 'text') return content;
      if (stableStringify(formattingAttrs(item.attributes)) !== stableStringify(expected)) {
        return content;
      }
      const available = item.text.length - itemOffset;
      const consumed = Math.min(remaining, available);
      remaining -= consumed;
      itemOffset += consumed;
      if (itemOffset === item.text.length) {
        itemIndex += 1;
        itemOffset = 0;
      }
    }
  }

  if (itemIndex !== items.length || itemOffset !== 0) return content;

  return boundaries.map((boundary, index) => {
    // restoreOriginalRuns restores the original segmentation/property-change
    // cache, but non-empty runs keep formatting reconstructed from their live
    // marks. Only empty runs have no node and therefore take cached formatting.
    const formatting =
      boundary.text.length === 0
        ? boundary.formatting
        : attrsToTextFormatting(restoredAttrs[index]);
    const run: Run = {
      type: 'run',
      content: boundaryContent(boundary, formatting ?? {}),
    };
    if (formatting && Object.keys(formatting).length > 0) run.formatting = formatting;
    if (boundary.propertyChanges?.length) run.propertyChanges = boundary.propertyChanges;
    return run;
  });
}

function runTextLength(run: Run): number {
  return run.content.reduce((length, content) => {
    if (content.type === 'text' || content.type === 'instrText')
      return length + content.text.length;
    if (content.type === 'symbol') return length + content.char.length;
    if (
      content.type === 'tab' ||
      content.type === 'softHyphen' ||
      content.type === 'noBreakHyphen' ||
      content.type === 'footnoteRef' ||
      content.type === 'endnoteRef' ||
      content.type === 'horizontalRule'
    ) {
      return length + 1;
    }
    return length;
  }, 0);
}

function paragraphContentLength(content: ParagraphContent): number {
  switch (content.type) {
    case 'run':
      return runTextLength(content);
    case 'hyperlink':
      return content.children.reduce(
        (sum, child) => sum + (child.type === 'run' ? runTextLength(child) : 0),
        0
      );
    case 'simpleField':
      return content.content.reduce((sum, child) => sum + paragraphContentLength(child), 0);
    case 'complexField':
      return content.fieldResult.reduce((sum, run) => sum + runTextLength(run), 0);
    case 'inlineSdt':
      return content.content.reduce((sum, child) => sum + paragraphContentLength(child), 0);
    case 'insertion':
    case 'deletion':
    case 'moveFrom':
    case 'moveTo':
      return content.content.reduce(
        (sum, child) => sum + (child.type === 'run' ? runTextLength(child) : 0),
        0
      );
    case 'mathEquation':
      return content.plainText?.length ?? 0;
    default:
      return 0;
  }
}

function runContentUnits(content: RunContent): number {
  switch (content.type) {
    case 'text':
      return content.text.length;
    case 'symbol': {
      const code = Number.parseInt(content.char, 16);
      return code >= 0 && code <= 0x10ffff ? String.fromCodePoint(code).length : 0;
    }
    case 'break':
      return content.breakType === undefined ||
        content.breakType === 'textWrapping' ||
        unitBreaks.has(content)
        ? 1
        : 0;
    case 'tab':
    case 'softHyphen':
    case 'noBreakHyphen':
    case 'horizontalRule':
    case 'commentReference':
    case 'drawing':
    case 'shape':
    case 'chart':
    case 'opaqueDrawing':
    case 'footnoteRef':
    case 'endnoteRef':
      return 1;
    default:
      return 0;
  }
}

/** Story units `content` seeds as (seed.rs `paragraph_units`), where comment anchors live. */
function storyUnits(content: ParagraphContent): number {
  switch (content.type) {
    case 'run':
      return content.content.reduce((sum, entry) => sum + runContentUnits(entry), 0);
    case 'hyperlink': {
      let units = 0;
      for (const child of content.structuredChildren ?? content.children) {
        if (child.type === 'run') units += storyUnits(child);
        else if (['simpleField', 'complexField', 'mathEquation'].includes(child.type)) units += 1;
      }
      return units;
    }
    case 'complexField': {
      // seed.rs `field_to_units`: result hyperlinks and nested simple fields seed their own units.
      const projected = [
        ...(content.structuredCode?.inline ?? []),
        ...(content.structuredResult?.inline ?? []),
      ].filter((child) => child.type === 'hyperlink' || child.type === 'simpleField');
      if (/^\d+$/.test(content.instruction.trim()) || projected.length === 0) return 1;
      return projected.reduce((sum, child) => sum + storyUnits(child), 1);
    }
    case 'simpleField':
    case 'inlineSdt':
    case 'mathEquation':
      return 1;
    case 'insertion':
    case 'deletion':
    case 'moveFrom':
    case 'moveTo':
      return content.content.reduce((sum, child) => sum + storyUnits(child), 0);
    default:
      return 0;
  }
}

/** Bookmark offsets count story units, but an inline content control as 2 (seed.rs `paragraph_attrs`). */
const bookmarkUnits = (content: ParagraphContent): number =>
  content.type === 'inlineSdt' ? 2 : storyUnits(content);

/**
 * `content` split `offset` story units in, or null where it cannot split: a
 * run splits between its entries or inside text, a tracked change around its
 * run; a hyperlink or field stays whole.
 */
function splitContent(
  content: ParagraphContent,
  offset: number
): [ParagraphContent, ParagraphContent] | null {
  if (
    (content.type === 'insertion' ||
      content.type === 'deletion' ||
      content.type === 'moveFrom' ||
      content.type === 'moveTo') &&
    content.content.length === 1 &&
    content.content[0]!.type === 'run'
  ) {
    const halves = splitContent(content.content[0]!, offset) as [Run, Run] | null;
    if (!halves) return null;
    return [
      { ...content, content: [halves[0]] },
      { ...content, content: [halves[1]] },
    ];
  }
  if (content.type !== 'run') return null;
  const left: RunContent[] = [];
  const right: RunContent[] = [];
  let units = 0;
  for (const entry of content.content) {
    const width = runContentUnits(entry);
    if (units + width <= offset) left.push(entry);
    else if (units >= offset) right.push(entry);
    else if (entry.type === 'text') {
      left.push({ type: 'text', text: entry.text.slice(0, offset - units) });
      right.push({ type: 'text', text: entry.text.slice(offset - units) });
    } else return null;
    units += width;
  }
  const half = (entries: RunContent[]): Run => ({
    type: 'run',
    ...(content.formatting ? { formatting: content.formatting } : {}),
    ...(content.propertyChanges ? { propertyChanges: content.propertyChanges } : {}),
    content: entries,
  });
  return [half(left), half(right)];
}

/**
 * Places each boundary `offset` units (by `measure`) into `content`. A
 * boundary inside content that cannot split moves to that content's edge, so
 * its range widens to hold the content whole.
 */
/** Boundaries by offset: an empty range opens before it closes, ranges meeting close first. */
function boundaryOrder(left: CommentBoundary, right: CommentBoundary): number {
  return (
    left.offset - right.offset ||
    (left.kind === right.kind
      ? left.id - right.id
      : left.id === right.id
        ? left.kind === 'start'
          ? -1
          : 1
        : left.kind === 'end'
          ? -1
          : 1)
  );
}

function insertBoundaries(
  content: ParagraphContent[],
  boundaries: CommentBoundary[],
  measure: (content: ParagraphContent) => number,
  makeMarker: (boundary: CommentBoundary) => ParagraphContent = (boundary) =>
    boundary.kind === 'start'
      ? { type: 'commentRangeStart', id: boundary.id }
      : { type: 'commentRangeEnd', id: boundary.id }
): ParagraphContent[] {
  if (boundaries.length === 0) return content;
  const sorted = [...boundaries].sort(boundaryOrder);
  const result: ParagraphContent[] = [];
  let cursor = 0;
  let boundaryIndex = 0;
  const emit = (offset: number): void => {
    while (boundaryIndex < sorted.length && sorted[boundaryIndex].offset === offset) {
      const boundary = sorted[boundaryIndex++];
      result.push(makeMarker(boundary));
    }
  };

  emit(0);
  for (const item of content) {
    const end = cursor + measure(item);
    let rest = item;
    let restStart = cursor;
    const widened: ParagraphContent[] = [];
    while (boundaryIndex < sorted.length && sorted[boundaryIndex].offset < end) {
      const boundary = sorted[boundaryIndex];
      const halves = splitContent(rest, boundary.offset - restStart);
      if (halves) {
        result.push(halves[0]);
        emit(boundary.offset);
        [rest, restStart] = [halves[1], boundary.offset];
      } else {
        boundaryIndex += 1;
        if (boundary.kind === 'start') result.push(makeMarker(boundary));
        else widened.push(makeMarker(boundary));
      }
    }
    result.push(rest, ...widened);
    cursor = end;
    emit(cursor);
  }
  emit(cursor);
  return result;
}

function bookmarkBoundaries(properties: Attrs): BookmarkBoundary[] {
  if (!Array.isArray(properties.bookmarks)) return [];
  const result: CommentBoundary[] = [];
  for (const raw of properties.bookmarks) {
    const bookmark = asObject(raw);
    const id = asFiniteNumber(bookmark?.id);
    if (!bookmark || id === undefined) continue;
    const offset = asFiniteNumber(bookmark.offset) ?? 0;
    const metadata = {
      name: asString(bookmark.name),
      colFirst: asFiniteNumber(bookmark.colFirst),
      colLast: asFiniteNumber(bookmark.colLast),
    };
    if (bookmark.kind === 'start') result.push({ id, kind: 'start', offset, ...metadata });
    else if (bookmark.kind === 'end') result.push({ id, kind: 'end', offset, ...metadata });
    else {
      result.push({ id, kind: 'start', offset: 0, ...metadata });
      result.push({ id, kind: 'end', offset: Number.MAX_SAFE_INTEGER, ...metadata });
    }
  }
  return result;
}

function restoreRawInlines(content: ParagraphContent[], base: Paragraph | undefined): ParagraphContent[] {
  if (!base) return content;
  const length = content.reduce((sum, child) => sum + paragraphContentLength(child), 0);
  const boundaries: CommentBoundary[] = [];
  let offset = 0;
  for (const [index, child] of base.content.entries()) {
    if (child.type === 'rawXml') boundaries.push({ id: index, kind: 'start', offset: Math.min(offset, length) });
    offset += paragraphContentLength(child);
  }
  return insertBoundaries(
    content,
    boundaries,
    paragraphContentLength,
    (boundary) => base.content[boundary.id]!
  );
}

function paragraphAttrs(properties: Attrs): ParagraphSaveAttrs {
  const attrs = { ...PARAGRAPH_ATTR_DEFAULTS, ...properties } as Attrs;
  attrs.styleId = properties.pStyle ?? null;
  attrs._sectionProperties = properties.sectPr ?? null;
  delete attrs.pStyle;
  delete attrs.sectPr;
  return attrs as ParagraphSaveAttrs;
}

function paragraphFromStory(
  paraId: string,
  properties: Attrs,
  items: InlineItem[],
  commentBoundaries: CommentBoundary[],
  baseParagraph: Paragraph | undefined
): Paragraph {
  const attrs = paragraphAttrs(properties);
  let content = buildParagraphContent(items);
  content = restoreOriginalRuns(
    content,
    items,
    Array.isArray(attrs._originalRunBoundaries)
      ? (attrs._originalRunBoundaries as OriginalRunBoundary[])
      : undefined
  );
  content = restoreRawInlines(content, baseParagraph);
  content = insertBoundaries(content, commentBoundaries, storyUnits);

  const bookmarks = bookmarkBoundaries(properties).map((boundary) => ({
    ...boundary,
    offset:
      boundary.offset === Number.MAX_SAFE_INTEGER
        ? content.reduce((sum, child) => sum + bookmarkUnits(child), 0)
        : boundary.offset,
  }));
  if (bookmarks.length > 0) {
    content = insertBoundaries(content, bookmarks, bookmarkUnits, (rawBoundary) => {
      const boundary = rawBoundary as BookmarkBoundary;
      return boundary.kind === 'start'
        ? {
            type: 'bookmarkStart',
            id: boundary.id,
            name: boundary.name || '',
            ...(boundary.colFirst !== undefined ? { colFirst: boundary.colFirst } : {}),
            ...(boundary.colLast !== undefined ? { colLast: boundary.colLast } : {}),
            position: { offset: boundary.offset },
          }
        : { type: 'bookmarkEnd', id: boundary.id, position: { offset: boundary.offset } };
    });
  }

  const paragraph: Paragraph = {
    type: 'paragraph',
    paraId: paraId || undefined,
    textId: baseParagraph?.textId,
    formatting: paragraphAttrsToFormatting(attrs),
    content,
  };
  if (baseParagraph?.renderedPageBreakBefore) paragraph.renderedPageBreakBefore = true;

  const pPrIns = trackedInfo(properties.pPrIns, true);
  const pPrDel = trackedInfo(properties.pPrDel, true);
  if (pPrIns) paragraph.pPrIns = pPrIns;
  if (pPrDel) paragraph.pPrDel = pPrDel;
  if (Array.isArray(properties.pPrChange) && properties.pPrChange.length > 0) {
    paragraph.propertyChanges = properties.pPrChange as Paragraph['propertyChanges'];
  }
  if (properties.sectPr) {
    paragraph.sectionProperties = properties.sectPr as Paragraph['sectionProperties'];
  } else if (properties.sectionBreakType) {
    paragraph.sectionProperties = {
      sectionStart: properties.sectionBreakType as NonNullable<
        Paragraph['sectionProperties']
      >['sectionStart'],
    };
  }
  return paragraph;
}

/**
 * Lifts a `w:tblBorders` back out of the per-cell edges seeding pushed down:
 * outer sides come from the cells owning the table's boundary, `insideH`/
 * `insideV` from an interior edge. Sides no cell authors stay absent.
 */
function inferTableBorders(rows: TableRow[]): TableBorders | undefined {
  const firstRow = rows[0]?.cells;
  const lastRow = rows[rows.length - 1]?.cells;
  const corner = firstRow?.[0]?.formatting?.borders;
  if (!firstRow || !lastRow || !corner) return undefined;
  const borders: TableBorders = {
    top: corner.top,
    left: corner.left,
    bottom: lastRow[0]?.formatting?.borders?.bottom,
    right: firstRow[firstRow.length - 1]?.formatting?.borders?.right,
    insideH: rows.length > 1 ? corner.bottom : undefined,
    insideV: firstRow.length > 1 ? corner.right : undefined,
  };
  const authored = Object.entries(borders).filter(([, value]) => value !== undefined);
  return authored.length > 0 ? (Object.fromEntries(authored) as TableBorders) : undefined;
}

function normalizeVMergeRuns(rows: TableRow[]): void {
  const columns = new Map<number, Array<{ rowIndex: number; cell: TableCell }>>();
  rows.forEach((row, rowIndex) => {
    let column = 0;
    for (const cell of row.cells) {
      const start = column;
      column += cell.formatting?.gridSpan ?? 1;
      if (cell.formatting?.vMerge) {
        const entries = columns.get(start) ?? [];
        entries.push({ rowIndex, cell });
        columns.set(start, entries);
      }
    }
  });

  const clear = (cell: TableCell): void => {
    if (!cell.formatting) return;
    delete cell.formatting.vMerge;
    if (Object.keys(cell.formatting).length === 0) cell.formatting = undefined;
  };

  for (const entries of columns.values()) {
    let start: TableCell | null = null;
    let length = 0;
    let lastRow = -1;
    const close = (): void => {
      if (start && length < 2) clear(start);
      start = null;
      length = 0;
    };
    for (const entry of entries) {
      const marker = entry.cell.formatting?.vMerge;
      if (marker === 'restart') {
        close();
        start = entry.cell;
        length = 1;
        lastRow = entry.rowIndex;
      } else if (marker === 'continue') {
        if (start && entry.rowIndex === lastRow + 1) {
          length += 1;
          lastRow = entry.rowIndex;
        } else {
          close();
          clear(entry.cell);
        }
      }
    }
    close();
  }
}

function tableCellFromPayload(context: SaveContext, payload: TableCellPayload): TableCell {
  const attrs = {
    ...TABLE_CELL_ATTR_DEFAULTS,
    ...(payload.tcPr ?? {}),
  } as unknown as TableCellSaveAttrs;
  const content =
    payload.story && context.storyIds.has(payload.story)
      ? context.storyToBlocks(payload.story)
      : [];
  const cell: TableCell = {
    type: 'tableCell',
    formatting: tableCellAttrsToFormatting(attrs),
    content,
  };
  const marker = attrs.cellMarker;
  if (marker) {
    const info = trackedInfo(marker.info, true) ?? { id: 0, author: 'Unknown' };
    if (marker.kind === 'ins') cell.structuralChange = { type: 'tableCellInsertion', info };
    else if (marker.kind === 'del') cell.structuralChange = { type: 'tableCellDeletion', info };
    else {
      cell.structuralChange = {
        type: 'tableCellMerge',
        info,
        ...(marker.vMerge ? { vMerge: marker.vMerge } : {}),
        ...(marker.vMergeOrig ? { vMergeOrig: marker.vMergeOrig } : {}),
      };
    }
  }
  if (Array.isArray(attrs.tcPrChange) && attrs.tcPrChange.length > 0) {
    cell.propertyChanges = attrs.tcPrChange;
  }
  return cell;
}

function tableFromPayload(context: SaveContext, payload: TablePayload): Table {
  const rowPayloads = Array.isArray(payload.rows) ? payload.rows : [];
  const occupied: boolean[][] = [];
  const anchors: Array<{
    row: number;
    col: number;
    rowspan: number;
    colspan: number;
    cell: TableCell;
  }> = [];
  let totalColumns = 0;

  rowPayloads.forEach((row, rowIndex) => {
    let column = 0;
    for (const cellPayload of Array.isArray(row.cells) ? row.cells : []) {
      while (occupied[rowIndex]?.[column]) column += 1;
      const tcPr = cellPayload.tcPr ?? {};
      const rowspan = asFiniteNumber(tcPr.rowspan) || 1;
      const colspan = asFiniteNumber(tcPr.colspan) || 1;
      anchors.push({
        row: rowIndex,
        col: column,
        rowspan,
        colspan,
        cell: tableCellFromPayload(context, cellPayload),
      });
      for (let r = rowIndex; r < rowIndex + rowspan; r += 1) {
        occupied[r] ??= [];
        for (let c = column; c < column + colspan; c += 1) occupied[r][c] = true;
      }
      column += colspan;
      totalColumns = Math.max(totalColumns, column);
    }
  });

  const byStart = new Map(anchors.map((anchor) => [`${anchor.row}-${anchor.col}`, anchor]));
  const byCovered = new Map<string, (typeof anchors)[number]>();
  for (const anchor of anchors) {
    for (let row = anchor.row; row < anchor.row + anchor.rowspan; row += 1) {
      for (let col = anchor.col; col < anchor.col + anchor.colspan; col += 1) {
        byCovered.set(`${row}-${col}`, anchor);
      }
    }
  }

  const rows: TableRow[] = rowPayloads.map((rowPayload, rowIndex) => {
    const cells: TableCell[] = [];
    for (let col = 0; col < totalColumns; ) {
      const anchor = byStart.get(`${rowIndex}-${col}`);
      if (anchor) {
        const formatting = { ...(anchor.cell.formatting ?? {}) };
        if (anchor.colspan > 1) formatting.gridSpan = anchor.colspan;
        else delete formatting.gridSpan;
        if (anchor.rowspan > 1) formatting.vMerge = 'restart';
        else if (formatting.vMerge !== 'restart' && formatting.vMerge !== 'continue') {
          delete formatting.vMerge;
        }
        cells.push({
          ...anchor.cell,
          formatting: Object.keys(formatting).length > 0 ? formatting : undefined,
        });
        col += anchor.colspan;
        continue;
      }
      const covering = byCovered.get(`${rowIndex}-${col}`);
      if (!covering) {
        col += 1;
        continue;
      }
      const formatting = { ...(covering.cell.formatting ?? {}) };
      if (covering.colspan > 1) formatting.gridSpan = covering.colspan;
      else delete formatting.gridSpan;
      formatting.vMerge = 'continue';
      cells.push({ ...covering.cell, content: [], formatting });
      col += covering.colspan;
    }

    const attrs = { ...TABLE_ROW_ATTR_DEFAULTS, ...(rowPayload.trPr ?? {}) } as TableRowSaveAttrs;
    const row: TableRow = {
      type: 'tableRow',
      formatting: tableRowAttrsToFormatting(attrs),
      cells,
    };
    const ins = trackedInfo(attrs.trIns, true);
    const del = trackedInfo(attrs.trDel, true);
    if (ins) row.structuralChange = { type: 'tableRowInsertion', info: ins };
    else if (del) row.structuralChange = { type: 'tableRowDeletion', info: del };
    if (Array.isArray(attrs.trPrChange) && attrs.trPrChange.length > 0) {
      row.propertyChanges = attrs.trPrChange;
    }
    return row;
  });

  normalizeVMergeRuns(rows);
  const grid = Array.isArray(payload.grid)
    ? payload.grid.filter((width): width is number => typeof width === 'number')
    : [];
  const attrs = {
    ...TABLE_ATTR_DEFAULTS,
    ...(payload.tblPr ?? {}),
    columnWidths: grid.length > 0 ? grid : undefined,
  } as TableSaveAttrs;
  let formatting = tableAttrsToFormatting(attrs);
  if (!formatting?.borders) {
    const borders = inferTableBorders(rows);
    if (borders) formatting = { ...(formatting ?? {}), borders };
  }
  const table: Table = {
    type: 'table',
    columnWidths: attrs.columnWidths || undefined,
    formatting,
    rows,
  };
  if (Array.isArray(attrs.tblPrChange) && attrs.tblPrChange.length > 0) {
    table.propertyChanges = attrs.tblPrChange;
  }
  return table;
}

function pageBreakParagraph(): Paragraph {
  return {
    type: 'paragraph',
    content: [{ type: 'run', content: [{ type: 'break', breakType: 'page' }] }],
  };
}

type FlowBreak = 'page' | 'column';
type Tracked = 'ins' | 'del' | undefined;
interface FlowToken {
  kind: FlowBreak | 'visible';
  run?: Run;
  index?: number;
  /** The tracked change around a break. */
  tracked?: Tracked;
}
/** A break unit opening a paragraph slot. */
interface SlotBreak {
  kind: FlowBreak;
  /** Its story offset. */
  at: number;
  /** Seeded from a page break opening its paragraph's text: it stays that text's first run. */
  leading: boolean;
  attributes: Attrs;
}

const breakRun = (kind: FlowBreak): Run => ({
  type: 'run',
  content: [{ type: 'break', breakType: kind }],
});

/** Page and column breaks that are story units; one the run cache restores holds none. */
const unitBreaks = new WeakSet<RunContent>();
function unitBreakRun(kind: FlowBreak): Run {
  const run = breakRun(kind);
  unitBreaks.add(run.content[0]!);
  return run;
}

const trackedKind = (attributes: Attrs): Tracked =>
  attributes.ins ? 'ins' : attributes.del ? 'del' : undefined;

/** A slot's break as a run, inside the tracked change it carries. */
function slotBreakContent({ kind, attributes }: SlotBreak): ParagraphContent {
  const info = trackedInfo(attributes.ins ?? attributes.del);
  if (!info) return breakRun(kind);
  const item: InlineItem = {
    kind: 'embed',
    embedKind: 'flowBreak',
    payload: { breakType: kind },
    attributes,
  };
  return trackedContentForItem(item, info);
}

/**
 * A paragraph's page/column breaks and visible content in the order the seed
 * reads them (`inline_tokens` in crates/docx-edit/src/seed.rs).
 */
function flowTokens(content: readonly ParagraphContent[]): FlowToken[] {
  const tokens: FlowToken[] = [];
  const run = (item: Run, tracked?: Tracked): void =>
    item.content.forEach((entry, index) => {
      if (entry.type === 'break' && (entry.breakType === 'page' || entry.breakType === 'column'))
        tokens.push({ kind: entry.breakType, run: item, index, tracked });
      else if (entry.type !== 'text' || entry.text !== '') tokens.push({ kind: 'visible' });
    });
  const runs = (items: readonly { type: string }[], tracked?: Tracked): void => {
    for (const item of items) if (item.type === 'run') run(item as Run, tracked);
  };
  const inline = (items: readonly ParagraphContent[]): void => {
    for (const item of items) {
      if (item.type === 'run') run(item);
      else if (item.type === 'hyperlink') runs(item.children);
      else if (item.type === 'simpleField') runs(item.content);
      else if (item.type === 'complexField') runs([...item.fieldCode, ...item.fieldResult]);
      else if (item.type === 'inlineSdt') inline(item.content);
      else if (item.type === 'insertion' || item.type === 'moveTo') runs(item.content, 'ins');
      else if (item.type === 'deletion' || item.type === 'moveFrom') runs(item.content, 'del');
      else if (item.type === 'mathEquation') tokens.push({ kind: 'visible' });
    }
  };
  inline(content);
  return tokens;
}

/**
 * Splits a paragraph's breaks as the seed places them (`paragraph_flow_breaks`):
 * `leading` become units before its text, `trailing` units after its pilcrow.
 * Without text, the breaks up to its last column break lead.
 */
function splitFlow(tokens: FlowToken[]): { leading: FlowToken[]; trailing: FlowToken[] } {
  let split = tokens.findIndex((token) => token.kind === 'visible');
  if (split < 0) {
    split = 0;
    tokens.forEach((token, index) => {
      if (token.kind === 'column') split = index + 1;
    });
  }
  const breaks = (part: FlowToken[]) => part.filter((token) => token.kind !== 'visible');
  return { leading: breaks(tokens.slice(0, split)), trailing: breaks(tokens.slice(split)) };
}

/** `content` without the given break tokens; runs left empty are dropped. */
function withoutBreaks(
  content: readonly ParagraphContent[],
  tokens: readonly FlowToken[]
): ParagraphContent[] {
  const drop = new Map<Run, Set<number>>();
  for (const token of tokens) {
    if (!token.run || token.index === undefined) continue;
    const indexes = drop.get(token.run) ?? new Set<number>();
    indexes.add(token.index);
    drop.set(token.run, indexes);
  }
  const run = (item: Run): Run | null => {
    const indexes = drop.get(item);
    if (!indexes) return item;
    const kept = item.content.filter((_, index) => !indexes.has(index));
    return kept.length > 0 ? { ...item, content: kept } : null;
  };
  const runs = <T extends { type: string }>(items: readonly T[]): T[] =>
    items.flatMap((item): T[] => {
      if (item.type !== 'run') return [item];
      const kept = run(item as unknown as Run);
      return kept ? [kept as unknown as T] : [];
    });
  const inline = <T extends ParagraphContent>(items: readonly T[]): T[] =>
    items.flatMap((item): T[] => {
      if (item.type === 'run') return runs([item]);
      if (item.type === 'hyperlink') return [{ ...item, children: runs(item.children) }];
      if (item.type === 'simpleField') return [{ ...item, content: runs(item.content) }];
      if (item.type === 'complexField')
        return [{ ...item, fieldCode: runs(item.fieldCode), fieldResult: runs(item.fieldResult) }];
      if (item.type === 'inlineSdt') return [{ ...item, content: inline(item.content) }];
      if (
        item.type === 'insertion' ||
        item.type === 'deletion' ||
        item.type === 'moveFrom' ||
        item.type === 'moveTo'
      ) {
        const kept = runs(item.content);
        return kept.length > 0 || item.content.length === 0 ? [{ ...item, content: kept }] : [];
      }
      return [item];
    });
  return drop.size === 0 ? [...content] : inline(content);
}

/** A note number mark run, or a tracked-change wrapper holding only those. */
function isNoteMark(content: ParagraphContent): boolean {
  if (
    content.type === 'insertion' ||
    content.type === 'deletion' ||
    content.type === 'moveFrom' ||
    content.type === 'moveTo'
  ) {
    return content.content.length > 0 && content.content.every(isNoteMark);
  }
  return (
    content.type === 'run' &&
    content.content.length > 0 &&
    content.content.every(
      (entry) => entry.type === 'footnoteRefMark' || entry.type === 'endnoteRefMark'
    )
  );
}

function hasNoteMark(blocks: readonly BlockContent[]): boolean {
  return blocks.some((block) => block.type === 'paragraph' && block.content.some(isNoteMark));
}

/**
 * Reinstates the `w:footnoteRef` / `w:endnoteRef` number mark on a projected
 * note. The mark carries no story unit, so an edit that invalidates the run
 * boundary cache would otherwise drop it; the source note is the only record.
 */
function restoreNoteMarks(
  projected: BlockContent[],
  base: readonly BlockContent[]
): BlockContent[] {
  if (hasNoteMark(projected)) return projected;
  const opening = base.find((block) => block.type === 'paragraph')?.content ?? [];
  const end = opening.findIndex((child) => !isNoteMark(child));
  const marks = opening.slice(0, end < 0 ? opening.length : end);
  if (marks.length === 0) return projected;
  const index = projected.findIndex((block) => block.type === 'paragraph');
  if (index < 0) return projected;
  const target = projected[index] as Paragraph;
  const restored = [...projected];
  restored[index] = { ...target, content: [...marks, ...target.content] };
  return restored;
}

function collectBaseParagraphs(document: Document): Map<string, Paragraph> {
  const paragraphs = new Map<string, Paragraph>();
  const visit = (blocks: readonly BlockContent[]): void => {
    for (const block of blocks) {
      if (block.type === 'paragraph') {
        if (block.paraId && !paragraphs.has(block.paraId)) paragraphs.set(block.paraId, block);
      } else if (block.type === 'table') {
        for (const row of block.rows) for (const cell of row.cells) visit(cell.content);
      } else if (block.type === 'blockSdt') {
        visit(block.content);
      }
    }
  };
  visit(document.package.document.content);
  for (const part of document.package.headers?.values() ?? []) visit(part.content);
  for (const part of document.package.footers?.values() ?? []) visit(part.content);
  for (const note of document.package.footnotes ?? []) visit(note.content);
  for (const note of document.package.endnotes ?? []) visit(note.content);
  return paragraphs;
}

function collectBaseStories(document: Document): Map<string, readonly BlockContent[]> {
  const stories = new Map<string, readonly BlockContent[]>();
  const visit = (storyId: string, blocks: readonly BlockContent[]): void => {
    stories.set(storyId, blocks);
    let tableIndex = 0;
    let sdtIndex = 0;
    for (const block of blocks) {
      if (block.type === 'blockSdt') {
        visit(`${storyId}:sdt${sdtIndex++}`, block.content);
        continue;
      }
      if (block.type !== 'table') continue;
      const currentTableIndex = tableIndex++;
      block.rows.forEach((row, rowIndex) => {
        row.cells.forEach((cell, cellIndex) => {
          visit(`${storyId}:t${currentTableIndex}:r${rowIndex}c${cellIndex}`, cell.content);
        });
      });
    }
  };

  visit('body', document.package.document.content);
  for (const [rId, part] of document.package.headers ?? []) visit(`hf:${rId}`, part.content);
  for (const [rId, part] of document.package.footers ?? []) visit(`hf:${rId}`, part.content);
  for (const note of document.package.footnotes ?? []) visit(`fn:${note.id}`, note.content);
  for (const note of document.package.endnotes ?? []) visit(`en:${note.id}`, note.content);
  return stories;
}

function commentRanges(
  session: YrsSession,
  comments: readonly Comment[] | undefined
): Map<string, Array<{ id: number; start: number; end: number }>> {
  const byStory = new Map<string, Array<{ id: number; start: number; end: number }>>();
  for (const comment of comments ?? []) {
    let anchors: ReturnType<YrsSession['resolveComment']>;
    try {
      anchors = session.resolveComment(commentSharedId(comment));
    } catch {
      continue;
    }
    const storyGroups = new Map<string, typeof anchors>();
    for (const anchor of anchors) {
      const group = storyGroups.get(anchor.story) ?? [];
      group.push(anchor);
      storyGroups.set(anchor.story, group);
    }
    for (const [story, group] of storyGroups) {
      if (group.length === 0) continue;
      const start = Math.min(...group.map((anchor) => anchor.start));
      const end = Math.max(...group.map((anchor) => anchor.end));
      const ranges = byStory.get(story) ?? [];
      ranges.push({ id: comment.id, start, end });
      byStory.set(story, ranges);
    }
  }
  return byStory;
}

function restoreRawBlocks(projected: BlockContent[], base: readonly BlockContent[]): BlockContent[] {
  let offset = 0;
  for (let index = 0; index < base.length; index += 1) {
    const block = base[index]!;
    if (!isRawXml(block)) continue;
    const following = base.slice(index + 1).find((candidate) =>
      candidate.type === 'paragraph' && candidate.paraId && projected.some((entry) =>
        entry.type === 'paragraph' && entry.paraId === candidate.paraId));
    const anchor = following?.type === 'paragraph'
      ? projected.findIndex((entry) => entry.type === 'paragraph' && entry.paraId === following.paraId)
      : -1;
    const position = anchor >= 0 ? anchor : Math.min(index + offset, projected.length);
    projected.splice(position, 0, block);
    offset = Math.max(0, position - index);
  }
  return projected;
}

/** A projected block plus a fingerprint of the story inputs it was built from. */
interface ProjectedBlockMemo {
  /** Bucket key for locating same-container candidates (first cell story / child story). */
  key?: string;
  /** Snapshot of the raw story inputs the block was built from. */
  inputs?: readonly unknown[];
  /** Container cell/child content arrays in payload order (tables, block SDTs). */
  children?: readonly BlockContent[][];
}

interface ProjectedStory {
  commentsKey: string;
  blocks: BlockContent[];
}

/** Per-session projection reuse state shared across `yrsToDocument` calls. */
interface SessionProjectionMemo {
  /** An op without story scope ran; every story counts dirty except `clean` ones. */
  wholesale: boolean;
  clean: Set<string>;
  dirty: Set<string>;
  stories: Map<string, ProjectedStory>;
}

const projectedBlocks = new WeakMap<BlockContent, ProjectedBlockMemo>();
const sessionProjectionMemos = new WeakMap<YrsSession, SessionProjectionMemo>();

function sessionProjectionMemo(session: YrsSession): SessionProjectionMemo {
  let memo = sessionProjectionMemos.get(session);
  if (!memo) {
    memo = { wholesale: true, clean: new Set(), dirty: new Set(), stories: new Map() };
    sessionProjectionMemos.set(session, memo);
  }
  return memo;
}

/** Structural equality over JSON-shaped story inputs — allocation-free sig check. */
function sameJson(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (Array.isArray(a) && Array.isArray(b)) {
    return a.length === b.length && a.every((entry, index) => sameJson(entry, b[index]));
  }
  const left = asObject(a);
  const right = asObject(b);
  if (!left || !right) return false;
  const keys = Object.keys(left);
  return (
    keys.length === Object.keys(right).length &&
    keys.every((key) => key in right && sameJson(left[key], right[key]))
  );
}

const NESTED_STORY_ID = /^(.*?)(?::t\d+:r\d+c\d+|:sdt\d+)$/;

/** Stories a committed op touched; `all` invalidates the whole session. */
export function noteYrsStoriesDirty(
  session: YrsSession,
  stories: 'all' | string | Iterable<string>
): void {
  const memo = sessionProjectionMemo(session);
  if (stories === 'all') {
    memo.wholesale = true;
    memo.clean.clear();
    memo.dirty.clear();
    memo.stories.clear();
    return;
  }
  const queue = typeof stories === 'string' ? [stories] : [...stories];
  for (let index = 0; index < queue.length; index += 1) {
    const story = queue[index]!;
    memo.dirty.add(story);
    memo.clean.delete(story);
    memo.stories.delete(story);
    const parent = NESTED_STORY_ID.exec(story)?.[1];
    if (parent) queue.push(parent);
  }
}

const EMPTY_BLOCKS: BlockContent[] = [];

/** `w14:paraId` is eight hex digits below this; docx-parse re-mints any other value. */
const PARA_ID_LIMIT = 0x7fffffff;
const isParaId = (id: string): boolean =>
  /^[0-9A-Fa-f]{8}$/.test(id) && parseInt(id, 16) < PARA_ID_LIMIT;

class SaveContext {
  readonly storyIds: Set<string>;
  readonly projectedComments: Comment[];
  private readonly baseParagraphs: Map<string, Paragraph>;
  /** Every `w14:paraId` the base holds, then each one minted by {@link savedParaId}. */
  private readonly paraIds: Set<number>;
  private readonly baseStories: Map<string, readonly BlockContent[]>;
  private readonly comments: Map<string, Array<{ id: number; start: number; end: number }>>;
  private readonly storyOwners = new WeakMap<object, string>();
  private readonly projectedStories = new Set<string>();
  private readonly memo: SessionProjectionMemo;
  private readonly bypassMemo: boolean;

  constructor(
    private readonly session: YrsSession,
    base: Document,
    private readonly onEmbed?: YrsToDocumentOptions['onEmbed'],
    private readonly onParagraph?: YrsToDocumentOptions['onParagraph'],
    trackStories = false
  ) {
    this.storyIds = new Set(session.storyIds());
    this.baseParagraphs = collectBaseParagraphs(base);
    this.paraIds = new Set([...this.baseParagraphs.keys()].map((id) => parseInt(id, 16)));
    this.baseStories = collectBaseStories(base);
    this.projectedComments = projectYrsComments(session, base.package.document.comments);
    this.comments = commentRanges(session, this.projectedComments);
    this.memo = sessionProjectionMemo(session);
    // Hooked projections (checkpoint export, rebase) run once and must see
    // every block, so they neither read nor fill the session cache.
    this.bypassMemo = trackStories || onEmbed !== undefined || onParagraph !== undefined;
  }

  /**
   * The `w14:paraId` a story paragraph saves with. A source id is kept. An
   * editor id (`<client>:<clock>`, or a seeded `body:pN` that moved) is no
   * ST_LongHexNumber, so it saves as a hex id hashed from it and probed past
   * every id already taken; the same state always saves the same ids.
   */
  private savedParaId(id: string): string {
    if (isParaId(id)) return id;
    let hash = 2166136261;
    for (let index = 0; index < id.length; index += 1)
      hash = Math.imul(hash ^ id.charCodeAt(index), 16777619);
    let value = ((hash >>> 0) % (PARA_ID_LIMIT - 1)) + 1;
    while (this.paraIds.has(value)) value = (value % (PARA_ID_LIMIT - 1)) + 1;
    this.paraIds.add(value);
    return value.toString(16).toUpperCase().padStart(8, '0');
  }

  private storyIsClean(storyId: string): boolean {
    return this.memo.wholesale
      ? this.memo.clean.has(storyId)
      : !this.memo.dirty.has(storyId);
  }

  private cellContents(payload: TablePayload): BlockContent[][] {
    const contents: BlockContent[][] = [];
    for (const row of Array.isArray(payload.rows) ? payload.rows : []) {
      for (const cell of Array.isArray(row.cells) ? row.cells : []) {
        contents.push(
          cell.story !== undefined && this.storyIds.has(cell.story)
            ? this.storyToBlocks(cell.story)
            : EMPTY_BLOCKS
        );
      }
    }
    return contents;
  }

  storyToBlocks(storyId: string): BlockContent[] {
    this.projectedStories.add(storyId);
    const baseBlocks = this.baseStories.get(storyId);
    const storyComments = this.comments.get(storyId) ?? [];
    const commentsKey = storyComments.length === 0 ? '' : stableStringify(storyComments);
    const priorStory = this.memo.stories.get(storyId);
    // Nested stories are mapped positionally, so a different array means
    // drift, not divergence — mutation paths mark them dirty anyway. A root
    // story on a different array is genuine divergence and forces a fresh
    // projection.
    if (
      !this.bypassMemo &&
      priorStory !== undefined &&
      priorStory.commentsKey === commentsKey &&
      this.storyIsClean(storyId) &&
      (NESTED_STORY_ID.test(storyId) ||
        baseBlocks === undefined ||
        baseBlocks === priorStory.blocks)
    ) {
      return priorStory.blocks;
    }
    const blocks: BlockContent[] = [];
    const baseParagraphBlocks = baseBlocks?.filter((block): block is Paragraph => block.type === 'paragraph');
    const segments = this.session.storySegments(storyId);
    let items: InlineItem[] = [];
    let paragraphStart = 0;
    // Past a table, content control or break that opens the paragraph's slot.
    let contentStart = 0;
    let paragraphIndex = 0;
    let storyOffset = 0;
    let candidatesByKey: Map<string, BlockContent[]> | null = null;

    const candidatesFor = (key: string): BlockContent[] => {
      if (this.bypassMemo) return EMPTY_BLOCKS;
      if (!candidatesByKey) {
        candidatesByKey = new Map();
        for (const block of baseBlocks ?? []) {
          const blockKey = projectedBlocks.get(block)?.key;
          if (blockKey === undefined) continue;
          const bucket = candidatesByKey.get(blockKey);
          if (bucket) bucket.push(block);
          else candidatesByKey.set(blockKey, [block]);
        }
      }
      return candidatesByKey.get(key) ?? [];
    };

    const reuseContainer = (
      key: string,
      inputs: readonly unknown[],
      children: readonly BlockContent[][]
    ): BlockContent | undefined =>
      candidatesFor(key).find((block) => {
        const memo = projectedBlocks.get(block);
        return (
          memo?.inputs !== undefined &&
          memo.children !== undefined &&
          sameJson(memo.inputs, inputs) &&
          memo.children.length === children.length &&
          memo.children.every((child, index) => child === children[index])
        );
      });

    // A range that leaves the story it opened in (Word's range into the next
    // cell, say) covers its final pilcrow; it closes at its last paragraph's end.
    let lastPilcrow = -1;
    segments.reduce((offset, segment) => {
      if (segment.kind === 'pilcrow') lastPilcrow = offset;
      return offset + (segment.kind === 'text' ? segment.text.length : 1);
    }, 0);
    // Boundaries in the paragraph's text, and those at the units opening its
    // slot (offset: the story offset), which settle places among its breaks.
    const paragraphCommentBoundaries = (end: number) => {
      const boundaries: CommentBoundary[] = [];
      const slot: CommentBoundary[] = [];
      const add = (id: number, kind: 'start' | 'end', at: number) => {
        if (at < paragraphStart || at > end) return;
        if (at < contentStart) slot.push({ id, kind, offset: at });
        else boundaries.push({ id, kind, offset: at - contentStart });
      };
      for (const range of storyComments) {
        add(range.id, 'start', range.start);
        add(range.id, 'end', Math.min(range.end, lastPilcrow));
      }
      return { boundaries, slot: slot.sort(boundaryOrder) };
    };

    // Break units opening a slot are where the seed put the trailing breaks of
    // the paragraph before it, or the breaks leading the slot paragraph's text.
    // Paragraphs keep breaks that still match the units and are rewritten
    // from the units otherwise.
    let slotBreaks: SlotBreak[] = [];
    let slotInlineBreaks: SlotBreak[] = [];
    let carried: SlotBreak[] = [];
    let previous = -1;
    const settle = (next?: Paragraph, marks: CommentBoundary[] = []): Paragraph | undefined => {
      const expected = [...carried, ...slotBreaks];
      carried = [];
      slotBreaks = [];
      const before = previous >= 0 ? (blocks[previous] as Paragraph) : undefined;
      const beforeFlow = before && splitFlow(flowTokens(before.content));
      const nextFlow = next && splitFlow(flowTokens(next.content));
      const text =
        next !== undefined && flowTokens(next.content).some(({ kind }) => kind === 'visible');
      // From the first leading break on they open the text (all of them with no
      // paragraph before); the rest close the paragraph before.
      const leading = expected.findIndex((entry) => entry.leading);
      const split = !text || (before && leading < 0) ? expected.length : before ? leading : 0;
      const trailing = beforeFlow?.trailing.length ?? 0;
      const have = [...(beforeFlow?.trailing ?? []), ...(nextFlow?.leading ?? [])];
      if (
        have.length === expected.length &&
        have.every(
          (token, index) =>
            token.kind === expected[index]!.kind &&
            token.tracked === trackedKind(expected[index]!.attributes) &&
            ((token.kind === 'column' && !expected[index]!.leading) ||
              index < trailing === index < split)
        ) &&
        // Comment boundaries at the slot's breaks are placed among them below.
        marks.length === 0
      )
        return next;
      const rest =
        next && nextFlow && nextFlow.leading.length > 0
          ? { ...next, content: withoutBreaks(next.content, nextFlow.leading) }
          : next;
      // A boundary at a break the paragraph leads with goes just before it;
      // others at the slot (before a table, or at a break the paragraph
      // before closes with) open the paragraph.
      const lead = (paragraph: Paragraph | undefined, entries: SlotBreak[]) => {
        if (!paragraph || (entries.length === 0 && marks.length === 0)) return paragraph;
        const head: ParagraphContent[] = [];
        let mark = 0;
        const markTo = (at: number) => {
          while (mark < marks.length && marks[mark]!.offset <= at) {
            const { id, kind } = marks[mark++]!;
            head.push({ type: kind === 'start' ? 'commentRangeStart' : 'commentRangeEnd', id });
          }
        };
        for (const entry of entries) {
          markTo(entry.at);
          head.push(slotBreakContent(entry));
        }
        markTo(Number.POSITIVE_INFINITY);
        return { ...paragraph, content: [...head, ...paragraph.content] };
      };
      if (before && beforeFlow) {
        blocks[previous] = {
          ...before,
          content: [
            ...withoutBreaks(before.content, beforeFlow.trailing),
            ...expected.slice(0, split).map(slotBreakContent),
          ],
        };
        return lead(rest, expected.slice(split));
      }
      if (text) return lead(rest, expected);
      // No paragraph before them and no text to lead: a paragraph of their own.
      if (expected.length > 0)
        blocks.push({ type: 'paragraph', content: expected.map(slotBreakContent) });
      return lead(rest, []);
    };
    const openBlock = (): void => {
      settle();
      previous = -1;
    };

    const pushText = (text: string, attributes: Attrs): void => {
      let cursor = 0;
      for (let index = 0; index < text.length; index += 1) {
        if (text[index] !== '\t') continue;
        if (index > cursor)
          items.push({ kind: 'text', text: text.slice(cursor, index), attributes });
        items.push({ kind: 'embed', embedKind: 'tab', payload: {}, attributes });
        cursor = index + 1;
      }
      if (cursor < text.length) items.push({ kind: 'text', text: text.slice(cursor), attributes });
    };

    for (const segment of segments) {
      if (segment.kind === 'text') {
        pushText(segment.text, segment.attributes);
        storyOffset += segment.text.length;
        continue;
      }
      if (segment.kind === 'pilcrow') {
        const generatedId = `${storyId}:p${paragraphIndex}`;
        const savedParaId =
          segment.paraId === generatedId && !this.baseParagraphs.has(segment.paraId)
            ? ''
            : this.savedParaId(segment.paraId);
        const baseParagraph =
          this.baseParagraphs.get(segment.paraId) ??
          (segment.paraId === generatedId ? baseParagraphBlocks?.[paragraphIndex] : undefined);
        const { boundaries, slot } = paragraphCommentBoundaries(storyOffset);
        const inputs = [
          segment.paraId,
          savedParaId,
          segment.properties,
          segment.attributes,
          items,
          boundaries,
        ] as const;
        const priorMemo =
          baseParagraph !== undefined ? projectedBlocks.get(baseParagraph) : undefined;
        let paragraph: Paragraph;
        if (
          !this.bypassMemo &&
          baseParagraph !== undefined &&
          priorMemo?.inputs !== undefined &&
          sameJson(priorMemo.inputs, inputs)
        ) {
          paragraph = baseParagraph;
        } else {
          // Projection mutates `item.payload`; pin the pre-mutation inputs.
          const snapshot = inputs.map((input, index) =>
            index === 4 ? (input as InlineItem[]).map((item) => ({ ...item })) : input
          );
          paragraph = paragraphFromStory(
            savedParaId,
            segment.properties,
            items,
            boundaries,
            baseParagraph
          );
          projectedBlocks.set(paragraph, { inputs: snapshot });
        }
        paragraph = settle(paragraph, slot) ?? paragraph;
        carried = slotInlineBreaks;
        slotInlineBreaks = [];
        blocks.push(paragraph);
        previous = blocks.length - 1;
        this.onParagraph?.(storyId, storyOffset, paragraph, segment.paraId);
        items = [];
        paragraphIndex += 1;
        storyOffset += 1;
        paragraphStart = contentStart = storyOffset;
        continue;
      }

      let projectedEmbed: ParagraphContent | BlockContent | null = null;
      if (segment.embedKind === 'table') {
        const payload = segment.payload as TablePayload;
        const inputs = [segment.attributes, payload] as const;
        const firstCell = Array.isArray(payload.rows) ? payload.rows[0]?.cells?.[0] : undefined;
        const key = `T${firstCell?.story ?? ''}`;
        const contents =
          candidatesFor(key).length > 0 ? this.cellContents(payload) : undefined;
        const reused =
          contents !== undefined
            ? (reuseContainer(key, inputs, contents) as Table | undefined)
            : undefined;
        let table: Table;
        if (reused) {
          table = reused;
        } else {
          table = tableFromPayload(this, payload);
          if (!this.bypassMemo) {
            projectedBlocks.set(table, {
              key,
              inputs,
              children: contents ?? this.cellContents(payload),
            });
          }
        }
        projectedEmbed = table;
        openBlock();
        blocks.push(table);
      } else if (segment.embedKind === 'blockSdt') {
        const childStory = asString(segment.payload.story);
        openBlock();
        const childContent =
          childStory && this.storyIds.has(childStory)
            ? this.storyToBlocks(childStory)
            : EMPTY_BLOCKS;
        const inputs = [segment.attributes, segment.payload] as const;
        const reused = reuseContainer(`S${childStory ?? ''}`, inputs, [childContent]);
        if (reused) {
          projectedEmbed = reused;
          blocks.push(reused);
        } else {
          let properties = sdtAttrsToProps(segment.payload);
          let content = childContent;
          const authoredValue = contentControlValue(segment.payload.value);
          if (authoredValue) {
            try {
              const applied = applyContentControlValue(properties, authoredValue);
              properties = applied.properties;
              content = applied.content;
            } catch {
              // Retain the child story if the authored value is invalid.
            }
          }
          const block: BlockContent = {
            type: 'blockSdt',
            properties,
            content: content === EMPTY_BLOCKS ? [] : content,
          };
          // Rebind the child's entry to the embedded array for the next projection.
          if (!this.bypassMemo && content !== childContent && childStory !== undefined) {
            const childEntry = this.memo.stories.get(childStory);
            if (childEntry) childEntry.blocks = block.content;
          }
          projectedBlocks.set(block, {
            key: `S${childStory ?? ''}`,
            inputs,
            children: [block.content],
          });
          projectedEmbed = block;
          blocks.push(block);
        }
      } else if (segment.embedKind === 'opaque') {
        const blob = asObject(segment.payload.blob);
        if (blob?.type === 'pageBreak') {
          projectedEmbed = pageBreakParagraph();
          openBlock();
          blocks.push(projectedEmbed);
        } else {
          // Carry a same-position base block while block SDTs stay opaque.
          const baseBlock = baseBlocks?.[blocks.length];
          if (blob?.type === 'blockSdt' && baseBlock?.type === 'blockSdt') {
            projectedEmbed = baseBlock;
            openBlock();
            blocks.push(baseBlock);
          }
          // Standalone text-box blobs do not have a one-to-one base block (the
          // base model stores them inside paragraph runs), so they remain the
          // documented opaque carry gap.
        }
      } else if (segment.embedKind === 'pageBreak' || segment.embedKind === 'columnBreak') {
        const entry: SlotBreak = {
          at: storyOffset,
          kind: segment.embedKind === 'pageBreak' ? 'page' : 'column',
          leading: segment.payload.leading === true,
          attributes: segment.attributes,
        };
        if (items.length === 0) slotBreaks.push(entry);
        else {
          // Inside paragraph content, which the seed never produces: kept in
          // place, and seeded back as a break after the paragraph.
          items.push({
            kind: 'embed',
            embedKind: 'flowBreak',
            payload: { breakType: entry.kind },
            attributes: segment.attributes,
          });
          slotInlineBreaks.push({ ...entry, leading: false });
        }
      } else {
        items.push(segment as EmbedItem);
        if (this.onEmbed) projectedEmbed = ordinaryContentForItem(segment as EmbedItem);
      }
      if (projectedEmbed) this.onEmbed?.(storyId, storyOffset, projectedEmbed);
      storyOffset += 1;
      if (items.length === 0) contentStart = storyOffset;
    }
    settle();

    // Defensive recovery for malformed/legacy stories without a final pilcrow.
    // A story ending in a flow-break embed is well-formed, not a lost
    // paragraph, so only content the projection can carry opens one.
    if (items.length > 0) {
      const trailing = buildParagraphContent(items);
      if (trailing.length > 0) blocks.push({ type: 'paragraph', content: trailing });
    }
    const projected = restoreRawBlocks(blocks, baseBlocks ?? []);
    for (const block of projected) this.storyOwners.set(block, storyId);
    if (!this.bypassMemo) {
      this.memo.stories.set(storyId, { commentsKey, blocks: projected });
      this.memo.dirty.delete(storyId);
      this.memo.clean.add(storyId);
    }
    return projected;
  }

  visitReachableStories(document: Document, visit: (storyId: string) => void): void {
    const reachable = new Set<string>();
    const walk = (blocks: readonly BlockContent[]): void => {
      for (const block of blocks) {
        const owner = this.storyOwners.get(block);
        if (owner) reachable.add(owner);
        if (block.type === 'table') {
          for (const row of block.rows) for (const cell of row.cells) walk(cell.content);
        } else if (block.type === 'blockSdt') walk(block.content);
      }
    };
    walk(document.package.document.content);
    for (const parts of [document.package.headers, document.package.footers]) {
      for (const part of parts?.values() ?? []) walk(part.content);
    }
    for (const [prefix, notes] of [
      ['fn:', document.package.footnotes],
      ['en:', document.package.endnotes],
    ] as const) {
      for (const note of notes ?? []) {
        const storyId = `${prefix}${note.id}`;
        if (this.projectedStories.has(storyId)) reachable.add(storyId);
        walk(note.content);
      }
    }
    for (const storyId of reachable) visit(storyId);
  }
}

/**
 * Rebuild all yrs-owned editable stories into the serializer-facing Document
 * while preserving every package part the editor does not own.
 */
export interface YrsToDocumentOptions {
  /**
   * Root stories to project. Omit to rebuild every editable story. A `body`
   * projection includes its nested table/content-control stories.
   */
  storyIds?: ReadonlySet<string>;
  /** Stories whose projected blocks remain reachable in the authored output. */
  onStory?: (storyId: string) => void;
  /** Current serializer-facing content for a native embed at a UTF-16 offset. */
  onEmbed?: (storyId: string, offset: number, content: ParagraphContent | BlockContent) => void;
  /** Projected paragraph and the native pilcrow offset that owns it. */
  onParagraph?: (storyId: string, offset: number, paragraph: Paragraph, logicalId: string) => void;
}

export function yrsToDocument(
  session: YrsSession,
  base: Document,
  options: YrsToDocumentOptions = {}
): Document {
  const context = new SaveContext(
    session,
    base,
    options.onEmbed,
    options.onParagraph,
    options.onStory !== undefined
  );
  const shouldProject = (storyId: string): boolean =>
    options.storyIds === undefined || options.storyIds.has(storyId);
  const bodyContent = context.storyIds.has('body') && shouldProject('body')
    ? context.storyToBlocks('body')
    : base.package.document.content;

  let headers = base.package.headers;
  if (
    headers &&
    (options.storyIds === undefined || [...headers.keys()].some((rId) => shouldProject(`hf:${rId}`)))
  ) {
    headers = new Map(
      [...headers].map(([rId, part]) => {
        const storyId = `hf:${rId}`;
        return [
          rId,
          context.storyIds.has(storyId) && shouldProject(storyId)
            ? { ...part, content: context.storyToBlocks(storyId) }
            : part,
        ];
      })
    );
  }

  let footers = base.package.footers;
  if (
    footers &&
    (options.storyIds === undefined || [...footers.keys()].some((rId) => shouldProject(`hf:${rId}`)))
  ) {
    footers = new Map(
      [...footers].map(([rId, part]) => {
        const storyId = `hf:${rId}`;
        return [
          rId,
          context.storyIds.has(storyId) && shouldProject(storyId)
            ? { ...part, content: context.storyToBlocks(storyId) }
            : part,
        ];
      })
    );
  }

  const projectNotes = <T extends Footnote | Endnote>(
    notes: T[] | undefined,
    prefix: string
  ): T[] | undefined => {
    const shouldProjectNotes =
      options.storyIds === undefined ||
      notes?.some((note) => shouldProject(`${prefix}${note.id}`));
    if (!shouldProjectNotes) return notes;
    return notes?.map((note) => {
      const storyId = `${prefix}${note.id}`;
      return context.storyIds.has(storyId) && shouldProject(storyId)
        ? {
            ...note,
            content: restoreNoteMarks(context.storyToBlocks(storyId), note.content),
            verbatimXml: undefined,
          }
        : note;
    });
  };
  const footnotes = projectNotes(base.package.footnotes, 'fn:');
  const endnotes = projectNotes(base.package.endnotes, 'en:');

  const document: Document = {
    ...base,
    package: {
      ...base.package,
      document: {
        ...base.package.document,
        content: bodyContent,
        comments: context.projectedComments,
      },
      ...(headers ? { headers } : {}),
      ...(footers ? { footers } : {}),
      ...(footnotes ? { footnotes } : {}),
      ...(endnotes ? { endnotes } : {}),
    },
  };
  if (options.onStory) context.visitReachableStories(document, options.onStory);
  return document;
}
