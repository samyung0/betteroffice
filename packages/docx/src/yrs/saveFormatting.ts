import type {
  ParagraphFormatting,
  TableCellFormatting,
  TableCellPropertyChange,
  TableFormatting,
  TablePropertyChange,
  TableRowFormatting,
  TableRowPropertyChange,
} from '../types/document';
import type { RevisionInfo } from '../types/content/trackedChange';

export interface ParagraphSaveAttrs extends Record<string, unknown> {
  alignment?: ParagraphFormatting['alignment'];
  spaceBefore?: number;
  spaceAfter?: number;
  spaceBeforeLines?: number;
  spaceAfterLines?: number;
  beforeAutospacing?: boolean;
  afterAutospacing?: boolean;
  lineSpacing?: number;
  lineSpacingRule?: ParagraphFormatting['lineSpacingRule'];
  indentLeft?: number;
  indentRight?: number;
  indentFirstLine?: number;
  hangingIndent?: boolean;
  numPr?: ParagraphFormatting['numPr'];
  numPrFromStyle?: ParagraphFormatting['numPr'];
  styleId?: string;
  borders?: ParagraphFormatting['borders'];
  shading?: ParagraphFormatting['shading'];
  tabs?: ParagraphFormatting['tabs'];
  outlineLevel?: number;
  contextualSpacing?: boolean;
  keepNext?: boolean;
  keepLines?: boolean;
  snapToGrid?: boolean;
  pageBreakBefore?: boolean;
  widowControl?: boolean | null;
  autoSpaceDE?: boolean | null;
  autoSpaceDN?: boolean | null;
  bidi?: boolean;
  _originalFormatting?: ParagraphFormatting;
  _originalRunBoundaries?: unknown[];
}

function isStyleSourcedNumPr(attrs: ParagraphSaveAttrs): boolean {
  return (
    attrs.numPrFromStyle != null &&
    attrs.numPr != null &&
    JSON.stringify(attrs.numPr) === JSON.stringify(attrs.numPrFromStyle)
  );
}

/** The pPr properties the editor holds resolved (direct, else list level, else style). */
const RESOLVED_PROPERTIES = [
  'alignment',
  'spaceBefore',
  'spaceAfter',
  'spaceBeforeLines',
  'spaceAfterLines',
  'beforeAutospacing',
  'afterAutospacing',
  'lineSpacing',
  'lineSpacingRule',
  'indentLeft',
  'indentRight',
  'indentFirstLine',
  'hangingIndent',
  'borders',
  'shading',
  'tabs',
  'pageBreakBefore',
  'keepNext',
  'keepLines',
  'widowControl',
  'contextualSpacing',
  'snapToGrid',
  'autoSpaceDE',
  'autoSpaceDN',
  'outlineLevel',
  'bidi',
] as const satisfies readonly (keyof ParagraphFormatting)[];

type ResolvedProperty = (typeof RESOLVED_PROPERTIES)[number];

/** Properties one OOXML attribute set carries together: a change to one writes all. */
const LINKED: ReadonlyArray<readonly ResolvedProperty[]> = [
  ['lineSpacing', 'lineSpacingRule'],
  ['indentFirstLine', 'hangingIndent'],
];

/** Character-unit twins Word reads in place of a twips indent, dropped when the indent changes. */
const CHARACTER_TWINS: Partial<Record<ResolvedProperty, ReadonlyArray<keyof ParagraphFormatting>>> = {
  indentLeft: ['indentLeftChars'],
  indentRight: ['indentRightChars'],
  indentFirstLine: ['indentFirstLineChars', 'hangingIndentChars'],
};

/** JSON with sorted keys and without nulls, so stored and parsed values compare equal. */
function canonical(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(canonical).join(',')}]`;
  if (value !== null && typeof value === 'object') {
    const entries = Object.entries(value).filter(([, entry]) => entry != null);
    entries.sort(([a], [b]) => (a < b ? -1 : 1));
    return `{${entries.map(([key, entry]) => `${JSON.stringify(key)}:${canonical(entry)}`).join(',')}}`;
  }
  return JSON.stringify(value ?? null);
}

/** Tab stops in the model's shape; editor ops store `{pos, val}`. */
function modelTabs(value: unknown): ParagraphFormatting['tabs'] {
  if (!Array.isArray(value)) return undefined;
  return value.map((stop: Record<string, unknown>) => ({
    position: (stop.position ?? stop.pos) as number,
    alignment: (stop.alignment ?? stop.val) as NonNullable<ParagraphFormatting['tabs']>[number]['alignment'],
    ...(stop.leader != null ? { leader: stop.leader as NonNullable<ParagraphFormatting['tabs']>[number]['leader'] } : {}),
  }));
}

function modelValue(attrs: ParagraphSaveAttrs, key: ResolvedProperty): unknown {
  return key === 'tabs' ? modelTabs(attrs.tabs) : attrs[key];
}

/**
 * The pPr a paragraph saves with: its source formatting, with every resolved
 * property the editor holds differently from what the seed gave it
 * (`seeded`) written over it. Untouched properties stay as the source wrote
 * them, and style or list values are never copied into direct formatting.
 */
export function paragraphAttrsToFormatting(
  attrs: ParagraphSaveAttrs,
  seeded: Readonly<Record<string, unknown>>
): ParagraphFormatting | undefined {
  const orig = attrs._originalFormatting ?? {};
  const result: Record<string, unknown> = { ...orig };
  const changed = new Set(
    RESOLVED_PROPERTIES.filter(
      (key) => canonical(modelValue(attrs, key)) !== canonical(key === 'tabs' ? modelTabs(seeded.tabs) : seeded[key])
    )
  );
  for (const group of LINKED) {
    if (group.some((key) => changed.has(key))) for (const key of group) changed.add(key);
  }
  for (const key of changed) {
    const value = modelValue(attrs, key);
    if (value == null) delete result[key];
    else result[key] = value;
    for (const twin of CHARACTER_TWINS[key] ?? []) delete result[twin];
  }
  if (isStyleSourcedNumPr(attrs)) {
    delete result.numPr;
    delete result.numPrFromStyle;
  } else if (attrs.numPr !== orig.numPr && JSON.stringify(attrs.numPr) !== JSON.stringify(orig.numPr)) {
    result.numPr = attrs.numPr || undefined;
    delete result.numPrFromStyle;
  }
  if (attrs.styleId !== (orig.styleId || undefined)) {
    result.styleId = attrs.styleId || undefined;
  }
  return Object.values(result).some((value) => value !== undefined)
    ? (result as ParagraphFormatting)
    : undefined;
}

export interface TableSaveAttrs extends Record<string, unknown> {
  styleId?: string;
  width?: number;
  widthType?: string;
  justification?: TableFormatting['justification'];
  columnWidths?: number[];
  tableLayout?: TableFormatting['layout'];
  floating?: TableFormatting['floating'];
  cellMargins?: { top?: number; bottom?: number; left?: number; right?: number };
  look?: TableFormatting['look'];
  bidi?: boolean;
  _originalFormatting?: TableFormatting;
  tblPrChange?: TablePropertyChange[] | null;
}

export function tableAttrsToFormatting(attrs: TableSaveAttrs): TableFormatting | undefined {
  if (attrs._originalFormatting) {
    const orig = attrs._originalFormatting;
    const result = { ...orig };
    if (attrs.styleId !== (orig.styleId || undefined)) result.styleId = attrs.styleId || undefined;
    if (attrs.justification !== (orig.justification || undefined)) {
      result.justification = attrs.justification || undefined;
    }
    if (attrs.floating !== (orig.floating || undefined)) {
      result.floating = attrs.floating || undefined;
    }
    if (attrs.tableLayout !== (orig.layout || undefined)) {
      result.layout = attrs.tableLayout || undefined;
    }
    if (attrs.look !== (orig.look || undefined)) result.look = attrs.look || undefined;
    if (attrs.bidi !== (orig.bidi || undefined)) result.bidi = attrs.bidi || undefined;
    if (attrs.width !== orig.width?.value || attrs.widthType !== orig.width?.type) {
      result.width =
        attrs.width != null || attrs.widthType
          ? {
              value: attrs.width ?? 0,
              type: (attrs.widthType as 'auto' | 'dxa' | 'pct' | 'nil') || 'dxa',
            }
          : undefined;
    }
    if (attrs.cellMargins) {
      result.cellMargins = measurementMargins(attrs.cellMargins);
    }
    return result;
  }

  const hasFormatting =
    attrs.styleId ||
    attrs.width != null ||
    attrs.widthType ||
    attrs.justification ||
    attrs.tableLayout ||
    attrs.floating ||
    attrs.cellMargins ||
    attrs.look ||
    attrs.bidi;
  if (!hasFormatting) return undefined;
  return {
    styleId: attrs.styleId || undefined,
    width:
      attrs.width != null || attrs.widthType
        ? {
            value: attrs.width ?? 0,
            type: (attrs.widthType as 'auto' | 'dxa' | 'pct' | 'nil') || 'dxa',
          }
        : undefined,
    justification: attrs.justification || undefined,
    layout: attrs.tableLayout || undefined,
    floating: attrs.floating || undefined,
    cellMargins: attrs.cellMargins ? measurementMargins(attrs.cellMargins) : undefined,
    look: attrs.look || undefined,
    bidi: attrs.bidi || undefined,
  };
}

function measurementMargins(margins: {
  top?: number;
  bottom?: number;
  left?: number;
  right?: number;
}): NonNullable<TableFormatting['cellMargins']> {
  return {
    top: margins.top != null ? { value: margins.top, type: 'dxa' } : undefined,
    bottom: margins.bottom != null ? { value: margins.bottom, type: 'dxa' } : undefined,
    left: margins.left != null ? { value: margins.left, type: 'dxa' } : undefined,
    right: margins.right != null ? { value: margins.right, type: 'dxa' } : undefined,
  };
}

export interface TableRowSaveAttrs extends Record<string, unknown> {
  height?: number;
  heightRule?: string;
  isHeader?: boolean;
  _originalFormatting?: TableRowFormatting;
  trIns?: RevisionInfo | null;
  trDel?: RevisionInfo | null;
  trPrChange?: TableRowPropertyChange[] | null;
}

export function tableRowAttrsToFormatting(
  attrs: TableRowSaveAttrs
): TableRowFormatting | undefined {
  if (attrs._originalFormatting) {
    const orig = attrs._originalFormatting;
    const result = { ...orig };
    if (attrs.height !== (orig.height?.value || undefined)) {
      result.height = attrs.height ? { value: attrs.height, type: 'dxa' } : undefined;
    }
    if (attrs.heightRule !== (orig.heightRule || undefined)) {
      result.heightRule = (attrs.heightRule as 'auto' | 'atLeast' | 'exact') || undefined;
    }
    if (attrs.isHeader !== (orig.header || undefined)) {
      result.header = attrs.isHeader || undefined;
    }
    return result;
  }
  if (!attrs.height && !attrs.isHeader) return undefined;
  return {
    height: attrs.height ? { value: attrs.height, type: 'dxa' } : undefined,
    heightRule: (attrs.heightRule as 'auto' | 'atLeast' | 'exact') || undefined,
    header: attrs.isHeader || undefined,
  };
}

export interface TableCellSaveAttrs extends Record<string, unknown> {
  colspan: number;
  rowspan: number;
  width?: number;
  widthType?: string;
  verticalAlign?: TableCellFormatting['verticalAlign'];
  backgroundColor?: string;
  textDirection?: TableCellFormatting['textDirection'];
  borders?: TableCellFormatting['borders'];
  margins?: { top?: number; bottom?: number; left?: number; right?: number };
  _originalFormatting?: TableCellFormatting;
  _originalResolvedFill?: string;
  cellMarker?:
    | { kind: 'ins'; info: RevisionInfo }
    | { kind: 'del'; info: RevisionInfo }
    | {
        kind: 'merge';
        info: RevisionInfo;
        vMerge: 'rest' | 'cont';
        vMergeOrig?: 'rest' | 'cont';
      }
    | null;
  tcPrChange?: TableCellPropertyChange[] | null;
}

export function tableCellAttrsToFormatting(
  attrs: TableCellSaveAttrs
): TableCellFormatting | undefined {
  if (attrs._originalFormatting) {
    const orig = attrs._originalFormatting;
    const result = { ...orig };
    if (attrs.colspan > 1) result.gridSpan = attrs.colspan;
    if (attrs.width != null) {
      result.width = {
        value: attrs.width,
        type: (attrs.widthType as 'auto' | 'dxa' | 'pct' | 'nil') || 'dxa',
      };
    }
    if (attrs.verticalAlign !== (orig.verticalAlign || undefined)) {
      result.verticalAlign = attrs.verticalAlign || undefined;
    }
    if (attrs.backgroundColor) {
      result.shading =
        attrs._originalResolvedFill === attrs.backgroundColor && orig.shading
          ? orig.shading
          : { fill: { rgb: attrs.backgroundColor } };
    } else if (orig.shading) {
      result.shading = undefined;
    }
    if (attrs.borders) result.borders = attrs.borders;
    if (attrs.margins) result.margins = cellMargins(attrs.margins);
    if (attrs.textDirection !== (orig.textDirection || undefined)) {
      result.textDirection = attrs.textDirection || undefined;
    }
    return result;
  }

  const hasFormatting =
    attrs.colspan > 1 ||
    attrs.rowspan > 1 ||
    attrs.width != null ||
    attrs.verticalAlign ||
    attrs.backgroundColor ||
    attrs.borders ||
    attrs.margins ||
    attrs.textDirection;
  if (!hasFormatting) return undefined;
  return {
    gridSpan: attrs.colspan > 1 ? attrs.colspan : undefined,
    width:
      attrs.width != null
        ? {
            value: attrs.width,
            type: (attrs.widthType as 'auto' | 'dxa' | 'pct' | 'nil') || 'dxa',
          }
        : undefined,
    verticalAlign: attrs.verticalAlign || undefined,
    textDirection: attrs.textDirection || undefined,
    shading: attrs.backgroundColor ? { fill: { rgb: attrs.backgroundColor } } : undefined,
    borders: attrs.borders,
    margins: attrs.margins ? cellMargins(attrs.margins) : undefined,
  };
}

function cellMargins(margins: {
  top?: number;
  bottom?: number;
  left?: number;
  right?: number;
}): NonNullable<TableCellFormatting['margins']> {
  const result: NonNullable<TableCellFormatting['margins']> = {};
  if (margins.top != null) result.top = { value: margins.top, type: 'dxa' };
  if (margins.bottom != null) result.bottom = { value: margins.bottom, type: 'dxa' };
  if (margins.left != null) result.left = { value: margins.left, type: 'dxa' };
  if (margins.right != null) result.right = { value: margins.right, type: 'dxa' };
  return result;
}
