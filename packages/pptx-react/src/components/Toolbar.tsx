import { Fragment, useContext, useEffect, useMemo, useRef, useState } from 'react';
import type { CSSProperties, ReactNode } from 'react';
import type { ParagraphAlignment, TextAnchorValue } from '@betteroffice/pptx';
import type { TranslationKey } from '@betteroffice/pptx-i18n';
import { useTranslation } from '../i18n';
import {
  BULLET_PRESETS,
  NUMBER_PRESETS,
  presetLabel,
  type ListKind,
  type ListPresetId,
  type ParagraphFormatting,
} from '../paragraphFormatting';
import { EditorToolbarContext } from './EditorToolbarContext';
import { ColorPicker } from './ui/ColorPicker';
import { EditableCombobox } from './ui/EditableCombobox';
import { ToolbarIcon } from './ui/ToolbarIcon';
import { useToolbarRowScroll } from './useToolbarRowScroll';
import type { ToolbarIconName } from './ui/ToolbarIcon';
import {
  ToolbarButton,
  ToolbarDropdown,
  ToolbarGroup,
  ToolbarMenuItem,
  ToolbarMenuLabel,
  ToolbarMenuSeparator,
  ToolbarSeparator,
  toolbarColors,
} from './ui/ToolbarPrimitives';

export const SHAPE_PRESETS = [
  { geometry: 'rect', labelKey: 'toolbar.shapes.rect' },
  { geometry: 'roundRect', labelKey: 'toolbar.shapes.roundRect' },
  { geometry: 'ellipse', labelKey: 'toolbar.shapes.ellipse' },
  { geometry: 'triangle', labelKey: 'toolbar.shapes.triangle' },
  { geometry: 'rtTriangle', labelKey: 'toolbar.shapes.rtTriangle' },
  { geometry: 'diamond', labelKey: 'toolbar.shapes.diamond' },
  { geometry: 'parallelogram', labelKey: 'toolbar.shapes.parallelogram' },
  { geometry: 'trapezoid', labelKey: 'toolbar.shapes.trapezoid' },
  { geometry: 'pentagon', labelKey: 'toolbar.shapes.pentagon' },
  { geometry: 'hexagon', labelKey: 'toolbar.shapes.hexagon' },
  { geometry: 'octagon', labelKey: 'toolbar.shapes.octagon' },
  { geometry: 'star5', labelKey: 'toolbar.shapes.star5' },
  { geometry: 'rightArrow', labelKey: 'toolbar.shapes.rightArrow' },
  { geometry: 'leftArrow', labelKey: 'toolbar.shapes.leftArrow' },
  { geometry: 'upArrow', labelKey: 'toolbar.shapes.upArrow' },
  { geometry: 'downArrow', labelKey: 'toolbar.shapes.downArrow' },
  { geometry: 'chevron', labelKey: 'toolbar.shapes.chevron' },
] as const satisfies ReadonlyArray<{ geometry: string; labelKey: TranslationKey }>;

export type PptxShapePreset = (typeof SHAPE_PRESETS)[number]['geometry'];
export type PptxEditorTool = 'select' | 'textBox' | `shape:${PptxShapePreset}`;
export type PptxZoom = number | 'fit';

export interface SelectionFormatting {
  fontFamily?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** Superscript or subscript; null for neither, undefined for mixed. */
  script?: 'super' | 'sub' | null;
  textColor?: string;
  /** null: no highlight; undefined: mixed. */
  highlight?: string | null;
  align?: ParagraphAlignment;
}

export type FormattingAction =
  | 'bold'
  | 'italic'
  | 'underline'
  | 'strikethrough'
  | 'clearFormatting'
  | { type: 'fontFamily'; value: string }
  | { type: 'fontSize'; value: number }
  | { type: 'textColor'; value: string }
  | { type: 'highlight'; value: string | null }
  | { type: 'align'; value: ParagraphAlignment }
  | { type: 'verticalAlign'; value: TextAnchorValue }
  /** Without a style, toggles the list kind. */
  | { type: 'list'; kind: ListKind; preset?: ListPresetId }
  | { type: 'indent'; delta: 1 | -1 }
  | { type: 'lineSpacing'; value: string }
  | { type: 'spaceBefore' }
  | { type: 'spaceAfter' };

export interface ShapeFormatting {
  geometry?: string;
  fillColor?: string | null;
  strokeColor?: string | null;
  strokeWidthPt?: number | null;
  adjustments?: Record<string, number>;
}

export type ShapeZOrder = 'front' | 'forward' | 'backward' | 'back';

export type ShapeFormattingAction =
  | { type: 'fillColor'; value: string | null }
  | { type: 'strokeColor'; value: string | null }
  | { type: 'strokeWidth'; value: number | null }
  | { type: 'adjust'; name: string; value: number }
  | { type: 'zOrder'; value: ShapeZOrder };

export interface SlideLayoutOption {
  partPath: string | null;
  label?: string;
}

export interface ToolbarProps {
  currentFormatting?: SelectionFormatting;
  /** The touched paragraphs' lists, spacing and the text box's anchor. */
  currentParagraph?: ParagraphFormatting;
  textSelectionActive?: boolean;
  onFormat?: (action: FormattingAction) => void;
  currentShapeFormatting?: ShapeFormatting;
  shapeSelectionActive?: boolean;
  /** Enables the arrange (z-order) menu for any selected object. */
  shapeArrangeActive?: boolean;
  onShapeFormat?: (action: ShapeFormattingAction) => void;
  onInsertSlide?: (layoutPartPath?: string | null) => void;
  onInsertImage?: () => void;
  slideLayouts?: readonly SlideLayoutOption[];
  currentLayoutPartPath?: string | null;
  onSave?: () => void;
  onExportPng?: () => void;
  onUndo?: () => void;
  onRedo?: () => void;
  canUndo?: boolean;
  canRedo?: boolean;
  zoom?: PptxZoom;
  onZoomChange?: (zoom: PptxZoom) => void;
  activeTool?: PptxEditorTool;
  onToolChange?: (tool: PptxEditorTool) => void;
  fontFamilies?: readonly string[];
  fontSizes?: readonly number[];
  disabled?: boolean;
  /**
   * One flat full-height row in Google Slides' order whose groups scroll
   * sideways. Text and shape controls show only for a matching selection; save,
   * PNG export and arrange are left to the host's menus.
   */
  singleRow?: boolean;
  /** Single row only; hosts hide these on narrow screens. */
  showFontPicker?: boolean;
  showFontSizePicker?: boolean;
  showZoomControl?: boolean;
  className?: string;
  style?: CSSProperties;
  children?: ReactNode;
}

interface ToolbarSection {
  key: string;
  width: number;
  node: ReactNode;
}

const ZOOM_LEVELS = [0.5, 0.75, 1, 1.25, 1.5, 2] as const;
const DEFAULT_FONT_FAMILIES = [
  'Arial',
  'Calibri',
  'Cambria',
  'Georgia',
  'Roboto',
  'Times New Roman',
  'Verdana',
] as const;
export const DEFAULT_FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72] as const;
const VERTICAL_ALIGNMENTS = [
  { value: 't', icon: 'alignTop', labelKey: 'toolbar.align.top', testId: 'top' },
  { value: 'ctr', icon: 'alignMiddle', labelKey: 'toolbar.align.middle', testId: 'middle' },
  { value: 'b', icon: 'alignBottom', labelKey: 'toolbar.align.bottom', testId: 'bottom' },
] as const satisfies ReadonlyArray<{
  value: TextAnchorValue;
  icon: ToolbarIconName;
  labelKey: TranslationKey;
  testId: string;
}>;
/** Google Slides' line spacings, as shares of single. */
export const LINE_SPACINGS = [
  { value: '1', labelKey: 'toolbar.lineSpacing.single' },
  { value: '1.15', label: '1.15' },
  { value: '1.5', label: '1.5' },
  { value: '2', labelKey: 'toolbar.lineSpacing.double' },
] as const satisfies ReadonlyArray<{ value: string; labelKey?: TranslationKey; label?: string }>;
const ALIGNMENTS = [
  { value: 'l', icon: 'alignLeft', labelKey: 'toolbar.align.left', testId: 'left' },
  { value: 'ctr', icon: 'alignCenter', labelKey: 'toolbar.align.center', testId: 'center' },
  { value: 'r', icon: 'alignRight', labelKey: 'toolbar.align.right', testId: 'right' },
  { value: 'just', icon: 'alignJustify', labelKey: 'toolbar.align.justify', testId: 'justify' },
] as const satisfies ReadonlyArray<{
  value: ParagraphAlignment;
  icon: ToolbarIconName;
  labelKey: TranslationKey;
  testId: string;
}>;
const BORDER_WIDTHS = [1, 2, 3, 4, 8] as const;
const CORNER_RADIUS_OPTIONS = [0, 10, 17, 25, 33, 50] as const;
const SHAPE_ADJUSTMENT_OPTIONS = [0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100] as const;

function stripUndefined<T extends object>(value: T): Partial<T> {
  const result: Partial<T> = {};
  for (const key of Object.keys(value) as Array<keyof T>) {
    if (value[key] !== undefined) result[key] = value[key];
  }
  return result;
}

function useToolbarProps(props: ToolbarProps): ToolbarProps {
  const context = useContext(EditorToolbarContext);
  return context ? { ...context, ...stripUndefined(props) } : props;
}

/** A fade over a row edge that has more to scroll (useToolbarRowScroll). */
function EdgeFade({ side, visible }: { side: 'start' | 'end'; visible: boolean }) {
  return (
    <div
      aria-hidden="true"
      style={{
        position: 'absolute',
        top: 0,
        bottom: 0,
        [side === 'start' ? 'left' : 'right']: 0,
        width: 28,
        pointerEvents: 'none',
        background: `linear-gradient(to ${side === 'start' ? 'right' : 'left'}, var(--pptx-surface, #ffffff), transparent)`,
        opacity: visible ? 1 : 0,
        transition: 'opacity 0.15s',
      }}
    />
  );
}

export function nextFontSize(value: number, sizes: readonly number[], direction: -1 | 1): number {
  if (direction > 0) return sizes.find((size) => size > value) ?? value + 1;
  return [...sizes].reverse().find((size) => size < value) ?? Math.max(1, value - 1);
}

function primaryAdjustment(
  adjustments: Record<string, number> | undefined
): [string, number] | null {
  if (!adjustments) return null;
  if (adjustments.adj !== undefined) return ['adj', adjustments.adj];
  if (adjustments.adj1 !== undefined) return ['adj1', adjustments.adj1];
  const name = Object.keys(adjustments).sort()[0];
  return name ? [name, adjustments[name]] : null;
}

export function Toolbar(explicitProps: ToolbarProps) {
  const { t } = useTranslation();
  const {
    currentFormatting = {},
    currentParagraph,
    textSelectionActive = false,
    onFormat,
    currentShapeFormatting = {},
    shapeSelectionActive = false,
    shapeArrangeActive = false,
    onShapeFormat,
    onInsertSlide,
    onInsertImage,
    slideLayouts = [],
    currentLayoutPartPath,
    onSave,
    onExportPng,
    onUndo,
    onRedo,
    canUndo = false,
    canRedo = false,
    zoom = 'fit',
    onZoomChange,
    activeTool = 'select',
    onToolChange,
    fontFamilies = DEFAULT_FONT_FAMILIES,
    fontSizes = DEFAULT_FONT_SIZES,
    disabled = false,
    singleRow = false,
    showFontPicker = true,
    showFontSizePicker = true,
    showZoomControl = true,
    className,
    style,
    children,
  } = useToolbarProps(explicitProps);
  const rootRef = useRef<HTMLDivElement>(null);
  const rowRef = useRef<HTMLDivElement>(null);
  const rowEdges = useToolbarRowScroll(rowRef, singleRow);
  const [rootWidth, setRootWidth] = useState(Number.POSITIVE_INFINITY);
  const formattingEnabled = !disabled && textSelectionActive && Boolean(onFormat);
  const shapeFormattingEnabled = !disabled && shapeSelectionActive && Boolean(onShapeFormat);
  const arrangeEnabled = !disabled && shapeArrangeActive && Boolean(onShapeFormat);
  const slideEnabled = !disabled && Boolean(onInsertSlide);
  const toolEnabled = !disabled && Boolean(onToolChange);
  const insertImageEnabled = !disabled && Boolean(onInsertImage);
  const fontSize = currentFormatting.fontSize ?? 24;
  const fitLabel = t('toolbar.fit');
  const zoomValue = zoom === 'fit' ? fitLabel : `${Math.round(zoom * 100)}%`;
  const shapeAdjustment = primaryAdjustment(currentShapeFormatting.adjustments);
  const roundRectAdjustment =
    currentShapeFormatting.geometry === 'roundRect' && shapeAdjustment?.[0] === 'adj';

  useEffect(() => {
    const root = rootRef.current;
    if (!root) return;
    const update = () => setRootWidth(root.clientWidth);
    update();
    const observer = new ResizeObserver(update);
    observer.observe(root);
    return () => observer.disconnect();
  }, []);

  const zoomOptions = useMemo(
    () => [
      { value: fitLabel, label: fitLabel },
      ...ZOOM_LEVELS.map((level) => ({
        value: `${level * 100}%`,
        label: `${level * 100}%`,
      })),
    ],
    [fitLabel]
  );
  const fontSizeOptions = useMemo(
    () => fontSizes.map((size) => ({ value: String(size), label: String(size) })),
    [fontSizes]
  );

  const apply = (action: FormattingAction) => {
    if (formattingEnabled) onFormat?.(action);
  };
  const applyShape = (action: ShapeFormattingAction) => {
    if (action.type === 'zOrder') {
      if (arrangeEnabled) onShapeFormat?.(action);
      return;
    }
    if (shapeFormattingEnabled) onShapeFormat?.(action);
  };

  const fileButtons = (
    <ToolbarGroup label={t('toolbar.groups.file')}>
      <ToolbarButton
        title={t('toolbar.save')}
        disabled={disabled || !onSave}
        onClick={onSave}
        testId="pptx-save"
      >
        <ToolbarIcon name="save" size={18} />
      </ToolbarButton>
      <ToolbarButton
        title={t('toolbar.exportPng')}
        disabled={disabled || !onExportPng}
        onClick={onExportPng}
        testId="pptx-export-png"
      >
        <ToolbarIcon name="image" size={18} />
      </ToolbarButton>
    </ToolbarGroup>
  );

  const newSlideButtons = (
    <>
      <ToolbarButton
        title={t('toolbar.newSlide')}
        disabled={!slideEnabled}
        onClick={() => onInsertSlide?.()}
        style={singleRow ? undefined : { borderRadius: '4px 0 0 4px' }}
        testId="pptx-new-slide"
      >
        <ToolbarIcon name="newSlide" />
      </ToolbarButton>
      <ToolbarDropdown
        title={t('toolbar.newSlideWithLayout')}
        disabled={!slideEnabled || slideLayouts.length === 0}
        menuWidth={230}
        testId="pptx-new-slide-layout"
        style={{
          minWidth: 20,
          width: 20,
          padding: 0,
          ...(singleRow ? {} : { borderRadius: '0 4px 4px 0' }),
        }}
        trigger={<ToolbarIcon name="chevronDown" size={13} />}
      >
        {(close) => (
          <>
            {slideLayouts.map((layout, index) => (
              <ToolbarMenuItem
                key={layout.partPath ?? `default-${index}`}
                label={layout.label ?? t('toolbar.layoutOption', { number: index + 1 })}
                selected={(layout.partPath ?? null) === (currentLayoutPartPath ?? null)}
                onClick={() => onInsertSlide?.(layout.partPath)}
                close={close}
              />
            ))}
          </>
        )}
      </ToolbarDropdown>
    </>
  );

  const historyButtons = (
    <>
      <ToolbarButton
        title={t('toolbar.undoShortcut')}
        disabled={disabled || !canUndo || !onUndo}
        onClick={onUndo}
        testId="pptx-undo"
      >
        <ToolbarIcon name="undo" />
      </ToolbarButton>
      <ToolbarButton
        title={t('toolbar.redoShortcut')}
        disabled={disabled || !canRedo || !onRedo}
        onClick={onRedo}
        testId="pptx-redo"
      >
        <ToolbarIcon name="redo" />
      </ToolbarButton>
    </>
  );

  const commitZoom = (value: string) => {
    if (value === fitLabel) {
      onZoomChange?.('fit');
      return;
    }
    const percent = Number.parseFloat(value.replace('%', ''));
    if (Number.isFinite(percent) && percent >= 25 && percent <= 400) {
      onZoomChange?.(percent / 100);
    }
  };

  const zoomControl = singleRow ? (
    <ToolbarDropdown
      title={t('toolbar.zoomValue', { value: zoomValue })}
      disabled={disabled || !onZoomChange}
      menuWidth={120}
      testId="pptx-zoom"
      trigger={
        <>
          <span>{zoomValue}</span>
          <ToolbarIcon name="chevronDown" size={13} />
        </>
      }
    >
      {(close) => (
        <>
          {zoomOptions.map((option) => (
            <ToolbarMenuItem
              key={option.value}
              label={option.label}
              selected={option.value === zoomValue}
              onClick={() => commitZoom(option.value)}
              close={close}
            />
          ))}
        </>
      )}
    </ToolbarDropdown>
  ) : (
    <EditableCombobox
      value={zoomValue}
      options={zoomOptions}
      label={t('toolbar.zoomValue', { value: zoomValue })}
      disabled={disabled || !onZoomChange}
      onCommit={commitZoom}
      width={76}
      testId="pptx-zoom"
    />
  );

  const toolButtons = (
    <>
      <ToolbarButton
        title={t('toolbar.selectToolShortcut')}
        active={activeTool === 'select'}
        disabled={!toolEnabled}
        onClick={() => onToolChange?.('select')}
        testId="pptx-tool-select"
      >
        <ToolbarIcon name="select" />
      </ToolbarButton>
      <ToolbarButton
        title={t('toolbar.textBoxTool')}
        active={activeTool === 'textBox'}
        disabled={!toolEnabled}
        onClick={() => onToolChange?.('textBox')}
        testId="pptx-tool-text-box"
      >
        <ToolbarIcon name="textBox" />
      </ToolbarButton>
      <ToolbarButton
        title={t('toolbar.insertImage')}
        disabled={!insertImageEnabled}
        onClick={() => onInsertImage?.()}
        testId="pptx-insert-image"
      >
        <ToolbarIcon name="insertImage" />
      </ToolbarButton>
      <ToolbarDropdown
        title={t('toolbar.shapeTool')}
        active={activeTool.startsWith('shape:')}
        disabled={!toolEnabled}
        menuWidth={228}
        testId="pptx-tool-shape"
        style={{ minWidth: 46, padding: '0 4px' }}
        trigger={
          <>
            <ToolbarIcon name="shape" />
            <ToolbarIcon name="chevronDown" size={11} />
          </>
        }
      >
        {(close) => (
          <>
            <ToolbarMenuLabel first>{t('toolbar.shapeTool')}</ToolbarMenuLabel>
            <div
              style={{
                display: 'grid',
                gridTemplateColumns:
                  'repeat(6, var(--office-menu-button-size, var(--pptx-button-size, 28px)))',
                gap: 4,
                padding: '0 4px 2px',
              }}
            >
              {SHAPE_PRESETS.map((preset) => (
                <ToolbarButton
                  key={preset.geometry}
                  title={t(preset.labelKey)}
                  active={activeTool === `shape:${preset.geometry}`}
                  onClick={() => {
                    onToolChange?.(`shape:${preset.geometry}`);
                    close();
                  }}
                  testId={`pptx-shape-${preset.geometry}`}
                  inMenu
                >
                  <ShapePresetIcon geometry={preset.geometry} />
                </ToolbarButton>
              ))}
            </div>
          </>
        )}
      </ToolbarDropdown>
    </>
  );

  const fontFamilyPicker = (
    <ToolbarDropdown
      title={t('toolbar.fontFamily')}
      disabled={!formattingEnabled}
      menuWidth={210}
      testId="pptx-font-family"
      style={singleRow ? { maxWidth: 160 } : { width: 120, justifyContent: 'space-between' }}
      trigger={
        <>
          <span style={{ overflow: 'hidden', textOverflow: 'ellipsis' }}>
            {currentFormatting.fontFamily ?? t('toolbar.mixed')}
          </span>
          <ToolbarIcon name="chevronDown" size={13} />
        </>
      }
    >
      {(close) => (
        <>
          {fontFamilies.map((font) => (
            <ToolbarMenuItem
              key={font}
              label={font}
              selected={currentFormatting.fontFamily === font}
              onClick={() => apply({ type: 'fontFamily', value: font })}
              close={close}
            />
          ))}
        </>
      )}
    </ToolbarDropdown>
  );

  const fontSizeBox = (
    <EditableCombobox
      value={currentFormatting.fontSize === undefined ? '' : String(fontSize)}
      options={fontSizeOptions}
      label={t('toolbar.fontSize')}
      disabled={!formattingEnabled}
      onCommit={(value) => {
        const size = Number.parseFloat(value);
        if (Number.isFinite(size) && size >= 1 && size <= 400) {
          apply({ type: 'fontSize', value: size });
        }
      }}
      width={singleRow ? 40 : 50}
      inputStyle={{ textAlign: 'center', ...(singleRow ? { padding: 0 } : {}) }}
      chevron={!singleRow}
      testId="pptx-font-size"
    />
  );

  const fontSizeSteps = (
    <>
      <ToolbarButton
        title={t('toolbar.decreaseFontSize')}
        disabled={!formattingEnabled}
        onClick={() =>
          apply({
            type: 'fontSize',
            value: nextFontSize(fontSize, fontSizes, -1),
          })
        }
      >
        <ToolbarIcon name="remove" />
      </ToolbarButton>
      {fontSizeBox}
      <ToolbarButton
        title={t('toolbar.increaseFontSize')}
        disabled={!formattingEnabled}
        onClick={() =>
          apply({
            type: 'fontSize',
            value: nextFontSize(fontSize, fontSizes, 1),
          })
        }
      >
        <ToolbarIcon name="add" />
      </ToolbarButton>
    </>
  );

  const textButtons = (
    <>
      <ToolbarButton
        title={t('toolbar.boldShortcut')}
        active={currentFormatting.bold}
        disabled={!formattingEnabled}
        onClick={() => apply('bold')}
        testId="pptx-bold"
      >
        <ToolbarIcon name="bold" />
      </ToolbarButton>
      <ToolbarButton
        title={t('toolbar.italicShortcut')}
        active={currentFormatting.italic}
        disabled={!formattingEnabled}
        onClick={() => apply('italic')}
        testId="pptx-italic"
      >
        <ToolbarIcon name="italic" />
      </ToolbarButton>
      <ToolbarButton
        title={t('toolbar.underlineShortcut')}
        active={currentFormatting.underline}
        disabled={!formattingEnabled}
        onClick={() => apply('underline')}
        testId="pptx-underline"
      >
        <ToolbarIcon name="underline" />
      </ToolbarButton>
      <ColorPicker
        value={currentFormatting.textColor ?? '#000000'}
        label={t('toolbar.textColor')}
        disabled={!formattingEnabled}
        onChange={(value) => apply({ type: 'textColor', value })}
        testId="pptx-text-color"
      />
      <ColorPicker
        value={currentFormatting.highlight ?? '#fde047'}
        label={t('toolbar.highlightColor')}
        clearLabel={t('toolbar.noHighlight')}
        icon="highlight"
        none={!currentFormatting.highlight}
        disabled={!formattingEnabled}
        onChange={(value) => apply({ type: 'highlight', value })}
        onClear={() => apply({ type: 'highlight', value: null })}
        testId="pptx-highlight-color"
      />
    </>
  );

  const alignButtons = ALIGNMENTS.map((alignment) => (
    <ToolbarButton
      key={alignment.value}
      title={t(alignment.labelKey)}
      active={currentFormatting.align === alignment.value}
      disabled={!formattingEnabled}
      onClick={() => apply({ type: 'align', value: alignment.value })}
      testId={`pptx-align-${alignment.testId}`}
    >
      <ToolbarIcon name={alignment.icon} />
    </ToolbarButton>
  ));

  const currentAlignment =
    ALIGNMENTS.find((alignment) => alignment.value === currentFormatting.align) ?? ALIGNMENTS[0];
  const alignDropdown = (
    <ToolbarDropdown
      title={t('toolbar.groups.alignment')}
      disabled={!formattingEnabled}
      menuWidth={42}
      // A button column's panel is tighter than a list's, as the note toolbar's
      // (p-1) and XLSX's; no shared variable covers it.
      menuStyle={{ padding: 4 }}
      testId="pptx-align"
      trigger={
        <>
          <ToolbarIcon name={currentAlignment.icon} />
          <ToolbarIcon name="chevronDown" size={13} />
        </>
      }
    >
      {(close) => (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
          {ALIGNMENTS.map((alignment) => (
            <ToolbarButton
              key={alignment.value}
              title={t(alignment.labelKey)}
              active={currentFormatting.align === alignment.value}
              onClick={() => {
                apply({ type: 'align', value: alignment.value });
                close();
              }}
              testId={`pptx-align-${alignment.testId}`}
              inMenu
            >
              <ToolbarIcon name={alignment.icon} />
            </ToolbarButton>
          ))}
          {/* Google Slides' align popover holds the text box's vertical
              alignment too, after the horizontal one. */}
          <ToolbarSeparator style={{ width: 'auto', height: 1, margin: 0 }} />
          {VERTICAL_ALIGNMENTS.map((alignment) => (
            <ToolbarButton
              key={alignment.value}
              title={t(alignment.labelKey)}
              active={currentParagraph?.anchor === alignment.value}
              disabled={!currentParagraph?.anchor}
              onClick={() => {
                apply({ type: 'verticalAlign', value: alignment.value });
                close();
              }}
              testId={`pptx-align-${alignment.testId}`}
              inMenu
            >
              <ToolbarIcon name={alignment.icon} />
            </ToolbarButton>
          ))}
        </div>
      )}
    </ToolbarDropdown>
  );

  const lineSpacingDropdown = (
    <ToolbarDropdown
      title={t('toolbar.lineSpacing.label')}
      disabled={!formattingEnabled}
      menuWidth={240}
      testId="pptx-line-spacing"
      trigger={<ToolbarIcon name="lineSpacing" />}
    >
      {(close) => (
        <>
          {LINE_SPACINGS.map((spacing) => (
            <ToolbarMenuItem
              key={spacing.value}
              label={'labelKey' in spacing ? t(spacing.labelKey) : spacing.label}
              selected={currentParagraph?.lineSpacing === spacing.value}
              onClick={() => apply({ type: 'lineSpacing', value: spacing.value })}
              close={close}
            />
          ))}
          <ToolbarMenuSeparator />
          <ToolbarMenuItem
            label={t(
              currentParagraph?.spaceBefore
                ? 'toolbar.lineSpacing.removeSpaceBefore'
                : 'toolbar.lineSpacing.addSpaceBefore'
            )}
            onClick={() => apply({ type: 'spaceBefore' })}
            close={close}
          />
          <ToolbarMenuItem
            label={t(
              currentParagraph?.spaceAfter
                ? 'toolbar.lineSpacing.removeSpaceAfter'
                : 'toolbar.lineSpacing.addSpaceAfter'
            )}
            onClick={() => apply({ type: 'spaceAfter' })}
            close={close}
          />
        </>
      )}
    </ToolbarDropdown>
  );

  const listButton = (kind: ListKind) => {
    const bullets = kind === 'bullet';
    const presets = bullets ? BULLET_PRESETS : NUMBER_PRESETS;
    const label = t(bullets ? 'toolbar.bulletedList' : 'toolbar.numberedList');
    return (
      <>
        <ToolbarButton
          title={label}
          active={currentParagraph?.list === kind}
          disabled={!formattingEnabled}
          onClick={() => apply({ type: 'list', kind })}
          testId={`pptx-${bullets ? 'bulleted' : 'numbered'}-list`}
        >
          <ToolbarIcon name={bullets ? 'bulletedList' : 'numberedList'} />
        </ToolbarButton>
        <ToolbarDropdown
          title={t(bullets ? 'toolbar.bulletStyles' : 'toolbar.numberingStyles')}
          disabled={!formattingEnabled}
          menuWidth={150}
          testId={`pptx-${bullets ? 'bulleted' : 'numbered'}-list-styles`}
          style={{ minWidth: 16, width: 16, padding: 0 }}
          trigger={<ToolbarIcon name="chevronDown" size={13} />}
        >
          {(close) => (
            <>
              {presets.map((preset) => (
                <ToolbarMenuItem
                  key={preset.id}
                  label={presetLabel(preset.id)}
                  selected={
                    currentParagraph?.list === kind && currentParagraph.listPreset === preset.id
                  }
                  onClick={() => apply({ type: 'list', kind, preset: preset.id })}
                  close={close}
                />
              ))}
            </>
          )}
        </ToolbarDropdown>
      </>
    );
  };

  const listButtons = (
    <>
      {listButton('bullet')}
      {listButton('number')}
      <ToolbarButton
        title={t('toolbar.decreaseIndent')}
        disabled={!formattingEnabled || !currentParagraph?.canOutdent}
        onClick={() => apply({ type: 'indent', delta: -1 })}
        testId="pptx-decrease-indent"
      >
        <ToolbarIcon name="indentDecrease" />
      </ToolbarButton>
      <ToolbarButton
        title={t('toolbar.increaseIndent')}
        disabled={!formattingEnabled || currentParagraph?.canIndent === false}
        onClick={() => apply({ type: 'indent', delta: 1 })}
        testId="pptx-increase-indent"
      >
        <ToolbarIcon name="indentIncrease" />
      </ToolbarButton>
    </>
  );

  const clearButton = (
    <ToolbarButton
      title={t('toolbar.clearFormatting')}
      disabled={!formattingEnabled}
      onClick={() => apply('clearFormatting')}
      testId="pptx-clear-formatting"
    >
      <ToolbarIcon name="clearFormatting" />
    </ToolbarButton>
  );

  const shapeControls = (
    <>
      <ColorPicker
        value={currentShapeFormatting.fillColor ?? '#d9eaf7'}
        label={t('toolbar.fillColor')}
        clearLabel={t('toolbar.noFill')}
        icon="fillColor"
        none={!currentShapeFormatting.fillColor}
        disabled={!shapeFormattingEnabled}
        onChange={(value) => applyShape({ type: 'fillColor', value })}
        onClear={() => applyShape({ type: 'fillColor', value: null })}
        testId="pptx-shape-fill"
      />
      <ColorPicker
        value={currentShapeFormatting.strokeColor ?? '#202124'}
        label={t('toolbar.borderColor')}
        clearLabel={t('toolbar.noBorder')}
        icon="borderColor"
        none={!currentShapeFormatting.strokeColor}
        disabled={!shapeFormattingEnabled}
        onChange={(value) => applyShape({ type: 'strokeColor', value })}
        onClear={() => applyShape({ type: 'strokeColor', value: null })}
        testId="pptx-shape-border-color"
      />
      <ToolbarDropdown
        title={t('toolbar.borderWidth')}
        disabled={!shapeFormattingEnabled}
        menuWidth={170}
        testId="pptx-shape-border-width"
        trigger={<ToolbarIcon name="borderWidth" />}
      >
        {(close) => (
          <>
            <ToolbarMenuItem
              label={t('toolbar.noBorder')}
              icon={<span aria-hidden="true" style={{ width: 18 }} />}
              selected={currentShapeFormatting.strokeWidthPt === null}
              onClick={() => applyShape({ type: 'strokeWidth', value: null })}
              close={close}
            />
            {BORDER_WIDTHS.map((width) => (
              <ToolbarMenuItem
                key={width}
                label={t('toolbar.borderWidthValue', { width })}
                selected={currentShapeFormatting.strokeWidthPt === width}
                icon={
                  <span
                    aria-hidden="true"
                    style={{
                      width: 18,
                      borderTop: `${Math.min(width, 5)}px solid currentColor`,
                    }}
                  />
                }
                onClick={() => applyShape({ type: 'strokeWidth', value: width })}
                close={close}
              />
            ))}
          </>
        )}
      </ToolbarDropdown>
      {shapeAdjustment ? (
        <EditableCombobox
          value={`${Math.round(shapeAdjustment[1] * 100)}%`}
          options={(roundRectAdjustment ? CORNER_RADIUS_OPTIONS : SHAPE_ADJUSTMENT_OPTIONS).map(
            (value) => ({
              value: String(value),
              label: `${value}%`,
            })
          )}
          label={t(roundRectAdjustment ? 'toolbar.cornerRadius' : 'toolbar.shapeAdjustment')}
          disabled={!shapeFormattingEnabled}
          onCommit={(value) => {
            const percent = Number.parseFloat(value.replace('%', ''));
            if (Number.isFinite(percent)) {
              const maximum = roundRectAdjustment ? 50 : 100;
              applyShape({
                type: 'adjust',
                name: shapeAdjustment[0],
                value: Math.max(0, Math.min(maximum, percent)) / 100,
              });
            }
          }}
          width={68}
          inputStyle={{ textAlign: 'center' }}
          testId={roundRectAdjustment ? 'pptx-shape-corner-radius' : 'pptx-shape-adjustment'}
        />
      ) : null}
    </>
  );

  const arrangeDropdown = (
    <ToolbarDropdown
      title={t('toolbar.arrange')}
      disabled={!arrangeEnabled}
      menuWidth={190}
      testId="pptx-shape-arrange"
      trigger={<ToolbarIcon name="bringToFront" />}
    >
      {(close) => (
        <>
          <ToolbarMenuItem
            label={t('toolbar.bringForward')}
            icon={<ToolbarIcon name="bringForward" size={16} />}
            onClick={() => applyShape({ type: 'zOrder', value: 'forward' })}
            close={close}
          />
          <ToolbarMenuItem
            label={t('toolbar.sendBackward')}
            icon={<ToolbarIcon name="sendBackward" size={16} />}
            onClick={() => applyShape({ type: 'zOrder', value: 'backward' })}
            close={close}
          />
          <ToolbarMenuItem
            label={t('toolbar.bringToFront')}
            icon={<ToolbarIcon name="bringToFront" size={16} />}
            onClick={() => applyShape({ type: 'zOrder', value: 'front' })}
            close={close}
          />
          <ToolbarMenuItem
            label={t('toolbar.sendToBack')}
            icon={<ToolbarIcon name="sendToBack" size={16} />}
            onClick={() => applyShape({ type: 'zOrder', value: 'back' })}
            close={close}
          />
        </>
      )}
    </ToolbarDropdown>
  );

  if (singleRow) {
    const textActive = formattingEnabled;
    const shapeActive = shapeFormattingEnabled;
    const groups: Array<{ key: string; label: string; node: ReactNode } | false> = [
      { key: 'new-slide', label: t('toolbar.groups.slides'), node: newSlideButtons },
      { key: 'history', label: t('toolbar.groups.history'), node: historyButtons },
      showZoomControl && { key: 'zoom', label: t('toolbar.groups.zoom'), node: zoomControl },
      { key: 'tools', label: t('toolbar.groups.tools'), node: toolButtons },
      textActive &&
        (showFontPicker || showFontSizePicker) && {
          key: 'font',
          label: t('toolbar.groups.font'),
          node: (
            <>
              {showFontPicker && fontFamilyPicker}
              {showFontSizePicker && fontSizeBox}
            </>
          ),
        },
      textActive && { key: 'text', label: t('toolbar.groups.text'), node: textButtons },
      textActive && {
        key: 'align',
        label: t('toolbar.groups.alignment'),
        node: (
          <>
            {alignDropdown}
            {lineSpacingDropdown}
          </>
        ),
      },
      textActive && { key: 'lists', label: t('toolbar.groups.lists'), node: listButtons },
      textActive && { key: 'clear', label: t('toolbar.clearFormatting'), node: clearButton },
      shapeActive && { key: 'shape', label: t('toolbar.groups.shape'), node: shapeControls },
      Boolean(children) && { key: 'custom', label: t('toolbar.more'), node: children },
    ];
    return (
      <div
        ref={rootRef}
        className={className}
        role="toolbar"
        aria-label={t('toolbar.label')}
        data-testid="pptx-formatting-toolbar"
        data-layout="single-row"
        style={{
          position: 'relative',
          display: 'flex',
          minWidth: 0,
          height: '100%',
          color: toolbarColors.text,
          ...style,
        }}
      >
        <div
          ref={rowRef}
          data-testid="pptx-toolbar-row"
          style={{
            display: 'flex',
            alignItems: 'center',
            flex: 1,
            minWidth: 0,
            height: '100%',
            overflowX: 'auto',
            overflowY: 'hidden',
            scrollbarWidth: 'none',
          }}
        >
          {groups
            .filter((group) => group !== false)
            .map((group, index) => (
              <Fragment key={group.key}>
                {index > 0 && <ToolbarSeparator style={{ height: 28, margin: '0 6px' }} />}
                <ToolbarGroup label={group.label}>{group.node}</ToolbarGroup>
              </Fragment>
            ))}
        </div>
        <EdgeFade side="start" visible={rowEdges.start} />
        <EdgeFade side="end" visible={rowEdges.end} />
      </div>
    );
  }

  const sections: ToolbarSection[] = [
    { key: 'file', width: 72, node: fileButtons },
    {
      key: 'new-slide',
      width: 59,
      node: <ToolbarGroup label={t('toolbar.groups.slides')}>{newSlideButtons}</ToolbarGroup>,
    },
    {
      key: 'history',
      width: 75,
      node: (
        <>
          <ToolbarSeparator />
          <ToolbarGroup label={t('toolbar.groups.history')}>{historyButtons}</ToolbarGroup>
        </>
      ),
    },
    {
      key: 'zoom',
      width: 82,
      node: <ToolbarGroup label={t('toolbar.groups.zoom')}>{zoomControl}</ToolbarGroup>,
    },
    {
      key: 'tools',
      width: 169,
      node: (
        <>
          <ToolbarSeparator />
          <ToolbarGroup label={t('toolbar.groups.tools')}>{toolButtons}</ToolbarGroup>
        </>
      ),
    },
    {
      key: 'font-family',
      width: 137,
      node: (
        <>
          <ToolbarSeparator />
          <ToolbarGroup label={t('toolbar.groups.font')}>{fontFamilyPicker}</ToolbarGroup>
        </>
      ),
    },
    {
      key: 'font-size',
      width: 116,
      node: <ToolbarGroup label={t('toolbar.groups.font')}>{fontSizeSteps}</ToolbarGroup>,
    },
    {
      key: 'text',
      width: 166,
      node: (
        <>
          <ToolbarSeparator />
          <ToolbarGroup label={t('toolbar.groups.text')}>{textButtons}</ToolbarGroup>
        </>
      ),
    },
    {
      key: 'align',
      width: 159,
      node: (
        <>
          <ToolbarSeparator />
          <ToolbarGroup label={t('toolbar.groups.alignment')}>
            {alignButtons}
            {lineSpacingDropdown}
          </ToolbarGroup>
        </>
      ),
    },
    {
      key: 'shape-formatting',
      width: (shapeAdjustment ? 244 : 172) + 40,
      node: (
        <>
          <ToolbarSeparator />
          <ToolbarGroup label={t('toolbar.groups.shape')}>
            {shapeControls}
            {arrangeDropdown}
          </ToolbarGroup>
        </>
      ),
    },
  ];

  // The rail keeps its first sections inline; lists and clear formatting
  // overflow into More before them.
  sections.push(
    {
      key: 'lists',
      width: 145,
      node: (
        <>
          <ToolbarSeparator />
          <ToolbarGroup label={t('toolbar.groups.lists')}>{listButtons}</ToolbarGroup>
        </>
      ),
    },
    {
      key: 'clear',
      width: 44,
      node: (
        <>
          <ToolbarSeparator />
          <ToolbarGroup label={t('toolbar.clearFormatting')}>{clearButton}</ToolbarGroup>
        </>
      ),
    }
  );
  if (children) sections.push({ key: 'custom', width: 40, node: children });

  const availableWidth = Math.max(0, rootWidth - 48);
  let usedWidth = 0;
  let visibleCount = sections.length;
  for (let index = 0; index < sections.length; index++) {
    usedWidth += sections[index].width;
    if (usedWidth > availableWidth) {
      visibleCount = index;
      break;
    }
  }
  const visibleSections = sections.slice(0, visibleCount);
  const overflowSections = sections.slice(visibleCount);

  return (
    <div
      ref={rootRef}
      className={className}
      role="toolbar"
      aria-label={t('toolbar.label')}
      data-testid="pptx-formatting-toolbar"
      style={{
        display: 'flex',
        alignItems: 'center',
        minWidth: 0,
        ...(singleRow
          ? { height: '100%' }
          : {
              minHeight: 36,
              margin: '0 8px 5px',
              padding: '4px 7px',
              borderRadius: 18,
              background: toolbarColors.rail,
            }),
        color: toolbarColors.text,
        overflow: 'hidden',
        boxSizing: 'border-box',
        ...style,
      }}
    >
      {visibleSections.map((section) => (
        <span key={section.key} style={{ display: 'contents' }}>
          {section.node}
        </span>
      ))}
      <span style={{ marginLeft: 'auto', flex: '0 0 auto' }}>
        <ToolbarDropdown
          title={t('toolbar.more')}
          disabled={overflowSections.length === 0}
          menuWidth={390}
          testId="pptx-toolbar-more"
          trigger={<ToolbarIcon name="more" />}
        >
          {() => (
            <div
              style={{
                display: 'flex',
                alignItems: 'center',
                flexWrap: 'wrap',
                gap: 2,
              }}
            >
              {overflowSections.map((section) => (
                <span key={section.key} style={{ display: 'contents' }}>
                  {section.node}
                </span>
              ))}
            </div>
          )}
        </ToolbarDropdown>
      </span>
    </div>
  );
}

export { Toolbar as PptxToolbar };

function ShapePresetIcon({ geometry }: { geometry: PptxShapePreset }) {
  const path = {
    rect: 'M3 5h18v14H3Z',
    roundRect: 'M7 5h10a4 4 0 0 1 4 4v6a4 4 0 0 1-4 4H7a4 4 0 0 1-4-4V9a4 4 0 0 1 4-4Z',
    ellipse: 'M3 12a9 7 0 1 0 18 0 9 7 0 1 0-18 0',
    triangle: 'm12 4 9 16H3Z',
    rtTriangle: 'M4 4v16h16Z',
    diamond: 'm12 3 9 9-9 9-9-9Z',
    parallelogram: 'M7 5h14l-4 14H3Z',
    trapezoid: 'M7 5h10l4 14H3Z',
    pentagon: 'm12 3 9 7-4 11H7L3 10Z',
    hexagon: 'm7 4 10 0 5 8-5 8H7l-5-8Z',
    octagon: 'm7 3 10 0 4 4v10l-4 4H7l-4-4V7Z',
    star5: 'm12 2.5 2.8 6 6.5.6-5 4.3 1.6 6.4-5.7-3.4-5.7 3.4 1.6-6.4-5-4.3 6.5-.6Z',
    rightArrow: 'M3 8h11V4l7 8-7 8v-4H3Z',
    leftArrow: 'm21 8H10V4l-7 8 7 8v-4h11Z',
    upArrow: 'M8 21V10H4l8-7 8 7h-4v11Z',
    downArrow: 'M8 3v11H4l8 7 8-7h-4V3Z',
    chevron: 'M4 4h9l7 8-7 8H4l7-8Z',
  }[geometry];
  return (
    <svg
      width="30"
      height="24"
      viewBox="0 0 24 24"
      fill="rgba(60, 64, 67, 0.08)"
      stroke="currentColor"
      strokeWidth="1.4"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d={path} />
    </svg>
  );
}
