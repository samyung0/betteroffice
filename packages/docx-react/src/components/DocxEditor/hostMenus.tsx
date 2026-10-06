/**
 * The editor's menus as data, for a host that draws its own menu bar
 * (`DocxEditor`'s `onMenus`): ids, labels in the editor's locale, shortcuts and
 * state. The host calls `run(id, value)` for a click; nothing here renders.
 */

import { useEffect, useMemo, useRef } from 'react';
import { useTranslation } from '../../i18n';
import { useEditorToolbar } from '../EditorToolbarContext';
import { paragraphStyleOptions } from '../ui/StylePicker';
import type { TextContextAction } from '../TextContextMenu';

export type HostMenuEntry =
  | {
      kind: 'item';
      id: string;
      label: string;
      /** Running it changes the document (or saves it): a read-only editor disables it. */
      edits: boolean;
      shortcut?: string;
      checked?: boolean;
      /** One of an exclusive set (drawn as a radio item); needs `checked`. */
      radio?: boolean;
      disabled?: boolean;
    }
  | { kind: 'separator' }
  | { kind: 'submenu'; id: string; label: string; items: HostMenuEntry[] }
  /** A table size picker; `run` gets the size as "<rows>x<cols>". */
  | { kind: 'grid'; id: string };

export interface HostMenu {
  id: string;
  label: string;
  items: HostMenuEntry[];
}

export interface DocxMenuModel {
  menus: HostMenu[];
  /** `file` comes with Insert › Image when the host picked the image. */
  run: (id: string, value?: string, file?: File) => void;
}

/** Commands that live outside the toolbar context. */
export interface HostMenuActions {
  onEditAction: (action: TextContextAction) => void;
  onFindReplace?: () => void;
  onAddComment: () => void;
  onInsertImageFile: (file: File) => void;
  showComments: boolean;
  onToggleComments: () => void;
  showRuler: boolean;
  onToggleRuler: () => void;
}

const ZOOMS = [50, 75, 90, 100, 125, 150, 200];
/** Format › Paragraph styles lists this many; the toolbar's style picker has all. */
export const MAX_MENU_STYLES = 40;
const LINE_SPACINGS = [
  { twips: 240, key: 'lineSpacing.single' as const },
  { twips: 276, label: '1.15' },
  { twips: 360, label: '1.5' },
  { twips: 480, key: 'lineSpacing.double' as const },
];

function shortcutFormatter() {
  const mac = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform);
  // Mac lists modifiers as ⌃⌥⇧⌘; elsewhere Ctrl+Alt+Shift+key.
  const order = ['Ctrl', 'Alt', 'Shift', 'Mod'];
  const mark: Record<string, string> = { Ctrl: '⌃', Alt: '⌥', Shift: '⇧', Mod: '⌘' };
  return (keys: string) => {
    const parts = keys.split('+');
    const key = parts.pop() ?? '';
    const named = mac ? parts : parts.map((part) => (part === 'Mod' ? 'Ctrl' : part));
    named.sort((a, b) => order.indexOf(a) - order.indexOf(b));
    if (mac) return named.map((part) => mark[part]).join('') + key;
    return [...named, key].join('+');
  };
}

function findItem(
  menus: { items: HostMenuEntry[] }[],
  id: string
): Extract<HostMenuEntry, { kind: 'item' }> | undefined {
  for (const menu of menus)
    for (const entry of menu.items) {
      if (entry.kind === 'item' && entry.id === id) return entry;
      if (entry.kind === 'submenu') {
        const found = findItem([entry], id);
        if (found) return found;
      }
    }
  return undefined;
}

/** Reports the menus while mounted inside `EditorToolbar`; renders nothing. */
export function HostMenus({
  onMenus,
  actions,
}: {
  onMenus: (model: DocxMenuModel | null) => void;
  actions: HostMenuActions;
}) {
  const { t } = useTranslation();
  const ctx = useEditorToolbar();
  const latest = useRef({ actions, ctx });
  latest.current = { actions, ctx };
  const formatting = ctx.currentFormatting ?? {};
  const disabled = ctx.disabled ?? false;
  const table = ctx.tableContext;

  const menus = useMemo((): HostMenu[] => {
    const key = shortcutFormatter();
    const item = (
      id: string,
      label: string,
      extra: Omit<Extract<HostMenuEntry, { kind: 'item' }>, 'kind' | 'id' | 'label'>
    ): HostMenuEntry => ({
      kind: 'item',
      id,
      label,
      ...extra,
      disabled: (disabled && extra.edits) || extra.disabled,
    });
    const separator: HostMenuEntry = { kind: 'separator' };
    const submenu = (id: string, label: string, items: HostMenuEntry[]): HostMenuEntry => ({
      kind: 'submenu',
      id,
      label,
      items,
    });

    const file: HostMenuEntry[] = [
      ...(ctx.onSave
        ? [item('save', t('toolbar.save'), { edits: true, shortcut: key('Mod+S') })]
        : []),
      ...(ctx.onPageSetup ? [item('page-setup', t('toolbar.pageSetup'), { edits: true })] : []),
    ];
    const edit: HostMenuEntry[] = [
      item('undo', t('formattingBar.undo'), {
        edits: true,
        shortcut: key('Mod+Z'),
        disabled: !ctx.canUndo,
      }),
      item('redo', t('formattingBar.redo'), {
        edits: true,
        shortcut: key('Mod+Y'),
        disabled: !ctx.canRedo,
      }),
      separator,
      item('select-all', t('hostMenus.selectAll'), { edits: false, shortcut: key('Mod+A') }),
      item('delete', t('hostMenus.delete'), { edits: true }),
      ...(actions.onFindReplace
        ? [
            separator,
            item('find-replace', t('hostMenus.findReplace'), {
              edits: false,
              shortcut: key('Mod+H'),
            }),
          ]
        : []),
    ];
    const zoom = Math.round((ctx.zoom ?? 1) * 100);
    // Google Docs' order: ruler, outline, then comments.
    const view: HostMenuEntry[] = [
      item('show-ruler', t('hostMenus.showRuler'), { edits: false, checked: actions.showRuler }),
      ...(ctx.onToggleOutline
        ? [
            item('show-outline', t('hostMenus.showOutline'), {
              edits: false,
              checked: !!ctx.outlineOpen,
            }),
          ]
        : []),
      item('show-comments', t('hostMenus.showComments'), {
        edits: false,
        checked: actions.showComments,
      }),
      ...(ctx.onZoomChange
        ? [
            separator,
            submenu(
              'zoom',
              t('hostMenus.zoom'),
              ZOOMS.map((value) =>
                item(`zoom:${value}`, `${value}%`, { edits: false, checked: value === zoom })
              )
            ),
          ]
        : []),
    ];
    const breaks = [
      ...(ctx.onInsertPageBreak
        ? [item('insert-page-break', t('toolbar.pageBreak'), { edits: true })]
        : []),
      ...(ctx.onInsertSectionBreakNextPage
        ? [item('insert-section-next', t('toolbar.sectionBreakNextPage'), { edits: true })]
        : []),
      ...(ctx.onInsertSectionBreakContinuous
        ? [item('insert-section-continuous', t('toolbar.sectionBreakContinuous'), { edits: true })]
        : []),
    ];
    const insert: HostMenuEntry[] = [
      ...(ctx.onInsertImage ? [item('insert-image', t('toolbar.image'), { edits: true })] : []),
      ...(ctx.onInsertTable
        ? [submenu('insert-table', t('toolbar.table'), [{ kind: 'grid', id: 'insert-table' }])]
        : []),
      item('insert-link', t('hostMenus.link'), { edits: true, shortcut: key('Mod+K') }),
      item('insert-comment', t('hostMenus.comment'), { edits: true }),
      ...(ctx.onWatermark
        ? [item('insert-watermark', t('toolbar.watermark'), { edits: true })]
        : []),
      ...(breaks.length ? [separator, submenu('insert-break', t('toolbar.break'), breaks)] : []),
      ...(ctx.onInsertTOC
        ? [item('insert-toc', t('toolbar.tableOfContents'), { edits: true })]
        : []),
      ...(ctx.onUpdateTOC
        ? [item('update-toc', t('hostMenus.updateTableOfContents'), { edits: true })]
        : []),
    ];
    const spacing = formatting.lineSpacing ?? 240;
    const format: HostMenuEntry[] = [
      submenu('format-text', t('hostMenus.text'), [
        item('bold', t('formattingBar.bold'), { edits: true, shortcut: key('Mod+B') }),
        item('italic', t('formattingBar.italic'), { edits: true, shortcut: key('Mod+I') }),
        item('underline', t('formattingBar.underline'), { edits: true, shortcut: key('Mod+U') }),
        item('strikethrough', t('formattingBar.strikethrough'), { edits: true }),
        item('superscript', t('formattingBar.superscript'), {
          edits: true,
          shortcut: key('Mod+Shift+='),
        }),
        item('subscript', t('formattingBar.subscript'), { edits: true, shortcut: key('Mod+=') }),
      ]),
      submenu(
        'format-styles',
        t('hostMenus.paragraphStyles'),
        paragraphStyleOptions(ctx.documentStyles)
          .slice(0, MAX_MENU_STYLES)
          .map((style) =>
            item(`style:${style.styleId}`, style.nameKey ? t(style.nameKey) : style.name, {
              edits: true,
              checked: (formatting.styleId || 'Normal') === style.styleId,
            })
          )
      ),
      submenu('format-align', t('hostMenus.alignIndent'), [
        item('align:left', t('hostMenus.left'), { edits: true, shortcut: key('Mod+L') }),
        item('align:center', t('hostMenus.center'), { edits: true, shortcut: key('Mod+E') }),
        item('align:right', t('hostMenus.right'), { edits: true, shortcut: key('Mod+R') }),
        item('align:both', t('hostMenus.justify'), { edits: true, shortcut: key('Mod+J') }),
        separator,
        item('indent', t('hostMenus.increaseIndent'), { edits: true }),
        item('outdent', t('hostMenus.decreaseIndent'), { edits: true }),
      ]),
      submenu(
        'format-spacing',
        t('lineSpacing.label'),
        LINE_SPACINGS.map((option) =>
          item(`spacing:${option.twips}`, option.key ? t(option.key) : option.label, {
            edits: true,
            checked: option.twips === spacing,
          })
        )
      ),
      submenu('format-lists', t('hostMenus.bulletsNumbering'), [
        item('bulletList', t('hostMenus.bulletedList'), { edits: true }),
        item('numberedList', t('hostMenus.numberedList'), { edits: true }),
      ]),
      submenu('format-direction', t('hostMenus.textDirection'), [
        item('ltr', t('hostMenus.leftToRight'), { edits: true, checked: !formatting.bidi }),
        item('rtl', t('hostMenus.rightToLeft'), { edits: true, checked: !!formatting.bidi }),
      ]),
      ...(table?.isInTable && ctx.onTableAction
        ? [
            separator,
            // Google Docs' Format › Table, with Word's alignment, wrap and auto-fit.
            submenu('format-table', t('hostMenus.table'), [
              submenu('table-vertical-align', t('tableAdvanced.verticalAlignment'), [
                item('table-valign:top', t('tableAdvanced.top'), {
                  edits: true,
                  radio: true,
                  checked: (table.verticalAlign ?? 'top') === 'top',
                }),
                item('table-valign:center', t('tableAdvanced.middle'), {
                  edits: true,
                  radio: true,
                  checked: table.verticalAlign === 'center',
                }),
                item('table-valign:bottom', t('tableAdvanced.bottom'), {
                  edits: true,
                  radio: true,
                  checked: table.verticalAlign === 'bottom',
                }),
              ]),
              submenu('table-alignment', t('tableAdvanced.tableAlignment'), [
                item('table-align:left', t('hostMenus.left'), {
                  edits: true,
                  radio: true,
                  checked: (table.tableAlignment ?? 'left') === 'left',
                }),
                item('table-align:center', t('hostMenus.center'), {
                  edits: true,
                  radio: true,
                  checked: table.tableAlignment === 'center',
                }),
                item('table-align:right', t('hostMenus.right'), {
                  edits: true,
                  radio: true,
                  checked: table.tableAlignment === 'right',
                }),
              ]),
              separator,
              item('table-header-row', t('tableAdvanced.pinHeaderRow'), {
                edits: true,
                checked: !!table.headerRow,
              }),
              item('table-wrap-text', t('tableAdvanced.wrapText'), {
                edits: true,
                checked: table.wrapText !== false,
              }),
              item('table-distribute', t('tableAdvanced.distributeColumns'), { edits: true }),
              item('table-autofit', t('tableAdvanced.autoFit'), { edits: true }),
              separator,
              item('table-properties', t('hostMenus.tableProperties'), { edits: true }),
            ]),
          ]
        : []),
      ...(ctx.imageContext && ctx.onOpenImageProperties
        ? [separator, item('image-options', t('hostMenus.imageOptions'), { edits: true })]
        : []),
      separator,
      item('clearFormatting', t('formattingBar.clearFormatting'), { edits: true }),
    ];
    return [
      { id: 'file', label: t('toolbar.file'), items: file },
      { id: 'edit', label: t('hostMenus.edit'), items: edit },
      { id: 'view', label: t('hostMenus.view'), items: view },
      { id: 'insert', label: t('toolbar.insert'), items: insert },
      { id: 'format', label: t('toolbar.format'), items: format },
    ];
  }, [
    t,
    ctx.onSave,
    ctx.onPageSetup,
    ctx.canUndo,
    ctx.canRedo,
    ctx.zoom,
    ctx.onZoomChange,
    ctx.outlineOpen,
    ctx.onToggleOutline,
    ctx.onInsertImage,
    ctx.onInsertTable,
    ctx.onWatermark,
    ctx.onInsertPageBreak,
    ctx.onInsertSectionBreakNextPage,
    ctx.onInsertSectionBreakContinuous,
    ctx.onInsertTOC,
    ctx.onUpdateTOC,
    ctx.documentStyles,
    table?.isInTable,
    table?.verticalAlign,
    table?.tableAlignment,
    table?.headerRow,
    table?.wrapText,
    ctx.onTableAction,
    ctx.imageContext,
    ctx.onOpenImageProperties,
    actions.onFindReplace,
    actions.showComments,
    actions.showRuler,
    formatting.lineSpacing,
    formatting.styleId,
    formatting.bidi,
    disabled,
  ]);

  const latestMenus = useRef(menus);
  latestMenus.current = menus;
  const run = useMemo(
    () => (id: string, value?: string, file?: File) => {
      // A disabled item does nothing, whoever sends its id.
      if (findItem(latestMenus.current, id)?.disabled) return;
      const { actions: act, ctx: c } = latest.current;
      const [command, argument] = id.split(':');
      // Formatting returns focus to the page; dialogs and pickers keep theirs.
      const format: NonNullable<typeof c.onFormat> = (action) => {
        c.onFormat?.(action);
        requestAnimationFrame(() => c.onRefocusEditor?.());
      };
      const tableAction: NonNullable<typeof c.onTableAction> = (action) => {
        c.onTableAction?.(action);
        requestAnimationFrame(() => c.onRefocusEditor?.());
      };
      switch (command) {
        case 'save':
          void c.onSave?.();
          return;
        case 'page-setup':
          c.onPageSetup?.();
          return;
        case 'undo':
          c.onUndo?.();
          requestAnimationFrame(() => c.onRefocusEditor?.());
          return;
        case 'redo':
          c.onRedo?.();
          requestAnimationFrame(() => c.onRefocusEditor?.());
          return;
        case 'select-all':
          act.onEditAction('selectAll');
          return;
        case 'delete':
          act.onEditAction('delete');
          return;
        case 'find-replace':
          act.onFindReplace?.();
          return;
        case 'show-outline':
          c.onToggleOutline?.();
          return;
        case 'show-comments':
          act.onToggleComments();
          return;
        case 'show-ruler':
          act.onToggleRuler();
          return;
        case 'zoom':
          c.onZoomChange?.(Number(argument) / 100);
          return;
        case 'insert-image':
          if (file) act.onInsertImageFile(file);
          else c.onInsertImage?.();
          return;
        case 'insert-table': {
          const [rows, cols] = (value ?? '').split('x').map(Number);
          if (rows > 0 && cols > 0) c.onInsertTable?.(rows, cols);
          return;
        }
        case 'insert-link':
          format('insertLink');
          return;
        case 'insert-comment':
          act.onAddComment();
          return;
        case 'insert-watermark':
          c.onWatermark?.();
          return;
        case 'insert-page-break':
          c.onInsertPageBreak?.();
          return;
        case 'insert-section-next':
          c.onInsertSectionBreakNextPage?.();
          return;
        case 'insert-section-continuous':
          c.onInsertSectionBreakContinuous?.();
          return;
        case 'insert-toc':
          c.onInsertTOC?.();
          return;
        case 'update-toc':
          c.onUpdateTOC?.();
          return;
        case 'style':
          format({ type: 'applyStyle', value: argument });
          return;
        case 'align':
          format({
            type: 'alignment',
            value: argument as 'left' | 'center' | 'right' | 'both',
          });
          return;
        case 'spacing':
          format({ type: 'lineSpacing', value: Number(argument) });
          return;
        case 'ltr':
          format('setLtr');
          return;
        case 'rtl':
          format('setRtl');
          return;
        case 'table-properties':
          c.onTableAction?.({ type: 'openTableProperties' });
          return;
        case 'table-valign':
          tableAction({
            type: 'cellVerticalAlign',
            align: argument as 'top' | 'center' | 'bottom',
          });
          return;
        case 'table-align':
          tableAction({
            type: 'tableAlignment',
            alignment: argument as 'left' | 'center' | 'right',
          });
          return;
        case 'table-header-row':
          tableAction({ type: 'pinHeaderRow', pinned: !c.tableContext?.headerRow });
          return;
        case 'table-wrap-text':
          tableAction({ type: 'wrapText', wrap: c.tableContext?.wrapText === false });
          return;
        case 'table-distribute':
          tableAction({ type: 'distributeColumns' });
          return;
        case 'table-autofit':
          tableAction({ type: 'autoFitContents' });
          return;
        case 'image-options':
          c.onOpenImageProperties?.();
          return;
        case 'bold':
        case 'italic':
        case 'underline':
        case 'strikethrough':
        case 'superscript':
        case 'subscript':
        case 'clearFormatting':
        case 'bulletList':
        case 'numberedList':
        case 'indent':
        case 'outdent':
          format(command);
          return;
      }
    },
    []
  );

  useEffect(() => {
    onMenus({ menus, run });
  }, [menus, onMenus, run]);
  useEffect(() => () => onMenus(null), [onMenus]);
  return null;
}
