/**
 * `@betteroffice/xlsx-react` — the React chrome for the xlsx editor. Framework
 * glue only; all compute lives in `@betteroffice/xlsx`.
 */

export { XlsxEditor } from './XlsxEditor';
export type {
  XlsxEditorProps,
  XlsxEditorApi,
  XlsxEditorCollaborationOptions,
} from './XlsxEditor';
export { EditorToolbar } from './components/EditorToolbar';
export {
  XLSX_COMMANDS,
  ZOOM_PERCENTS,
  isXlsxCommand,
  type XlsxCommand,
  type XlsxCommandState,
} from './commands';
export {
  Toolbar,
  XlsxToolbar,
  type ToolbarProps,
  type SelectionFormatting,
  type SelectionShape,
  type FormattingAction,
  type NumberFormat,
  type BorderPreset,
  type BorderStyle,
  type HorizontalAlignment,
  type VerticalAlignment,
  type TextWrapping,
  type MergeAction,
} from './components/Toolbar';
export {
  EditorToolbarContext,
  useEditorToolbar,
  type EditorToolbarProps,
} from './components/EditorToolbarContext';
export {
  ToolbarButton,
  ToolbarDropdown,
  ToolbarGroup,
  ToolbarMenuItem,
  ToolbarMenuSeparator,
  ToolbarSeparator,
} from './components/ui/ToolbarPrimitives';
export {
  DRAWN_ICON_NAMES,
  TOOLBAR_ICON_NAMES,
  type DrawnIconName,
  type IconProps,
  type IconSet,
  type ToolbarIconName,
} from './components/ui/ToolbarIcon';
export { LocaleProvider, useTranslation, type LocaleProviderProps } from './i18n';
