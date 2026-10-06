export { PptxEditor } from './PptxEditor';
export type {
  PptxEditorApi,
  PptxEditorCollaborationOptions,
  PptxEditorProps,
  PptxPointPosition,
  PptxTextSelection,
  PptxTextSelectionTarget,
} from './PptxEditor';
export {
  NotesWindow,
  NOTES_WINDOW_FEATURES,
  ownNotesWindow,
  type NotesWindowHost,
  type NotesWindowState,
} from './notesWindow';
export { DEFAULT_NOTES_SIZE, NOTES_SIZES, PRESENT_ITEMS, notesSize } from './presenter';
export {
  PPTX_COMMAND_EDITS,
  PPTX_COMMAND_IDS,
  type PptxCommandId,
  type PptxCommandState,
} from './commands';
export {
  BULLET_PRESETS,
  NUMBER_PRESETS,
  presetLabel,
  type BulletPresetId,
  type ListPresetId,
  type NumberPresetId,
} from './paragraphFormatting';
export { EditorToolbar } from './components/EditorToolbar';
export {
  Toolbar,
  PptxToolbar,
  SHAPE_PRESETS,
  type ToolbarProps,
  type SelectionFormatting,
  type FormattingAction,
  type ShapeFormatting,
  type ShapeFormattingAction,
  type SlideLayoutOption,
  type PptxEditorTool,
  type PptxShapePreset,
  type PptxZoom,
} from './components/Toolbar';
export {
  EditorToolbarContext,
  useEditorToolbar,
  type EditorToolbarProps,
} from './components/EditorToolbarContext';
export {
  DRAWN_ICON_NAMES,
  TOOLBAR_ICON_NAMES,
  type IconProps,
  type IconSet,
  type ToolbarIconName,
} from './components/ui/ToolbarIcon';
export {
  ToolbarButton,
  ToolbarDropdown,
  ToolbarGroup,
  ToolbarMenuItem,
  ToolbarMenuSeparator,
  ToolbarSeparator,
} from './components/ui/ToolbarPrimitives';
export { LocaleProvider, useTranslation, type LocaleProviderProps } from './i18n';
