/** The show and Presenter view alone, for viewers that load no editor code. */
export { PresentationOverlay, type PresentationSource } from './components/PresentationOverlay';
export { IconSetContext, type IconSet } from './components/ui/ToolbarIcon';
export { LocaleProvider } from './i18n';
export {
  NotesWindow,
  NOTES_WINDOW_FEATURES,
  ownNotesWindow,
  type NotesWindowHost,
  type NotesWindowState,
} from './notesWindow';
export {
  DEFAULT_NOTES_SIZE,
  NOTES_SIZES,
  PRESENT_ITEMS,
  notesSize,
} from './presenter';
