/**
 * Commands a host's menus run through `PptxEditorApi.runCommand`. A few take a
 * value: `view.present` ('start' presents from the first slide, otherwise
 * from the current one), `view.zoom` ('fit' or a scale such as '1.5'),
 * `insert.shape` (a preset geometry), `slide.newWithLayout` (a layout part
 * path, '' for none), `format.borderWeight` (points, '' for no border),
 * `format.lineSpacing` (a share: '1', '1.15', '1.5', '2'), `format.bulletedList` and
 * `format.numberedList` (a style from `BULLET_PRESETS`/`NUMBER_PRESETS`;
 * without one they toggle the list), `arrange.align` ('left', 'center',
 * 'right', 'top', 'middle', 'bottom'), `arrange.distribute` and
 * `arrange.centerOnPage` ('horizontal', 'vertical').
 */
export const PPTX_COMMAND_IDS = [
  'file.save',
  'file.exportPng',
  'edit.undo',
  'edit.redo',
  'edit.selectAll',
  'edit.delete',
  'view.present',
  'view.presenterView',
  'view.zoom',
  'view.speakerNotes',
  'insert.textBox',
  'insert.image',
  'insert.shape',
  'slide.new',
  'slide.newWithLayout',
  'slide.delete',
  'slide.moveUp',
  'slide.moveDown',
  'slide.moveToStart',
  'slide.moveToEnd',
  'format.bold',
  'format.italic',
  'format.underline',
  'format.strikethrough',
  'format.superscript',
  'format.subscript',
  'format.increaseFontSize',
  'format.decreaseFontSize',
  'format.alignLeft',
  'format.alignCenter',
  'format.alignRight',
  'format.alignJustify',
  'format.alignTop',
  'format.alignMiddle',
  'format.alignBottom',
  'format.increaseIndent',
  'format.decreaseIndent',
  'format.lineSpacing',
  'format.spaceBefore',
  'format.spaceAfter',
  'format.bulletedList',
  'format.numberedList',
  'format.clearFormatting',
  'format.borderWeight',
  'arrange.bringToFront',
  'arrange.bringForward',
  'arrange.sendBackward',
  'arrange.sendToBack',
  'arrange.align',
  'arrange.distribute',
  'arrange.centerOnPage',
] as const;

export type PptxCommandId = (typeof PPTX_COMMAND_IDS)[number];

/** Whether a command changes the deck (saving counts): a read-only editor runs only those that don't. */
export const PPTX_COMMAND_EDITS: Record<PptxCommandId, boolean> = {
  'file.save': true,
  'file.exportPng': false,
  'edit.undo': true,
  'edit.redo': true,
  'edit.selectAll': false,
  'edit.delete': true,
  'view.present': false,
  'view.presenterView': false,
  'view.zoom': false,
  'view.speakerNotes': false,
  'insert.textBox': true,
  'insert.image': true,
  'insert.shape': true,
  'slide.new': true,
  'slide.newWithLayout': true,
  'slide.delete': true,
  'slide.moveUp': true,
  'slide.moveDown': true,
  'slide.moveToStart': true,
  'slide.moveToEnd': true,
  'format.bold': true,
  'format.italic': true,
  'format.underline': true,
  'format.strikethrough': true,
  'format.superscript': true,
  'format.subscript': true,
  'format.increaseFontSize': true,
  'format.decreaseFontSize': true,
  'format.alignLeft': true,
  'format.alignCenter': true,
  'format.alignRight': true,
  'format.alignJustify': true,
  'format.alignTop': true,
  'format.alignMiddle': true,
  'format.alignBottom': true,
  'format.increaseIndent': true,
  'format.decreaseIndent': true,
  'format.lineSpacing': true,
  'format.spaceBefore': true,
  'format.spaceAfter': true,
  'format.bulletedList': true,
  'format.numberedList': true,
  'format.clearFormatting': true,
  'format.borderWeight': true,
  'arrange.bringToFront': true,
  'arrange.bringForward': true,
  'arrange.sendBackward': true,
  'arrange.sendToBack': true,
  'arrange.align': true,
  'arrange.distribute': true,
  'arrange.centerOnPage': true,
};

/** What a host menu needs to draw the commands, reported on every change. */
export interface PptxCommandState {
  enabled: Record<PptxCommandId, boolean>;
  /**
   * Toggles that are on: bold, italic, underline, strikethrough, superscript,
   * subscript, the current alignments, the list kind, space before/after and
   * speaker notes.
   */
  checked: PptxCommandId[];
  /** `view.zoom`'s current value. */
  zoom: string;
  /** `format.borderWeight`'s current value; null without a selected shape. */
  borderWeight: string | null;
  /** `format.lineSpacing`'s current value; null when it differs or is not a share. */
  lineSpacing: string | null;
  /** The list style every selected list item is in; null otherwise. */
  listStyle: string | null;
  /** `slide.newWithLayout`'s values, labelled. */
  slideLayouts: ReadonlyArray<{ value: string; label: string }>;
  /** The current slide, 0-based. */
  slideIndex: number;
}
