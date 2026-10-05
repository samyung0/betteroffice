/**
 * Commands a host's menus run through `PptxEditorApi.runCommand`. A few take a
 * value: `view.zoom` ('fit' or a scale such as '1.5'), `insert.shape` (a
 * preset geometry), `slide.newWithLayout` (a layout part path, '' for none)
 * and `format.borderWeight` (points, '' for no border).
 */
export const PPTX_COMMAND_IDS = [
  'file.save',
  'file.exportPng',
  'edit.undo',
  'edit.redo',
  'edit.delete',
  'view.present',
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
  'format.alignLeft',
  'format.alignCenter',
  'format.alignRight',
  'format.alignJustify',
  'format.borderWeight',
  'arrange.bringToFront',
  'arrange.bringForward',
  'arrange.sendBackward',
  'arrange.sendToBack',
] as const;

export type PptxCommandId = (typeof PPTX_COMMAND_IDS)[number];

/** Whether a command changes the deck (saving counts): a read-only editor runs only those that don't. */
export const PPTX_COMMAND_EDITS: Record<PptxCommandId, boolean> = {
  'file.save': true,
  'file.exportPng': false,
  'edit.undo': true,
  'edit.redo': true,
  'edit.delete': true,
  'view.present': false,
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
  'format.alignLeft': true,
  'format.alignCenter': true,
  'format.alignRight': true,
  'format.alignJustify': true,
  'format.borderWeight': true,
  'arrange.bringToFront': true,
  'arrange.bringForward': true,
  'arrange.sendBackward': true,
  'arrange.sendToBack': true,
};

/** What a host menu needs to draw the commands, reported on every change. */
export interface PptxCommandState {
  enabled: Record<PptxCommandId, boolean>;
  /** Toggles that are on: bold, italic, underline, the current alignment and speaker notes. */
  checked: PptxCommandId[];
  /** `view.zoom`'s current value. */
  zoom: string;
  /** `format.borderWeight`'s current value; null without a selected shape. */
  borderWeight: string | null;
  /** `slide.newWithLayout`'s values, labelled. */
  slideLayouts: ReadonlyArray<{ value: string; label: string }>;
  /** The current slide, 0-based. */
  slideIndex: number;
}
