import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { en } from '@betteroffice/docx-i18n';
import { useRef } from 'react';
import { useFindReplace } from '../../hooks/useFindReplace';
import type { FindResult } from '../dialogs/FindReplaceDialog';
import { useHyperlinkDialog } from '../dialogs/HyperlinkDialog';
import { DocxEditorDialogs } from './DocxEditorDialogs';

const { createRoot } = await import('react-dom/client');
const { flushSync } = await import('react-dom');

afterAll(async () => {
  await new Promise((resolve) => setTimeout(resolve, 10));
  if (ownsDom) await GlobalRegistrator.unregister();
});

const noop = () => {};

/** The editor's dialogs with every other dialog closed; `onFindReplace` hands out Find's state. */
function Dialogs({ onFindReplace }: { onFindReplace: (find: ReturnType<typeof useFindReplace>) => void }) {
  const findReplace = useFindReplace();
  const hyperlinkDialog = useHyperlinkDialog();
  const findResultRef = useRef<FindResult | null>(null);
  onFindReplace(findReplace);
  return (
    <DocxEditorDialogs
      findReplace={findReplace}
      findResultRef={findResultRef}
      onFindClose={noop}
      onFind={() => null}
      onFindNext={() => null}
      onFindPrevious={() => null}
      onReplace={() => false}
      onReplaceAll={() => 0}
      readOnly={false}
      hyperlinkDialog={hyperlinkDialog}
      onHyperlinkSubmit={noop}
      onHyperlinkRemove={noop}
      tablePropsOpen={false}
      onTablePropsClose={noop}
      tableProperties={undefined}
      onTablePropertiesApply={noop}
      splitCellDialogState={{ isOpen: false, initialRows: 1, initialCols: 1, minRows: 1, minCols: 1 }}
      onSplitCellDialogClose={noop}
      onSplitCellDialogApply={noop}
      imagePositionOpen={false}
      onImagePositionClose={noop}
      onApplyImagePosition={noop}
      imagePropsOpen={false}
      onImagePropsClose={noop}
      onApplyImageProperties={noop}
      pmImageContext={null}
      showPageSetup={false}
      onPageSetupClose={noop}
      onPageSetupApply={noop}
      showWatermark={false}
      onWatermarkClose={noop}
      onWatermarkApply={noop}
      currentWatermark={undefined}
      document={null}
      footnotePropsOpen={false}
      onFootnotePropsClose={noop}
      onApplyFootnoteProperties={noop}
    />
  );
}

/**
 * Renders the dialogs and hands back Find's state, its box and a field outside
 * Find standing in for the document's focus target. None of the tests below
 * advances a timer: whatever Find does happens in the update that asks for it.
 */
function renderDialogs() {
  const container = document.createElement('div');
  const outside = document.createElement('textarea');
  document.body.append(outside, container);
  const root = createRoot(container);
  let findReplace: ReturnType<typeof useFindReplace> | undefined;
  flushSync(() => {
    root.render(<Dialogs onFindReplace={(find) => (findReplace = find)} />);
  });
  const find = () => {
    if (!findReplace) throw new Error('Find state missing');
    return findReplace;
  };
  const box = () =>
    container.querySelector<HTMLInputElement>(
      `[aria-label="${en.dialogs.findReplace.findAriaLabel}"]`
    );
  const replaceBox = () =>
    container.querySelector<HTMLInputElement>(
      `[aria-label="${en.dialogs.findReplace.replaceAriaLabel}"]`
    );
  /** Types `text` into Find's box, as a user does (React's onChange). */
  const type = (text: string) =>
    flushSync(() => {
      const input = box();
      if (!input) throw new Error('Find box missing');
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set?.call(input, text);
      input.dispatchEvent(new Event('input', { bubbles: true }));
    });
  const unmount = () => {
    root.unmount();
    container.remove();
    outside.remove();
  };
  return { box, container, find, outside, replaceBox, type, unmount };
}

const selection = (input: HTMLInputElement | null) => [input?.selectionStart, input?.selectionEnd];
/** The focused field's label: a failing check prints it, not the element's whole tree. */
const focused = () =>
  document.activeElement?.getAttribute('aria-label') ?? document.activeElement?.tagName;
const { findAriaLabel, replaceAriaLabel } = en.dialogs.findReplace;

// Q825 A: every key typed after Ctrl+F reaches Find, so the box is focused with
// its seed selected in the update that opens the dialog: Find loads with the
// editor (no lazy chunk) and focuses without a timer.
test('Find opens focused with its seed selected, and a request from the document takes its new selection', () => {
  const dialogs = renderDialogs();
  try {
    flushSync(() => dialogs.find().openFind('Seed'));
    expect(focused()).toBe(findAriaLabel);
    expect(dialogs.box()?.value).toBe('Seed');
    expect(selection(dialogs.box())).toEqual([0, 4]);

    // Find and replace from the ribbon, with another selection in the document.
    dialogs.outside.focus();
    flushSync(() => dialogs.find().openReplace('Other seed'));
    expect(focused()).toBe(findAriaLabel);
    expect(dialogs.box()?.value).toBe('Other seed');
    expect(selection(dialogs.box())).toEqual([0, 10]);
  } finally {
    dialogs.unmount();
  }
});

// Q828 A: Ctrl+F while Find is open moves the focus to its box and selects the
// query there, so the keys typed next replace it rather than reach the document.
test('a repeat Ctrl+F from the document focuses the open box with its query selected', () => {
  const dialogs = renderDialogs();
  try {
    flushSync(() => dialogs.find().openFind('Seed'));
    dialogs.type('mitosis');
    dialogs.outside.focus();
    flushSync(() => dialogs.find().openFind(''));
    expect(focused()).toBe(findAriaLabel);
    expect(dialogs.box()?.value).toBe('mitosis');
    expect(selection(dialogs.box())).toEqual([0, 7]);
  } finally {
    dialogs.unmount();
  }
});

// Q828 A: switching to Replace keeps the typed query. Ctrl+H in the box comes
// with the document's selection, Find's own current match, which stays out.
test('switching to Replace keeps the typed query', () => {
  const dialogs = renderDialogs();
  try {
    flushSync(() => dialogs.find().openFind('Seed'));
    dialogs.type('mitosis');
    flushSync(() => dialogs.find().openReplace('Mitosis'));
    expect(dialogs.box()?.value).toBe('mitosis');
    expect(dialogs.replaceBox()).not.toBeNull();
    expect(focused()).toBe(findAriaLabel);
  } finally {
    dialogs.unmount();
  }
});

// Q828 A: the Replace box opened from Find's own link focuses at once, as
// Find's box does, so the keys typed right after the click reach it.
test("Find's Replace link focuses the Replace box at once", () => {
  const dialogs = renderDialogs();
  try {
    flushSync(() => dialogs.find().openFind('Seed'));
    const link = [...dialogs.container.querySelectorAll('button')].find(
      (button) => button.textContent === en.dialogs.findReplace.toggleReplace
    );
    flushSync(() => link?.click());
    expect(dialogs.replaceBox()).not.toBeNull();
    expect(focused()).toBe(replaceAriaLabel);
  } finally {
    dialogs.unmount();
  }
});
