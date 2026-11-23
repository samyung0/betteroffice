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

// Q825 A: every key typed after Ctrl+F reaches Find, so the box is focused with
// its seed selected in the update that opens the dialog: Find loads with the
// editor (no lazy chunk) and focuses without a timer. None is advanced here.
test('Find opens focused with its seed selected, also when reopened with a new seed', () => {
  const container = document.createElement('div');
  document.body.appendChild(container);
  const root = createRoot(container);
  let findReplace: ReturnType<typeof useFindReplace> | undefined;

  try {
    flushSync(() => {
      root.render(<Dialogs onFindReplace={(find) => (findReplace = find)} />);
    });
    flushSync(() => findReplace?.openFind('Seed'));
    const box = container.querySelector<HTMLInputElement>(
      `[aria-label="${en.dialogs.findReplace.findAriaLabel}"]`
    );
    expect(document.activeElement).toBe(box);
    expect(box?.value).toBe('Seed');
    expect([box?.selectionStart, box?.selectionEnd]).toEqual([0, 4]);

    // Find and replace from the ribbon, with another selection, while Find is open.
    box?.setSelectionRange(4, 4);
    flushSync(() => findReplace?.openReplace('Other seed'));
    expect(document.activeElement).toBe(box);
    expect(box?.value).toBe('Other seed');
    expect([box?.selectionStart, box?.selectionEnd]).toEqual([0, 10]);
  } finally {
    root.unmount();
    container.remove();
  }
});
