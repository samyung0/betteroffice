import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { en } from '@betteroffice/docx-i18n';
import { useEffect } from 'react';
import { useFindReplace } from '../../hooks/useFindReplace';
import { useTableSelection } from '../../hooks/useTableSelection';
import FindReplaceDialog from '../dialogs/FindReplaceDialog';
import { useHyperlinkDialog } from '../dialogs/HyperlinkDialog';
import type { PagedEditorRef } from './PagedEditor';
import { useFindReplaceBridge } from './hooks/useFindReplaceBridge';
import { useKeyboardShortcuts } from './hooks/useKeyboardShortcuts';
import {
  destroySessions,
  docxPackage,
  loadEditor,
  openSession,
  renderPagedEditor,
} from './pagedEditor.testing';

const { act, cleanup, fireEvent, render, screen } = await import('@testing-library/react');

beforeAll(loadEditor);
afterEach(async () => {
  await act(async () => {
    await new Promise((done) => setTimeout(done, 50));
  });
  cleanup();
  destroySessions();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

/** Find, wired to the editor as DocxEditor and DocxEditorDialogs wire it. */
function Find({ editor }: { editor: React.RefObject<PagedEditorRef | null> }) {
  const findReplace = useFindReplace();
  const bridge = useFindReplaceBridge({ pagedEditorRef: editor, findReplace });
  const { openFind } = findReplace;
  useEffect(() => openFind(), [openFind]);
  return findReplace.state.isOpen ? (
    <FindReplaceDialog
      isOpen
      onClose={bridge.handleClose}
      onFind={bridge.handleFind}
      onFindNext={bridge.handleFindNext}
      onFindPrevious={bridge.handleFindPrevious}
      onReplace={bridge.handleReplace}
      onReplaceAll={bridge.handleReplaceAll}
      currentResult={bridge.findResultRef.current}
    />
  ) : null;
}

// #1504: closing find returns focus to the document, where navigation left the current match
// selected, so typing continues at the match.
test('Escape in find returns focus to the document and keeps the match selected', async () => {
  const session = await openSession(
    docxPackage('<w:p w14:paraId="10000001"><w:r><w:t>one word, two word, three</w:t></w:r></w:p>'),
    7501
  );
  const editor = await renderPagedEditor(session);
  render(<Find editor={editor} />);
  const search = await screen.findByPlaceholderText(en.dialogs.findReplace.findPlaceholder);
  // the selection find navigation leaves on the second match
  const match = {
    anchor: { story: 'body', paraId: '10000001', offset: 14 },
    head: { story: 'body', paraId: '10000001', offset: 18 },
  };
  session.setSelection(match.anchor, match.head);
  search.focus();

  await act(async () => {
    fireEvent.keyDown(search, { key: 'Escape' });
  });
  expect(screen.queryByPlaceholderText(en.dialogs.findReplace.findPlaceholder)).toBeNull();
  expect(document.activeElement).toBe(screen.getByTestId('yrs-input'));
  expect(session.selection()).toMatchObject(match);
});

/** Find opened by the editor's shortcuts, wired as DocxEditor and DocxEditorDialogs wire it. */
function Shortcuts({ editor }: { editor: React.RefObject<PagedEditorRef | null> }) {
  const findReplace = useFindReplace();
  const bridge = useFindReplaceBridge({ pagedEditorRef: editor, findReplace });
  const hyperlinkDialog = useHyperlinkDialog();
  const tableSelection = useTableSelection({ document: null });
  useKeyboardShortcuts({
    pagedEditorRef: editor,
    disableFindReplaceShortcuts: false,
    showFileOpen: false,
    findReplace,
    hyperlinkDialog,
    tableSelection,
    readOnly: false,
    showHiddenText: false,
  });
  return findReplace.state.isOpen ? (
    <FindReplaceDialog
      isOpen
      onClose={bridge.handleClose}
      onFind={bridge.handleFind}
      onFindNext={bridge.handleFindNext}
      onFindPrevious={bridge.handleFindPrevious}
      onReplace={bridge.handleReplace}
      onReplaceAll={bridge.handleReplaceAll}
      initialSearchText={findReplace.state.searchText}
      replaceMode={findReplace.state.replaceMode}
      openRequest={findReplace.state.openRequest}
      selectedText={findReplace.state.selectedText}
      currentResult={bridge.findResultRef.current}
    />
  ) : null;
}

// Q828 A: Ctrl+F and Ctrl+H fill Find with the document's selected text, read
// from the editor (its focus target, a hidden textarea, holds no selection).
test('Ctrl+F and Ctrl+H fill Find with the document selection', async () => {
  const session = await openSession(
    docxPackage('<w:p w14:paraId="10000002"><w:r><w:t>one word, two word, three</w:t></w:r></w:p>'),
    7502
  );
  const editor = await renderPagedEditor(session);
  render(<Shortcuts editor={editor} />);
  /** Selects [start, end) of the paragraph and presses Ctrl/Cmd+`key` (both modifiers, wherever the test runs). */
  const press = (key: string, start: number, end: number) =>
    act(async () => {
      session.setSelection(
        { story: 'body', paraId: '10000002', offset: start },
        { story: 'body', paraId: '10000002', offset: end }
      );
      fireEvent.keyDown(document, { key, ctrlKey: true, metaKey: true });
    });

  await press('f', 4, 8);
  const search = screen.getByPlaceholderText(
    en.dialogs.findReplace.findPlaceholder
  ) as HTMLInputElement;
  expect(search.value).toBe('word');
  expect(document.activeElement === search).toBe(true);

  // Back in the document with another selection, Ctrl+H takes it.
  await act(async () => screen.getByTestId('yrs-input').focus());
  await press('h', 10, 13);
  expect(search.value).toBe('two');
  expect(screen.getByLabelText(en.dialogs.findReplace.replaceAriaLabel)).toBeTruthy();
  expect(document.activeElement === search).toBe(true);
});
