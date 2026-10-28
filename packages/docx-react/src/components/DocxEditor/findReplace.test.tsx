import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { en } from '@betteroffice/docx-i18n';
import { useEffect } from 'react';
import { useFindReplace } from '../../hooks/useFindReplace';
import FindReplaceDialog from '../dialogs/FindReplaceDialog';
import type { PagedEditorRef } from './PagedEditor';
import { useFindReplaceBridge } from './hooks/useFindReplaceBridge';
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
