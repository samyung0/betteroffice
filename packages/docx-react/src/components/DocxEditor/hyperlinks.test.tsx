import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { en } from '@betteroffice/docx-i18n';
import type { YrsSession } from '@betteroffice/docx/yrs';
import { useFindReplace } from '../../hooks/useFindReplace';
import type { useTableSelection } from '../../hooks/useTableSelection';
import { HyperlinkDialog, useHyperlinkDialog } from '../dialogs/HyperlinkDialog';
import type { PagedEditorRef } from './PagedEditor';
import { useHyperlinkActions } from './hooks/useHyperlinkActions';
import { useKeyboardShortcuts } from './hooks/useKeyboardShortcuts';
import {
  destroySessions,
  docxPackage,
  loadEditor,
  openSession,
  renderPagedEditor,
  savedParts,
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

const noTableSelection = {
  state: { tableIndex: null },
  handleAction: () => {},
} as unknown as ReturnType<typeof useTableSelection>;

/** Ctrl/Cmd+K and the link dialog, wired to the editor as DocxEditor wires them. */
function LinkDialog({ editor }: { editor: React.RefObject<PagedEditorRef | null> }) {
  const hyperlinkDialog = useHyperlinkDialog();
  const findReplace = useFindReplace();
  const actions = useHyperlinkActions({
    hyperlinkDialog,
    pagedEditorRef: editor,
    focusActiveEditor: () => {},
  });
  useKeyboardShortcuts({
    pagedEditorRef: editor,
    disableFindReplaceShortcuts: true,
    showFileOpen: false,
    findReplace,
    hyperlinkDialog,
    tableSelection: noTableSelection,
    readOnly: false,
  });
  const { state } = hyperlinkDialog;
  return state.isOpen ? (
    <HyperlinkDialog
      isOpen
      onClose={hyperlinkDialog.close}
      onSubmit={actions.handleHyperlinkSubmit}
      onRemove={state.isEditing ? actions.handleHyperlinkRemove : undefined}
      initialData={state.initialData}
      selectedText={state.selectedText}
      isEditing={state.isEditing}
    />
  ) : null;
}

const linkKey = () =>
  act(() => {
    fireEvent.keyDown(document, { key: 'k', ctrlKey: true, metaKey: true });
  });

const fill = (placeholder: string, value: string) =>
  fireEvent.change(screen.getByPlaceholderText(placeholder), { target: { value } });

/** The body as text, `[text](href)` around linked runs. */
function linked(session: YrsSession): string {
  return session
    .storySegments('body')
    .map((segment) => {
      if (segment.kind !== 'text') return segment.kind === 'pilcrow' ? '¶' : '';
      const link = segment.attributes.hyperlink as { href?: string } | null | undefined;
      return link?.href ? `[${segment.text}](${link.href})` : segment.text;
    })
    .join('');
}

test('Ctrl/Cmd+K inserts, edits and removes a link, and the save writes it', async () => {
  const session = await openSession(
    docxPackage('<w:p w14:paraId="10000001"><w:r><w:t>Read the guide today.</w:t></w:r></w:p>'),
    7401
  );
  const editor = await renderPagedEditor(session);
  render(<LinkDialog editor={editor} />);
  const at = (offset: number) => ({ story: 'body', paraId: '10000001', offset });

  // Insert over the selection.
  session.setSelection(at(9), at(14));
  linkKey();
  fill(en.dialogs.hyperlink.urlPlaceholder, 'example.com/guide');
  fireEvent.click(screen.getByRole('button', { name: en.common.insert }));
  expect(linked(session)).toBe('Read the [guide](https://example.com/guide) today.¶');

  // Insert at the caret with its own text.
  session.setSelection(at(21));
  linkKey();
  fill(en.dialogs.hyperlink.urlPlaceholder, 'https://capy.example/');
  fill(en.dialogs.hyperlink.displayTextPlaceholder, ' Capy');
  fireEvent.click(screen.getByRole('button', { name: en.common.insert }));
  expect(linked(session)).toBe(
    'Read the [guide](https://example.com/guide) today.[Capy](https://capy.example/)¶'
  );

  // An address alone at the caret becomes the link's text.
  session.setSelection(at(0));
  linkKey();
  fill(en.dialogs.hyperlink.urlPlaceholder, 'capy.example/start ');
  fireEvent.click(screen.getByRole('button', { name: en.common.insert }));
  expect(linked(session)).toStartWith('[https://capy.example/start](https://capy.example/start)Read');
  session.deleteRange({ story: 'body', start: at(0), end: at(26) });

  // Edit the link under the caret: new address and text.
  session.setSelection(at(11));
  linkKey();
  expect(
    (screen.getByPlaceholderText(en.dialogs.hyperlink.urlPlaceholder) as HTMLInputElement).value
  ).toBe('https://example.com/guide');
  fill(en.dialogs.hyperlink.urlPlaceholder, 'https://example.com/v2');
  fill(en.dialogs.hyperlink.displayTextPlaceholder, 'new guide');
  fireEvent.click(screen.getByRole('button', { name: en.common.update }));
  const edited =
    'Read the [new guide](https://example.com/v2) today.[Capy](https://capy.example/)¶';
  expect(linked(session)).toBe(edited);

  let parts = await savedParts(session);
  const rels = parts['word/_rels/document.xml.rels'];
  for (const [text, target] of [
    ['new guide', 'https://example.com/v2'],
    ['Capy', 'https://capy.example/'],
  ]) {
    const id = new RegExp(`Id="([^"]+)"[^>]*Target="${target}"[^>]*TargetMode="External"`).exec(
      rels
    )?.[1];
    expect(id).toBeDefined();
    expect(parts['word/document.xml']).toMatch(
      new RegExp(`<w:hyperlink r:id="${id}"[^>]*>(?:(?!</w:hyperlink>).)*${text}`)
    );
  }

  // Remove the link under the caret from the dialog.
  session.setSelection(at(12));
  linkKey();
  fireEvent.click(screen.getByRole('button', { name: en.dialogs.hyperlink.removeLink }));
  expect(linked(session)).toBe('Read the new guide today.[Capy](https://capy.example/)¶');
  parts = await savedParts(session);
  expect(parts['word/document.xml'].match(/<w:hyperlink /g)).toHaveLength(1);
  expect(parts['word/document.xml']).toContain('Read the new guide today.');
});
