import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { en } from '@betteroffice/docx-i18n';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession, type YrsLoc, type YrsSession } from '@betteroffice/docx/yrs';
import { useFindReplace } from '../../../hooks/useFindReplace';
import { FindReplaceDialog } from '../../dialogs/FindReplaceDialog';
import type { PagedEditorRef } from '../PagedEditor';
import { useFindReplaceBridge } from './useFindReplaceBridge';

const { cleanup, fireEvent, render } = await import('@testing-library/react');

const WASM = resolve(
  import.meta.dir,
  '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm'
);
const sessions: YrsSession[] = [];

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));
afterEach(() => {
  cleanup();
  for (const session of sessions.splice(0)) session.destroy();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

/** What the editor offers the bridge: its session, and display positions in body order. */
function editorFor(session: YrsSession): PagedEditorRef {
  const position = (loc: YrsLoc) =>
    session.paragraphs('body').findIndex((paragraph) => paragraph.paraId === loc.paraId) * 1000 +
    loc.offset;
  return {
    getYrsSession: () => session,
    yrsLocToDisplayPosition: position,
    syncYrsInputState: () => true,
    scrollToPosition: () => {},
  } as unknown as PagedEditorRef;
}

/** The find and replace dialog wired to the bridge as DocxEditorDialogs wires it. */
function FindAndReplace({ editor }: { editor: PagedEditorRef }) {
  const findReplace = useFindReplace({ initialReplaceMode: true });
  const bridge = useFindReplaceBridge({ pagedEditorRef: { current: editor }, findReplace });
  return (
    <FindReplaceDialog
      isOpen
      replaceMode
      onClose={() => {}}
      onFind={bridge.handleFind}
      onFindNext={bridge.handleFindNext}
      onFindPrevious={bridge.handleFindPrevious}
      onReplace={bridge.handleReplace}
      onReplaceAll={bridge.handleReplaceAll}
      currentResult={bridge.findResultRef.current}
    />
  );
}

const count = (current: number, total: number) =>
  en.dialogs.findReplace.matchCount
    .replace('{current}', String(current))
    .replace('{total}', String(total));

test('Replace and Replace all change the matches, a picture before one included', async () => {
  const session = await createYrsSession({ clientId: 7301 });
  sessions.push(session);
  const first = session.createStory('body', 'The cat and the Cat.').paraId;
  const { secondParaId } = session.splitParagraph({ story: 'body', paraId: first, offset: 20 });
  session.insertText({ story: 'body', paraId: secondParaId, offset: 0 }, 'See  the cat.');
  session.insertImage(
    { story: 'body', paraId: secondParaId, offset: 4 },
    { embedId: 'figure', widthEmu: 914400, heightEmu: 914400 }
  );
  session.formatRange(
    { story: 'body', start: { paraId: first, offset: 4 }, end: { paraId: first, offset: 7 } },
    { bold: true }
  );

  const view = render(<FindAndReplace editor={editorFor(session)} />);
  const find = view.getByLabelText(en.dialogs.findReplace.findAriaLabel);
  fireEvent.change(find, { target: { value: 'cat' } });
  fireEvent.keyDown(find, { key: 'Enter' });
  view.getByText(count(1, 3));
  expect(session.selection()).toMatchObject({
    anchor: { paraId: first, offset: 4 },
    head: { paraId: first, offset: 7 },
  });

  fireEvent.change(view.getByLabelText(en.dialogs.findReplace.replaceAriaLabel), {
    target: { value: 'dog' },
  });
  fireEvent.click(view.getByRole('button', { name: en.dialogs.findReplace.replaceButton }));
  view.getByText(count(1, 2));
  const texts = () => session.paragraphs('body').map((paragraph) => paragraph.text);
  expect(texts()).toEqual(['The dog and the Cat.', 'See  the cat.']);
  const dog = session
    .storySegments('body')
    .find((segment) => segment.kind === 'text' && segment.text === 'dog');
  expect(dog?.attributes.bold).toBe(true);

  fireEvent.click(view.getByRole('button', { name: en.dialogs.findReplace.replaceAllButton }));
  expect(texts()).toEqual(['The dog and the dog.', 'See  the dog.']);
  const second = session.storySegments('body').slice(
    session.storySegments('body').findIndex((segment) => segment.kind === 'pilcrow') + 1
  );
  expect(second.map((segment) => (segment.kind === 'text' ? segment.text : segment.kind))).toEqual(
    ['See ', 'embed', ' the dog.', 'pilcrow']
  );
});
