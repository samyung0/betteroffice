import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import { createRef } from 'react';
import type { YrsSession } from '@betteroffice/docx/yrs';
import type { PagedEditorRef } from '../PagedEditor';
import type { YrsInputRef } from '../YrsInput';
import { usePagedEditorRefApi, type UsePagedEditorRefApiOptions } from './usePagedEditorRefApi';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, render } = await import('@testing-library/react');

afterEach(() => cleanup());
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function Host(props: UsePagedEditorRefApiOptions) {
  usePagedEditorRefApi(props);
  return null;
}

// Q833 B: Undo, Redo and Find's replace edit outside the document input; each marks the input
// as the reader's edit, so one made before the first caret pass brings the caret into view.
test("Undo, Redo and the reader-edit sync mark the input as the reader's edit", () => {
  const marks: string[] = [];
  const synced: boolean[] = [];
  const ref = createRef<PagedEditorRef>();
  const input = { noteLocalEdit: () => marks.push('edit') } as unknown as YrsInputRef;
  const session = { selection: () => null, undo: () => true, redo: () => true };
  render(
    <Host
      {...({} as UsePagedEditorRefApiOptions)}
      ref={ref}
      yrsInputRef={{ current: input }}
      onReadyRef={{ current: undefined }}
      yrsSession={session as unknown as YrsSession}
      syncYrsInputState={(docChanged) => synced.push(docChanged) > 0}
    />
  );
  ref.current!.undo();
  ref.current!.redo();
  ref.current!.syncYrsReaderEdit();
  expect(marks).toEqual(['edit', 'edit', 'edit']);
  expect(synced).toEqual([true, true, true]);
});
