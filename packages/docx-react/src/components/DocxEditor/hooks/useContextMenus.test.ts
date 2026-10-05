import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, test } from 'bun:test';
import type { YrsSession } from '@betteroffice/docx/yrs';
import type { PagedEditorRef } from '../PagedEditor';
import { useContextMenus } from './useContextMenus';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');

function options(
  partEditOpen: boolean,
  onAddComment: (range: { from: number; to: number; yPos: number | null }) => void = () => {},
  readOnly = false
) {
  return {
    pagedEditorRef: {
      current: {
        getYrsSession: () => null,
        getSelectionRange: () => ({ from: 1, to: 2 }),
      } as PagedEditorRef,
    },
    focusActiveEditor: () => {},
    openSplitCellDialog: () => {},
    editorContentRef: { current: null },
    displayListQueries: null,
    interactionPageHostRef: { current: null },
    i18n: undefined,
    partEditOpen,
    onAddComment,
    readOnly,
  };
}

function openSelectionMenu(result: { current: ReturnType<typeof useContextMenus> }): void {
  act(() => result.current.handleContextMenu({ x: 10, y: 20, hasSelection: true }));
}

afterEach(cleanup);
afterAll(() => {
  if (ownsDom) GlobalRegistrator.unregister();
});

describe('selection context menu', () => {
  test('offers Comment for a body selection', () => {
    const { result } = renderHook(() => useContextMenus(options(false)));
    openSelectionMenu(result);

    expect(result.current.contextMenuItems.map((item) => item.action)).toContain('addComment');
  });

  test("labels the menu from the editor's i18n, which its LocaleProvider does not reach", () => {
    const i18n = { contextMenu: { selectAll: 'Tout sélectionner' } };
    const { result } = renderHook(() => useContextMenus({ ...options(false), i18n }));
    openSelectionMenu(result);

    const selectAll = result.current.contextMenuItems.find((item) => item.action === 'selectAll');
    expect(selectAll?.label).toBe('Tout sélectionner');
  });

  test('does not offer Comment for a selection in an open note', () => {
    const { result } = renderHook(() => useContextMenus(options(true)));
    openSelectionMenu(result);

    expect(result.current.contextMenuItems.map((item) => item.action)).not.toContain('addComment');
  });

  test('rejects a displayed Comment action when a part opens before dispatch', async () => {
    let added = 0;
    const onAddComment = () => {
      added += 1;
    };
    const { result, rerender } = renderHook(
      ({ partEditOpen }) => useContextMenus(options(partEditOpen, onAddComment)),
      { initialProps: { partEditOpen: false } }
    );
    openSelectionMenu(result);
    expect(result.current.contextMenuItems.map((item) => item.action)).toContain('addComment');

    rerender({ partEditOpen: true });
    await act(async () => {
      await result.current.handleContextMenuAction('addComment');
    });

    expect(added).toBe(0);
  });
});

describe('read-only context menu', () => {
  test('offers only Copy and Select all, and runs nothing else', async () => {
    let deleted = 0;
    const base = options(false, () => {}, true);
    const pagedEditorRef = {
      current: {
        ...base.pagedEditorRef.current,
        deleteSelection: () => {
          deleted += 1;
        },
        insertText: () => {
          deleted += 1;
        },
      } as PagedEditorRef,
    };
    const { result } = renderHook(() => useContextMenus({ ...base, pagedEditorRef }));
    openSelectionMenu(result);
    expect(result.current.contextMenuItems.map((item) => item.action)).toEqual([
      'copy',
      'selectAll',
    ]);
    for (const action of ['cut', 'delete', 'paste'] as const)
      await act(async () => {
        await result.current.handleContextMenuAction(action);
      });
    expect(deleted).toBe(0);
  });
});

describe('context menu cut', () => {
  test('deletes only after the clipboard accepts the text', async () => {
    // [Seed]¶ selected
    const session = {
      selection: () => ({
        anchor: { story: 'body', paraId: 'p', offset: 0 },
        head: { story: 'body', paraId: 'p', offset: 4 },
      }),
      locateParagraph: () => ({ start: 0, end: 4 }),
      storySegments: () => [{ kind: 'text', text: 'Seed', attributes: {} }],
    } as unknown as YrsSession;
    let deleted = 0;
    const base = options(false);
    const pagedEditorRef = {
      current: {
        ...base.pagedEditorRef.current,
        getYrsSession: () => session,
        deleteSelection: () => {
          deleted += 1;
        },
      } as PagedEditorRef,
    };
    const { result } = renderHook(() => useContextMenus({ ...base, pagedEditorRef }));
    const clipboard = navigator.clipboard;
    const writeText = clipboard.writeText;
    const written: string[] = [];
    try {
      clipboard.writeText = () => Promise.reject(new Error('refused'));
      await act(async () => {
        await result.current.handleContextMenuAction('cut');
      });
      expect(deleted).toBe(0);

      clipboard.writeText = async (text) => {
        written.push(text);
      };
      await act(async () => {
        await result.current.handleContextMenuAction('cut');
      });
      expect(written).toEqual(['Seed']);
      expect(deleted).toBe(1);
    } finally {
      clipboard.writeText = writeText;
    }
  });
});
