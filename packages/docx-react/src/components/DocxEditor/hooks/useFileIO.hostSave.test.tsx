import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '@betteroffice/docx/docx';
import type { Document } from '@betteroffice/docx/types/document';
import type { PagedEditorRef } from '../PagedEditor';
import { useFileIO } from './useFileIO';
import { useKeyboardShortcuts } from './useKeyboardShortcuts';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');
let fixture: Document;
beforeAll(async () => {
  fixture = await parseDocx(
    new Uint8Array(readFileSync(resolve(import.meta.dir, '__fixtures__/probe-linked-header.docx'))),
    { preloadFonts: false }
  );
});
afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function setup(
  options: {
    onSaveRequest?: () => boolean | void | Promise<boolean | void>;
    flush?: () => Promise<void>;
    focused?: boolean;
  } = {}
) {
  const events: string[] = [];
  const errors: Error[] = [];
  const saved: ArrayBuffer[] = [];
  const document = structuredClone(fixture);
  const session = {};
  const editor = {
    getYrsSession: () => session,
    isFocused: () => options.focused ?? true,
    flushPendingInput: async () => {
      events.push('flush');
      await options.flush?.();
    },
    getDocument: () => {
      events.push('snapshot');
      return document;
    },
  } as unknown as PagedEditorRef;
  const pagedEditorRef = { current: editor };
  const hook = renderHook(() => {
    const io = useFileIO({
      pagedEditorRef,
      displayList: null,
      resolveImage: () => null,
      documentName: 'saved',
      onSave: (buffer) => {
        events.push('saved');
        saved.push(buffer);
      },
      onSaveRequest: options.onSaveRequest,
      downloadOnSave: false,
      onError: (error) => errors.push(error),
      onOpen: undefined,
      onPrint: undefined,
      onDocumentNameChange: undefined,
      loadBuffer: async () => {},
      focusActiveEditor: () => {},
    });
    useKeyboardShortcuts({
      pagedEditorRef,
      onSaveDocument: io.handleDownloadDocument,
      disableFindReplaceShortcuts: true,
      showFileOpen: false,
      readOnly: false,
      findReplace: {} as never,
      hyperlinkDialog: {} as never,
      tableSelection: { state: { tableIndex: null } } as never,
    });
    return io;
  });
  return { hook, events, errors, saved, editor, pagedEditorRef };
}

test('a host owns Save before serialization and can save explicitly without reentry', async () => {
  let requests = 0;
  const state = setup({
    onSaveRequest: async () => {
      requests += 1;
      expect(state.events).toEqual([]);
      await state.hook.result.current.handleSave();
    },
  });
  await act(async () => {
    await state.hook.result.current.handleDownloadDocument();
  });
  expect(requests).toBe(1);
  expect(state.events).toEqual(['flush', 'snapshot', 'saved']);
  expect(state.saved).toHaveLength(1);
  expect(state.saved[0].byteLength).toBeGreaterThan(0);
  expect(state.errors).toEqual([]);
});

test('an awaited request can continue the built-in export exactly once', async () => {
  let release!: () => void;
  let requests = 0;
  const gate = new Promise<void>((resolve) => {
    release = resolve;
  });
  const state = setup({
    onSaveRequest: async () => {
      requests += 1;
      await gate;
      return true;
    },
  });
  const first = state.hook.result.current.handleDownloadDocument();
  const second = state.hook.result.current.handleDownloadDocument();
  expect(first).toBe(second);
  await Promise.resolve();
  expect(state.events).toEqual([]);
  await act(async () => {
    release();
    await first;
  });
  expect(requests).toBe(1);
  expect(state.events).toEqual(['flush', 'snapshot', 'saved']);
  expect(state.errors).toEqual([]);
});

test('cancellation and callback failures prevent export', async () => {
  const cancelled = setup({ onSaveRequest: () => false });
  await cancelled.hook.result.current.handleDownloadDocument();
  expect(cancelled.events).toEqual([]);
  const failure = new Error('revision mismatch');
  const failed = setup({
    onSaveRequest: () => {
      throw failure;
    },
  });
  await failed.hook.result.current.handleDownloadDocument();
  expect(failed.events).toEqual([]);
  expect(failed.errors).toEqual([failure]);
});

test('built-in export waits for input and aborts if the document changes', async () => {
  let release!: () => void;
  const gate = new Promise<void>((resolve) => {
    release = resolve;
  });
  const state = setup({ flush: () => gate });
  const saved = state.hook.result.current.handleSave();
  expect(state.events).toEqual(['flush']);
  state.pagedEditorRef.current = { ...state.editor, getYrsSession: () => ({}) as never };
  release();
  expect(await saved).toBeNull();
  expect(state.events).toEqual(['flush']);
  expect(state.errors[0].message).toContain('document changed');
});

test('a repaint during export keeps the save', async () => {
  let release!: () => void;
  const gate = new Promise<void>((resolve) => {
    release = resolve;
  });
  const state = setup({ flush: () => gate });
  const saved = state.hook.result.current.handleSave();
  // PagedEditor rebuilds its ref object on every repaint, same session.
  state.pagedEditorRef.current = { ...state.editor };
  release();
  expect(await saved).not.toBeNull();
  expect(state.events).toEqual(['flush', 'snapshot', 'saved']);
  expect(state.errors).toEqual([]);
});

test('failed input flush prevents serialization', async () => {
  const state = setup({
    flush: async () => {
      throw new Error('input failed');
    },
  });
  await state.hook.result.current.handleDownloadDocument();
  expect(state.events).toEqual(['flush']);
  expect(state.errors[0].message).toBe('input failed');
});

test('the save shortcut invokes only the focused editor and ignores repeat events', async () => {
  let inactiveRequests = 0;
  let activeRequests = 0;
  setup({
    focused: false,
    onSaveRequest: () => {
      inactiveRequests += 1;
    },
  });
  setup({
    onSaveRequest: () => {
      activeRequests += 1;
    },
  });
  const event = new KeyboardEvent('keydown', {
    key: 's',
    ctrlKey: true,
    metaKey: true,
    bubbles: true,
    cancelable: true,
  });
  await act(async () => {
    document.dispatchEvent(event);
  });
  expect(event.defaultPrevented).toBe(true);
  expect(activeRequests).toBe(1);
  expect(inactiveRequests).toBe(0);
  await act(async () => {
    document.dispatchEvent(
      new KeyboardEvent('keydown', {
        key: 's',
        ctrlKey: true,
        metaKey: true,
        repeat: true,
        bubbles: true,
        cancelable: true,
      })
    );
  });
  expect(activeRequests).toBe(1);
});

test('an inserted image is the file prepareImage returns: one at a time, pending until it lands, focus left where the user went, a refusal inserting nothing', async () => {
  const originalImage = globalThis.Image;
  const originalHasFocus = document.hasFocus;
  class FakeImage {
    onload: (() => void) | null = null;
    onerror: (() => void) | null = null;
    naturalWidth = 1000;
    naturalHeight = 500;
    set src(_value: string) {
      queueMicrotask(() => this.onload?.());
    }
  }
  globalThis.Image = FakeImage as unknown as typeof Image;
  try {
    const commands: { image: Record<string, unknown> }[] = [];
    const errors: Error[] = [];
    const picked: string[] = [];
    const pending: boolean[] = [];
    let focused = 0;
    let refuse = false;
    let release!: () => void;
    const editor = {
      applyYrsCommand: (command: { image: Record<string, unknown> }) => {
        commands.push(command);
        return true;
      },
    } as unknown as PagedEditorRef;
    const { result } = renderHook(() =>
      useFileIO({
        pagedEditorRef: { current: editor },
        displayList: null,
        resolveImage: () => null,
        documentName: 'doc',
        onSave: undefined,
        onError: (error) => errors.push(error),
        onOpen: undefined,
        onPrint: undefined,
        onDocumentNameChange: undefined,
        loadBuffer: async () => {},
        focusActiveEditor: () => {
          focused += 1;
        },
        onImagePending: (value) => pending.push(value),
        prepareImage: async (file) => {
          picked.push(file.name);
          await new Promise<void>((resolve) => {
            release = resolve;
          });
          if (refuse) throw new Error('image too large');
          return new File(['jpeg'], 'photo.jpg', { type: 'image/jpeg' });
        },
      })
    );
    document.hasFocus = () => true;
    let first!: Promise<void>;
    act(() => {
      first = result.current.insertImageFile(new File(['png bytes'], 'photo.png', { type: 'image/png' }));
    });
    // A second pick while the first is prepared is ignored.
    await act(async () => {
      await result.current.insertImageFile(new File(['png'], 'again.png', { type: 'image/png' }));
    });
    expect(picked).toEqual(['photo.png']);
    expect(pending).toEqual([true]);
    await act(async () => {
      release();
      await first;
    });
    expect(commands[0].image).toMatchObject({
      alt: 'photo.jpg',
      src: `data:image/jpeg;base64,${btoa('jpeg')}`,
      width: 612,
      height: 306,
    });
    expect(pending).toEqual([true, false]);
    expect(focused).toBe(1);

    // The user went elsewhere (another frame or window) meanwhile: no focus.
    document.hasFocus = () => false;
    let second!: Promise<void>;
    act(() => {
      second = result.current.insertImageFile(new File(['png'], 'later.png', { type: 'image/png' }));
    });
    await act(async () => {
      release();
      await second;
    });
    expect(commands).toHaveLength(2);
    expect(focused).toBe(1);

    refuse = true;
    let third!: Promise<void>;
    act(() => {
      third = result.current.insertImageFile(new File(['png'], 'big.png', { type: 'image/png' }));
    });
    await act(async () => {
      release();
      await third;
    });
    expect(commands).toHaveLength(2);
    expect(picked).toEqual(['photo.png', 'later.png', 'big.png']);
    expect(errors.map((error) => error.message)).toEqual(['image too large']);
    expect(pending).toEqual([true, false, true, false, true, false]);
  } finally {
    globalThis.Image = originalImage;
    document.hasFocus = originalHasFocus;
  }
});
