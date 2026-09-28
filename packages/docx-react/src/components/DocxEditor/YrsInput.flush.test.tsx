import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRef } from 'react';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import {
  createYrsInputPositionMap,
  createYrsSession,
  displayPositionToYrsLoc,
  yrsLocToDisplayPosition,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { YrsInput, type YrsInputProps, type YrsInputRef } from './YrsInput';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render } = await import('@testing-library/react');
const sessions: YrsSession[] = [];

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(
        resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
      )
    )
  )
);
afterEach(() => {
  cleanup();
  for (const session of sessions.splice(0)) session.destroy();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

async function mount(
  applyResidentInput?: YrsInputProps['applyResidentInput'],
  onPendingChange?: YrsInputProps['onPendingChange'],
  nextParagraphStyleId?: YrsInputProps['nextParagraphStyleId'],
) {
  const session = await createYrsSession();
  sessions.push(session);
  const { paraId } = session.createStory('body', 'Seed');
  session.setSelection({ story: 'body', paraId, offset: 4 });
  const map = () =>
    createYrsInputPositionMap(
      'body',
      session.paragraphs('body').map((p) => ({
        paraId: p.paraId,
        length: p.text.length,
      }))
    );
  const input = createRef<YrsInputRef>();
  const component = (readOnly = false) => (
    <YrsInput
      ref={input}
      enabled
      readOnly={readOnly}
      session={session}
      inputPositionMap={map}
      displayPositionToLoc={(position) => displayPositionToYrsLoc(map(), position)}
      locToDisplayPosition={(loc) => yrsLocToDisplayPosition(map(), loc)}
      onStateChange={() => {}}
      onDirectInput={() => {}}
      onPendingChange={onPendingChange}
      applyResidentInput={applyResidentInput}
      nextParagraphStyleId={nextParagraphStyleId}
    />
  );
  const view = render(component());
  return { session, input, view, setReadOnly: () => view.rerender(component(true)) };
}

test('flush waits for resident input and publishes the latest selection', async () => {
  const pending: boolean[] = [];
  let release!: () => void;
  const blocked = new Promise<void>((resolve) => {
    release = resolve;
  });
  const { session, input } = await mount(async () => {
    await blocked;
    return null;
  }, (value) => pending.push(value));
  act(() => input.current!.insertText(' accepted'));
  expect(pending).toEqual([true]);
  let done = false;
  const flush = input.current!.flushPendingInput().then(() => {
    done = true;
  });
  await Promise.resolve();
  expect(done).toBe(false);
  expect(session.paragraphs('body')[0].text).toBe('Seed');
  await act(async () => {
    release();
    await flush;
  });
  expect(session.paragraphs('body')[0].text).toBe('Seed accepted');
  expect(session.selection()?.head.offset).toBe(13);
  expect(pending).toEqual([true, false]);
});

test('flush seals a queued text batch before subsequent input', async () => {
  const calls: string[] = [];
  const { session, input } = await mount(async (text) => {
    calls.push(text);
    return null;
  });
  act(() => input.current!.insertText('A'));
  const first = input.current!.flushPendingInput();
  act(() => input.current!.insertText('B'));
  await act(async () => {
    await first;
    await input.current!.flushPendingInput();
  });
  expect(calls).toEqual(['A', 'B']);
  expect(session.paragraphs('body')[0].text).toBe('SeedAB');
});

test('an unresolved pointer placement blocks input until a valid fresh placement', async () => {
  const { session, input, view } = await mount();
  const textarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
  act(() => {
    input.current!.beginPointerSelection();
    input.current!.insertText(' wrong');
  });
  fireEvent.keyDown(textarea, { key: 'Enter' });
  fireEvent.paste(textarea, { clipboardData: { getData: () => 'wrong paste' } });
  await act(async () => { await input.current!.flushPendingInput(); });
  expect(session.paragraphs('body')[0].text).toBe('Seed');
  expect(textarea.readOnly).toBe(true);
  expect(textarea.dataset.pointerPlacement).toBe('pending');
  act(() => input.current!.setSelectionFromDisplay(1));
  act(() => input.current!.insertText('Right '));
  await act(async () => { await input.current!.flushPendingInput(); });
  expect(session.paragraphs('body')[0].text).toBe('Right Seed');
  expect(textarea.dataset.pointerPlacement).toBe('ready');
});

test('flush includes a completed IME composition exactly once', async () => {
  const { session, input, view } = await mount();
  const textarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
  fireEvent.compositionStart(textarea);
  textarea.value = '日本';
  act(() => {
    expect(input.current!.beginPointerSelection()).toBe(false);
    input.current!.setSelectionFromDisplay(1);
  });
  fireEvent.compositionEnd(textarea, { data: '日本' });
  await act(async () => {
    await input.current!.flushPendingInput();
  });
  fireEvent.input(textarea);
  await act(async () => {
    await input.current!.flushPendingInput();
  });
  expect(session.paragraphs('body')[0].text).toBe('Seed日本');
  expect(session.selection()?.head.offset).toBe(6);
});

for (const inFlight of [false, true]) {
  test(`pointer placement preserves ${inFlight ? 'in-flight' : 'queued'} accepted input at its original caret`, async () => {
    let release!: () => void;
    const blocked = new Promise<void>((resolve) => { release = resolve; });
    const { session, input, view } = await mount(async () => { await blocked; return null; });
    const text = () => session.paragraphs('body').map((p) => p.text).join('\n');
    const textarea = view.getByTestId('yrs-input');
    act(() => input.current!.insertText(' accepted'));
    if (inFlight) await Promise.resolve();
    fireEvent.keyDown(textarea, { key: 'Enter' });
    act(() => input.current!.insertText('😀Z'));
    fireEvent.keyDown(textarea, { key: 'Backspace' });
    act(() => {
      expect(input.current!.beginPointerSelection()).toBe(false);
      input.current!.setSelectionFromDisplay(1);
      input.current!.insertText(' wrong');
    });
    expect(textarea.getAttribute('data-pointer-placement')).toBe('pending');
    await act(async () => { release(); await input.current!.flushPendingInput(); });
    expect(text()).toBe('Seed accepted\n😀');
    expect(textarea.getAttribute('data-pointer-placement')).toBe('pending');
    act(() => {
      expect(input.current!.beginPointerSelection()).toBe(true);
      input.current!.setSelectionFromDisplay(1);
      input.current!.insertText('Right ');
    });
    await act(async () => { await input.current!.flushPendingInput(); });
    expect(text()).toBe('Right Seed accepted\n😀');
  });
}

test('flush waits for an active composition and rejects if the input is removed', async () => {
  const { input, view } = await mount();
  fireEvent.compositionStart(view.getByTestId('yrs-input'));
  let done = false;
  const flush = input.current!.flushPendingInput().finally(() => {
    done = true;
  });
  await Promise.resolve();
  expect(done).toBe(false);
  view.unmount();
  await expect(flush).rejects.toThrow('unavailable');
});

test('flush rejects failed resident input instead of claiming it was committed', async () => {
  const failure = new Error('resident failure');
  const { session, input } = await mount(async () => {
    throw failure;
  });
  act(() => input.current!.insertText('lost'));
  await expect(input.current!.flushPendingInput()).rejects.toBe(failure);
  expect(session.paragraphs('body')[0].text).toBe('Seed');
});

test('read-only interrupts active composition and rejects flush without waiting for compositionend', async () => {
  const { session, input, view, setReadOnly } = await mount();
  const textarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
  fireEvent.compositionStart(textarea);
  textarea.value = '日本';
  const flush = input.current!.flushPendingInput();
  setReadOnly();
  await expect(flush).rejects.toThrow('Composition interrupted');
  fireEvent.compositionEnd(textarea, { data: '日本' });
  expect(textarea.value).toBe('');
  expect(session.paragraphs('body')[0].text).toBe('Seed');
  await expect(input.current!.flushPendingInput()).rejects.toThrow('Composition interrupted');
});

test("Enter before a table leaves the table paragraph's style alone", async () => {
  const { session, input, view } = await mount(undefined, undefined, () => 'Normal');
  const [seed] = session.paragraphs('body');
  // Seed¶[table]¶: the paragraph after the table is empty.
  const { secondParaId: slot } = session.splitParagraph({ story: 'body', paraId: seed.paraId, offset: 4 });
  session.insertTable({ story: 'body', paraId: slot, offset: 0 }, 1, 1);
  const styled: string[] = [];
  session.applyParagraphStyle = (range) => {
    styled.push(range.start.paraId);
  };
  const textarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
  const enter = async (paraId: string, offset: number) => {
    act(() => session.setSelection({ story: 'body', paraId, offset }));
    fireEvent.keyDown(textarea, { key: 'Enter' });
    await act(async () => {
      await input.current!.flushPendingInput();
    });
  };
  // The new paragraph opens ahead of the table's, which keeps its id and style.
  await enter(slot, 0);
  expect(styled).toEqual([]);
  expect(session.paragraphs('body').map((p) => p.paraId)[2]).toBe(slot);
  // At the end of a paragraph the new one takes the next style.
  await enter(seed.paraId, 4);
  expect(styled).toHaveLength(1);
  expect(styled[0]).not.toBe(seed.paraId);
});
