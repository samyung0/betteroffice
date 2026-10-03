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
  // Story units per paragraph, embeds included, as the editor's core session maps them.
  const map = () => createYrsInputPositionMap('body', session.paragraphSpans('body'));
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

test('Delete before a field that shows nothing at the story end leaves input working', async () => {
  const { session, input, view } = await mount();
  const [seed] = session.paragraphs('body');
  // Seed[TC]¶: a field that draws nothing ends the story, and the caret sits before it.
  session.applyRawOps('body', [
    { op: 'insertEmbed', index: 4, kind: 'field', payload: { fieldType: 'TC', instruction: ' TC x ', displayText: '' } },
  ]);
  act(() => session.setSelection({ story: 'body', paraId: seed.paraId, offset: 4 }));
  const textarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
  fireEvent.keyDown(textarea, { key: 'Delete' });
  await act(async () => {
    await input.current!.flushPendingInput();
  });
  act(() => input.current!.insertText('!'));
  await act(async () => {
    await input.current!.flushPendingInput();
  });
  expect(session.paragraphs('body')[0].text.replace(/￼/g, '')).toBe('Seed!');
  expect(session.storySegments('body').filter((segment) => segment.kind === 'embed')).toHaveLength(1);
});

function clipboardSpy() {
  const written: Record<string, string> = {};
  const clipboardData = { setData: (type: string, value: string) => { written[type] = value; } };
  return { written, clipboardData };
}

async function flush(input: { current: YrsInputRef | null }) {
  await act(async () => { await input.current!.flushPendingInput(); });
}

async function mountAcrossParagraphs() {
  const mounted = await mount();
  const { session } = mounted;
  const [seed] = session.paragraphs('body');
  const { secondParaId } = session.splitParagraph({ story: 'body', paraId: seed.paraId, offset: 4 });
  session.insertText({ story: 'body', paraId: secondParaId, offset: 0 }, 'Next');
  session.addUndoBoundary();
  // Se[ed¶Ne]xt
  act(() =>
    session.setSelection(
      { story: 'body', paraId: seed.paraId, offset: 2 },
      { story: 'body', paraId: secondParaId, offset: 2 }
    )
  );
  const written: Record<string, string> = {};
  const clipboardData = { setData: (type: string, value: string) => { written[type] = value; } };
  const texts = () => session.paragraphs('body').map((p) => p.text);
  return { ...mounted, textarea: mounted.view.getByTestId('yrs-input'), clipboardData, written, texts };
}

test('copy writes the selection as plain text with paragraph newlines; read-only cut only copies', async () => {
  const { textarea, clipboardData, written, texts, setReadOnly } = await mountAcrossParagraphs();
  expect(fireEvent.copy(textarea, { clipboardData })).toBe(false);
  expect(written).toEqual({ 'text/plain': 'ed\nNe' });

  setReadOnly();
  delete written['text/plain'];
  expect(fireEvent.cut(textarea, { clipboardData })).toBe(false);
  expect(written).toEqual({ 'text/plain': 'ed\nNe' });
  expect(texts()).toEqual(['Seed', 'Next']);
});

test('cut copies and deletes the selection as one undo step', async () => {
  const { input, textarea, clipboardData, written, texts } = await mountAcrossParagraphs();
  fireEvent.cut(textarea, { clipboardData });
  await act(async () => { await input.current!.flushPendingInput(); });
  expect(written).toEqual({ 'text/plain': 'ed\nNe' });
  expect(texts()).toEqual(['Sext']);

  fireEvent.keyDown(textarea, { key: 'z', metaKey: true });
  await act(async () => { await input.current!.flushPendingInput(); });
  expect(texts()).toEqual(['Seed', 'Next']);
});

// S[e<embed>e]d: plain text carries a soft line break, so a cut over one
// deletes; over anything else it only copies.
for (const { name, kind, payload, copied, after } of [
  { name: 'a field', kind: 'field', payload: { fieldType: 'REF', instruction: ' REF fig ', displayText: 'Figure 1' }, copied: 'eFigure 1e', after: null },
  { name: 'an image', kind: 'image', payload: {}, copied: 'ee', after: null },
  { name: 'a soft line break', kind: 'break', payload: {}, copied: 'e\ne', after: 'Sd' },
]) {
  test(`cut over ${name} ${after ? 'deletes' : 'only copies'}`, async () => {
    const { session, input, view } = await mount();
    const [seed] = session.paragraphs('body');
    session.applyRawOps('body', [{ op: 'insertEmbed', index: 2, kind, payload }]);
    const before = session.storySegments('body');
    act(() => session.setSelection({ story: 'body', paraId: seed.paraId, offset: 1 }, { story: 'body', paraId: seed.paraId, offset: 4 }));
    const { written, clipboardData } = clipboardSpy();
    fireEvent.cut(view.getByTestId('yrs-input'), { clipboardData });
    await flush(input);
    expect(written).toEqual({ 'text/plain': copied });
    if (after) expect(session.paragraphs('body').map((p) => p.text)).toEqual([after]);
    else expect(session.storySegments('body')).toEqual(before);
  });
}

test('cut over a table only copies', async () => {
  const { session, input, view } = await mount();
  const [seed] = session.paragraphs('body');
  // Seed¶[table]¶, selected from Se|ed through the table.
  const { secondParaId: slot } = session.splitParagraph({ story: 'body', paraId: seed.paraId, offset: 4 });
  session.insertTable({ story: 'body', paraId: slot, offset: 0 }, 1, 1);
  const before = session.storySegments('body');
  // The table is the slot paragraph's one unit.
  act(() => session.setSelection({ story: 'body', paraId: seed.paraId, offset: 2 }, { story: 'body', paraId: slot, offset: 1 }));
  const { written, clipboardData } = clipboardSpy();
  fireEvent.cut(view.getByTestId('yrs-input'), { clipboardData });
  await flush(input);
  expect(written).toEqual({ 'text/plain': 'ed\n' });
  expect(session.storySegments('body')).toEqual(before);
});
