import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, it } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { initWasm, type TextBoxPrimitive } from '@betteroffice/pptx';
import { PptxEditor, type PptxEditorApi } from './PptxEditor';

const root = resolve(import.meta.dir, '../../..');
const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render, waitFor } = await import('@testing-library/react');

let fixture: Uint8Array;
let fontBytes: Uint8Array;
beforeAll(async () => {
  const [wasm, deck, font] = await Promise.all([
    readFile(resolve(root, 'packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
    readFile(resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')),
  ]);
  await initWasm(wasm);
  fixture = deck;
  fontBytes = font;
});
afterEach(cleanup);
afterAll(async () => {
  if (ownsDom && GlobalRegistrator.isRegistered) await GlobalRegistrator.unregister();
});

// A host's citation mark is display only: it fills the story's lines on its
// slide and changes nothing in the deck, its history or what peers receive.
it("fills a marked story's lines on its slide and leaves the deck untouched", async () => {
  let api: PptxEditorApi | undefined;
  let changes = 0;
  const props = {
    file: fixture,
    fonts: [{ family: 'Liberation Sans', bytes: fontBytes }],
    onChange: () => {
      changes += 1;
    },
    onReady: (ready: PptxEditorApi) => {
      api = ready;
    },
  };
  const view = render(<PptxEditor {...props} />);
  await waitFor(() => expect(api).toBeDefined(), { timeout: 15_000 });
  const { handle } = api!;
  // A text story of slide 2's own, with lines.
  const own = new Set(handle.snapshot().slides[1].shapes.flatMap((shape) => shape.textStories.map((story) => story.id)));
  const box = handle
    .layoutSlide(1)
    .primitives.find(
      (item): item is TextBoxPrimitive =>
        item.kind === 'textBox' && own.has(item.storyId ?? '') && item.lines.length > 0
    )!;
  const state = handle.encodeStateAsUpdate();
  const marks = () => view.baseElement.querySelectorAll<HTMLElement>('[data-citation-highlight]');
  const text = box.paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('\n');
  const mark = {
    color: 'rgba(255, 196, 0, 0.3)',
    holds: (shown: string) => shown === text,
    storyId: box.storyId!,
  };

  view.rerender(<PptxEditor {...props} textMark={mark} />);
  // Slide 1 shows: nothing of slide 2's story is drawn.
  expect(marks()).toHaveLength(0);
  // A host's move leaves the focus outside the frame where it is.
  const outside = document.createElement('input');
  document.body.append(outside);
  outside.focus();
  const hasFocus = document.hasFocus;
  document.hasFocus = () => false;
  try {
    await act(async () => {
      expect(api!.goToSlide(2)).toBe(true);
    });
  } finally {
    document.hasFocus = hasFocus;
  }
  expect(document.activeElement).toBe(outside);
  outside.remove();
  await waitFor(() => expect(marks()).toHaveLength(box.lines.length));
  expect(marks()[0].style.background).toBe(mark.color);
  // Q809 A: a story that no longer holds what it marks shows nothing.
  view.rerender(<PptxEditor {...props} textMark={{ ...mark, holds: () => false }} />);
  expect(marks()).toHaveLength(0);

  view.rerender(<PptxEditor {...props} textMark={null} />);
  expect(marks()).toHaveLength(0);
  expect(handle.encodeStateAsUpdate()).toEqual(state);
  expect(changes).toBe(0);
}, 60_000);
