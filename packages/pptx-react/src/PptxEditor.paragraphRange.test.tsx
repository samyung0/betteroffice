import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, it } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import * as pptx from '@betteroffice/pptx';
import type { StorySnapshot } from '@betteroffice/pptx';
import type { PptxEditorApi } from './PptxEditor';
import { PptxEditor } from './PptxEditor';

const root = resolve(import.meta.dir, '../../..');
// The demo deck's first slide title: "Office files," then "without the office.".
const STORY = 'story:slide:0:256:shape:2:0';
// "fice files," + the paragraph break + "with".
const START = 2;
const END = 'Office files,'.length + 5;

// the registrator writes one process-wide global set, so only the file that
// installed it may tear it down.
const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');

let fixture: Uint8Array;
let fonts: pptx.PptxFontFace[];

beforeAll(async () => {
  const [wasm, deck, font] = await Promise.all([
    readFile(resolve(root, 'packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
    readFile(resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')),
  ]);
  await pptx.initWasm(wasm);
  fixture = deck;
  fonts = [{ family: 'Liberation Sans', bytes: font }];
});

afterEach(cleanup);
afterAll(async () => {
  if (ownsDom && GlobalRegistrator.isRegistered) await GlobalRegistrator.unregister();
});

const plain = (story: StorySnapshot) =>
  story.paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join('')).join('\n');

describe('PptxEditor edits across paragraphs', () => {
  it('types, pastes, deletes, splits and composes over the selection, undoes each in one step, and keeps saving after a refusal', async () => {
    const opened: PptxEditorApi[] = [];
    const errors: Error[] = [];
    const view = render(<PptxEditor file={fixture} fonts={fonts} onReady={(api) => opened.push(api)} onError={(error) => errors.push(error)} />);
    await waitFor(() => expect(opened.length).toBe(1), { timeout: 15_000 });
    const api = opened[0];
    const shape = api.handle.snapshot().slides[0].shapes.find((item) => item.textStories.some((story) => story.id === STORY))!;
    const before = api.handle.story(STORY);
    expect(plain(before)).toBe('Office files,\nwithout the office.');
    const input = view.getByTestId('pptx-text-input') as HTMLTextAreaElement;
    const select = () => act(async () => {
      expect(api.selectText({ slide: 1, shapeId: shape.id, storyId: STORY, start: START, end: END })).toBe(true);
    });

    const edits: Array<[string, () => void | Promise<void>, string]> = [
      ['typing', () => { fireEvent.input(input, { target: { value: 'x' } }); }, 'Ofxout the office.'],
      ['paste', () => { fireEvent.paste(input, { clipboardData: { getData: () => 'pasted ' } }); }, 'Ofpasted out the office.'],
      ['multi-line paste', () => { fireEvent.paste(input, { clipboardData: { getData: () => 'one\r\ntwo ' } }); }, 'Ofone\ntwo out the office.'],
      ['Backspace', () => { fireEvent.keyDown(input, { key: 'Backspace' }); }, 'Ofout the office.'],
      ['Delete', () => { fireEvent.keyDown(input, { key: 'Delete' }); }, 'Ofout the office.'],
      ['Enter', () => { fireEvent.keyDown(input, { key: 'Enter' }); }, 'Of\nout the office.'],
      ['composition', async () => {
        fireEvent.compositionStart(input);
        input.value = '日本';
        fireEvent.compositionEnd(input, { data: '日本' });
        await api.flushPendingInput();
      }, 'Of日本out the office.'],
    ];
    for (const [name, edit, expected] of edits) {
      await select();
      await act(async () => { await edit(); });
      expect([name, plain(api.handle.story(STORY))]).toEqual([name, expected]);
      const { paragraphs } = api.handle.story(STORY);
      expect([name, paragraphs[0].id]).toEqual([name, before.paragraphs[0].id]);
      expect([name, ...paragraphs.map((paragraph) => paragraph.alignment)]).toEqual([name, ...paragraphs.map(() => before.paragraphs[0].alignment)]);
      fireEvent.keyDown(input, { key: 'z', ctrlKey: true });
      expect([name, api.handle.story(STORY)]).toEqual([name, before]);
    }
    expect(errors).toEqual([]);

    await select();
    fireEvent.paste(input, { clipboardData: { getData: () => 'one\ntwo\u0001' } });
    expect(errors.map((error) => error.message)).toEqual([expect.stringContaining('cannot be stored')]);
    expect(api.handle.story(STORY)).toEqual(before);
    await act(async () => { await api.flushPendingInput(); });
    expect(api.save().byteLength).toBeGreaterThan(0);

    fireEvent.paste(input, { clipboardData: { getData: () => 'one\ntwo ' } });
    await act(async () => { await api.flushPendingInput(); });
    const reopened = pptx.openPresentation(api.save(), { fonts });
    try {
      expect(plain(reopened.story(STORY))).toBe('Ofone\ntwo out the office.');
    } finally {
      reopened.dispose();
    }
  }, 60_000);

  it('restores the paragraph and the saved slide exactly after Enter then Backspace', async () => {
    const opened: PptxEditorApi[] = [];
    const errors: Error[] = [];
    const view = render(<PptxEditor file={fixture} fonts={fonts} onReady={(api) => opened.push(api)} onError={(error) => errors.push(error)} />);
    await waitFor(() => expect(opened.length).toBe(1), { timeout: 15_000 });
    const api = opened[0];
    const shape = api.handle.snapshot().slides[0].shapes.find((item) => item.textStories.some((story) => story.id === STORY))!;
    const before = api.handle.story(STORY);
    const slide = async (bytes: Uint8Array) => (await JSZip.loadAsync(bytes)).file('ppt/slides/slide1.xml')!.async('string');
    const source = await slide(api.save());
    const input = view.getByTestId('pptx-text-input') as HTMLTextAreaElement;
    const caret = 'Office'.length;
    await act(async () => {
      expect(api.selectText({ slide: 1, shapeId: shape.id, storyId: STORY, start: caret, end: caret })).toBe(true);
    });

    fireEvent.keyDown(input, { key: 'Enter' });
    const split = api.handle.story(STORY);
    expect(plain(split)).toBe('Office\n files,\nwithout the office.');
    expect(split.paragraphs[0].id).toBe(before.paragraphs[0].id);
    fireEvent.keyDown(input, { key: 'Backspace' });
    expect(api.handle.story(STORY)).toEqual(before);
    await act(async () => { await api.flushPendingInput(); });
    expect(await slide(api.save())).toBe(source);
    expect(errors).toEqual([]);
  }, 60_000);
});
