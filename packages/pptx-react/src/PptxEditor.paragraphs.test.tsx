import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, it } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { initWasm } from '@betteroffice/pptx';
import type { PptxCommandState } from './commands';
import type { PptxEditorApi } from './PptxEditor';
import { PptxEditor } from './PptxEditor';

const root = resolve(import.meta.dir, '../../..');
const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');

let fixture: Uint8Array;
let lecture: Uint8Array;
let listStyles: Uint8Array;
let table: Uint8Array;
let fonts: { family: string; bytes: Uint8Array }[];

beforeAll(async () => {
  const [wasm, pptx, font, lectureDeck, listDeck, tableDeck] = await Promise.all([
    readFile(resolve(root, 'packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
    readFile(resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')),
    readFile(resolve(root, 'poc/fixtures/lecture.pptx')),
    readFile(resolve(root, 'crates/pptx-render/tests/fixtures/list-style-bullets.pptx')),
    readFile(resolve(root, 'crates/pptx-render/tests/fixtures/table-basic.pptx')),
  ]);
  await initWasm(wasm);
  fixture = pptx;
  lecture = lectureDeck;
  listStyles = listDeck;
  table = tableDeck;
  fonts = [{ family: 'Arial', bytes: font }];
});

afterEach(async () => {
  cleanup();
  // The caret blinks on a timer; let React finish what it scheduled while the DOM exists.
  await new Promise((done) => setTimeout(done, 50));
});
afterAll(async () => {
  if (ownsDom && GlobalRegistrator.isRegistered) await GlobalRegistrator.unregister();
});

async function open(file = fixture) {
  const opened: PptxEditorApi[] = [];
  const states: PptxCommandState[] = [];
  const errors: Error[] = [];
  const view = render(
    <PptxEditor
      file={file}
      fonts={fonts}
      onReady={(api) => opened.push(api)}
      onCommandState={(state) => states.push(state)}
      onError={(error) => errors.push(error)}
    />
  );
  await waitFor(() => expect(opened.length).toBe(1), { timeout: 15_000 });
  await waitFor(() => expect(states[states.length - 1]?.enabled['view.zoom']).toBe(true));
  const api = opened[0];
  const title = api.handle.snapshot().slides[0].shapes.find((shape) => shape.name === 'Title');
  const storyId = title?.textStories[0].id ?? '';
  const story = () => api.handle.story(storyId);
  const run = (id: Parameters<PptxEditorApi['runCommand']>[0], value?: string) => {
    let ran = false;
    act(() => {
      ran = api.runCommand(id, value);
    });
    return ran;
  };
  const caret = (start: number, end = start) =>
    act(() => {
      api.selectText({ slide: 1, shapeId: title!.id, storyId, start, end });
    });
  const input = () => view.getByTestId('pptx-text-input');
  const key = (init: { key: string; code?: string; shiftKey?: boolean; metaKey?: boolean }) =>
    act(() => {
      fireEvent.keyDown(input(), init);
    });
  const bullets = () =>
    story().paragraphs.map((paragraph) =>
      paragraph.bulletJson ? (JSON.parse(paragraph.bulletJson) as { value?: string; type: string }) : null
    );
  const markers = () =>
    bullets().map((bullet) => (bullet?.type === 'character' ? bullet.value : bullet?.type ?? null));
  return {
    api,
    caret,
    errors,
    input,
    key,
    markers,
    run,
    state: () => states[states.length - 1],
    story,
    title,
    view,
  };
}

/** "Office files," then "without the office.": the first starts at 0, the second at 14. */
const SECOND = 14;

describe('PptxEditor lists and paragraph formatting', () => {
  it('toggles a list, picks a style and reports both to host menus', async () => {
    const { caret, errors, markers, run, state } = await open();
    caret(0, SECOND + 3);
    expect(run('format.bulletedList')).toBe(true);
    expect(markers()).toEqual(['●', '●']);
    await waitFor(() => expect(state().checked).toContain('format.bulletedList'));
    expect(state().listStyle).toBe('disc');

    expect(run('format.numberedList', 'paren')).toBe(true);
    expect(markers()).toEqual(['autoNumber', 'autoNumber']);
    await waitFor(() => expect(state().listStyle).toBe('paren'));
    expect(run('format.numberedList', 'emoji')).toBe(false);

    // Without a style the button toggles the list off again.
    expect(run('format.numberedList')).toBe(true);
    expect(markers()).toEqual(['none', 'none']);
    await waitFor(() => expect(state().checked).not.toContain('format.numberedList'));
    expect(errors).toEqual([]);
  }, 60_000);

  it('ends the list on Enter in an empty item and removes a marker on Backspace at its start', async () => {
    const { caret, errors, key, markers, run, story } = await open();
    caret(0, SECOND + 3);
    run('format.bulletedList');
    const end = story().length - 1;
    caret(end);
    key({ key: 'Enter' });
    expect(markers()).toEqual(['●', '●', '●']);
    // Enter again on the new, empty item leaves the list without a new paragraph.
    key({ key: 'Enter' });
    expect(markers()).toEqual(['●', '●', 'none']);
    expect(story().paragraphs.length).toBe(3);

    // Backspace at an item's start takes its marker first, then joins.
    caret(SECOND);
    key({ key: 'Backspace' });
    expect(markers()).toEqual(['●', 'none', 'none']);
    expect(story().paragraphs.length).toBe(3);
    key({ key: 'Backspace' });
    expect(story().paragraphs.length).toBe(2);
    expect(story().paragraphs[0].runs.map((run) => run.text).join('')).toBe(
      'Office files,without the office.'
    );
    expect(errors).toEqual([]);
  }, 60_000);

  it('indents a list item with Tab at its start and steps a nested empty item out on Enter', async () => {
    const { caret, errors, key, markers, run, story } = await open();
    caret(0, SECOND + 3);
    run('format.bulletedList');
    caret(SECOND);
    key({ key: 'Tab' });
    expect(story().paragraphs[1].level).toBe(1);
    expect(markers()).toEqual(['●', '○']);
    key({ key: 'Tab', shiftKey: true });
    expect(story().paragraphs[1].level).toBe(0);
    expect(markers()).toEqual(['●', '●']);

    // Mid-line, Tab types a tab.
    caret(3);
    key({ key: 'Tab' });
    expect(story().paragraphs[0].runs.map((run) => run.text).join('')).toBe('Off\tice files,');

    // An empty nested item steps out a level before it leaves the list.
    caret(story().length - 1);
    key({ key: 'Enter' });
    key({ key: 'Tab' });
    expect(story().paragraphs[2].level).toBe(1);
    key({ key: 'Enter' });
    expect(story().paragraphs[2].level).toBe(0);
    expect(markers()[2]).toBe('●');
    key({ key: 'Enter' });
    expect(markers()[2]).toBe('none');
    expect(errors).toEqual([]);
  }, 60_000);

  it('runs the Google Slides shortcuts for lists, indent and strikethrough', async () => {
    const { caret, errors, key, markers, story } = await open();
    caret(0, SECOND + 3);
    key({ key: '*', code: 'Digit8', shiftKey: true, metaKey: true });
    expect(markers()).toEqual(['●', '●']);
    key({ key: ']', code: 'BracketRight', metaKey: true });
    expect(story().paragraphs.map((paragraph) => paragraph.level)).toEqual([1, 1]);
    key({ key: '[', code: 'BracketLeft', metaKey: true });
    expect(story().paragraphs.map((paragraph) => paragraph.level)).toEqual([0, 0]);
    key({ key: 'x', code: 'KeyX', shiftKey: true, metaKey: true });
    expect(story().paragraphs[0].runs[0].style.strike).toBe('sngStrike');
    key({ key: '.', code: 'Period', metaKey: true });
    expect(story().paragraphs[0].runs[0].style.baselinePct).toBe(30);
    key({ key: '\\', code: 'Backslash', metaKey: true });
    expect(story().paragraphs[0].runs[0].style.strike ?? null).toBeNull();
    expect(story().paragraphs[0].runs[0].style.baselinePct ?? null).toBeNull();
    expect(errors).toEqual([]);
  }, 60_000);

  it('sets spacing, vertical alignment and highlight from host commands', async () => {
    const { api, caret, errors, run, state, story, title } = await open();
    caret(2);
    expect(run('format.lineSpacing', '1.5')).toBe(true);
    expect(run('format.lineSpacing', '3')).toBe(false);
    expect(story().paragraphs[0].lineSpacing).toEqual({ type: 'percent', value: 1.5 });
    expect(story().paragraphs[1].lineSpacing).toBeUndefined();
    await waitFor(() => expect(state().lineSpacing).toBe('1.5'));
    expect(run('format.spaceBefore')).toBe(true);
    expect(story().paragraphs[0].spaceBefore).toEqual({ type: 'points', value: 10 });
    await waitFor(() => expect(state().checked).toContain('format.spaceBefore'));
    expect(run('format.spaceBefore')).toBe(true);
    expect(story().paragraphs[0].spaceBefore).toEqual({ type: 'points', value: 0 });

    expect(run('format.alignBottom')).toBe(true);
    const shape = () =>
      api.handle.snapshot().slides[0].shapes.find((candidate) => candidate.id === title!.id)!;
    expect(shape().textAnchor).toBe('b');
    await waitFor(() => expect(state().checked).toContain('format.alignBottom'));
    expect(errors).toEqual([]);
  }, 60_000);

  it('selects all text while typing, else every object, which delete and undo as one step', async () => {
    const { api, caret, errors, input, key, run, story, view } = await open();
    caret(3);
    expect(run('edit.selectAll')).toBe(true);
    expect(run('format.bold')).toBe(true);
    expect(run('format.bold')).toBe(true);
    // Bold toggled over the whole story: every run is set the same way.
    const weights = story().paragraphs.flatMap((paragraph) => paragraph.runs.map((run) => run.style.bold));
    expect(new Set(weights).size).toBe(1);

    act(() => api.clearSelection());
    const visible = api.handle.snapshot().slides[0].shapes.filter((shape) => !shape.hidden);
    act(() => {
      fireEvent.keyDown(view.getByRole('application'), { key: 'a', code: 'KeyA', metaKey: true });
    });
    expect(view.getAllByTestId('pptx-multi-selection').length).toBeGreaterThan(1);
    expect(run('edit.delete')).toBe(true);
    expect(api.handle.snapshot().slides[0].shapes.length).toBe(0);
    expect(run('edit.undo')).toBe(true);
    expect(api.handle.snapshot().slides[0].shapes.length).toBe(visible.length);
    void input;
    void key;
    expect(errors).toEqual([]);
  }, 60_000);

  it('aligns, centres and distributes the selected objects', async () => {
    const { api, errors, run, view } = await open();
    const shapes = () => api.handle.snapshot().slides[0].shapes.filter((shape) => shape.width > 0);
    act(() => {
      fireEvent.keyDown(view.getByRole('application'), { key: 'a', code: 'KeyA', metaKey: true });
    });
    expect(run('arrange.align', 'diagonal')).toBe(false);
    expect(run('arrange.align', 'left')).toBe(true);
    const lefts = new Set(shapes().map((shape) => shape.x));
    expect(lefts.size).toBe(1);
    // Distributed in the order the objects start in, with one gap between each.
    const order = [...shapes()].sort((a, b) => a.y - b.y).map((shape) => shape.id);
    expect(run('arrange.distribute', 'vertical')).toBe(true);
    const placed = order.map((id) => shapes().find((shape) => shape.id === id)!);
    const gaps = placed.slice(1).map((shape, index) => shape.y - (placed[index].y + placed[index].height));
    expect(Math.max(...gaps) - Math.min(...gaps)).toBeLessThanOrEqual(2);
    // One step undoes the whole distribution.
    expect(run('edit.undo')).toBe(true);
    expect(new Set(shapes().map((shape) => shape.x)).size).toBe(1);
    expect(errors).toEqual([]);
  }, 60_000);
});

describe('PptxEditor list styles on existing list items', () => {
  /** Selects `start` in the first story of `slide` whose text contains `text`. */
  function selectIn(api: PptxEditorApi, slide: number, text: string) {
    const shape = api.handle
      .snapshot()
      .slides[slide - 1].shapes.find((candidate) =>
        candidate.textStories.some((story) =>
          story.paragraphs.some((paragraph) => paragraph.runs.some((run) => run.text.includes(text)))
        )
      )!;
    const story = shape.textStories[0];
    let start = 0;
    for (const paragraph of story.paragraphs) {
      const content = paragraph.runs.map((run) => run.text).join('');
      if (content.includes(text)) break;
      start += content.length + 1;
    }
    act(() => {
      api.selectText({ slide, shapeId: shape.id, storyId: story.id, start, end: start });
    });
    return { index: story.paragraphs.findIndex((paragraph) => paragraph.runs.some((run) => run.text.includes(text))), storyId: story.id };
  }

  it("changes a file list item's marker, not its indents", async () => {
    const { api, errors, run } = await open(lecture);
    act(() => {
      api.goToSlide(2);
    });
    const { index, storyId } = selectIn(api, 2, 'Ad minim veniam');
    expect(run('format.bulletedList', 'dash')).toBe(true);
    expect(run('format.numberedList', 'decimal')).toBe(true);
    const paragraph = api.handle.story(storyId).paragraphs[index];
    expect(JSON.parse(paragraph.bulletJson!).type).toBe('autoNumber');
    expect(paragraph.marginLeft).toBeUndefined();
    expect(paragraph.indent).toBeUndefined();
    expect(errors).toEqual([]);
  }, 60_000);

  it('keeps the indents of an item that inherits its marker', async () => {
    const { api, errors, run } = await open(listStyles);
    const { index, storyId } = selectIn(api, 1, 'First level');
    expect(run('format.bulletedList', 'square')).toBe(true);
    const paragraph = api.handle.story(storyId).paragraphs[index];
    expect(JSON.parse(paragraph.bulletJson!)).toEqual({ type: 'character', value: '■' });
    expect(paragraph.marginLeft).toBeUndefined();
    expect(errors).toEqual([]);
  }, 60_000);

  it('leaves Tab in a table cell to the browser', async () => {
    const { api, errors, view } = await open(table);
    const frame = api.handle.snapshot().slides[0].shapes.find((shape) => shape.kind === 'graphicFrame')!;
    const cell = frame.textStories[0];
    act(() => {
      api.selectText({ slide: 1, shapeId: frame.id, storyId: cell.id, start: 0, end: 0 });
    });
    const before = JSON.stringify(api.handle.story(cell.id));
    let handled = false;
    act(() => {
      handled = !fireEvent.keyDown(view.getByTestId('pptx-text-input'), { key: 'Tab' });
    });
    expect(handled).toBe(false);
    expect(JSON.stringify(api.handle.story(cell.id))).toBe(before);
    expect(errors).toEqual([]);
  }, 60_000);
});
