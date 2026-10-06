import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, it, spyOn } from 'bun:test';
import type { SlideDisplayList } from '@betteroffice/pptx/viewer';
import { Window as HappyWindow } from 'happy-dom';
import { NotesWindow } from '../notesWindow';
import { PresentationOverlay } from './PresentationOverlay';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => {
  if (ownsDom && GlobalRegistrator.isRegistered) await GlobalRegistrator.unregister();
});

const blank: SlideDisplayList = { contractVersion: 1, width: 320, height: 180, primitives: [] };

/** A notes window host that records what the show asks of it. */
function setup({
  slideCount = 3,
  defaultNotesSize,
}: { slideCount?: number; defaultNotesSize?: number } = {}) {
  const calls = { opened: 0, closed: [] as Window[], exits: [] as number[], sizes: [] as number[] };
  const notes = new NotesWindow({
    open: () => {
      calls.opened++;
    },
    close: (window) => calls.closed.push(window),
  });
  const view = render(
    <PresentationOverlay
      handle={{ layoutSlide: () => blank }}
      slideCount={slideCount}
      startIndex={0}
      resolveImage={() => null}
      notesWindow={notes}
      notesFor={(index) => `Notes for slide ${index + 1}`}
      defaultNotesSize={defaultNotesSize}
      onNotesSizeChange={(size) => calls.sizes.push(size)}
      onExit={(index) => calls.exits.push(index)}
      onError={(error) => {
        throw error;
      }}
    />
  );
  const popup = new HappyWindow({ url: 'about:blank' }) as unknown as Window;
  const attach = () => act(() => notes.set(popup));
  const inPopup = (selector: string) => popup.document.body.querySelector(selector);
  const popupText = () => popup.document.body.textContent ?? '';
  // Keys come from the window the element lives in.
  const press = (target: Element, key: string) =>
    act(() => {
      const realm = target.ownerDocument.defaultView as unknown as typeof globalThis;
      target.dispatchEvent(new realm.KeyboardEvent('keydown', { key, bubbles: true, cancelable: true }));
    });
  const button = (name: string) =>
    [...popup.document.body.querySelectorAll('button')].find((each) => each.textContent === name)!;
  return { calls, notes, view, popup, attach, inPopup, popupText, press, button };
}

it('draws Presenter view in the notes window, and both windows drive one slide', () => {
  const { view, attach, inPopup, popupText, popup, press } = setup();
  attach();
  expect(popup.document.title).toBe('Presenter view');
  expect(popupText()).toContain('Slide 1 of 3');
  expect(inPopup('[data-testid="pptx-presenter-notes"]')?.textContent).toBe('Notes for slide 1');
  expect(view.getByText('1 / 3')).toBeTruthy();

  fireEvent.click(inPopup('[aria-label="Next slide"]')!);
  expect(view.getByText('2 / 3')).toBeTruthy();
  expect(popupText()).toContain('Slide 2 of 3');
  expect(inPopup('[data-testid="pptx-presenter-notes"]')?.textContent).toBe('Notes for slide 2');

  press(popup.document.body, 'ArrowRight');
  expect(view.getByText('3 / 3')).toBeTruthy();
  expect(popupText()).toContain('End of slides');

  press(view.getByRole('dialog'), 'ArrowLeft');
  expect(popupText()).toContain('Slide 2 of 3');
  press(popup.document.body, 'Home');
  expect(view.getByText('1 / 3')).toBeTruthy();
  press(view.getByRole('dialog'), 'End');
  expect(popupText()).toContain('Slide 3 of 3');
  press(popup.document.body, 'PageUp');
  expect(view.getByText('2 / 3')).toBeTruthy();

  // The size dropdown keeps its own arrow keys.
  press(inPopup('select')!, 'ArrowDown');
  expect(view.getByText('2 / 3')).toBeTruthy();
});

it('ends the show at its slide from End or Esc in either window, and the notes window closes', () => {
  const first = setup();
  first.attach();
  first.press(first.popup.document.body, 'ArrowRight');
  fireEvent.click(first.button('End'));
  expect(first.calls.exits).toEqual([1]);
  first.view.unmount();
  expect(first.calls.closed.length === 1 && first.calls.closed[0] === first.popup).toBe(true);
  expect(first.notes.get() === null).toBe(true);
  cleanup();

  const second = setup();
  second.attach();
  second.press(second.popup.document.body, 'Escape');
  expect(second.calls.exits).toEqual([0]);
  cleanup();

  const third = setup();
  third.attach();
  fireEvent.click(third.view.getByRole('button', { name: 'Exit presentation (Esc)' }));
  expect(third.calls.exits).toEqual([0]);
});

it('keeps the show when the notes window is closed, and ⋮ opens it again', () => {
  const { view, attach, calls, popup } = setup();
  attach();
  expect(Boolean(view.queryByRole('button', { name: 'More actions' }))).toBe(false);
  act(() => {
    popup.dispatchEvent(new (popup as unknown as typeof globalThis).Event('pagehide'));
  });
  expect(Boolean(view.getByRole('dialog'))).toBe(true);
  expect(calls.exits).toEqual([]);
  fireEvent.click(view.getByRole('button', { name: 'More actions' }));
  fireEvent.click(view.getByRole('menuitem', { name: 'Open speaker notes' }));
  expect(calls.opened).toBe(1);
  expect(Boolean(view.queryByRole('menu'))).toBe(false);
});

it('runs the show when the browser blocks the notes window, with a notice that opens it from a click', () => {
  const { view, notes, calls } = setup();
  act(() => notes.set('blocked'));
  expect(view.getByRole('alert').textContent).toContain(
    'The browser blocked the speaker notes window.'
  );
  fireEvent.click(view.getByRole('button', { name: 'Open speaker notes' }));
  expect(calls.opened).toBe(1);
  expect(calls.exits).toEqual([]);
  expect(view.getByText('1 / 3')).toBeTruthy();
});

it('times from the start of the show, with pause, resume and reset', async () => {
  let now = 0;
  const clock = spyOn(Date, 'now').mockImplementation(() => now);
  try {
    const { attach, inPopup } = setup();
    now = 30_000;
    attach();
    const timer = () => inPopup('[role="timer"]')?.textContent;
    expect(timer()).toBe('0:30');
    now = 65_000;
    fireEvent.click(inPopup('[aria-label="Pause timer"]')!);
    expect(timer()).toBe('1:05');
    now = 300_000;
    fireEvent.click(inPopup('[aria-label="Next slide"]')!);
    expect(timer()).toBe('1:05');
    fireEvent.click(inPopup('[aria-label="Resume timer"]')!);
    now = 310_000;
    fireEvent.click(inPopup('[aria-label="Reset timer"]')!);
    expect(timer()).toBe('0:00');
    now = 312_000;
    await waitFor(() => expect(timer()).toBe('0:02'), { timeout: 2_500 });
  } finally {
    clock.mockRestore();
  }
});

it('shows the notes at the remembered size and reports a new one', () => {
  const { attach, inPopup, calls } = setup({ defaultNotesSize: 40 });
  attach();
  const notes = () => inPopup('[data-testid="pptx-presenter-notes"]') as HTMLElement;
  expect(notes().style.fontSize).toBe('40px');
  fireEvent.change(inPopup('select')!, { target: { value: '20' } });
  expect(calls.sizes).toEqual([20]);
  expect(notes().style.fontSize).toBe('20px');
  cleanup();

  const fallback = setup({ defaultNotesSize: 18 });
  fallback.attach();
  expect((fallback.inPopup('select') as HTMLSelectElement).value).toBe('24');
});

it('ends when full screen is left, except when the show leaves it to open the notes window', async () => {
  let fullscreen: Element | null = null;
  const change = () => document.dispatchEvent(new Event('fullscreenchange'));
  const descriptor = Object.getOwnPropertyDescriptor(Document.prototype, 'fullscreenElement');
  Object.defineProperty(document, 'fullscreenElement', {
    configurable: true,
    get: () => fullscreen,
  });
  const request = HTMLElement.prototype.requestFullscreen;
  const exitFullscreen = document.exitFullscreen;
  HTMLElement.prototype.requestFullscreen = async function (this: HTMLElement) {
    fullscreen = this;
    change();
  };
  document.exitFullscreen = async () => {
    fullscreen = null;
    change();
  };
  try {
    const { view, calls, attach, popupText, notes } = setup();
    await waitFor(() => expect(fullscreen !== null).toBe(true));
    expect(Boolean(view.queryByRole('button', { name: 'Full screen' }))).toBe(false);
    attach();
    expect(popupText()).toContain('Slides in full screen');
    act(() => notes.set(null));

    fireEvent.click(view.getByRole('button', { name: 'More actions' }));
    fireEvent.click(view.getByRole('menuitem', { name: 'Open speaker notes' }));
    await waitFor(() => expect(fullscreen === null).toBe(true));
    expect(calls.opened).toBe(1);
    expect(calls.exits).toEqual([]);
    const button = await view.findByRole('button', { name: 'Full screen' });
    attach();
    expect(popupText()).toContain('Slides not in full screen');

    fireEvent.click(button);
    await waitFor(() => expect(popupText()).toContain('Slides in full screen'));
    // Esc in full screen: the browser leaves it, and the show ends.
    act(() => {
      fullscreen = null;
      change();
    });
    expect(calls.exits).toEqual([0]);
  } finally {
    HTMLElement.prototype.requestFullscreen = request;
    document.exitFullscreen = exitFullscreen;
    if (descriptor) Object.defineProperty(document, 'fullscreenElement', descriptor);
    else delete (document as { fullscreenElement?: unknown }).fullscreenElement;
  }
});
