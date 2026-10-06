import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, it, spyOn } from 'bun:test';
import { Window as HappyWindow } from 'happy-dom';
import { NotesWindow, ownNotesWindow } from './notesWindow';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
afterAll(async () => {
  if (ownsDom && GlobalRegistrator.isRegistered) await GlobalRegistrator.unregister();
});

const popup = () => new HappyWindow({ url: 'about:blank' }) as unknown as Window;

it('hands a window to its subscribers, forgets it when the user closes it, and closes it through the host', () => {
  const opened: NotesWindow[] = [];
  const closed: Window[] = [];
  const notes = new NotesWindow({
    open: (store) => opened.push(store),
    close: (window) => closed.push(window),
  });
  const seen: unknown[] = [];
  const unsubscribe = notes.subscribe(() => seen.push(notes.get()));
  notes.open();
  expect(opened).toEqual([notes]);
  notes.set('opening');
  const first = popup();
  notes.set(first);
  expect(notes.get()).toBe(first);
  first.dispatchEvent(new (first as unknown as typeof globalThis).Event('pagehide'));
  expect(notes.get()).toBeNull();
  expect(closed).toEqual([]);
  const second = popup();
  notes.set(second);
  // A late pagehide from the first window changes nothing.
  first.dispatchEvent(new (first as unknown as typeof globalThis).Event('pagehide'));
  expect(notes.get()).toBe(second);
  notes.close();
  expect(closed).toEqual([second]);
  expect(notes.get()).toBeNull();
  expect(seen).toEqual(['opening', first, null, second, null]);
  unsubscribe();
  notes.set('blocked');
  expect(seen).toHaveLength(5);
});

it("opens a page's own pop-up, and reports a blocked one", () => {
  const window_ = popup();
  const open = spyOn(window, 'open').mockImplementation(() => window_);
  try {
    const notes = ownNotesWindow();
    notes.open();
    expect(open).toHaveBeenCalledWith('', 'betteroffice-notes', 'popup,width=860,height=640');
    expect(notes.get()).toBe(window_);
    const close = spyOn(window_, 'close');
    notes.close();
    expect(close).toHaveBeenCalled();
    open.mockImplementation(() => null);
    notes.open();
    expect(notes.get()).toBe('blocked');
  } finally {
    open.mockRestore();
  }
});
