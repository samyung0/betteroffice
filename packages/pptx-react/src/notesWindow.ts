/**
 * Presenter view's speaker notes window, which the show draws into (one deck,
 * painted by the same code in both windows). Opening and closing it belong to
 * the host: a page opens its own (`ownNotesWindow`), a sandboxed runtime asks
 * the page around it, which hands the window over with `set`.
 */

/** A window, 'opening' while one is on its way, 'blocked' when the browser refused it. */
export type NotesWindowState = Window | 'opening' | 'blocked' | null;

export interface NotesWindowHost {
  /** Called from a click in the show (user activation); the window comes through `set`. */
  open(notes: NotesWindow): void;
  /** Closes a window `set` gave. */
  close(window: Window): void;
}

export class NotesWindow {
  private state: NotesWindowState = null;
  private readonly listeners = new Set<() => void>();

  constructor(private readonly host: NotesWindowHost) {}

  get = (): NotesWindowState => this.state;

  subscribe = (listener: () => void) => {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  };

  /** A window arrived, is coming, or was blocked; one the user closes resets to null. */
  set(state: NotesWindowState) {
    if (state === this.state) return;
    this.state = state;
    if (isWindow(state))
      state.addEventListener(
        'pagehide',
        () => {
          if (this.state === state) this.set(null);
        },
        { once: true }
      );
    for (const listener of this.listeners) listener();
  }

  open() {
    this.host.open(this);
  }

  /** The show ended: the window closes and nothing more is expected. */
  close() {
    const current = this.state;
    this.set(null);
    if (isWindow(current)) this.host.close(current);
  }
}

/** Windows from another realm fail `instanceof Window`. */
export function isWindow(state: NotesWindowState): state is Window {
  return typeof state === 'object' && state !== null;
}

/** For a page that is its own top window: a same-origin pop-up it opens and closes. */
export function ownNotesWindow(name = 'betteroffice-notes'): NotesWindow {
  return new NotesWindow({
    open: (notes) => notes.set(window.open('', name, NOTES_WINDOW_FEATURES) ?? 'blocked'),
    close: (opened) => opened.close(),
  });
}

/** The pop-up's size: Google Slides' presenter window, laid out notes first. */
export const NOTES_WINDOW_FEATURES = 'popup,width=860,height=640';
