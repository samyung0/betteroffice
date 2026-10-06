/**
 * Presenter view's plain parts: the items behind Present's split button, the
 * elapsed timer and the notes text sizes. Nothing here draws.
 */

import type { PptxCommandId } from './commands';

/**
 * Present's split button and View › Present, as Google Slides' Slideshow ▾:
 * a command, its value, and the label's i18n key. The button's main part runs
 * the first.
 */
export const PRESENT_ITEMS = [
  { command: 'view.present', labelKey: 'presentation.fromCurrent' },
  { command: 'view.present', value: 'start', labelKey: 'presentation.fromStart' },
  { command: 'view.presenterView', labelKey: 'presentation.presenterView' },
] as const satisfies ReadonlyArray<{
  command: PptxCommandId;
  value?: string;
  labelKey: string;
}>;

/** The speaker notes' text sizes in px, as the mock's dropdown. */
export const NOTES_SIZES = [16, 20, 24, 32, 40] as const;
export const DEFAULT_NOTES_SIZE = 24;

/** A remembered size, or the default when it is not one of `NOTES_SIZES`. */
export function notesSize(value: unknown): number {
  const size = Number(value);
  return (NOTES_SIZES as readonly number[]).includes(size) ? size : DEFAULT_NOTES_SIZE;
}

/**
 * The elapsed timer: it runs from the show's start (Google Slides), pauses and
 * resets. `since` is when it last resumed, null while paused.
 */
export interface PresenterTimer {
  elapsed: number;
  since: number | null;
}

export const startTimer = (now: number): PresenterTimer => ({ elapsed: 0, since: now });

export const elapsedMs = (timer: PresenterTimer, now: number) =>
  timer.elapsed + (timer.since === null ? 0 : now - timer.since);

export const pauseTimer = (timer: PresenterTimer, now: number): PresenterTimer =>
  timer.since === null ? timer : { elapsed: elapsedMs(timer, now), since: null };

export const resumeTimer = (timer: PresenterTimer, now: number): PresenterTimer =>
  timer.since === null ? { elapsed: timer.elapsed, since: now } : timer;

/** Back to 0:00, still running if it was. */
export const resetTimer = (timer: PresenterTimer, now: number): PresenterTimer => ({
  elapsed: 0,
  since: timer.since === null ? null : now,
});

/** "4:07", or "1:02:09" from an hour on. */
export function formatElapsed(ms: number): string {
  const total = Math.max(0, Math.floor(ms / 1000));
  const hours = Math.floor(total / 3600);
  const minutes = Math.floor((total % 3600) / 60);
  const seconds = String(total % 60).padStart(2, '0');
  return hours > 0
    ? `${hours}:${String(minutes).padStart(2, '0')}:${seconds}`
    : `${minutes}:${seconds}`;
}
