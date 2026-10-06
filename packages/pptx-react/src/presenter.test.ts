import { expect, it } from 'bun:test';
import { PPTX_COMMAND_EDITS, PPTX_COMMAND_IDS } from './commands';
import {
  DEFAULT_NOTES_SIZE,
  NOTES_SIZES,
  PRESENT_ITEMS,
  elapsedMs,
  formatElapsed,
  notesSize,
  pauseTimer,
  resetTimer,
  resumeTimer,
  startTimer,
} from './presenter';

it("lists Present's split-button items as Google Slides' Slideshow menu, none of them editing", () => {
  expect(PRESENT_ITEMS.map((item) => [item.command, 'value' in item ? item.value : null])).toEqual([
    ['view.present', null],
    ['view.present', 'start'],
    ['view.presenterView', null],
  ]);
  for (const item of PRESENT_ITEMS) {
    expect(PPTX_COMMAND_IDS).toContain(item.command);
    expect(PPTX_COMMAND_EDITS[item.command]).toBe(false);
  }
});

it('times from the start, pauses, resumes and resets without stopping', () => {
  let timer = startTimer(1_000);
  expect(elapsedMs(timer, 66_000)).toBe(65_000);
  timer = pauseTimer(timer, 66_000);
  expect(elapsedMs(timer, 999_000)).toBe(65_000);
  expect(pauseTimer(timer, 999_000)).toBe(timer);
  timer = resumeTimer(timer, 100_000);
  expect(elapsedMs(timer, 110_000)).toBe(75_000);
  expect(resumeTimer(timer, 120_000)).toBe(timer);
  timer = resetTimer(timer, 110_000);
  expect(elapsedMs(timer, 112_000)).toBe(2_000);
  const paused = resetTimer(pauseTimer(timer, 112_000), 113_000);
  expect(paused).toEqual({ elapsed: 0, since: null });
});

it('formats elapsed time as m:ss, with hours from the first hour on', () => {
  expect(formatElapsed(0)).toBe('0:00');
  expect(formatElapsed(65_999)).toBe('1:05');
  expect(formatElapsed(3_600_000 + 2 * 60_000 + 9_000)).toBe('1:02:09');
  expect(formatElapsed(-5)).toBe('0:00');
});

it('takes a remembered notes size only when it is one of the offered sizes', () => {
  expect(NOTES_SIZES).toEqual([16, 20, 24, 32, 40]);
  expect(notesSize('32')).toBe(32);
  expect(notesSize(40)).toBe(40);
  for (const value of [null, undefined, '', '18', 'big', 400]) expect(notesSize(value)).toBe(DEFAULT_NOTES_SIZE);
  expect(DEFAULT_NOTES_SIZE).toBe(24);
});
