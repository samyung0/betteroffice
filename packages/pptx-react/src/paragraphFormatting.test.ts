import { describe, expect, it } from 'bun:test';
import type { StorySnapshot, TextBoxPrimitive } from '@betteroffice/pptx';
import {
  BULLET_PRESETS,
  NUMBER_PRESETS,
  indentLevels,
  paragraphFormatting,
  presetLabel,
  presetLevels,
  presetOf,
  selectedParagraphs,
} from './paragraphFormatting';

describe('pptx list styles', () => {
  it('name each marker at its level in one style only, so indenting can find the style', () => {
    const presets = [...BULLET_PRESETS, ...NUMBER_PRESETS];
    for (let slot = 0; slot < 3; slot += 1) {
      const markers = presets.map((preset) => preset.levels[slot]);
      expect(new Set(markers).size).toBe(markers.length);
    }
    for (const preset of presets) {
      presetLevels(preset.id).forEach((bullet, level) => {
        expect(presetOf(bullet, level)).toBe(preset.id);
        expect(presetOf(bullet, level + 3)).toBe(preset.id);
      });
    }
  });

  it('label a style with its three markers', () => {
    expect(presetLabel('disc')).toBe('● ○ ■');
    expect(presetLabel('decimal')).toBe('1. a. i.');
    expect(presetLabel('paren')).toBe('1) a) i)');
    expect(presetLevels('decimal')[0]).toEqual({
      type: 'autoNumber',
      scheme: 'arabicPeriod',
      startAt: 1,
    });
  });

  it('leave a file marker in no style', () => {
    expect(presetOf({ type: 'character', value: '§' }, 0)).toBeUndefined();
    expect(presetOf({ type: 'character', value: '●' }, 1)).toBeUndefined();
    expect(
      presetOf({ type: 'autoNumber', scheme: 'arabicPeriod', startAt: 3, restart: true }, 0)
    ).toBeUndefined();
    expect(presetOf(null, 0)).toBeUndefined();
  });
});

describe('pptx paragraph formatting', () => {
  it('picks the paragraphs a range touches as the engine does', () => {
    const list = story();
    // "One" 0-3, "Two" 4-7, "" 8-8.
    expect(selectedParagraphs(list, 5, 5)).toEqual([1]);
    expect(selectedParagraphs(list, 3, 3)).toEqual([0]);
    expect(selectedParagraphs(list, 0, 4)).toEqual([0]);
    expect(selectedParagraphs(list, 2, 5)).toEqual([0, 1]);
    expect(selectedParagraphs(list, 8, 8)).toEqual([2]);
  });

  it('reports the list, style, spacing, levels and anchor of the touched paragraphs', () => {
    const list = story();
    const laidOut = textBox();
    expect(paragraphFormatting(list, laidOut, [0, 1])).toEqual({
      list: 'bullet',
      listPreset: 'disc',
      lineSpacing: '1',
      spaceBefore: false,
      spaceAfter: false,
      canIndent: true,
      canOutdent: true,
      anchor: 'ctr',
    });
    expect(paragraphFormatting(list, laidOut, [0])).toMatchObject({
      lineSpacing: '1',
      canOutdent: false,
    });
    // The third paragraph lays out at 1.5 with space before it and no list.
    expect(paragraphFormatting(list, laidOut, [2])).toMatchObject({
      list: undefined,
      listPreset: undefined,
      lineSpacing: '1.5',
      spaceBefore: true,
    });
    expect(paragraphFormatting(list, laidOut, [1, 2])).toMatchObject({
      list: undefined,
      lineSpacing: undefined,
    });
  });

  it('keeps an indent in the style of the first touched list item', () => {
    expect(indentLevels(story(), [2, 1])).toEqual(presetLevels('disc'));
    expect(indentLevels(story(), [2])).toBeNull();
  });
});

function story(): StorySnapshot {
  const run = (text: string) => ({
    text,
    style: {
      bold: null,
      italic: null,
      fontSizePt: null,
      color: null,
      fontFamily: null,
      underline: null,
    },
  });
  return {
    id: 'story',
    length: 9,
    paragraphs: [
      {
        id: 'p0',
        alignment: null,
        level: 0,
        bulletJson: '{"type":"character","value":"●"}',
        runs: [run('One')],
      },
      {
        id: 'p1',
        alignment: null,
        level: 1,
        bulletJson: '{"type":"character","value":"○"}',
        runs: [run('Two')],
      },
      { id: 'p2', alignment: null, level: 0, bulletJson: null, runs: [] },
    ],
  };
}

function textBox(): TextBoxPrimitive {
  return {
    kind: 'textBox',
    objectId: 1,
    storyId: 'story',
    x: 0,
    y: 0,
    w: 100,
    h: 100,
    anchor: 'center',
    paragraphs: [
      { level: 0, list: 'bullet', runs: [] },
      { level: 1, list: 'bullet', runs: [] },
      {
        level: 0,
        lineSpacing: { type: 'percent', value: 1.5 },
        spaceBefore: { type: 'points', value: 10 },
        runs: [],
      },
    ],
    lines: [],
  } as TextBoxPrimitive;
}
