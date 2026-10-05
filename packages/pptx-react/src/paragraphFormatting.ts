import type {
  Bullet,
  LineSpacing,
  StorySnapshot,
  TextBoxPrimitive,
} from '@betteroffice/pptx';
import { storyTextRanges } from './textFormatting';

/**
 * Bulleted list styles, one marker per nesting level (repeating past the
 * third), as Google Slides' bullet menu. A marker is in one style only at its
 * level, so indenting can tell which style a paragraph is in. The glyphs are
 * all in Arial (WGL4), the face a bullet is written with.
 */
export const BULLET_PRESETS = [
  { id: 'disc', levels: ['●', '○', '■'] },
  { id: 'bullet', levels: ['•', '◦', '▪'] },
  { id: 'square', levels: ['■', '□', '▫'] },
  { id: 'arrow', levels: ['►', '–', '•'] },
  { id: 'dash', levels: ['–', '»', '◦'] },
] as const;

/** Numbered list styles as `ST_TextAutonumberScheme` per level. */
export const NUMBER_PRESETS = [
  { id: 'decimal', levels: ['arabicPeriod', 'alphaLcPeriod', 'romanLcPeriod'] },
  { id: 'paren', levels: ['arabicParenR', 'alphaLcParenR', 'romanLcParenR'] },
  { id: 'upperAlpha', levels: ['alphaUcPeriod', 'arabicPeriod', 'alphaLcParenR'] },
  { id: 'upperRoman', levels: ['romanUcPeriod', 'alphaUcPeriod', 'arabicPeriod'] },
] as const;

export type BulletPresetId = (typeof BULLET_PRESETS)[number]['id'];
export type NumberPresetId = (typeof NUMBER_PRESETS)[number]['id'];
export type ListPresetId = BulletPresetId | NumberPresetId;
export type ListKind = 'bullet' | 'number';

const NUMBER_SAMPLES: Record<string, string> = {
  arabicPeriod: '1.',
  arabicParenR: '1)',
  alphaLcPeriod: 'a.',
  alphaLcParenR: 'a)',
  alphaUcPeriod: 'A.',
  romanLcPeriod: 'i.',
  romanLcParenR: 'i)',
  romanUcPeriod: 'I.',
};

/** The markers a style writes, level by level. */
export function presetLevels(id: ListPresetId): Bullet[] {
  const bullets = BULLET_PRESETS.find((preset) => preset.id === id);
  if (bullets) return bullets.levels.map((value) => ({ type: 'character', value }));
  const numbers = NUMBER_PRESETS.find((preset) => preset.id === id);
  if (!numbers) throw new Error(`unknown list style ${id}`);
  return numbers.levels.map((scheme) => ({ type: 'autoNumber', scheme, startAt: 1 }));
}

/** A style's three level markers as text, e.g. "● ○ ■" or "1. a. i.". */
export function presetLabel(id: ListPresetId): string {
  return presetLevels(id)
    .map((bullet) =>
      bullet.type === 'character'
        ? bullet.value
        : bullet.type === 'autoNumber'
          ? (NUMBER_SAMPLES[bullet.scheme] ?? bullet.scheme)
          : ''
    )
    .join(' ');
}

function sameBullet(left: Bullet, right: Bullet): boolean {
  if (left.type === 'character' && right.type === 'character') return left.value === right.value;
  if (left.type === 'autoNumber' && right.type === 'autoNumber') {
    return left.scheme === right.scheme && left.startAt === right.startAt && !left.restart;
  }
  return false;
}

/** The style whose marker for `level` is `bullet`, if any. */
export function presetOf(bullet: Bullet | null, level: number): ListPresetId | undefined {
  if (!bullet) return undefined;
  for (const preset of [...BULLET_PRESETS, ...NUMBER_PRESETS]) {
    const levels = presetLevels(preset.id);
    if (sameBullet(levels[level % levels.length], bullet)) return preset.id;
  }
  return undefined;
}

export function paragraphBullet(story: StorySnapshot, index: number): Bullet | null {
  const json = story.paragraphs[index]?.bulletJson;
  if (!json) return null;
  try {
    return JSON.parse(json) as Bullet;
  } catch {
    return null;
  }
}

/** The paragraphs `[anchor, focus]` touches, as the engine's paragraph edits
 *  pick them: a caret its own, a range not the one it stops at the start of. */
export function selectedParagraphs(story: StorySnapshot, anchor: number, focus: number): number[] {
  const ranges = storyTextRanges(story);
  const start = Math.min(anchor, focus);
  const end = Math.max(anchor, focus);
  const touched = ranges
    .map((range, index) => ({ range, index }))
    .filter(({ range }) =>
      start === end ? start >= range.start && start <= range.end : start <= range.end && end > range.start
    )
    .map(({ index }) => index);
  return touched.length > 0 ? touched : [Math.max(0, ranges.length - 1)];
}

/** What the toolbar and menus show for the touched paragraphs of a text box. */
export interface ParagraphFormatting {
  list?: ListKind;
  /** The style every touched list item is in, for the style menus' ticks. */
  listPreset?: ListPresetId;
  /** Line spacing shared by every touched paragraph, as a share ('1.5'). */
  lineSpacing?: string;
  spaceBefore: boolean;
  spaceAfter: boolean;
  canIndent: boolean;
  canOutdent: boolean;
  anchor?: 't' | 'ctr' | 'b';
}

function common<T>(values: T[]): T | undefined {
  return values.length > 0 && values.every((value) => value === values[0]) ? values[0] : undefined;
}

function hasSpace(spacing: LineSpacing | undefined): boolean {
  return spacing !== undefined && spacing.value > 0;
}

export function paragraphFormatting(
  story: StorySnapshot,
  textBox: TextBoxPrimitive | undefined,
  indices: number[]
): ParagraphFormatting {
  const laidOut = indices.map((index) => textBox?.paragraphs[index]);
  const levels = indices.map((index) => story.paragraphs[index]?.level ?? 0);
  const list = common(laidOut.map((paragraph) => paragraph?.list));
  const presets = indices.map((index, position) =>
    presetOf(paragraphBullet(story, index), levels[position])
  );
  const spacing = laidOut.map((paragraph) =>
    // No `a:lnSpc` anywhere lays out single.
    paragraph?.lineSpacing === undefined
      ? '1'
      : paragraph.lineSpacing.type === 'percent'
        ? String(Math.round(paragraph.lineSpacing.value * 100) / 100)
        : undefined
  );
  return {
    list,
    listPreset: list ? common(presets) : undefined,
    lineSpacing: common(spacing),
    spaceBefore: laidOut.every((paragraph) => hasSpace(paragraph?.spaceBefore)),
    spaceAfter: laidOut.every((paragraph) => hasSpace(paragraph?.spaceAfter)),
    canIndent: levels.some((level) => level < 8),
    canOutdent: levels.some((level) => level > 0),
    anchor: textBox
      ? ({ top: 't', center: 'ctr', bottom: 'b' } as const)[textBox.anchor]
      : undefined,
  };
}

/** The style an indent keeps a list item in: the first touched item's. */
export function indentLevels(story: StorySnapshot, indices: number[]): Bullet[] | null {
  for (const index of indices) {
    const preset = presetOf(paragraphBullet(story, index), story.paragraphs[index]?.level ?? 0);
    if (preset) return presetLevels(preset);
  }
  return null;
}
