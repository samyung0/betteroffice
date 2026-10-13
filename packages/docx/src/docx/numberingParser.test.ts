import { describe, expect, it } from 'bun:test';
import { createNumberingMap } from './numberingParser';

describe('createNumberingMap', () => {
  it('draws no list for a num without an abstract numbering, even with a full level', () => {
    const map = createNumberingMap({
      abstractNums: [],
      nums: [
        {
          numId: 5,
          abstractNumId: -1,
          levelOverrides: [{ ilvl: 0, lvl: { ilvl: 0, numFmt: 'decimal', lvlText: '%1.' } }],
        },
      ],
    });
    expect(map.hasNumbering(5)).toBe(false);
    expect(map.getLevel(5, 0)).toBeNull();
  });
});
