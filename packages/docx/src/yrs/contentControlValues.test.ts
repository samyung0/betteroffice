import { expect, test } from 'bun:test';
import { applyContentControlValue, formatSdtDate } from './contentControlValues';

test('month names follow the language, English without locale data', () => {
  expect(formatSdtDate('2026-03-04', 'd MMMM yyyy (MMM)', 'ja-JP')).toBe('4 3月 2026 (3月)');
  for (const tag of [undefined, '', 'xx-XX', 'und', 'qaa', 'x-none', 'not a tag'])
    expect(formatSdtDate('2026-03-04', 'd MMMM yyyy (MMM)', tag)).toBe('4 March 2026 (Mar)');
});

test("a date control's value is shown in the language of its w:lid", () => {
  const applied = applyContentControlValue(
    {
      sdtType: 'date',
      rawPropertiesXml:
        '<w:sdtPr><w:date w:fullDate="2026-03-04T00:00:00Z"><w:dateFormat w:val="d MMMM yyyy"/><w:lid w:val="ja-JP"/></w:date></w:sdtPr>',
    },
    { kind: 'date', date: '2027-09-06' }
  );
  expect(JSON.stringify(applied.content)).toContain('6 9月 2027');
});
