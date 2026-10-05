import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import type { DisplayList, DisplayPage } from './displayList';
import { buildMirrorPage, buildMirrorTextPage, MIRROR_TEXT_PAGE_CLASS } from './mirrorDom';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const tablePage = (
  JSON.parse(
    readFileSync(
      resolve(
        import.meta.dir,
        '../../../../../crates/docx-layout/tests/fixtures/table-splits-with-repeated-header.displaylist.json'
      ),
      'utf8'
    )
  ) as DisplayList
).pages[0];

function run(text: string, docStart: number, extra: Record<string, unknown> = {}) {
  return {
    kind: 'text',
    text,
    x: 100 + docStart * 8,
    baselineY: 100,
    width: text.length * 8,
    font: '400 16px Calibri',
    color: '#000000',
    docStart,
    docEnd: docStart + text.length,
    blockId: 1,
    lineIndex: 0,
    ...extra,
  };
}

// a paragraph whose runs carry a link, a second language and a list marker
const paragraphPage = {
  pageIndex: 0,
  width: 500,
  height: 500,
  primitives: [
    run('1.', 0, { listMarker: true }),
    run('Read ', 2),
    run('the guide', 7, { href: 'https://example.com/guide' }),
    run(' 交換學生', 16, { lang: 'zh-TW' }),
    run(' now.', 21),
  ],
} as unknown as DisplayPage;

// each text node with the language and direction it is read in
const textNodes = (root: Node): string[] => {
  const out: string[] = [];
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const parent = node.parentElement;
    const lang = parent?.closest('[lang]')?.getAttribute('lang') ?? '';
    const dir = parent?.closest('[dir]')?.getAttribute('dir') ?? '';
    out.push(`${node.nodeValue ?? ''}|${lang}|${dir}`);
  }
  return out;
};

const meaningful = (root: Element): string[] =>
  [root, ...root.querySelectorAll('*')]
    .filter(
      (el) =>
        el.tagName === 'A' ||
        el.getAttributeNames().some((name) => name === 'role' || name.startsWith('aria-'))
    )
    .map((el) =>
      [
        el.tagName,
        ...el
          .getAttributeNames()
          .filter((name) => name === 'role' || name === 'href' || name.startsWith('aria-'))
          .sort()
          .map((name) => `${name}=${el.getAttribute(name)}`),
      ].join(' ')
    );

describe('buildMirrorTextPage', () => {
  for (const [name, page] of [
    ['a paragraph with a link, a second language and a list marker', paragraphPage],
    ['a page of a split table', tablePage],
  ] as const) {
    test(`keeps the accessible tree of ${name} without positioned runs`, () => {
      const full = buildMirrorPage(page, { labels: { page: 'Page 1' } });
      const text = buildMirrorTextPage(page, { labels: { page: 'Page 1' } });
      expect(text.classList.contains(MIRROR_TEXT_PAGE_CLASS)).toBe(true);
      expect(textNodes(text)).toEqual(textNodes(full));
      expect(meaningful(text)).toEqual(meaningful(full));
      expect(text.querySelectorAll('[style]')).toHaveLength(0);
      expect(text.querySelectorAll('*').length).toBeLessThan(full.querySelectorAll('*').length);
    });
  }
});

describe('buildMirrorPage text runs', () => {
  test('sit on the painted text: spacing replayed, clipped-out runs hidden', () => {
    const page = {
      pageIndex: 0,
      width: 500,
      height: 500,
      primitives: [
        run('Spaced out', 0, { letterSpacing: 1.5, wordSpacing: 4 }),
        // a split row's cut line, repeated on the next page
        run('cut', 10, { clipGroup: { clip: { x: 0, y: 0, w: 500, h: 50 } } }),
        run('kept', 13, { clipGroup: { clip: { x: 0, y: 50, w: 500, h: 100 } } }),
      ],
    } as unknown as DisplayPage;
    const [spaced, cut, kept] = buildMirrorPage(page).querySelectorAll<HTMLElement>('.layout-run-text');
    expect(spaced.style.letterSpacing).toBe('1.5px');
    expect(spaced.style.wordSpacing).toBe('4px');
    expect(cut.getAttribute('aria-hidden')).toBe('true');
    expect(kept.hasAttribute('aria-hidden')).toBe(false);
  });
});
