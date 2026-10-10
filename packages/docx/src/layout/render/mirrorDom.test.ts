import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { buildResidentRegionLayoutRequest } from '../../editor/computeLayout';
import { preloadEditWasm } from '../../wasm/edit';
import { preloadOpcWasm } from '../../wasm/opc';
import { preloadParseWasm } from '../../wasm/parse';
import { createYrsSession } from '../../yrs';
import type { ResidentFontRequirement } from '../measure';
import type { DisplayList, DisplayPage, DisplayPrimitive } from './displayList';
import { applyFrameDelta, decodeFrameDelta } from './frameDelta';
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

// each character with the language and direction it is read in, and whether
// screen readers skip it
const characters = (root: Node): string[] => {
  const out: string[] = [];
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const parent = node.parentElement;
    const lang = parent?.closest('[lang]')?.getAttribute('lang') ?? '';
    const dir = parent?.closest('[dir]')?.getAttribute('dir') ?? '';
    const hidden = parent?.closest('[aria-hidden="true"]') ? 'hidden' : '';
    for (const char of node.nodeValue ?? '') out.push(`${char}|${lang}|${dir}|${hidden}`);
  }
  return out;
};

const nodeCount = (root: Node): { elements: number; texts: number } => {
  let elements = 0;
  let texts = 0;
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT);
  for (let node: Node | null = root; node; node = walker.nextNode()) {
    if (node.nodeType === 1) elements += 1;
    else texts += 1;
  }
  return { elements, texts };
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
      expect(characters(text)).toEqual(characters(full));
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

describe('buildMirrorPage joined glyph runs', () => {
  // one glyph run per cluster, as an authoritatively measured line paints them
  const clusters = ['A', 'b', '交', '😀', '.'];
  const glyphRun = (text: string, docStart: number, order: number, x: number, extra = {}) => ({
    kind: 'glyphRun',
    fontId: 0,
    size: 14.667,
    color: '#000000',
    text,
    glyphs: [{ id: 7 + order, x, y: 100, cluster: 0, advance: 9 }],
    docStart,
    docEnd: docStart + text.length,
    blockKey: 'body:p2',
    lineIndex: 0,
    logicalOrder: order,
    bidiLevel: 0,
    fallbackFont: '400 14.667px Calibri, sans-serif',
    ...extra,
  });
  const pieces = clusters.map((text, index) =>
    glyphRun(
      text,
      10 + clusters.slice(0, index).join('').length,
      40 + index,
      50 + index * 9,
      index === 0 ? { tabsBefore: 1 } : index === clusters.length - 1 ? { breaksAfter: 1 } : {}
    )
  );
  const bytes = (text: string) => new TextEncoder().encode(text).length;
  const joined = {
    ...glyphRun(clusters.join(''), 10, 40, 50, { tabsBefore: 1, breaksAfter: 1 }),
    glyphs: clusters.map((_, index) => ({
      id: 47 + index,
      x: 50 + index * 9,
      y: 100,
      cluster: bytes(clusters.slice(0, index).join('')),
      advance: 9,
    })),
    clusterRuns: true,
  };
  const page = (primitives: unknown[]) =>
    ({ pageIndex: 0, width: 500, height: 500, primitives }) as unknown as DisplayPage;

  test('mirror one element per cluster, as the runs they were joined from', () => {
    expect(buildMirrorPage(page([joined])).outerHTML).toBe(
      buildMirrorPage(page(pieces)).outerHTML
    );
    expect(buildMirrorPage(page([joined])).querySelectorAll('.layout-run-text')).toHaveLength(5);
  });

  // `count` clusters of `text` joined into one run starting at docStart
  const joinedRun = (text: string, count: number, docStart: number, extra = {}) => {
    const chars = Array.from(text.repeat(count));
    return {
      ...glyphRun(chars.join(''), docStart, docStart, 50 + docStart * 9, extra),
      glyphs: chars.map((_, index) => ({
        id: 7,
        x: 50 + (docStart + index) * 9,
        y: 100,
        cluster: bytes(chars.slice(0, index).join('')),
        advance: 9,
      })),
      clusterRuns: true,
    };
  };
  const linked = (count: number) =>
    page([
      joinedRun('a', count, 0),
      joinedRun('b', count, count, { href: 'https://example.com/toc' }),
      joinedRun('c', count, 2 * count),
    ]);

  test('keep one link per run on a text page, reading the same characters', () => {
    const full = buildMirrorPage(linked(12));
    const text = buildMirrorTextPage(linked(12));
    expect(characters(text)).toEqual(characters(full));
    expect(full.querySelectorAll('a')).toHaveLength(12);
    expect(text.querySelectorAll('a')).toHaveLength(1);
    expect(text.querySelector('a')?.textContent).toBe('b'.repeat(12));
  });

  test('hold a text page to a node count that does not grow with its characters', () => {
    // the page, its paragraph and the link; the text before, in and after it
    const bound = { elements: 3, texts: 3 };
    expect(nodeCount(buildMirrorTextPage(linked(1)))).toEqual(bound);
    expect(nodeCount(buildMirrorTextPage(linked(400)))).toEqual(bound);
    expect(nodeCount(buildMirrorPage(linked(400))).elements).toBeGreaterThan(1200);
  });

  test('split per cluster on a text page only where the clip hides part of the run', () => {
    const clipped = (clip: object) =>
      page([joinedRun('a', 10, 0, { href: '#cut', clipGroup: { clip } })]);
    // the clip ends between the fifth and sixth glyph centres
    const cut = clipped({ x: 0, y: 0, w: 50 + 5 * 9, h: 500 });
    expect(characters(buildMirrorTextPage(cut))).toEqual(characters(buildMirrorPage(cut)));
    expect(buildMirrorTextPage(cut).querySelectorAll('[aria-hidden="true"]')).toHaveLength(5);
    const inside = clipped({ x: 0, y: 0, w: 500, h: 500 });
    expect(characters(buildMirrorTextPage(inside))).toEqual(characters(buildMirrorPage(inside)));
    expect(buildMirrorTextPage(inside).querySelectorAll('a')).toHaveLength(1);
  });
});

describe('buildMirrorTextPage on laid-out documents', () => {
  const root = resolve(import.meta.dir, '../../../../..');
  beforeAll(() =>
    Promise.all([
      preloadEditWasm(
        new Uint8Array(
          readFileSync(resolve(import.meta.dir, '../../wasm/generated/edit/docx_edit_bg.wasm'))
        )
      ),
      preloadOpcWasm(),
      preloadParseWasm(),
    ])
  );

  // The display list the editing engine paints, with joined glyph runs.
  async function laidOut(fixture: string): Promise<DisplayList> {
    const session = await createYrsSession({ clientId: 71_000 });
    try {
      session.openDocx(new Uint8Array(readFileSync(resolve(root, 'poc/fixtures', fixture))), true);
      const fontId = session.registerFont(
        new Uint8Array(
          readFileSync(resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'))
        )
      );
      const request = buildResidentRegionLayoutRequest(session.materializeDocx()!, 24, {});
      const requirements = JSON.parse(
        session.layoutFontRequirementsJson(JSON.stringify(request))
      ) as ResidentFontRequirement[];
      const fontChains = Object.fromEntries(requirements.map(({ key }) => [key, [fontId]]));
      request.measurement = {
        fontChains,
        defaults: { fontSize: 11, fontFamily: 'Calibri' },
        compat: { noLeading: false, doNotExpandShiftReturn: false },
        authoritativeShaping: true,
      };
      session.layoutDocumentWithRegionsRetainedJson(JSON.stringify(request));
      const frame = session.buildDisplayListFrame(JSON.stringify({ fontChains }), 0);
      return applyFrameDelta(null, decodeFrameDelta(frame)).displayList;
    } finally {
      session.destroy();
    }
  }

  const primitivesOf = (page: DisplayPage): DisplayPrimitive[] => [
    ...page.primitives,
    ...(page.header?.primitives ?? []),
    ...(page.footer?.primitives ?? []),
    ...(page.noteAreas ?? []).flatMap((area) => [
      ...(area.separatorPrimitives ?? []),
      ...(area.primitives ?? []),
    ]),
  ];

  for (const fixture of ['wordprocessingml-comprehensive.docx', 'exchange-plan.docx']) {
    test(`reads every page of ${fixture} as the positioned mirror does`, async () => {
      const list = await laidOut(fixture);
      expect(list.pages.length).toBeGreaterThan(0);
      const joined = list.pages.flatMap((page) =>
        primitivesOf(page).filter((p) => p.kind === 'glyphRun' && p.clusterRuns)
      );
      expect(joined.length).toBeGreaterThan(0);
      for (const page of list.pages) {
        expect(characters(buildMirrorTextPage(page))).toEqual(characters(buildMirrorPage(page)));
      }
    });
  }

  test('keeps the table of contents one link per run', async () => {
    const list = await laidOut('wordprocessingml-comprehensive.docx');
    // runs and pictures carrying a link (decorations under a link carry it too)
    const links = list.pages.flatMap((page) =>
      primitivesOf(page).filter(
        (p) => p.href && (p.kind === 'text' || p.kind === 'glyphRun' || p.kind === 'image')
      )
    );
    expect(links.some((p) => p.kind === 'glyphRun' && p.clusterRuns)).toBe(true);
    const text = list.pages.flatMap((page) => [
      ...buildMirrorTextPage(page).querySelectorAll('a[data-href]'),
    ]);
    const full = list.pages.flatMap((page) => [...buildMirrorPage(page).querySelectorAll('a[data-href]')]);
    expect(text).toHaveLength(links.length);
    expect(full.length).toBeGreaterThan(text.length);
  });
});
