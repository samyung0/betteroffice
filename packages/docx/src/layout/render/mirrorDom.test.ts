import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, describe, expect, spyOn, test } from 'bun:test';
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
import {
  buildMirrorPage,
  buildMirrorTextPage,
  MIRROR_TEXT_PAGE_CLASS,
  patchMirrorPage,
  syncMirrorPositions,
} from './mirrorDom';

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

  // one line of touching runs: English, two in Traditional Chinese, English
  const line = (lang: (index: number) => string | undefined) => ({
    pageIndex: 0,
    width: 500,
    height: 500,
    primitives: ['Read ', 'the ', '交換', '學生', ' now.'].map((text, index, all) => {
      const docStart = all.slice(0, index).join('').length;
      const extra = lang(index) ? { lang: lang(index) } : {};
      return run(text, docStart, extra);
    }),
  }) as unknown as DisplayPage;

  test('compares the styles only of runs it keeps as elements', () => {
    const reads = spyOn(CSSStyleDeclaration.prototype, 'getPropertyValue');
    // Where a run sits (left, width) decides whether it touches the next one.
    const compared = () => reads.mock.calls.filter(([name]) => name !== 'left' && name !== 'width');
    try {
      const plain = buildMirrorTextPage(line(() => undefined));
      expect(compared()).toEqual([]);
      expect(plain.querySelector('[role="paragraph"]')?.innerHTML).toBe('Read the 交換學生 now.');
      // Runs in another language than their paragraph stay elements, joined.
      const mixed = buildMirrorTextPage(line((index) => (index === 2 || index === 3 ? 'zh-TW' : 'en-US')));
      expect(compared().length).toBeGreaterThan(0);
      expect(mixed.querySelector('[role="paragraph"]')?.innerHTML).toBe(
        'Read the <span class="layout-run layout-run-text" lang="zh-TW" data-doc-start="9" data-doc-end="13">交換學生</span> now.'
      );
    } finally {
      reads.mockRestore();
    }
  });
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

  const clustered = { clusters: true };

  test('mirror one element per cluster in a text layer, as the runs they were joined from', () => {
    expect(buildMirrorPage(page([joined]), clustered).outerHTML).toBe(
      buildMirrorPage(page(pieces), clustered).outerHTML
    );
    expect(
      buildMirrorPage(page([joined]), clustered).querySelectorAll('.layout-run-text')
    ).toHaveLength(5);
  });

  test("mirror one element elsewhere, so the paragraph reads as the run's text", () => {
    const mirror = buildMirrorPage(page([joined]));
    const runs = mirror.querySelectorAll('.layout-run-text');
    expect(runs).toHaveLength(1);
    expect(runs[0]?.childNodes).toHaveLength(1);
    expect(runs[0]?.textContent).toBe(clusters.join(''));
    expect(mirror.querySelector('.layout-paragraph')?.textContent).toBe(clusters.join(''));
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

  test('keep one link per run outside a text layer, a text page reading the same characters', () => {
    const full = buildMirrorPage(linked(12));
    const text = buildMirrorTextPage(linked(12));
    expect(characters(text)).toEqual(characters(full));
    expect(buildMirrorPage(linked(12), clustered).querySelectorAll('a')).toHaveLength(12);
    expect(full.querySelectorAll('a')).toHaveLength(1);
    expect(text.querySelectorAll('a')).toHaveLength(1);
    expect(text.querySelector('a')?.textContent).toBe('b'.repeat(12));
  });

  test('hold a page outside a text layer to a node count that does not grow with its characters', () => {
    // the page, its paragraph and the link; the text before, in and after it
    const bound = { elements: 3, texts: 3 };
    expect(nodeCount(buildMirrorTextPage(linked(1)))).toEqual(bound);
    expect(nodeCount(buildMirrorTextPage(linked(400)))).toEqual(bound);
    expect(nodeCount(buildMirrorPage(linked(400)))).toEqual(nodeCount(buildMirrorPage(linked(1))));
    expect(nodeCount(buildMirrorPage(linked(400), clustered)).elements).toBeGreaterThan(1200);
  });

  test('split outside a text layer only where the clip hides part of the run', () => {
    const clipped = (clip: object) =>
      page([joinedRun('a', 10, 0, { href: '#cut', clipGroup: { clip } })]);
    // the clip ends between the fifth and sixth glyph centres
    const cut = clipped({ x: 0, y: 0, w: 50 + 5 * 9, h: 500 });
    expect(characters(buildMirrorTextPage(cut))).toEqual(characters(buildMirrorPage(cut)));
    for (const mirror of [buildMirrorPage(cut), buildMirrorTextPage(cut)]) {
      expect(
        [...mirror.querySelectorAll('a')].map((link) => [
          link.textContent,
          link.getAttribute('aria-hidden'),
        ])
      ).toEqual([
        ['aaaaa', null],
        ['aaaaa', 'true'],
      ]);
    }
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

  test('keeps a link one element per line of its text', async () => {
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
    const clusteredLinks = list.pages.flatMap((page) => [
      ...buildMirrorPage(page, { clusters: true }).querySelectorAll('a[data-href]'),
    ]);
    // the table of contents entry, and links the line shaped glyph by glyph
    const read = (links: Element[]) => links.map((link) => link.textContent);
    expect(read(full)).toEqual(['1. Character formatting', 'betteroffice.dev', 'back to section 1']);
    expect(read(text)).toEqual(read(full));
    expect(links.length).toBeGreaterThan(full.length);
    expect(clusteredLinks.length).toBeGreaterThan(full.length);
  });

  test('mirrors each laid-out run whole outside a text layer', async () => {
    let whole = 0;
    for (const fixture of ['wordprocessingml-comprehensive.docx', 'exchange-plan.docx']) {
      for (const page of (await laidOut(fixture)).pages) {
        const runs = [
          ...buildMirrorPage(page).querySelectorAll<HTMLElement>('.layout-run-text[data-doc-start]'),
        ];
        for (const p of page.primitives) {
          if (p.kind !== 'glyphRun' || !p.clusterRuns || p.clipGroup) continue;
          // one element holds the run's text, not one per glyph cluster
          const holder = runs.find(
            (run) =>
              Number(run.dataset.docStart) <= p.docStart! &&
              Number(run.dataset.docEnd) >= p.docEnd! &&
              run.textContent!.includes(p.text)
          );
          expect(holder).toBeDefined();
          whole += 1;
        }
      }
    }
    expect(whole).toBeGreaterThan(0);
  });

  // every field the owned frame delta moves (applyFrameDeltaOwned)
  const shiftPrimitive = (p: DisplayPrimitive, delta: number) => {
    for (const field of ['docStart', 'docEnd', 'fragmentDocStart', 'fragmentDocEnd'] as const) {
      const value = p[field];
      if (value !== undefined) p[field] = value + delta;
    }
    if (p.inlineSdtWidget) p.inlineSdtWidget.pos += delta;
  };
  const stamps = (root: Element): string[] =>
    [...root.querySelectorAll<HTMLElement>('[data-doc-start], [data-doc-end], [data-sdt-pos]')].map(
      (el) => `${el.className}|${el.dataset.docStart}|${el.dataset.docEnd}|${el.dataset.sdtPos}`
    );

  test('moves table cells, structural anchors, joined runs and controls in place', async () => {
    const seen = { cells: 0, anchors: 0, joined: 0, controls: 0 };
    for (const fixture of ['wordprocessingml-comprehensive.docx', 'exchange-plan.docx']) {
      for (const page of (await laidOut(fixture)).pages) {
        // a check box on the page's first body run, so the mirror stamps data-sdt-pos
        const first = page.primitives.find((p) => p.kind === 'glyphRun' && p.docStart !== undefined);
        if (first) first.inlineSdtWidget = { kind: 'checkbox', groupId: 'sdt:1:1', pos: first.docStart! };
        seen.joined += page.primitives.filter((p) => p.kind === 'glyphRun' && p.clusterRuns).length;
        for (const build of [buildMirrorPage, buildMirrorTextPage]) {
          const mirror = build(page);
          if (build === buildMirrorPage) {
            seen.cells += mirror.querySelectorAll('.layout-table-cell[data-doc-start]').length;
            seen.anchors += mirror.querySelectorAll('.layout-mirror-structural-anchor').length;
            seen.controls += mirror.querySelectorAll('[data-sdt-pos]').length;
          }
          const nodes = [...mirror.querySelectorAll('*')];
          for (const p of page.primitives) shiftPrimitive(p, 5);
          expect(syncMirrorPositions(mirror)).toBe(true);
          expect(stamps(mirror)).toEqual(stamps(build(page)));
          expect([...mirror.querySelectorAll('*')].every((node, i) => node === nodes[i])).toBe(true);
          for (const p of page.primitives) shiftPrimitive(p, -5);
        }
      }
    }
    for (const count of Object.values(seen)) expect(count).toBeGreaterThan(0);
  });
});

describe('syncMirrorPositions', () => {
  const positions = (root: Element): string[] =>
    [...root.querySelectorAll<HTMLElement>('[data-doc-start], [data-doc-end]')].map(
      (el) => `${el.dataset.docStart}-${el.dataset.docEnd}`
    );
  const shiftedPage = () =>
    ({
      ...structuredClone(paragraphPage),
      header: {
        kind: 'header',
        rId: 'rId1',
        y: 0,
        height: 40,
        primitives: [run('Head', 0, { blockId: 9 })],
      },
    }) as unknown as DisplayPage;
  const shiftBody = (page: DisplayPage, delta: number, fields = ['docStart', 'docEnd']) => {
    for (const primitive of page.primitives) {
      const record = primitive as unknown as Record<string, number>;
      for (const field of fields) if (record[field] !== undefined) record[field] += delta;
    }
  };

  for (const [name, build] of [
    ['positioned', buildMirrorPage],
    ['plain-text', buildMirrorTextPage],
  ] as const) {
    test(`moves a ${name} mirror's body positions in place and leaves the header's`, () => {
      const page = shiftedPage();
      const mirror = build(page);
      const nodes = [...mirror.querySelectorAll('*')];
      shiftBody(page, 3);
      expect(syncMirrorPositions(mirror)).toBe(true);
      expect(positions(mirror)).toEqual(positions(build(page)));
      expect([...mirror.querySelectorAll('*')].every((node, index) => node === nodes[index])).toBe(
        true
      );
    });
  }

  test('refuses a shift that moves one primitive by two amounts', () => {
    const page = shiftedPage();
    const mirror = buildMirrorPage(page);
    shiftBody(page, 3, ['docStart']);
    expect(syncMirrorPositions(mirror)).toBe(false);
  });

  test('moves a patched mirror the same way', () => {
    const page = shiftedPage();
    const mirror = buildMirrorPage(page);
    const next = structuredClone(page);
    expect(patchMirrorPage(mirror, buildMirrorPage(next))).toBe(true);
    shiftBody(next, 3);
    expect(syncMirrorPositions(mirror)).toBe(true);
    expect(positions(mirror)).toEqual(positions(buildMirrorPage(next)));
  });

  test('writes nothing when a later element gives up', () => {
    const page = shiftedPage();
    const mirror = buildMirrorPage(page);
    const before = positions(mirror);
    // each run moves as one, but the last by more: the runs' own elements
    // could move, the paragraph around them (bound after them) cannot
    shiftBody(page, 3);
    const last = page.primitives.at(-1) as unknown as { docStart: number; docEnd: number };
    last.docStart += 1;
    last.docEnd += 1;
    expect(syncMirrorPositions(mirror)).toBe(false);
    expect(positions(mirror)).toEqual(before);
  });
});

describe('patchMirrorPage', () => {
  // two paragraphs and a header, each run placed by its own x
  const twoParagraphs = () =>
    ({
      pageIndex: 0,
      width: 500,
      height: 500,
      header: {
        kind: 'header',
        rId: 'rId1',
        y: 0,
        height: 40,
        primitives: [run('Head', 0, { blockId: 9, x: 100 })],
      },
      primitives: [
        run('First ', 10, { blockId: 1, x: 100 }),
        run('paragraph.', 16, { blockId: 1, x: 148 }),
        run('Second ', 27, { blockId: 2, x: 100, baselineY: 130 }),
        run('one', 34, { blockId: 2, x: 156, baselineY: 130 }),
      ],
    }) as unknown as DisplayPage;
  const labels = { labels: { page: 'Page 1' } };
  const blockNodes = (root: Element, blockId: number): Element[] => [
    ...root.querySelectorAll(`[data-block-id="${blockId}"], [data-block-id="${blockId}"] *`),
  ];
  const typed = (page: DisplayPage, index: number, text: string): DisplayPage => {
    const next = structuredClone(page);
    type Run = { text: string; width: number; docStart: number; docEnd: number };
    const runs = next.primitives as unknown as Run[];
    runs[index]!.text += text;
    runs[index]!.width += text.length * 8;
    runs[index]!.docEnd += text.length;
    for (const later of runs.slice(index + 1)) {
      later.docStart += text.length;
      later.docEnd += text.length;
    }
    return next;
  };

  for (const [name, build] of [
    ['positioned', buildMirrorPage],
    ['plain-text', buildMirrorTextPage],
  ] as const) {
    test(`a key typed at the end of a ${name} page's last paragraph keeps the other nodes`, () => {
      const page = twoParagraphs();
      const mirror = build(page, labels);
      const first = blockNodes(mirror, 1);
      const next = typed(page, 3, 'a');
      expect(patchMirrorPage(mirror, build(next, labels))).toBe(true);
      expect(mirror.isEqualNode(build(next, labels))).toBe(true);
      expect(first.length).toBeGreaterThan(0);
      expect(blockNodes(mirror, 1).every((node, index) => node === first[index])).toBe(true);
    });
  }

  test('a key typed in the first paragraph moves the later runs in place', () => {
    const page = twoParagraphs();
    const mirror = buildMirrorPage(page, labels);
    const second = blockNodes(mirror, 2);
    const next = typed(page, 0, 'x');
    expect(patchMirrorPage(mirror, buildMirrorPage(next, labels))).toBe(true);
    expect(mirror.isEqualNode(buildMirrorPage(next, labels))).toBe(true);
    expect(blockNodes(mirror, 2).every((node, index) => node === second[index])).toBe(true);
    const moved = mirror.querySelector<HTMLElement>('[data-block-id="2"] .layout-run-text');
    expect(moved?.dataset.docStart).toBe('28');
  });

  // A page with the header above and one line per paragraph: each entry is a
  // block id and its runs (text, or text with extra fields), numbered from
  // `start` with one position between paragraphs.
  type Block = [number, ...(string | [string, Record<string, unknown>])[]];
  const blocks = (start: number, ...entries: Block[]): DisplayPage => {
    const primitives: unknown[] = [];
    let pos = start;
    entries.forEach(([blockId, ...runs], line) => {
      let x = 100;
      for (const entry of runs) {
        const [text, extra] = typeof entry === 'string' ? [entry, {}] : entry;
        primitives.push(run(text, pos, { blockId, x, baselineY: 100 + line * 30, ...extra }));
        x += text.length * 8;
        pos += text.length;
      }
      pos += 1;
    });
    return { ...twoParagraphs(), primitives } as unknown as DisplayPage;
  };
  // every body position the owned frame delta moves
  const shiftBody = (page: DisplayPage, delta: number) => {
    for (const p of page.primitives) {
      for (const field of ['docStart', 'docEnd', 'fragmentDocStart', 'fragmentDocEnd'] as const) {
        const value = p[field];
        if (value !== undefined) p[field] = value + delta;
      }
    }
  };
  // `before`'s mirror patched to `after` equals a fresh build, and its position
  // bindings reach the mounted nodes: a later shift moves them in place
  const textLayer = (page: DisplayPage, options: typeof labels) =>
    buildMirrorPage(page, { ...options, clusters: true });
  const expectPatched = (before: DisplayPage, after: DisplayPage) => {
    for (const build of [buildMirrorPage, buildMirrorTextPage, textLayer]) {
      const mirror = build(before, labels);
      const next = structuredClone(after);
      expect(patchMirrorPage(mirror, build(next, labels))).toBe(true);
      expect(mirror.isEqualNode(build(next, labels))).toBe(true);
      shiftBody(next, 5);
      expect(syncMirrorPositions(mirror)).toBe(true);
      expect(mirror.isEqualNode(build(next, labels))).toBe(true);
    }
  };
  const base = () => blocks(10, [1, 'First ', 'paragraph.'], [2, 'Second ', 'one']);

  test('deleting text equals a fresh build', () => {
    expectPatched(base(), blocks(10, [1, 'First ', 'para.'], [2, 'Second ', 'one']));
  });

  test('Enter splitting a paragraph equals a fresh build', () => {
    expectPatched(base(), blocks(10, [1, 'First '], [3, 'paragraph.'], [2, 'Second ', 'one']));
  });

  test('Backspace joining two paragraphs equals a fresh build', () => {
    expectPatched(base(), blocks(10, [1, 'First ', 'paragraph.', 'Second ', 'one']));
  });

  test('a paragraph leaving the page and one arriving at its top equals a fresh build', () => {
    expectPatched(
      blocks(10, [1, 'First ', 'paragraph.'], [2, 'Second ', 'one'], [3, 'Third.']),
      blocks(1, [8, 'Last.'], [1, 'First ', 'paragraph.'], [2, 'Second ', 'one'])
    );
  });

  test('an edit inside a table cell equals a fresh build', () => {
    const table = { tableId: 't', rowStart: 0, rowEnd: 1, rowCount: 1, columnCount: 2 };
    const cell = (col: number) => ({
      table,
      cell: { row: 0, col, rowSpan: 1, colSpan: 1, cellId: `t-r0-c${col}` },
      x: 100 + col * 200,
      baselineY: 130,
    });
    const page = (first: string) =>
      blocks(
        10,
        [1, 'Intro'],
        [2, [first, cell(0)]],
        [3, ['Beta', cell(1)]],
        [4, ['Outro', { baselineY: 160 }]]
      );
    expectPatched(page('Alpha'), page('Alphas'));
  });

  test('typing in a right-to-left run equals a fresh build', () => {
    const page = (word: string) =>
      blocks(
        10,
        [1, 'Say ', [word, { bidiLevel: 1, lang: 'he' }], ' now.'],
        [2, 'Second ', 'one']
      );
    expectPatched(page('שלום'), page('שלומי'));
  });

  test('typing in a line of shaped runs equals a fresh build', () => {
    // a line as the engine paints it: one glyph run joined from its clusters
    const shaped = (text: string, docStart: number, blockId: number, baselineY: number) => ({
      kind: 'glyphRun',
      fontId: 0,
      size: 16,
      color: '#000000',
      text,
      glyphs: Array.from(text, (_, index) => ({
        id: 7,
        x: 100 + index * 8,
        y: baselineY,
        cluster: index,
        advance: 8,
      })),
      docStart,
      docEnd: docStart + text.length,
      blockId,
      lineIndex: 0,
      clusterRuns: true,
    });
    const page = (first: string) =>
      ({
        ...twoParagraphs(),
        primitives: [shaped(first, 10, 1, 100), shaped('Second one', 11 + first.length, 2, 130)],
      }) as unknown as DisplayPage;
    expectPatched(page('Marker01 line 1 keeps'), page('Marker01 line 1 keeps the'));
    const mirror = buildMirrorPage(page('Marker01 line 1 keeps'), labels);
    const paragraphs = mirror.querySelectorAll('.layout-page-content .layout-paragraph');
    expect([...paragraphs].map((p) => p.textContent)).toEqual([
      'Marker01 line 1 keeps',
      'Second one',
    ]);
  });

  test('leaves a mirror of the other form to be swapped', () => {
    const page = twoParagraphs();
    expect(patchMirrorPage(buildMirrorPage(page), buildMirrorTextPage(page))).toBe(false);
  });
});
