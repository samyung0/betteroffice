import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { createRef } from 'react';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { repackDocx } from '@betteroffice/docx/docx/rezip';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { unzipContainer } from '@betteroffice/docx/wasm/opc';
import { createYrsSession, type YrsSession } from '@betteroffice/docx/yrs';
import { yrsToDocument } from '@betteroffice/docx/yrs/yrsToDocument';
import { PagedEditor, type PagedEditorRef } from './PagedEditor';
import type { YrsCoreSession } from './hooks/useYrsCoreSession';
import { samePages, tocLayout } from './tableOfContents';

const { act, cleanup, render } = await import('@testing-library/react');

const WASM = resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm');
const FONT = resolve(
  import.meta.dir,
  '../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
);
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const EMPTY_TEXT = 'No table of contents entries found.';

let fontBytes: ArrayBuffer;
const sessions: YrsSession[] = [];

beforeAll(async () => {
  if (!window.document.fonts) {
    Object.defineProperty(window.document, 'fonts', {
      value: { addEventListener: () => {}, removeEventListener: () => {} },
      configurable: true,
    });
  }
  await preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  fontBytes = readFileSync(FONT).buffer as ArrayBuffer;
});

afterEach(async () => {
  // Let the layouts the commands scheduled land before unmounting.
  await act(async () => {
    await new Promise((done) => setTimeout(done, 50));
  });
  cleanup();
  for (const session of sessions.splice(0)) session.destroy();
});

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const p = (id: string, text: string, style?: string, lead = '') =>
  `<w:p w14:paraId="${id}">${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ''}${lead}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const headingStyle = (n: number) =>
  `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/><w:pPr><w:outlineLvl w:val="${n - 1}"/></w:pPr></w:style>`;

/** An empty first paragraph, three headings (the third after a page break) and a level 4 one. */
function source(): Uint8Array {
  const body = [
    p('10000001', ''),
    p('10000002', 'Introduction', 'Heading1'),
    p('10000003', 'Some text.'),
    p('10000004', 'Background', 'Heading2'),
    p('10000005', 'Deep', 'Heading4'),
    p('10000006', 'Details', 'Heading3', '<w:r><w:br w:type="page"/></w:r>'),
  ].join('');
  const parts = new Map<string, Uint8Array>();
  for (const [name, xml] of [
    [
      '[Content_Types].xml',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
    ],
    [
      '_rels/.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
    ],
    [
      'word/_rels/document.xml.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/></Relationships>`,
    ],
    [
      'word/styles.xml',
      `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>${[1, 2, 3, 4].map(headingStyle).join('')}</w:styles>`,
    ],
    [
      'word/document.xml',
      `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
    ],
  ] as const)
    parts.set(name, toBytes(xml));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

async function open(bytes: Uint8Array, clientId: number): Promise<YrsSession> {
  const session = await createYrsSession({ clientId });
  sessions.push(session);
  session.openDocx(bytes, true);
  return session;
}

function yrsCore(session: YrsSession): YrsCoreSession {
  return {
    session,
    storyBlocks: () => null,
    bodyBlocks: () => null,
    inputPositionMap: () => null,
    displayPositionToLoc: () => null,
    locToDisplayPosition: () => null,
    documentFromYrs: () => null,
    publishDirectInput: () => {},
    sourceNumbering: () => undefined,
  };
}

/** The body as text, `[code|shown]` per field, ¶ per paragraph, `_Toc` names as `_TocN`. */
const units = (session: YrsSession) =>
  session
    .storySegments('body')
    .map((segment) =>
      segment.kind === 'text'
        ? segment.text
        : segment.kind === 'pilcrow'
          ? '¶'
          : segment.embedKind === 'field'
            ? `[${String(segment.payload.instruction)}|${String(segment.payload.displayText)}]`
            : `[${segment.embedKind}]`
    )
    .join('')
    .replace(/_Toc\d+/g, '_TocN');

const documentXml = async (session: YrsSession) => {
  const base = session.materializeDocx();
  if (!base) throw new Error('no source package');
  const saved = new Uint8Array(await repackDocx(yrsToDocument(session, base)));
  return { saved, xml: new TextDecoder().decode(unzipContainer(saved)['word/document.xml']) };
};

async function editor(session: YrsSession) {
  const ref = createRef<PagedEditorRef>();
  const layouts: Layout[] = [];
  render(
    <PagedEditor
      ref={ref}
      document={null}
      yrsCore={yrsCore(session)}
      measurementFontProvider={{ resolve: () => () => Promise.resolve(fontBytes) }}
      onLayoutComputed={(layout) => {
        if (layout) layouts.push(layout);
      }}
    />
  );
  const deadline = Date.now() + 2000;
  while (layouts.length === 0 && Date.now() < deadline)
    await act(async () => {
      await new Promise((done) => setTimeout(done, 10));
    });
  expect(layouts.length).toBeGreaterThan(0);
  return ref;
}

const command = (ref: React.RefObject<PagedEditorRef | null>, update: boolean) => {
  let applied = false;
  act(() => {
    applied =
      ref.current?.applyYrsCommand({ type: 'tableOfContents', update, emptyText: EMPTY_TEXT }) ??
      false;
  });
  return applied;
};

test('Insert table of contents writes Word field, links and page numbers in one Undo step', async () => {
  const session = await open(source(), 7101);
  const ref = await editor(session);
  const before = units(session);
  session.setSelection({ story: 'body', paraId: '10000001', offset: 0 });
  expect(command(ref, false)).toBe(true);

  const toc =
    'Introduction\t[PAGEREF _TocN \\h|1][TOC \\o "1-3" \\h \\z \\u|]¶Background\t[PAGEREF _TocN \\h|1]¶Details\t[PAGEREF _TocN \\h|2]¶';
  expect(units(session)).toBe(`${toc}${before}`);
  // The caret follows the table, in the paragraph it was in.
  expect(session.selection()?.head).toMatchObject({ paraId: '10000001', offset: 0 });

  const { saved, xml } = await documentXml(session);
  expect(xml).toContain(
    '<w:instrText xml:space="preserve"> TOC \\o &quot;1-3&quot; \\h \\z \\u </w:instrText>'
  );
  const names = [...xml.matchAll(/<w:bookmarkStart w:id="\d+" w:name="(_Toc\d{9})"\/>/g)].map(
    (match) => match[1]
  );
  expect(names).toHaveLength(3);
  for (const [index, name] of names.entries()) {
    expect(xml).toContain(`<w:hyperlink w:anchor="${name}"`);
    expect(xml).toContain(`PAGEREF ${name} \\h `);
    expect(xml).toMatch(new RegExp(`<w:pStyle w:val="TOC${index + 1}"/>`));
  }
  expect(xml).toContain('<w:tab w:val="right" w:pos="9350" w:leader="dot"/>');

  // What the save writes is what the editor shows once reopened.
  const reopened = await open(saved, 7102);
  expect(reopened.storySegments('body')).toEqual(session.storySegments('body'));

  expect(session.undo()).toBe(true);
  expect(units(session)).toBe(before);
});

test('Update table of contents rebuilds the entries after a heading is renamed', async () => {
  const session = await open(source(), 7103);
  const ref = await editor(session);
  session.setSelection({ story: 'body', paraId: '10000001', offset: 0 });
  expect(command(ref, false)).toBe(true);
  const heading = session.paragraphs('body').find((paragraph) => paragraph.text === 'Background');
  if (!heading) throw new Error('no heading');
  session.replaceRange(
    {
      story: 'body',
      start: { paraId: heading.paraId, offset: 0 },
      end: { paraId: heading.paraId, offset: 10 },
    },
    'Context'
  );
  session.setSelection({ story: 'body', paraId: heading.paraId, offset: 0 });
  expect(command(ref, true)).toBe(true);
  expect(units(session)).toStartWith(
    'Introduction\t[PAGEREF _TocN \\h|1][TOC \\o "1-3" \\h \\z \\u|]¶Context\t[PAGEREF _TocN \\h|1]¶'
  );
  expect(session.tableOfContentsCount()).toBe(1);
  const { saved } = await documentXml(session);
  const reopened = await open(saved, 7104);
  expect(reopened.storySegments('body')).toEqual(session.storySegments('body'));
});

test('a heading without a fragment takes the page before it, and the tab the text width', () => {
  const page = (number: number, ids: string[], pageLabel?: string) => ({
    number,
    pageLabel,
    size: { w: 816, h: 1056 },
    margins: { top: 96, right: 96, bottom: 96, left: 96 },
    fragments: ids.map((blockId) => ({ kind: 'paragraph', blockId })),
  });
  const layout = { pages: [page(1, ['a']), page(2, ['b'], 'ii')] } as unknown as Layout;
  const session = {
    tableOfContentsHeadings: () => [
      { paraId: 'b', level: 1, text: 'B' },
      { paraId: 'hidden', level: 1, text: 'H' },
    ],
    paragraphs: () => [{ paraId: 'a' }, { paraId: 'b' }, { paraId: 'hidden' }],
  } as unknown as YrsSession;
  const computed = tocLayout(session, layout, 'b', EMPTY_TEXT);
  expect(computed).toEqual({
    pages: { b: 'ii', hidden: 'ii' },
    tabTwips: 9350,
    emptyText: EMPTY_TEXT,
  });
  expect(samePages(computed, { ...computed, pages: { b: 'ii', hidden: '3' } })).toBe(false);
});
