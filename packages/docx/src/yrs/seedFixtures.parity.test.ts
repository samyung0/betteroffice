import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync, readdirSync } from 'node:fs';
import { join, resolve } from 'node:path';

import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import { projectYrsComments, commentNumericId } from './comments';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import type { Document } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm } from '../wasm/opc';
import { preloadParseWasm } from '../wasm/parse';
import { createYrsSession, type YrsSession } from './index';
import { documentToYrs } from './documentToYrs';
import { yrsToDocument } from './yrsToDocument';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FIXTURE_ROOT = resolve(import.meta.dir, '__fixtures__/seed-parity');

/** Required parsed node kinds by fixture. */
const REQUIRED_NODES: Record<string, string[]> = {
  'bookmarks-fields': ['bookmarkStart', 'bookmarkEnd', 'simpleField', 'complexField', 'instrText'],
  comments: ['commentRangeStart', 'commentRangeEnd', 'commentReference'],
  'content-controls': ['blockSdt', 'inlineSdt'],
  hyperlinks: ['hyperlink', 'bookmarkStart'],
  images: ['drawing', 'image', 'inline'],
  lists: ['paragraph'],
  math: ['mathEquation'],
  notes: ['footnoteRef', 'endnoteRef'],
  'rtl-cjk': ['tab', 'break', 'symbol', 'softHyphen', 'noBreakHyphen'],
  'sections-headers': ['paragraph'],
  shapes: ['shape', 'table'],
  tables: ['table', 'tableRow', 'tableCell'],
  'tracked-changes': ['insertion', 'deletion', 'paragraphPropertyChange', 'runPropertyChange'],
  'tracked-moves': ['moveFrom', 'moveTo'],
};

/** Package parts each fixture must carry once parsed. */
const REQUIRED_PARTS: Record<string, Array<(document: Document) => boolean>> = {
  comments: [(doc) => (doc.package.document.comments?.length ?? 0) === 4],
  images: [(doc) => (doc.package.media?.size ?? 0) > 0],
  lists: [(doc) => (doc.package.numbering?.abstractNums?.length ?? 0) === 2],
  notes: [
    (doc) => (doc.package.footnotes?.length ?? 0) === 2,
    (doc) => (doc.package.endnotes?.length ?? 0) === 1,
  ],
  'sections-headers': [
    (doc) => (doc.package.headers?.size ?? 0) === 3,
    (doc) => (doc.package.footers?.size ?? 0) === 2,
    (doc) => (doc.package.document.sections?.length ?? 0) === 2,
  ],
};

const PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPkg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DEFAULT_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:outlineLvl w:val="0"/><w:spacing w:before="240" w:after="0"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="EndnoteText"><w:name w:val="endnote text"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="CommentText"><w:name w:val="annotation text"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style>
  <w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
  <w:style w:type="character" w:styleId="EndnoteReference"><w:name w:val="endnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
  <w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/><w:rPr><w:sz w:val="16"/></w:rPr></w:style>
  <w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
  <w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style>
</w:styles>`;

const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument';
const CONTENT_TYPES: Record<string, string> = {
  'word/document.xml': `${OFFICE_DOC}.wordprocessingml.document.main+xml`,
  'word/styles.xml': `${OFFICE_DOC}.wordprocessingml.styles+xml`,
  'word/numbering.xml': `${OFFICE_DOC}.wordprocessingml.numbering+xml`,
  'word/settings.xml': `${OFFICE_DOC}.wordprocessingml.settings+xml`,
  'word/footnotes.xml': `${OFFICE_DOC}.wordprocessingml.footnotes+xml`,
  'word/endnotes.xml': `${OFFICE_DOC}.wordprocessingml.endnotes+xml`,
  'word/comments.xml': `${OFFICE_DOC}.wordprocessingml.comments+xml`,
  'word/commentsExtended.xml': `${OFFICE_DOC}.wordprocessingml.commentsExtended+xml`,
  'word/commentsIds.xml': `${OFFICE_DOC}.wordprocessingml.commentsIds+xml`,
  'word/fontTable.xml': `${OFFICE_DOC}.wordprocessingml.fontTable+xml`,
  'word/theme/theme1.xml': `${OFFICE_DOC}.theme+xml`,
};

function partContentType(path: string): string | undefined {
  if (CONTENT_TYPES[path]) return CONTENT_TYPES[path];
  if (/^word\/header\d+\.xml$/.test(path)) return `${OFFICE_DOC}.wordprocessingml.header+xml`;
  if (/^word\/footer\d+\.xml$/.test(path)) return `${OFFICE_DOC}.wordprocessingml.footer+xml`;
  return undefined;
}

function contentTypesXml(paths: string[]): string {
  const overrides = paths
    .map((path) => {
      const type = partContentType(path);
      return type ? `  <Override PartName="/${path}" ContentType="${type}"/>` : '';
    })
    .filter(Boolean)
    .join('\n');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
${overrides}
</Types>`;
}

function walk(dir: string, prefix = ''): string[] {
  return readdirSync(dir, { withFileTypes: true })
    .flatMap((entry) =>
      entry.isDirectory()
        ? walk(join(dir, entry.name), `${prefix}${entry.name}/`)
        : [`${prefix}${entry.name}`]
    )
    .sort();
}

/**
 * Assemble a fixture directory into DOCX bytes. Every file is a package part and a
 * `<path>.base64` file decodes to the binary part at `<path>`; the content types, package
 * rels and styles are supplied when the fixture does not author its own.
 */
function buildFixtureDocx(name: string): Uint8Array {
  const dir = join(FIXTURE_ROOT, name);
  const authored = new Map<string, Uint8Array>();
  for (const part of walk(dir)) {
    const bytes = readFileSync(join(dir, part));
    if (part.endsWith('.base64')) {
      const decoded = Buffer.from(bytes.toString('utf8').replace(/\s+/g, ''), 'base64');
      authored.set(part.slice(0, -'.base64'.length), new Uint8Array(decoded));
    } else {
      authored.set(part, new Uint8Array(bytes));
    }
  }
  if (!authored.has('word/styles.xml')) authored.set('word/styles.xml', toBytes(DEFAULT_STYLES));
  if (!authored.has('_rels/.rels')) authored.set('_rels/.rels', toBytes(PACKAGE_RELS));

  const pkg: PartsMap = new Map();
  pkg.set(
    '[Content_Types].xml',
    authored.get('[Content_Types].xml') ?? toBytes(contentTypesXml([...authored.keys()]))
  );
  for (const [path, bytes] of authored) {
    if (path === '[Content_Types].xml') continue;
    pkg.set(path, bytes);
  }
  return new Uint8Array(rezipPartsToArrayBuffer(pkg));
}

function nodeKinds(node: unknown, kinds = new Set<string>(), depth = 0): Set<string> {
  if (depth > 40 || node === null || typeof node !== 'object') return kinds;
  if (Array.isArray(node)) {
    for (const item of node) nodeKinds(item, kinds, depth + 1);
    return kinds;
  }
  if (node instanceof Map) {
    for (const value of node.values()) nodeKinds(value, kinds, depth + 1);
    return kinds;
  }
  const record = node as Record<string, unknown>;
  if (typeof record.type === 'string') kinds.add(record.type);
  for (const value of Object.values(record)) nodeKinds(value, kinds, depth + 1);
  return kinds;
}

function expectEquivalentStories(left: YrsSession, right: YrsSession): void {
  expect(left.storyIds()).toEqual(right.storyIds());
  for (const storyId of left.storyIds()) {
    expect(left.storySegments(storyId)).toEqual(right.storySegments(storyId));
  }
}

const fixtures = readdirSync(FIXTURE_ROOT, { withFileTypes: true })
  .filter((entry) => entry.isDirectory())
  .map((entry) => entry.name)
  .sort();

// Yrs formatting marker counts make seeded state vectors nondeterministic.
describe('DOCX seeding across document features', () => {
  beforeAll(() =>
    Promise.all([
      preloadEditWasm(new Uint8Array(readFileSync(WASM))),
      preloadOpcWasm(
        new Uint8Array(
          readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))
        )
      ),
      preloadParseWasm(
        new Uint8Array(
          readFileSync(resolve(import.meta.dir, '../wasm/generated/parse/docx_parse_bg.wasm'))
        )
      ),
    ])
  );

  for (const name of fixtures) {
    it(`preserves ${name} stories, comments, and save output`, async () => {
      const bytes = buildFixtureDocx(name);
      const parsed = await parseDocx(bytes.buffer as ArrayBuffer, {
        preloadFonts: false,
      });

      const kinds = nodeKinds(parsed.package.document);
      for (const required of REQUIRED_NODES[name] ?? []) expect([...kinds]).toContain(required);
      for (const [index, check] of (REQUIRED_PARTS[name] ?? []).entries()) {
        expect(check(parsed), `${name} package check ${index}`).toBe(true);
      }

      const projected = await createYrsSession({ clientId: 48001 });
      const engine = await createYrsSession({ clientId: 48001 });
      try {
        documentToYrs(projected, parsed);
        engine.seedFromDocx(bytes);

        expectEquivalentStories(engine, projected);
        for (const comment of parsed.package.document.comments ?? []) {
          const id = String(comment.id);
          expect(engine.resolveComment(id)).toEqual(projected.resolveComment(id));
        }
        expect(yrsToDocument(engine, parsed)).toEqual(yrsToDocument(projected, parsed));
      } finally {
        projected.destroy();
        engine.destroy();
      }
    });
  }
});

describe('shared DOCX comment lifecycle', () => {
  beforeAll(() =>
    Promise.all([
      preloadEditWasm(new Uint8Array(readFileSync(WASM))),
      preloadOpcWasm(
        new Uint8Array(
          readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))
        )
      ),
      preloadParseWasm(
        new Uint8Array(
          readFileSync(resolve(import.meta.dir, '../wasm/generated/parse/docx_parse_bg.wasm'))
        )
      ),
    ])
  );

  it('merges concurrent threads, replies and resolve metadata, then exports and reopens them', async () => {
    const bytes = buildFixtureDocx('comments');
    const left = await createYrsSession({ clientId: 62001 });
    const right = await createYrsSession({ clientId: 62002 });
    const reopened = await createYrsSession({ clientId: 62003 });
    const body = (text: string) => [
      {
        type: 'paragraph',
        content: [{ type: 'run', content: [{ type: 'text', text }] }],
      },
    ];
    try {
      const { document } = left.seedFromDocx(bytes);
      right.loadState(left.encodeState());
      const anchors = left.resolveComment('1');
      const anchor = anchors[0]!;
      for (const [session, id] of [
        [left, 'peer-left-thread'],
        [right, 'peer-right-thread'],
      ] as const) {
        session.applyRawOps(anchor.story, [
          {
            op: 'setComment',
            id,
            ranges: [[anchor.start, anchor.end]],
            author: id,
            date: '2026-09-06T00:00:00Z',
            body: body(id),
          },
        ]);
      }
      left.applyRawOps('body', [{ op: 'patchComment', id: '1', fields: { done: true } }]);
      right.applyRawOps('body', [
        {
          op: 'patchComment',
          id: 'peer-right-reply',
          fields: {
            parentId: '1',
            author: 'Peer',
            date: '2026-09-06T00:00:00Z',
            body: body('Shared reply'),
            done: false,
          },
        },
      ]);
      const leftState = left.encodeState();
      const rightState = right.encodeState();
      left.applyUpdate(rightState);
      right.applyUpdate(leftState);
      expect(left.listComments()).toEqual(right.listComments());
      expect(left.yrsBlocksForStory('body', {})).toBeTruthy();
      expect(left.resolveComment('1')).toEqual(anchors);
      const comments = projectYrsComments(left, document.package.document.comments);
      expect(comments.find((entry) => entry.sharedId === '1')?.done).toBe(true);
      expect(comments.find((entry) => entry.sharedId === 'peer-right-reply')?.parentId).toBe(1);
      expect(new Set(comments.map((entry) => entry.id)).size).toBe(comments.length);
      const output = await repackDocx(yrsToDocument(left, document));
      const parsed = await parseDocx(output, { preloadFonts: false });
      const exported = parsed.package.document.comments!;
      expect(exported).toHaveLength(comments.length);
      expect(exported.find((entry) => entry.id === 1)?.done).toBe(true);
      expect(
        exported.find((entry) => entry.id === commentNumericId('peer-right-reply'))?.parentId
      ).toBe(1);
      expect(JSON.stringify(exported)).toContain('Shared reply');
      expect(JSON.stringify(exported)).toContain('peer-left-thread');
      expect(JSON.stringify(exported)).toContain('peer-right-thread');
      reopened.seedFromDocx(new Uint8Array(output));
      expect(reopened.listComments()).toHaveLength(comments.length);
      expect(reopened.listComments().find((entry) => entry.id === '1')?.done).toBe(true);
      left.applyRawOps('body', [{ op: 'removeComment', id: 'peer-left-thread' }]);
      right.applyUpdate(left.encodeState());
      expect(right.listComments().some((entry) => entry.id === 'peer-left-thread')).toBe(false);
    } finally {
      left.destroy();
      right.destroy();
      reopened.destroy();
    }
  });

  it('preserves imported numeric IDs when a shared key hashes to the same export ID', () => {
    const shared = 'thread-collision';
    const occupied = commentNumericId(shared);
    const records = [String(occupied), shared].map((id) => ({
      id,
      author: 'Peer',
      date: '',
      done: false,
      parentId: null,
      body: [],
    }));
    const projected = projectYrsComments({ listComments: () => records });
    expect(projected.find((entry) => entry.sharedId === String(occupied))?.id).toBe(occupied);
    expect(projected.find((entry) => entry.sharedId === shared)?.id).not.toBe(occupied);
    expect(projectYrsComments({ listComments: () => [...records].reverse() })).toEqual(projected);
  });
});
