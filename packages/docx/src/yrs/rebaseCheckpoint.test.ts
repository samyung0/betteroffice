import { beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '../wasm/edit';
import { preloadParseWasm } from '../wasm/parse';
import { preloadOpcWasm, rezipContainer, unzipContainer } from '../wasm/opc';
import { writeDocumentWithRust } from '../docx/rustSaveFacade';
import { createYrsSession, type YrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';
import { rebaseDocxCheckpoint } from './rebaseCheckpoint';

const png =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a3ioAAAAASUVORK5CYII=';
const image = { src: `data:image/png;base64,${png}`, width: 12, height: 12 };
const encoder = new TextEncoder();

beforeAll(async () => {
  await Promise.all([
    preloadEditWasm(
      readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))
    ),
    preloadParseWasm(
      readFileSync(resolve(import.meta.dir, '../wasm/generated/parse/docx_parse_bg.wasm'))
    ),
    preloadOpcWasm(
      readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))
    ),
  ]);
});

function source(): Uint8Array {
  return rezipContainer(
    Object.fromEntries(
      Object.entries({
        '[Content_Types].xml':
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
        '_rels/.rels':
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
        'word/document.xml':
          '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="1234ABCD" w14:textId="87654321"><w:r><w:t>Anchor</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>',
      }).map(([path, xml]) => [path, encoder.encode(xml)])
    )
  );
}

async function save(session: YrsSession, base: Uint8Array): Promise<Uint8Array> {
  session.openDocx(base, false);
  const document = session.materializeDocx();
  if (!document) throw new Error('missing base');
  const result = await writeDocumentWithRust(
    yrsToDocument(session, document),
    base.slice().buffer,
    { updateModifiedDate: false },
    undefined,
    {
      seed: '1'.repeat(64),
      now: '2026-09-14T00:00:00.000Z',
    }
  );
  return new Uint8Array(result.buffer);
}

async function reopen(base: Uint8Array, state: Uint8Array): Promise<YrsSession> {
  const session = await createYrsSession();
  session.openDocx(base, false);
  session.loadState(state);
  return session;
}

it('keeps pre-capture table identities and later structural and paragraph edits using B alone', async () => {
  const a = source();
  const session = await createYrsSession();
  const opened: YrsSession[] = [session];
  try {
    session.openDocx(a, true);
    const para = session.paragraphs('body')[0];
    session.insertText({ story: 'body', paraId: para.paraId, offset: 0 }, 'captured ');
    const table = session.insertTable({ story: 'body', paraId: para.paraId, offset: 0 }, 2, 2);
    const firstCell = table.createdStoryIds[0];
    session.applyRawOps(firstCell, [{ op: 'insert', index: 0, text: 'cell10' }]);
    const captured = session.encodeState();
    const b = await save(session, a);
    session.insertRow({ story: 'body', tableIndex: 0, row: 0, column: 0 }, 'above');
    session.applyRawOps(firstCell, [{ op: 'insert', index: 0, text: 'saved11 ' }]);
    session.insertText({ story: 'body', paraId: para.paraId, offset: 0 }, 'later ');
    const latest = session.encodeState();
    const expectedOutput = await save(session, a);
    const result = await rebaseDocxCheckpoint({
      oldSource: a,
      capturedState: captured,
      latestState: latest,
      exportedSource: b,
    });
    const current = await reopen(b, result.state);
    opened.push(current);
    const indexed = await reopen(b, result.indexedState);
    opened.push(indexed);
    expect(current.storySegments(firstCell)).toEqual(
      session.storySegments(firstCell).map((segment) =>
        segment.kind === 'pilcrow'
          ? expect.objectContaining({
              kind: 'pilcrow',
              paraId: segment.paraId,
            })
          : segment
      )
    );
    expect(
      indexed
        .storySegments(firstCell)
        .filter((segment) => segment.kind === 'text')
        .map((segment) => segment.text)
        .join('')
    ).toBe('cell10');
    const output = await save(current, b);
    const xml = new TextDecoder().decode(unzipContainer(output)['word/document.xml']);
    expect(xml).toContain('saved11 cell10');
    expect(xml).toContain('later captured Anchor');
    expect(xml.match(/<w:tr>/g)?.length).toBe(3);
    expect(xml).toBe(new TextDecoder().decode(unzipContainer(expectedOutput)['word/document.xml']));
    current.splitParagraph({ story: 'body', paraId: para.paraId, offset: 6 });
    const splitXml = new TextDecoder().decode(
      unzipContainer(await save(current, b))['word/document.xml']
    );
    expect(splitXml.match(/w14:paraId="1234ABCD"/g)?.length).toBe(1);
  } finally {
    for (const replica of opened) replica.destroy();
  }
});

it('binds an image added at capture to B without allocating duplicate bytes on subsequent exports', async () => {
  const a = source();
  const session = await createYrsSession();
  let current: YrsSession | undefined;
  try {
    session.openDocx(a, true);
    session.applyRawOps('body', [{ op: 'insertEmbed', index: 0, kind: 'image', payload: image }]);
    const captured = session.encodeState();
    const b = await save(session, a);
    session.applyRawOps('body', [{ op: 'setEmbedAttr', index: 0, key: 'width', value: 33 }]);
    const result = await rebaseDocxCheckpoint({
      oldSource: a,
      capturedState: captured,
      latestState: session.encodeState(),
      exportedSource: b,
    });
    current = await reopen(b, result.state);
    const segment = current.storySegments('body')[0];
    expect(segment).toMatchObject({
      kind: 'embed',
      payload: { width: 33, src: image.src, rId: expect.any(String) },
    });
    const output = await save(current, b);
    const repeated = await save(current, output);
    for (const bytes of [b, output, repeated]) {
      const parts = unzipContainer(bytes);
      expect(Object.keys(parts).filter((path) => path.startsWith('word/media/')).length).toBe(1);
      expect([...parts['word/media/image1.png']]).toEqual([
        ...Uint8Array.from(atob(png), (char) => char.charCodeAt(0)),
      ]);
    }
  } finally {
    session.destroy();
    current?.destroy();
  }
});

it('preserves a saved restore of an image removed at capture, and refuses missing source dependencies', async () => {
  const seed = await createYrsSession();
  const session = await createYrsSession();
  let current: YrsSession | undefined;
  try {
    const blank = source();
    seed.openDocx(blank, true);
    seed.applyRawOps('body', [{ op: 'insertEmbed', index: 0, kind: 'image', payload: image }]);
    const a = await save(seed, blank);
    session.openDocx(a, true);
    const original = session.storySegments('body')[0];
    if (original.kind !== 'embed') throw new Error('missing image');
    session.applyRawOps('body', [{ op: 'delete', index: 0, len: 1 }]);
    const captured = session.encodeState();
    const b = await save(session, a);
    session.applyRawOps('body', [
      {
        op: 'insertEmbed',
        index: 0,
        kind: original.embedKind,
        payload: original.payload,
      },
    ]);
    const latest = session.encodeState();
    const result = await rebaseDocxCheckpoint({
      oldSource: a,
      capturedState: captured,
      latestState: latest,
      exportedSource: b,
    });
    current = await reopen(b, result.state);
    const output = await save(current, b);
    expect(new TextDecoder().decode(unzipContainer(output)['word/document.xml'])).toContain(
      'r:embed="rId1"'
    );
    const missing = unzipContainer(b);
    delete missing['word/media/image1.png'];
    await expect(
      rebaseDocxCheckpoint({
        oldSource: a,
        capturedState: captured,
        latestState: latest,
        exportedSource: rezipContainer(missing),
      })
    ).rejects.toThrow('unmodeled source part');
  } finally {
    seed.destroy();
    session.destroy();
    current?.destroy();
  }
});

it('rebinds a feature-rich package with saved header and body edits', async () => {
  const a = new Uint8Array(
    readFileSync(resolve(import.meta.dir, '../../../../poc/fixtures/feature-rich.docx'))
  );
  const session = await createYrsSession();
  let current: YrsSession | undefined;
  try {
    session.openDocx(a, true);
    const captured = session.encodeState();
    const b = await save(session, a);
    for (const story of session.storyIds().filter((id) => id === 'body' || id.startsWith('hf:'))) {
      const paragraph = session.paragraphs(story)[0];
      if (paragraph) session.insertText({ story, paraId: paragraph.paraId, offset: 0 }, 'saved11 ');
    }
    const expected = await save(session, a);
    const result = await rebaseDocxCheckpoint({
      oldSource: a,
      capturedState: captured,
      latestState: session.encodeState(),
      exportedSource: b,
    });
    current = await reopen(b, result.state);
    const actual = await save(current, b);
    const beforeParts = unzipContainer(expected);
    const afterParts = unzipContainer(actual);
    for (const path of Object.keys(beforeParts).filter((path) =>
      /^word\/(document|header\d+|footer\d+|footnotes|endnotes)\.xml$/.test(path)
    )) {
      expect(new TextDecoder().decode(afterParts[path]), path).toBe(
        new TextDecoder().decode(beforeParts[path])
      );
    }
  } finally {
    session.destroy();
    current?.destroy();
  }
});

it('carries legacy opaque blocks by their native embed when later structure changes their source position', async () => {
  const parts = unzipContainer(source());
  parts['word/document.xml'] = encoder.encode(
    new TextDecoder()
      .decode(parts['word/document.xml'])
      .replace(
        '<w:body>',
        '<w:body><w:sdt><w:sdtPr><w:alias w:val="Legacy"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>opaque block</w:t></w:r></w:p></w:sdtContent></w:sdt>'
      )
  );
  const a = rezipContainer(parts);
  const session = await createYrsSession();
  let current: YrsSession | undefined;
  try {
    session.openDocx(a, true);
    session.applyRawOps('body', [
      { op: 'delete', index: 0, len: 1 },
      {
        op: 'insertEmbed',
        index: 0,
        kind: 'opaque',
        payload: { blob: { type: 'blockSdt' } },
      },
    ]);
    const captured = session.encodeState();
    const b = await save(session, a);
    session.applyRawOps('body', [{ op: 'insert', index: 1, text: 'saved11 ' }]);
    const result = await rebaseDocxCheckpoint({
      oldSource: a,
      capturedState: captured,
      latestState: session.encodeState(),
      exportedSource: b,
    });
    current = await reopen(b, result.state);
    current.applyRawOps('body', [
      {
        op: 'insertEmbed',
        index: 0,
        kind: 'pilcrow',
        payload: { paraId: 'after-rebase' },
      },
      { op: 'insert', index: 0, text: 'before block' },
    ]);
    const xml = new TextDecoder().decode(
      unzipContainer(await save(current, b))['word/document.xml']
    );
    expect(xml).toContain('before block');
    expect(xml).toContain('saved11 Anchor');
    expect(xml.match(/opaque block/g)?.length).toBe(1);
  } finally {
    session.destroy();
    current?.destroy();
  }
});
