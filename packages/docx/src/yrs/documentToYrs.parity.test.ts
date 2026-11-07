import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsSession } from './index';
import { documentToYrs } from './documentToYrs';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FIXTURE = resolve(import.meta.dir, '../../../../apps/demo/public/betteroffice-demo.docx');
const OPAQUE_OBJECTS = resolve(import.meta.dir, '../../../../poc/fixtures/opaque-objects.docx');
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PNG = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAACklEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg=='), (char) => char.charCodeAt(0));

/** A body paragraph holding a VML horizontal-rule picture (`o:hr="t"`). */
function vmlRuleDocx(): Uint8Array {
  const parts = new Map(
    Object.entries({
      '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
      '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
      'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdRule" Type="${R}/image" Target="media/rule.png"/></Relationships>`,
      'word/document.xml': `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><w:body><w:p><w:r><w:pict><v:shape id="rule" style="width:487.05pt;height:2.85pt" o:hrpct="0" o:hralign="center" o:hr="t"><v:imagedata r:id="rIdRule" o:title="rule"/></v:shape></w:pict></w:r></w:p><w:sectPr/></w:body></w:document>`,
    }).map(([path, xml]) => [path, toBytes(xml)])
  );
  parts.set('word/media/rule.png', PNG);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function expectEquivalentStories(left: YrsSession, right: YrsSession): void {
  expect(left.storyIds()).toEqual(right.storyIds());
  for (const storyId of left.storyIds()) {
    expect(left.storySegments(storyId)).toEqual(right.storySegments(storyId));
  }
}

/**
 * Every embed's legacy VML attributes, story by story: the horizontal-rule flag (#1725) and the
 * text box's display shape (#1735).
 */
function legacyVmlAttrs(session: YrsSession): unknown[] {
  const found: unknown[] = [];
  const walk = (value: unknown): void => {
    if (Array.isArray(value)) value.forEach(walk);
    else if (value && typeof value === 'object') {
      const record = value as Record<string, unknown>;
      if (record.horizontalRule === true || typeof record.textBoxJson === 'string')
        found.push({ horizontalRule: record.horizontalRule === true, textBoxJson: record.textBoxJson ?? null });
      Object.values(record).forEach(walk);
    }
  };
  for (const storyId of session.storyIds()) walk(session.storySegments(storyId));
  return found;
}

describe('DOCX engine seeding', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))), 30_000);

  it('produces equivalent story structure and state updates', async () => {
    const bytes = Uint8Array.from(readFileSync(FIXTURE));
    const parsed = await parseDocx(bytes.buffer);
    // The byte seeder writes under the fixed seed client 0.
    const projected = await createYrsSession({ clientId: 0 });
    const engine = await createYrsSession({ clientId: 47001 });
    try {
      documentToYrs(projected, parsed);
      engine.seedFromDocx(bytes);

      expectEquivalentStories(engine, projected);
      expect(engine.encodeStateVector()).toEqual(projected.encodeStateVector());
      expect(engine.encodeState()).toEqual(projected.encodeState());

      const projectedStatePeer = await createYrsSession({ clientId: 47002 });
      const engineStatePeer = await createYrsSession({ clientId: 47003 });
      try {
        projectedStatePeer.loadState(projected.encodeState());
        engineStatePeer.loadState(engine.encodeState());
        const firstParagraph = projectedStatePeer.paragraphs('body')[0];
        projectedStatePeer.insertText(
          { story: 'body', paraId: firstParagraph.paraId, offset: 1 },
          'legacy'
        );
        engineStatePeer.loadState(
          projectedStatePeer.encodeStateAsUpdate(engineStatePeer.encodeStateVector())
        );
        const secondParagraph = engineStatePeer.paragraphs('body')[1];
        engineStatePeer.insertText(
          { story: 'body', paraId: secondParagraph.paraId, offset: 1 },
          'engine'
        );
        projectedStatePeer.loadState(
          engineStatePeer.encodeStateAsUpdate(projectedStatePeer.encodeStateVector())
        );
        expectEquivalentStories(engineStatePeer, projectedStatePeer);
      } finally {
        projectedStatePeer.destroy();
        engineStatePeer.destroy();
      }
    } finally {
      projected.destroy();
      engine.destroy();
    }
  });

  it('seeds a legacy VML text box (#1735) and a VML horizontal rule (#1725) as the engine does', async () => {
    for (const bytes of [Uint8Array.from(readFileSync(OPAQUE_OBJECTS)), vmlRuleDocx()]) {
      const projected = await createYrsSession({ clientId: 0 });
      const engine = await createYrsSession({ clientId: 47004 });
      try {
        documentToYrs(projected, await parseDocx(bytes.slice().buffer));
        engine.seedFromDocx(bytes);
        // Only these attributes: the two seeds still name a picture differently (a data URL
        // against the engine's media token) and round its size differently.
        const attrs = legacyVmlAttrs(engine);
        expect(attrs.length).toBeGreaterThan(0);
        expect(legacyVmlAttrs(projected)).toEqual(attrs);
      } finally {
        projected.destroy();
        engine.destroy();
      }
    }
  });

  it('fails on a story the document names twice, as the engine seed does', async () => {
    const parsed = await parseDocx(Uint8Array.from(readFileSync(FIXTURE)).buffer);
    const note = { type: 'footnote' as const, id: 1, content: [] };
    parsed.package.footnotes = [note, { ...note }];
    const session = await createYrsSession({ clientId: 0 });
    try {
      expect(() => documentToYrs(session, parsed)).toThrow('story "fn:1" already exists');
    } finally {
      session.destroy();
    }
  });

  it('returns thin host metadata and materializes the canonical package on demand', async () => {
    const bytes = Uint8Array.from(readFileSync(FIXTURE));
    const parsed = await parseDocx(bytes.buffer);
    const engine = await createYrsSession({ clientId: 47005 });
    const existingRoom = await createYrsSession({ clientId: 47006 });
    try {
      const host = engine.openDocx(bytes, false);

      expect(engine.storyIds()).toEqual([]);
      expect(host.document.package.document.content).toEqual([]);
      expect(
        host.document.package.document.sections?.every((section) => section.content.length === 0)
      ).toBe(true);
      expect(
        [...(host.document.package.headers?.values() ?? [])].every(
          (header) => header.content.length === 0
        )
      ).toBe(true);
      expect(host.document.package.media?.size).toBe(0);
      expect(host.document.package.charts?.size).toBe(0);
      expect(host.referencedFonts.length).toBeGreaterThan(0);

      const materialized = engine.materializeDocx();
      expect(materialized?.package.document.content).toEqual(parsed.package.document.content);
      expect(materialized?.package.document.sections).toEqual(parsed.package.document.sections);
      expect(materialized?.package.headers).toEqual(parsed.package.headers);
      expect(materialized?.package.footers).toEqual(parsed.package.footers);

      existingRoom.seedFromDocx(bytes);
      engine.loadState(existingRoom.encodeState());
      expectEquivalentStories(engine, existingRoom);
    } finally {
      engine.destroy();
      existingRoom.destroy();
    }
  });
});
