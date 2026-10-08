import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Document } from '@betteroffice/docx/types/document';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession, type YrsDocxHost, type YrsSession } from '@betteroffice/docx/yrs';
import {
  dirtyProjectionStory,
  mergeDocxHostMetadata,
  seedYrsSession,
  warmCompatibilityBase,
} from './useYrsCoreSession';

function fakeSeedSession(): {
  session: Pick<YrsSession, 'openDocx' | 'loadState' | 'renameDuplicateParaIds'>;
  host: YrsDocxHost;
  opened: { bytes: Uint8Array; seedStories: boolean }[];
  loaded: Uint8Array[];
} {
  const host = { name: 'host' } as unknown as YrsDocxHost;
  const opened: { bytes: Uint8Array; seedStories: boolean }[] = [];
  const loaded: Uint8Array[] = [];
  return {
    session: {
      openDocx: (bytes, seedStories) => {
        opened.push({ bytes, seedStories });
        return host;
      },
      loadState: (update) => {
        loaded.push(update);
      },
      // Each load is followed by the rename; the marker records it.
      renameDuplicateParaIds: () => {
        loaded.push(Uint8Array.of());
        return 0;
      },
    },
    host,
    opened,
    loaded,
  };
}

describe('seedYrsSession', () => {
  test('hydrates a pre-parsed document host from the shared initial update', () => {
    const { session, loaded } = fakeSeedSession();
    const initialUpdate = Uint8Array.of(7, 8, 9);
    const seeded: Document[] = [];

    const host = seedYrsSession(session, (document) => seeded.push(document), {
      bytes: null,
      document: { name: 'parsed' } as unknown as Document,
      initialUpdate,
    });

    expect(seeded).toEqual([]);
    expect(loaded).toEqual([initialUpdate, Uint8Array.of()]);
    expect(host).toBeNull();
  });

  test('seeds from the pre-parsed document when no shared state exists', () => {
    const { session, loaded } = fakeSeedSession();
    const document = { name: 'parsed' } as unknown as Document;
    const seeded: Document[] = [];

    const host = seedYrsSession(session, (next) => seeded.push(next), {
      bytes: null,
      document,
      initialUpdate: undefined,
    });

    expect(seeded).toEqual([document]);
    expect(loaded).toEqual([]);
    expect(host).toBeNull();
  });

  test('opens bytes for metadata without seeding stories when shared state exists', () => {
    const { session, host: expectedHost, opened, loaded } = fakeSeedSession();
    const bytes = Uint8Array.of(1, 2);
    const initialUpdate = Uint8Array.of(3, 4);

    const host = seedYrsSession(session, () => expect.unreachable(), {
      bytes,
      document: null,
      initialUpdate,
    });

    expect(opened).toEqual([{ bytes, seedStories: false }]);
    expect(loaded).toEqual([initialUpdate, Uint8Array.of()]);
    expect(host).toBe(expectedHost);
  });

  test('seeds stories from bytes when no shared state exists', () => {
    const { session, host: expectedHost, opened, loaded } = fakeSeedSession();
    const bytes = Uint8Array.of(1, 2);

    const host = seedYrsSession(session, () => expect.unreachable(), {
      bytes,
      document: null,
      initialUpdate: undefined,
    });

    expect(opened).toEqual([{ bytes, seedStories: true }]);
    expect(loaded).toEqual([]);
    expect(host).toBe(expectedHost);
  });
});

describe('seedYrsSession over the edit engine', () => {
  beforeAll(() =>
    preloadEditWasm(
      new Uint8Array(
        readFileSync(resolve(import.meta.dir, '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm'))
      )
    )
  );

  // Two peers split one paragraph mid-text and leave before exchanging, so
  // the stored state still holds the duplicate id; a client loading it
  // renames it, and typing by id reaches each paragraph.
  test('renames the duplicate paragraph id of a stored state so every half takes typing', async () => {
    const base = await createYrsSession({ clientId: 1 });
    const { paraId } = base.createStory('body', 'Alpha beta');
    const seed = base.encodeState();
    const peers = await Promise.all(
      ([[3001, 2], [3002, 7]] as const).map(async ([clientId, offset]) => {
        const peer = await createYrsSession({ clientId });
        peer.loadState(seed);
        peer.splitParagraph({ story: 'body', paraId, offset });
        return peer;
      })
    );
    const relay = await createYrsSession({ clientId: 9 });
    for (const state of [seed, ...peers.map((peer) => peer.encodeState())]) relay.loadState(state);
    expect(relay.paragraphs('body').filter((p) => p.paraId === paraId)).toHaveLength(2);
    const client = await createYrsSession({ clientId: 4001 });
    try {
      seedYrsSession(client, () => expect.unreachable(), {
        bytes: null,
        document: null,
        initialUpdate: relay.encodeState(),
      });
      const ids = client.paragraphs('body').map((p) => p.paraId);
      expect(new Set(ids).size).toBe(ids.length);
      for (const id of ids) client.insertText({ story: 'body', paraId: id, offset: 0 }, 'X');
      expect(client.paragraphs('body').map((p) => p.text)).toEqual(['XAl', 'Xpha b', 'Xeta']);
    } finally {
      for (const session of [base, ...peers, relay, client]) session.destroy();
    }
  });
});

describe('mergeDocxHostMetadata', () => {
  test('preserves recursive content while applying live host metadata', () => {
    const bodyContent = [
      { type: 'paragraph', content: [] },
    ] as unknown as Document['package']['document']['content'];
    const sectionContent = [
      { type: 'paragraph', content: [] },
    ] as unknown as Document['package']['document']['content'];
    const headerContent = [
      { type: 'paragraph', content: [] },
    ] as unknown as Document['package']['document']['content'];
    const noteContent = [
      { type: 'paragraph', content: [] },
    ] as unknown as Document['package']['document']['content'];
    const media = new Map([
      ['word/media/image1.png', { path: 'word/media/image1.png' }],
    ]) as unknown as NonNullable<Document['package']['media']>;
    const savedBuffer = Uint8Array.of(1).buffer;
    const sourceBuffer = Uint8Array.of(2).buffer;
    const full = {
      originalBuffer: savedBuffer,
      package: {
        document: {
          content: bodyContent,
          sections: [
            {
              id: 'section-1',
              properties: { marginTop: 100 },
              content: sectionContent,
            },
          ],
          finalSectionProperties: { marginTop: 100 },
        },
        headers: new Map([
          [
            'rId1',
            {
              type: 'header',
              hdrFtrType: 'default',
              content: headerContent,
            },
          ],
        ]),
        footnotes: [
          {
            type: 'footnote',
            id: 1,
            noteType: 'normal',
            content: noteContent,
          },
        ],
        media,
      },
    } as unknown as Document;
    const relationships = new Map();
    const host = {
      originalBuffer: sourceBuffer,
      package: {
        document: {
          content: [],
          sections: [
            {
              id: 'section-1',
              properties: { marginTop: 720 },
              content: [],
            },
          ],
          finalSectionProperties: { marginTop: 720 },
        },
        headers: new Map([
          [
            'rId1',
            {
              type: 'header',
              hdrFtrType: 'first',
              content: [],
            },
          ],
        ]),
        footnotes: [
          {
            type: 'footnote',
            id: 1,
            noteType: 'normal',
            content: [],
          },
        ],
        relationships,
      },
    } as unknown as Document;

    const merged = mergeDocxHostMetadata(full, host);

    expect(merged.package.document.content).toBe(bodyContent);
    expect(merged.package.document.sections?.[0].content).toBe(sectionContent);
    expect(merged.package.document.sections?.[0].properties.marginTop).toBe(720);
    expect(merged.package.headers?.get('rId1')?.content).toBe(headerContent);
    expect(merged.package.headers?.get('rId1')?.hdrFtrType).toBe('first');
    expect(merged.package.footnotes?.[0].content).toBe(noteContent);
    expect(merged.package.relationships).toBe(relationships);
    expect(merged.package.media).toBe(media);
    expect(merged.originalBuffer).toBe(savedBuffer);
  });
});

describe('warmCompatibilityBase', () => {
  test('materializes the projection base once', () => {
    let materializations = 0;
    const materialized = { name: 'materialized' } as unknown as Document;
    const session = {
      materializeDocx: () => {
        materializations += 1;
        return materialized;
      },
    };
    const compatibilityBase: { current: Document | null } = { current: null };

    warmCompatibilityBase(session, compatibilityBase);
    warmCompatibilityBase(session, compatibilityBase);

    expect(compatibilityBase.current).toBe(materialized);
    expect(materializations).toBe(1);
  });

  test('leaves an already-projected base untouched', () => {
    const projected = { name: 'projected' } as unknown as Document;
    const compatibilityBase: { current: Document | null } = { current: projected };

    warmCompatibilityBase({ materializeDocx: () => expect.unreachable() }, compatibilityBase);

    expect(compatibilityBase.current).toBe(projected);
  });
});

describe('dirtyProjectionStory', () => {
  test('marks the note story itself for header, footer, footnote and endnote input', () => {
    expect(dirtyProjectionStory('hf:rId5')).toBe('hf:rId5');
    expect(dirtyProjectionStory('fn:2')).toBe('fn:2');
    expect(dirtyProjectionStory('en:2')).toBe('en:2');
  });

  test('falls back to the body for a table cell nested in it', () => {
    expect(dirtyProjectionStory('body:t0:r0c0')).toBe('body');
  });

  test('marks the note or header root for a table cell nested in it', () => {
    expect(dirtyProjectionStory('hf:rId5:t0:r0c0')).toBe('hf:rId5');
    expect(dirtyProjectionStory('fn:2:t0:r0c0')).toBe('fn:2');
    expect(dirtyProjectionStory('en:2:t0:r0c0:sdt0')).toBe('en:2');
  });
});
