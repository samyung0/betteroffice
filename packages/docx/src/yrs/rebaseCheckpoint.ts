import type { Document } from '../types/document';
import { parseRelationshipsXmlWithRust } from '../docx/rustParseFacade';
import { serializeDocxS10Wire } from '../wasm/parse';
import { unzipContainer } from '../wasm/opc';
import { createYrsSession, type YrsRawOp, type YrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';

export interface DocxCheckpointRebase {
  oldSource: Uint8Array;
  capturedState: Uint8Array;
  latestState: Uint8Array;
  exportedSource: Uint8Array;
}

/** Rebind durable replicas to their own exported package without importing a new CRDT lineage. */
export async function rebaseDocxCheckpoint(input: DocxCheckpointRebase): Promise<{
  state: Uint8Array;
  indexedState: Uint8Array;
}> {
  const oldParts = unzipContainer(input.oldSource);
  const newParts = unzipContainer(input.exportedSource);
  validateExportBacking(oldParts, newParts);
  const captured = await createYrsSession();
  const latest = await createYrsSession();
  try {
    const indexedState = rebase(captured, input.capturedState, input, newParts);
    const state = rebase(latest, input.latestState, input, newParts);
    return { state, indexedState };
  } finally {
    captured.destroy();
    latest.destroy();
  }
}

type Parts = Record<string, Uint8Array>;
const decoder = new TextDecoder();

function equalBytes(left: Uint8Array | undefined, right: Uint8Array | undefined): boolean {
  return (
    !!left && !!right && left.length === right.length && left.every((byte, i) => byte === right[i])
  );
}

function ownerPath(relsPath: string): string {
  return relsPath.replace('/_rels/', '/').replace(/\.rels$/, '');
}

function targetPath(owner: string, target: string): string {
  const url = new URL(target, `https://docx.invalid/${owner}`);
  if (url.origin !== 'https://docx.invalid' || url.search || url.hash)
    throw new Error('DOCX rebase requires an internal package target');
  return decodeURIComponent(url.pathname.slice(1));
}

/** Current full DOCX export rewrites owned XML and preserves the remaining package parts. */
function validateExportBacking(oldParts: Parts, newParts: Parts): void {
  const owned = new Set([
    '[Content_Types].xml',
    'word/document.xml',
    'word/footnotes.xml',
    'word/endnotes.xml',
    'word/comments.xml',
    'word/commentsExtended.xml',
    'word/commentsIds.xml',
    'word/commentsExtensible.xml',
    'word/numbering.xml',
    'docProps/core.xml',
  ]);
  for (const [path, bytes] of Object.entries(oldParts)) {
    if (!path.endsWith('.rels')) continue;
    const next = newParts[path];
    if (!next) throw new Error(`DOCX rebase lost relationship part ${path}`);
    const previous = parseRelationshipsXmlWithRust(decoder.decode(bytes), path);
    const current = parseRelationshipsXmlWithRust(decoder.decode(next), path);
    for (const [id, relationship] of previous) {
      const replacement = current.get(id);
      if (
        !replacement ||
        replacement.type !== relationship.type ||
        replacement.target !== relationship.target ||
        replacement.targetMode !== relationship.targetMode
      )
        throw new Error(`DOCX rebase changed source relationship ${path}:${id}`);
      if (relationship.type.endsWith('/header') || relationship.type.endsWith('/footer'))
        owned.add(targetPath(ownerPath(path), relationship.target));
    }
  }
  for (const [path, bytes] of Object.entries(oldParts)) {
    if (owned.has(path) || path.endsWith('.rels')) continue;
    if (!equalBytes(bytes, newParts[path]))
      throw new Error(`DOCX rebase changed or lost unmodeled source part ${path}`);
  }
}

function packageDocument(session: YrsSession): Document {
  const document = session.materializeDocx();
  if (!document) throw new Error('DOCX rebase has no attached package');
  return document;
}

function storyOwner(story: string, base: Document): string {
  if (story === 'body' || story.startsWith('body:')) return 'word/document.xml';
  if (story.startsWith('fn:')) return 'word/footnotes.xml';
  if (story.startsWith('en:')) return 'word/endnotes.xml';
  for (const [id, relationship] of base.package.relationships ?? []) {
    const root = `hf:${id}`;
    if (story === root || story.startsWith(`${root}:`))
      return targetPath('word/document.xml', relationship.target);
  }
  throw new Error(`DOCX rebase cannot resolve story owner ${story}`);
}

function imageBindings(parts: Parts, owner: string): Map<string, string> {
  const slash = owner.lastIndexOf('/');
  const relsPath = `${owner.slice(0, slash + 1)}_rels/${owner.slice(slash + 1)}.rels`;
  const bytes = parts[relsPath];
  const result = new Map<string, string>();
  if (!bytes) return result;
  for (const [id, relationship] of parseRelationshipsXmlWithRust(decoder.decode(bytes), relsPath)) {
    if (!relationship.type.endsWith('/image') || relationship.targetMode === 'External') continue;
    const data = parts[targetPath(owner, relationship.target)];
    if (!data) throw new Error(`DOCX rebase image target missing for ${relsPath}:${id}`);
    let binary = '';
    for (const byte of data) binary += String.fromCharCode(byte);
    result.set(btoa(binary), id);
  }
  return result;
}

function authored(document: Document): string {
  const pkg = document.package;
  return JSON.stringify(
    {
      content: pkg.document.content,
      comments: pkg.document.comments,
      finalSectionProperties: (
        JSON.parse(
          serializeDocxS10Wire(
            JSON.stringify({
              family: 'section',
              properties: pkg.document.finalSectionProperties,
              determinism: {
                seed: '0'.repeat(64),
                now: '2026-01-01T00:00:00.000Z',
              },
            })
          )
        ) as { xml: string }
      ).xml,
      headers: [...(pkg.headers ?? [])],
      footers: [...(pkg.footers ?? [])],
      footnotes: pkg.footnotes,
      endnotes: pkg.endnotes,
    },
    (key, value: unknown) => {
      // Root namespace bindings are package metadata the writer normalizes.
      if (key === 'rId' || key === 'verbatimXml' || key === 'customRootBindings') return undefined;
      if (value && typeof value === 'object' && !Array.isArray(value))
        return Object.fromEntries(Object.entries(value).sort(([a], [b]) => a.localeCompare(b)));
      return value;
    }
  );
}

function rebase(
  session: YrsSession,
  state: Uint8Array,
  input: DocxCheckpointRebase,
  parts: Parts
): Uint8Array {
  session.openDocx(input.oldSource, false);
  session.loadState(state);
  const base = packageDocument(session);
  const ops = new Map<string, YrsRawOp[]>();
  const reachable = new Set<string>();
  const projectedOpaque = new Set<string>();
  const opaque = new Set<string>();
  for (const story of session.storyIds()) {
    let offset = 0;
    for (const segment of session.storySegments(story)) {
      if (segment.kind === 'embed' && segment.embedKind === 'opaque')
        opaque.add(`${story}:${offset}`);
      offset += segment.kind === 'text' ? segment.text.length : 1;
    }
  }
  const push = (story: string, op: YrsRawOp): void => {
    const batch = ops.get(story) ?? [];
    batch.push(op);
    ops.set(story, batch);
  };
  const before = yrsToDocument(session, base, {
    onStory: (story) => reachable.add(story),
    onParagraph: (story, index, paragraph, logicalId) =>
      push(story, {
        op: 'setEmbedAttr',
        index,
        key: 'sourceBinding',
        value: {
          ownerParaId: logicalId,
          paraId: paragraph.paraId ?? null,
          textId: paragraph.textId ?? null,
          renderedPageBreakBefore: paragraph.renderedPageBreakBefore ?? false,
        },
      }),
    onEmbed: (story, index, content) => {
      projectedOpaque.add(`${story}:${index}`);
      if (content.type === 'blockSdt' && opaque.has(`${story}:${index}`)) {
        push(story, {
          op: 'setEmbedAttr',
          index,
          key: 'sourceBlock',
          value: content,
        });
      }
    },
  });
  const bindings = new Map<string, Map<string, string>>();
  for (const story of reachable) {
    const owner = storyOwner(story, base);
    let images = bindings.get(owner);
    if (!images) {
      images = imageBindings(parts, owner);
      bindings.set(owner, images);
    }
    let offset = 0;
    for (const segment of session.storySegments(story)) {
      if (segment.kind === 'embed') {
        if (segment.embedKind === 'opaque' && !projectedOpaque.has(`${story}:${offset}`))
          throw new Error(`DOCX rebase cannot carry opaque content at ${story}:${offset}`);
        if (segment.embedKind === 'image') {
          const src = segment.payload.src;
          if (typeof src !== 'string' || !/^data:[^,]*;base64,/.test(src))
            throw new Error(`DOCX rebase requires image bytes at ${story}:${offset}`);
          const id = images.get(src.slice(src.indexOf(',') + 1));
          if (id)
            push(story, {
              op: 'setEmbedAttr',
              index: offset,
              key: 'rId',
              value: id,
            });
        }
      }
      offset += segment.kind === 'text' ? segment.text.length : 1;
    }
  }
  for (const comment of session.listComments()) {
    if (comment.body !== null) continue;
    const projected = before.package.document.comments?.find(
      (entry) => entry.sharedId === comment.id
    );
    if (!projected) throw new Error(`DOCX rebase cannot resolve comment ${comment.id}`);
    push('body', {
      op: 'patchComment',
      id: comment.id,
      fields: {
        body: projected.content,
        author: projected.author,
        date: projected.date,
      },
    });
  }
  for (const [story, batch] of ops) session.applyRawOps(story, batch);
  session.openDocx(input.exportedSource, false);
  const after = yrsToDocument(session, packageDocument(session));
  if (authored(before) !== authored(after))
    throw new Error('DOCX rebase changed authored content or package-owned story metadata');
  return session.encodeState();
}
