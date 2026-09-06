import type { Comment, Paragraph } from '../types/content';
import type { YrsSession } from './index';

export function commentSharedId(comment: Comment): string {
  return comment.sharedId ?? String(comment.id);
}

export function commentNumericId(id: string): number {
  if (/^\d+$/.test(id) && Number(id) <= 0x7fffffff) return Number(id);
  let hash = 2166136261;
  for (const byte of new TextEncoder().encode(id)) hash = Math.imul(hash ^ byte, 16777619);
  return hash >>> 1;
}

/** Projects stable shared keys into collision-free OOXML numeric IDs. */
export function projectYrsComments(
  session: Pick<YrsSession, 'listComments'>,
  base: readonly Comment[] = []
): Comment[] {
  const records = session.listComments().sort((a, b) => (a.id < b.id ? -1 : a.id > b.id ? 1 : 0));
  const ids = new Map<string, number>();
  const used = new Set<number>();
  for (const record of records) {
    if (/^\d+$/.test(record.id) && Number(record.id) <= 0x7fffffff) {
      ids.set(record.id, Number(record.id));
      used.add(Number(record.id));
    }
  }
  for (const record of records) {
    if (ids.has(record.id)) continue;
    let id = commentNumericId(record.id);
    while (used.has(id)) id = (id + 1) & 0x7fffffff;
    ids.set(record.id, id);
    used.add(id);
  }
  const originals = new Map(base.map((comment) => [commentSharedId(comment), comment]));
  return records
    .filter((record) => record.parentId === null || ids.has(record.parentId))
    .map((record) => {
      const original = originals.get(record.id);
      if (record.body !== null && !Array.isArray(record.body))
        throw new Error('Invalid shared comment body');
      const content =
        record.body === null ? (original?.content ?? []) : (record.body as Paragraph[]);
      const parentId = record.parentId === null ? undefined : ids.get(record.parentId);
      return {
        ...original,
        id: ids.get(record.id)!,
        sharedId: record.id,
        author: record.body === null && original ? original.author : record.author,
        date: record.body === null && original ? original.date : record.date,
        content,
        blockContent: undefined,
        parentId,
        done: record.done,
        status: record.done ? 'resolved' : 'active',
      };
    });
}
