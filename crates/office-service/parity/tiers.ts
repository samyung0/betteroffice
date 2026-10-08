// The comparisons replay.ts reports rebased and edited states by.
import { createHash } from "node:crypto";
import * as Y from "yjs";

const sha = (bytes: Uint8Array) => createHash("sha256").update(bytes).digest("hex");

/** Byte arrays as their SHA-256, object keys sorted. */
const canonical = (_key: string, value: unknown): unknown =>
  value instanceof Uint8Array
    ? `bytes:${sha(value)}`
    : value && typeof value === "object" && !Array.isArray(value)
      ? Object.fromEntries(Object.keys(value).sort().map((key) => [key, (value as Record<string, unknown>)[key]]))
      : value;

/** Unit `offset` of an item's content: a character, a format key and value, a type's kind, or a value. */
function unitContent(item: Y.Item, offset: number): string | null {
  const content = item.content as unknown as Record<string, unknown> | undefined;
  if (!content) return null;
  if (content instanceof Y.ContentString) return (content.str as string).slice(offset, offset + 1);
  if (content instanceof Y.ContentFormat) return JSON.stringify([content.key, content.value], canonical);
  if (content instanceof Y.ContentType) return `type:${(content.type as object).constructor.name}`;
  if (content instanceof Y.ContentEmbed) return JSON.stringify(content.embed, canonical);
  if (content instanceof Y.ContentBinary) return `bytes:${sha(content.content as Uint8Array)}`;
  if (content instanceof Y.ContentDeleted) return null;
  if (content instanceof Y.ContentDoc) return `doc:${(content.doc as Y.Doc).guid}`;
  const values = (content as unknown as { getContent(): unknown[] }).getContent();
  return JSON.stringify(values[offset], canonical);
}

/**
 * Every unit of a state: id, struct and content kind, its content, origin,
 * right origin, parent and map key, and the deleted units. Runs are split
 * per unit so differently merged encodings compare equal.
 */
export function structure(state: Uint8Array): string {
  const { structs, ds } = Y.decodeUpdate(state);
  const units: string[] = [];
  for (const struct of structs) {
    const item = struct as Y.Item;
    const kind = item.content?.constructor?.name ?? null;
    const id = (value: Y.ID | null) => (value ? [value.client, value.clock] : null);
    const parent =
      typeof item.parent === "string" ? item.parent : item.parent ? id(item.parent as unknown as Y.ID) : null;
    for (let offset = 0; offset < struct.length; offset++) {
      const unit = struct instanceof Y.Item ? unitContent(item, offset) : null;
      const origin = offset === 0 ? id(item.origin ?? null) : [struct.id.client, struct.id.clock + offset - 1];
      units.push(
        JSON.stringify([
          struct.id.client,
          struct.id.clock + offset,
          struct.constructor.name,
          kind,
          unit,
          origin,
          id(item.rightOrigin ?? null),
          parent,
          item.parentSub ?? null,
        ])
      );
    }
  }
  units.sort();
  const deleted: string[] = [];
  ds.clients.forEach((ranges, client) => {
    for (const range of ranges) for (let at = 0; at < range.len; at++) deleted.push(`${client}:${range.clock + at}`);
  });
  deleted.sort();
  return JSON.stringify({ units, deleted });
}

/**
 * What a state reads as: every text's delta with attributes, every map's
 * entries, every array's items, with relative positions resolved to indexes.
 * A root holding a sequence reads as an array (`pptx:slide-order`,
 * `xlsx:sheet-order`), any other as a map. Equal documents can still differ
 * in which redundant format items are deleted (yrs's cleanup after a
 * transaction compares attribute values deeply, Yjs by identity) and in the
 * order map entries were written.
 */
export function documentContent(state: Uint8Array): string {
  const doc = new Y.Doc();
  Y.applyUpdate(doc, state);
  const value = (item: unknown): unknown => {
    if (item instanceof Y.Text) {
      // Adjacent runs with equal attributes read as one run.
      const runs: Array<[unknown, unknown]> = [];
      for (const op of item.toDelta() as Array<{ insert: unknown; attributes?: unknown }>) {
        const attributes = value(op.attributes ?? null);
        const last = runs.at(-1);
        if (typeof op.insert === "string" && last && typeof last[0] === "string" && JSON.stringify(last[1]) === JSON.stringify(attributes))
          last[0] += op.insert;
        else runs.push([value(op.insert), attributes]);
      }
      return runs;
    }
    if (item instanceof Y.Map) return Object.fromEntries([...item.entries()].sort(([a], [b]) => (a < b ? -1 : 1)).map(([key, entry]) => [key, value(entry)]));
    if (item instanceof Y.Array) return item.toArray().map(value);
    if (item instanceof Uint8Array) {
      try {
        const position = Y.createAbsolutePositionFromRelativePosition(Y.decodeRelativePosition(item), doc);
        return position ? `@${position.index}:${position.assoc}` : "@unresolved";
      } catch {
        return `bytes:${sha(item)}`;
      }
    }
    if (Array.isArray(item)) return item.map(value);
    if (item && typeof item === "object") return Object.fromEntries(Object.entries(item).sort(([a], [b]) => (a < b ? -1 : 1)).map(([key, entry]) => [key, value(entry)]));
    return item;
  };
  const roots: Record<string, unknown> = {};
  for (const [name, type] of [...doc.share.entries()].sort(([a], [b]) => (a < b ? -1 : 1))) {
    const sequence = (type as unknown as { _start: unknown })._start !== null;
    roots[name] = sequence ? value(doc.getArray(name)) : value(doc.getMap(name));
  }
  return JSON.stringify(roots);
}
