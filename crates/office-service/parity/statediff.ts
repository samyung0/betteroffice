// Runs one recorded call (by method and index among those calls in a file)
// through the native crate and prints how the TS and native states differ.
//   bun crates/office-service/parity/statediff.ts <dir> <binary> <file> <method> <index>
import { readFileSync } from "node:fs";
import { join } from "node:path";
import * as Y from "yjs";

import { createHash } from "node:crypto";
const sha = (bytes: Uint8Array) => createHash("sha256").update(bytes).digest("hex");
const [dir, binary, file, method, nth] = process.argv.slice(2);
const blob = (sha: string) => new Uint8Array(readFileSync(join(dir, "blobs", sha)));
const wire = (value: any): any =>
  Array.isArray(value)
    ? value.map(wire)
    : value && typeof value === "object"
      ? value.$blob
        ? { $bytes: Buffer.from(blob(value.$blob)).toString("base64") }
        : Object.fromEntries(Object.entries(value).map(([key, item]) => [key, wire(item)]))
      : value;
const calls = readFileSync(join(dir, file), "utf8")
  .split("\n")
  .filter(Boolean)
  .map((line) => JSON.parse(line))
  .filter((call) => call.method === method);
const call = calls[Number(nth)];
const child = Bun.spawnSync([binary, "serve"], {
  stdin: Buffer.from(
    `${JSON.stringify({ method, args: call.args.map(wire), clients: call.clients, nowMs: call.nowMs, perfMs: call.perfMs })}\n`
  ),
});
const answer = JSON.parse(child.stdout.toString());
console.log("clients", call.clients, "TS error", call.error, "native error", answer.error);
if (!call.value || !answer.value) process.exit(0);
const ts = blob(call.value.state.$blob);
const native = new Uint8Array(Buffer.from(answer.value.state.$bytes, "base64"));
const units = (state: Uint8Array) => {
  const { structs, ds } = Y.decodeUpdate(state);
  const out = new Map<string, string>();
  for (const struct of structs) {
    const item = struct as Y.Item;
    const content = item.content?.getContent?.() ?? [];
    for (let offset = 0; offset < struct.length; offset++)
      out.set(
        `${struct.id.client}:${struct.id.clock + offset}`,
        JSON.stringify([
          struct.constructor.name,
          item.parentSub ?? null,
          typeof content[offset] === "string" ? content[offset] : content[offset] ? "type" : null,
        ])
      );
  }
  const deleted = new Set<string>();
  ds.clients.forEach((ranges, client) => {
    for (const range of ranges) for (let at = 0; at < range.len; at++) deleted.add(`${client}:${range.clock + at}`);
  });
  return { out, deleted };
};
const [a, b] = [units(ts), units(native)];
const only = (x: Map<string, string>, y: Map<string, string>) => [...x.keys()].filter((key) => !y.has(key) || y.get(key) !== x.get(key));
console.log("units only/different in TS", only(a.out, b.out).slice(0, 20).map((key) => [key, a.out.get(key), b.out.get(key)]));
console.log("units only/different in native", only(b.out, a.out).slice(0, 20).map((key) => [key, b.out.get(key), a.out.get(key)]));
console.log("deleted only in TS", [...a.deleted].filter((key) => !b.deleted.has(key)).slice(0, 20));
console.log("deleted only in native", [...b.deleted].filter((key) => !a.deleted.has(key)).slice(0, 20));
const text = (state: Uint8Array) => {
  const doc = new Y.Doc();
  Y.applyUpdate(doc, state);
  const stories = doc.getMap("stories");
  return Object.fromEntries([...stories.entries()].map(([key, value]) => [key, value instanceof Y.Text ? JSON.stringify(value.toDelta().map((op: any) => [typeof op.insert === "string" ? op.insert : `[${op.insert.get?.("_kind")}]`, op.attributes ?? null])) : null]));
};
const [x, y] = [text(ts), text(native)];
for (const key of new Set([...Object.keys(x), ...Object.keys(y)]))
  if (x[key] !== y[key]) console.log("story", key, "\n TS    ", JSON.stringify(x[key]), "\n native", JSON.stringify(y[key]));
const describe = (state: Uint8Array, keys: string[]) => {
  const { structs } = Y.decodeUpdate(state);
  for (const key of keys) {
    const [client, clock] = key.split(":").map(Number);
    const struct = structs.find((s) => s.id.client === client && s.id.clock <= clock && clock < s.id.clock + s.length) as Y.Item | undefined;
    if (!struct) continue;
    const content = struct.content?.getContent?.();
    console.log(key, {
      kind: struct.content?.constructor?.name,
      parent: struct.parent && typeof struct.parent === "object" && "client" in (struct.parent as object) ? `${(struct.parent as Y.ID).client}:${(struct.parent as Y.ID).clock}` : struct.parent,
      parentSub: struct.parentSub,
      origin: struct.origin && `${struct.origin.client}:${struct.origin.clock}`,
      rightOrigin: struct.rightOrigin && `${struct.rightOrigin.client}:${struct.rightOrigin.clock}`,
      content: JSON.stringify(content?.map((value: unknown) => (value instanceof Y.AbstractType ? "type" : value))).slice(0, 200),
    });
  }
};
const extra = [...b.deleted].filter((key) => !a.deleted.has(key)).slice(0, 6);
console.log("-- TS view of native-only deletions");
describe(ts, extra);

function documentContent(state: Uint8Array): string {
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
  for (const name of [...doc.share.keys()].sort()) {
    const map = doc.getMap(name);
    roots[name] = map.size ? value(map) : null;
  }
  return JSON.stringify(roots);
}
const [p, q] = [documentContent(ts), documentContent(native)];
let at = 0;
while (at < p.length && p[at] === q[at]) at++;
console.log("document differs at", at, "of", p.length);
console.log(" TS    ", p.slice(Math.max(0, at - 300), at + 200));
console.log(" native", q.slice(Math.max(0, at - 300), at + 200));
const structural = (x: any, y: any, path: string): boolean => {
  if (JSON.stringify(x) === JSON.stringify(y)) return false;
  if (Array.isArray(x) && Array.isArray(y)) {
    for (let i = 0; i < Math.max(x.length, y.length); i++) if (structural(x[i], y[i], `${path}[${i}]`)) return true;
    return false;
  }
  if (x && y && typeof x === "object" && typeof y === "object") {
    for (const key of new Set([...Object.keys(x), ...Object.keys(y)])) if (structural(x[key], y[key], `${path}.${key}`)) return true;
    return false;
  }
  console.log("first structural difference at", path, "\n TS    ", JSON.stringify(x)?.slice(0, 300), "\n native", JSON.stringify(y)?.slice(0, 300));
  return true;
};
structural(JSON.parse(p), JSON.parse(q), "");
