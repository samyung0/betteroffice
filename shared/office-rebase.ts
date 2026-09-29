import * as Y from "yjs";
import { commentOoxmlIds } from "../packages/docx/src/yrs/comments";

/**
 * Lands the edits saved after a capture (latest − captured) on seed(export).
 * Applying the later edits to the captured state yields Yjs changes at
 * captured positions; each is replayed at the corresponding position of the
 * seed, which is the parse of the captured state's export. Entities keep the
 * seed's ids, so a rebased state is the seed plus the later edits and its
 * effects baseline derives from the export. A change the seed cannot place
 * fails the rebase.
 */
export interface Lineage {
  maps: readonly string[];
  arrays: readonly string[];
  /** Root maps keyed by entity id: their keys, and id strings in values, are translated. */
  keyed: ReadonlySet<string>;
  /** A map value that names positions in texts, rewritten once every text has landed. */
  positions?: {
    root: string;
    key: string;
    rewrite(
      value: unknown,
      from: Y.Doc,
      to: Y.Doc,
      id: (value: string) => string
    ): unknown;
  };
  /** Brings the rebased state in line with what its export writes, once every change landed. */
  settle?(doc: Y.Doc): void;
}

export class RebaseError extends Error {}

const fail = (message: string): never => {
  throw new RebaseError(`Office rebase: ${message}`);
};

export const DOCX_LINEAGE: Lineage = {
  maps: ["stories", "comments"],
  arrays: [],
  keyed: new Set(["stories", "comments"]),
  positions: {
    root: "comments",
    key: "anchors",
    rewrite(value, from, to, id) {
      if (!Array.isArray(value)) fail("comment anchors are not a list");
      return (
        value as Array<{ story: string; start: Uint8Array; end: Uint8Array }>
      ).map(({ story, start, end }) => {
        const source = from.getMap("stories").get(story);
        const target = to.getMap("stories").get(id(story));
        if (!(source instanceof Y.Text) || !(target instanceof Y.Text))
          fail(`comment anchor story ${story} is missing`);
        const resolve = (bytes: Uint8Array) => {
          const position = Y.createAbsolutePositionFromRelativePosition(
            Y.decodeRelativePosition(bytes),
            from
          );
          if (!position) fail("a comment anchor no longer resolves");
          return position!;
        };
        const [first, last] = [resolve(start), resolve(end)];
        // Both stories hold the later edits, so they differ only where the
        // export wrote the text its own way; an anchor there cannot land.
        const f = aligned(source as Y.Text, target as Y.Text);
        let [startAt, endAt] = [
          pointAt(f, first.index),
          pointAt(f, last.index),
        ];
        if (last.index > first.index) {
          const covered = f.map.subarray(first.index, last.index);
          if (
            covered[0] < 0 ||
            covered.some(
              (unit, index) =>
                index > 0 && unit !== past(f, covered[index - 1] + 1)
            )
          )
            fail(
              "a comment anchor covers content the export wrote differently"
            );
          [startAt, endAt] = [covered[0], covered[covered.length - 1] + 1];
        }
        if (startAt === undefined || endAt === undefined)
          fail(
            "a comment anchor lands in content the export wrote differently"
          );
        const at = (index: number, assoc: number) =>
          Y.encodeRelativePosition(
            Y.createRelativePositionFromTypeIndex(
              target as Y.Text,
              index,
              assoc
            )
          );
        return {
          story: id(story),
          start: at(startAt!, first.assoc),
          end: at(endAt!, last.assoc),
        };
      });
    },
  },
  // A comment the later edits removed takes its reference field with it, as
  // removing it in the editor does: an export drops a field naming no comment.
  settle(doc) {
    const comments = doc.getMap("comments");
    for (const text of doc.getMap("stories").values())
      if (text instanceof Y.Text) {
        for (const [offset, embed] of embeds(text).reverse())
          if (
            embed.get("modelKind") === "commentReference" &&
            !comments.has(String(embed.get("commentId")))
          )
            text.delete(offset, 1);
        pairProjectedChildren(text);
      }
  },
};

/**
 * Gives each projected child of a field result the number of the field that
 * closes it, which the export pairs them by (`restoreProjectedFieldResults`):
 * text the later edits typed there carries the capture's number, and the
 * export's seed may number that field otherwise.
 */
function pairProjectedChildren(text: Y.Text): void {
  type Marker = { id: number; index: number };
  const pending: Array<[offset: number, length: number, marker: Marker]> = [];
  const fixes: Array<[offset: number, length: number, marker: Marker]> = [];
  let offset = 0;
  for (const { insert, attributes } of text.toDelta() as Array<{
    insert: unknown;
    attributes?: { fieldResult?: Marker };
  }>) {
    const length = typeof insert === "string" ? insert.length : 1;
    const marker = attributes?.fieldResult;
    if (marker) pending.push([offset, length, marker]);
    else if (insert instanceof Y.Map) {
      const projection = insert.get("resultProjection") as { id?: number } | undefined;
      if (insert.get("_kind") === "pilcrow") pending.length = 0;
      else if (typeof projection?.id === "number") {
        for (const [at, units, child] of pending)
          if (child.id !== projection.id)
            fixes.push([at, units, { ...child, id: projection.id }]);
        pending.length = 0;
      }
    }
    offset += length;
  }
  for (const [at, units, marker] of fixes) text.format(at, units, { fieldResult: marker });
}

export const PPTX_LINEAGE: Lineage = {
  maps: [
    "pptx:meta",
    "pptx:slides",
    "pptx:shapes",
    "pptx:stories",
    "pptx:comments",
  ],
  arrays: ["pptx:slide-order"],
  keyed: new Set([
    "pptx:slides",
    "pptx:shapes",
    "pptx:stories",
    "pptx:comments",
  ]),
};

type Path = (string | number)[];
interface Delta {
  insert?: unknown;
  retain?: number;
  delete?: number;
  attributes?: Record<string, unknown>;
}
type Step =
  | { kind: "text"; root: string; path: Path; delta: Delta[]; target: Y.Text }
  | { kind: "array"; root: string; path: Path; target: Y.Array<unknown> }
  | {
      kind: "map";
      root: string;
      path: Path;
      keys: Array<[string, "add" | "update" | "delete"]>;
      target: Y.Map<unknown>;
    };

function open(lineage: Lineage, update: Uint8Array, clientID?: number): Y.Doc {
  const doc = new Y.Doc();
  if (clientID !== undefined) doc.clientID = clientID;
  for (const name of lineage.maps) doc.getMap(name);
  for (const name of lineage.arrays) doc.getArray(name);
  Y.applyUpdate(doc, update);
  if (doc.store.pendingStructs || doc.store.pendingDs)
    fail("a state refers to content it does not hold");
  return doc;
}

function root(
  lineage: Lineage,
  doc: Y.Doc,
  name: string
): Y.Map<unknown> | Y.Array<unknown> {
  return lineage.arrays.includes(name) ? doc.getArray(name) : doc.getMap(name);
}

/** Visible items of a text with their offsets. */
function* items(text: Y.Text): Generator<[number, Y.Item]> {
  let offset = 0;
  for (let item = text._start; item; item = item.right) {
    if (item.deleted || !item.countable) continue;
    yield [offset, item];
    offset += item.length;
  }
}

/** A text's units, one per position: each UTF-16 code unit, or its embed's kind; with the Yjs id of each. */
interface Units {
  keys: string[];
  ids: string[];
}
function units(text: Y.Text): Units {
  const found: Units = { keys: [], ids: [] };
  for (const [, item] of items(text)) {
    // A string's content is one entry per UTF-16 code unit.
    const plain = item.content instanceof Y.ContentString;
    item.content.getContent().forEach((unit, i) => {
      found.keys.push(
        plain
          ? unit
          : unit instanceof Y.Map &&
            unit.get("modelKind") === "commentReference"
          ? COMMENT_REFERENCE
          : `\u0000${unit instanceof Y.Map ? unit.get("_kind") : "embed"}`
      );
      found.ids.push(`${item.id.client}:${item.id.clock + i}`);
    });
  }
  return found;
}

/**
 * The unit of the reference field a save writes at the end of a DOCX comment
 * made in the editor. The later edits never hold one the seed added, so such
 * a seed unit is transparent: content beside it lands beside it.
 */
const COMMENT_REFERENCE = "\u0000commentReference";

/** The seed unit at or after `s` that is not a comment reference only the seed holds. */
function past(f: Alignment, s: number): number {
  while (
    s < f.to.keys.length &&
    f.to.keys[s] === COMMENT_REFERENCE &&
    !f.matched[s]
  )
    s++;
  return s;
}

/** Embedded maps of a text with their positions. */
function embeds(text: Y.Text): Array<[number, Y.Map<unknown>]> {
  const found: Array<[number, Y.Map<unknown>]> = [];
  for (const [offset, item] of items(text)) {
    const [first] = item.content.getContent();
    if (first instanceof Y.Map) found.push([offset, first]);
  }
  return found;
}

/** An edit script longer than this means the export did not reproduce the text. */
const MAX_DIFFERENCES = 1000;

/**
 * Where each captured unit stands in the seed's text, or -1 where the export
 * wrote it differently: the shortest edit script between the two (Myers),
 * after their shared prefix and suffix.
 */
interface Alignment {
  from: Units;
  to: Units;
  map: Int32Array;
  /** Seed units some captured unit maps to. */
  matched: Uint8Array;
}
function align(from: Units, to: Units): Alignment {
  const f = myers(from, to);
  const matched = new Uint8Array(to.keys.length);
  for (const target of f.map) if (target >= 0) matched[target] = 1;
  return { ...f, matched };
}
function myers(from: Units, to: Units): Omit<Alignment, "matched"> {
  const a = from.keys;
  const b = to.keys;
  const map = new Int32Array(a.length).fill(-1);
  let start = 0;
  while (start < a.length && start < b.length && a[start] === b[start])
    map[start] = start++;
  let endA = a.length;
  let endB = b.length;
  while (endA > start && endB > start && a[endA - 1] === b[endB - 1])
    map[--endA] = --endB;
  const n = endA - start;
  const m = endB - start;
  if (!n || !m) return { from, to, map };
  const max = Math.min(n + m, MAX_DIFFERENCES);
  const offset = max + 1;
  const v = new Int32Array(2 * max + 3);
  const trace: Int32Array[] = [];
  for (let d = 0; d <= max; d++) {
    for (let k = -d; k <= d; k += 2) {
      let x =
        k === -d || (k !== d && v[offset + k - 1] < v[offset + k + 1])
          ? v[offset + k + 1]
          : v[offset + k - 1] + 1;
      let y = x - k;
      while (x < n && y < m && a[start + x] === b[start + y]) {
        x++;
        y++;
      }
      v[offset + k] = x;
      if (x < n || y < m) continue;
      trace.push(v.slice());
      // Walk the script back, matching every diagonal step.
      for (let e = d; e > 0; e--) {
        const previous = trace[e - 1];
        const key = x - y;
        const back =
          key === -e ||
          (key !== e && previous[offset + key - 1] < previous[offset + key + 1])
            ? key + 1
            : key - 1;
        const px = previous[offset + back];
        const py = px - back;
        while (x > px && y > py) map[start + --x] = start + --y;
        x = px;
        y = py;
      }
      while (x > 0 && y > 0) map[start + --x] = start + --y;
      return { from, to, map };
    }
    trace.push(v.slice());
  }
  return fail("the export wrote a text too differently");
}

/**
 * A story's alignment with its rebased counterpart, made once: comment
 * anchors are rewritten after every text has landed, so the rebased story no
 * longer changes, and each transplant's stories are new objects.
 */
const anchorAlignments = new WeakMap<Y.Text, Alignment>();
function aligned(source: Y.Text, target: Y.Text): Alignment {
  let f = anchorAlignments.get(target);
  if (!f)
    anchorAlignments.set(target, (f = align(units(source), units(target))));
  return f;
}

/** The seed position of an insertion before captured position `c`. */
function pointAt(f: Alignment, c: number): number | undefined {
  const left = c === 0 ? -1 : f.map[c - 1];
  const right = c === f.map.length ? f.to.keys.length : f.map[c];
  if (left >= 0 || c === 0)
    return right >= 0 || c === f.map.length
      ? right === past(f, left + 1)
        ? right
        : undefined
      : left + 1;
  return right >= 0 ? right : undefined;
}

/** A captured-position text delta as the same edit of the seed text. */
function landDelta(
  f: Alignment,
  delta: Delta[],
  copy: (value: unknown) => unknown,
  where: string
): Delta[] {
  const out: Delta[] = [];
  let c = 0;
  let s = 0;
  const move = (to: number) => {
    if (to > s) out.push({ retain: to - s });
    s = to;
  };
  // A range lands in runs of units that stay together in the seed.
  const range = (
    length: number,
    op: (length: number) => Delta,
    deleting = false
  ) => {
    for (let at = c; at < c + length; ) {
      const target = f.map[at];
      if (target < 0)
        fail(
          `a change at ${where}:${at} touches content the export wrote differently`
        );
      // Reference fields only the seed holds, inside a delete, go with it.
      if (deleting && at > c && target > s && past(f, s) === target) {
        out.push({ delete: target - s });
        s = target;
      }
      let run = 1;
      while (at + run < c + length && f.map[at + run] === target + run) run++;
      move(target);
      out.push(op(run));
      s += run;
      at += run;
    }
    c += length;
  };
  for (const op of delta) {
    if (op.insert !== undefined) {
      const at = pointAt(f, c);
      if (at === undefined)
        fail(
          `an insertion at ${where}:${c} lands in content the export wrote differently`
        );
      move(at!);
      out.push({
        ...op,
        insert: typeof op.insert === "string" ? op.insert : copy(op.insert),
      });
    } else if (op.delete !== undefined)
      range(op.delete, (length) => ({ delete: length }), true);
    else if (op.attributes)
      range(op.retain!, (length) => ({
        retain: length,
        attributes: op.attributes,
      }));
    else c += op.retain!;
  }
  return out;
}

/** A text's unit keys apart from the units it shares with nothing on the other side. */
const outside = (text: Units, apart: ReadonlySet<string>) =>
  text.keys.filter((_, index) => !apart.has(text.ids[index])).join("\u0001");

function translate(value: unknown, id: (value: string) => string): unknown {
  if (typeof value === "string") return id(value);
  if (
    value instanceof Uint8Array ||
    value === null ||
    typeof value !== "object"
  )
    return value;
  if (Array.isArray(value)) return value.map((item) => translate(item, id));
  return Object.fromEntries(
    Object.entries(value).map(([key, item]) => [key, translate(item, id)])
  );
}

/**
 * latest − captured landed on `seed`, written by `clientID`. `ids` maps
 * captured entity ids to the seed's; an entity the later edits added keeps
 * its id unless the seed uses it. Returns the state and every id it changed.
 */
export function transplant(
  lineage: Lineage,
  captured: Uint8Array,
  latest: Uint8Array,
  seed: Uint8Array,
  ids: ReadonlyMap<string, string>,
  clientID: number
): { state: Uint8Array; ids: Map<string, string> } {
  const before = open(lineage, captured);
  const edited = open(lineage, captured);
  const later = open(lineage, latest);
  const result = open(lineage, seed, clientID);
  try {
    const steps: Step[] = [];
    for (const name of [...lineage.maps, ...lineage.arrays])
      root(lineage, edited, name).observeDeep((events) => {
        for (const event of events) {
          const common = { root: name, path: event.path as Path };
          if (event.target instanceof Y.Text)
            steps.push({
              ...common,
              kind: "text",
              delta: event.delta as Delta[],
              target: event.target,
            });
          else if (event.target instanceof Y.Array)
            steps.push({ ...common, kind: "array", target: event.target });
          else if (event.target instanceof Y.Map)
            steps.push({
              ...common,
              kind: "map",
              keys: [...event.changes.keys].map(([key, change]) => [
                key,
                change.action,
              ]),
              target: event.target,
            });
          else fail(`unexpected ${event.target.constructor.name} change`);
        }
      });
    Y.applyUpdate(
      edited,
      Y.encodeStateAsUpdate(later, Y.encodeStateVector(edited))
    );

    const renamed = new Map<string, string>();
    for (const step of steps) {
      if (
        step.kind !== "map" ||
        step.path.length ||
        !lineage.keyed.has(step.root)
      )
        continue;
      const taken = result.getMap(step.root);
      for (const [key, action] of step.keys) {
        if (action !== "add" || !taken.has(key)) continue;
        let n = 1;
        while (
          taken.has(`${key}~${n}`) ||
          [...renamed.values()].includes(`${key}~${n}`)
        )
          n++;
        renamed.set(key, `${key}~${n}`);
      }
    }
    const id = (value: string) => ids.get(value) ?? renamed.get(value) ?? value;

    // Every captured entity a change touches has a counterpart in the seed;
    // a text's positions correspond through its alignment.
    const touched = new Map<string, Set<string>>();
    for (const step of steps) {
      if (!lineage.keyed.has(step.root)) continue;
      const keys = touched.get(step.root) ?? new Set<string>();
      touched.set(step.root, keys);
      if (step.path.length) keys.add(step.path[0] as string);
      else if (step.kind === "map")
        for (const [key] of step.keys) keys.add(key);
    }
    const alignments = new Map<string, Alignment>();
    for (const [name, keys] of touched)
      for (const key of keys) {
        const value = before.getMap(name).get(key);
        if (value === undefined || value === null) continue;
        const counterpart = ids.has(key)
          ? result.getMap(name).get(id(key))
          : undefined;
        if (
          (counterpart as object | undefined)?.constructor !== value.constructor
        )
          fail(`${name} ${key} has no counterpart in the export`);
        if (value instanceof Y.Text)
          alignments.set(
            `${name}\u0000${key}`,
            align(units(value), units(counterpart as Y.Text))
          );
      }

    const positions: Array<{ map: Y.Map<unknown>; value: unknown }> = [];
    const copy = (value: unknown, positioned = false): unknown => {
      if (value instanceof Y.Map) {
        const map = new Y.Map();
        for (const [key, item] of value.entries())
          if (positioned && key === lineage.positions?.key)
            positions.push({ map, value: item });
          else map.set(key, copy(item));
        return map;
      }
      if (value instanceof Y.Array) {
        const array = new Y.Array();
        array.push(value.toArray().map((item) => copy(item)));
        return array;
      }
      if (value instanceof Y.Text) {
        const text = new Y.Text();
        text.applyDelta(
          (value.toDelta() as Delta[]).map((op) => ({
            ...op,
            insert: typeof op.insert === "string" ? op.insert : copy(op.insert),
          })),
          { sanitize: false }
        );
        return text;
      }
      if (value instanceof Y.AbstractType)
        fail(`unexpected ${value.constructor.name} value`);
      return translate(value, id);
    };
    const byKey = (step: Step): unknown => {
      let target: unknown = root(lineage, result, step.root);
      step.path.forEach((segment, depth) => {
        target =
          target instanceof Y.Map
            ? target.get(
                depth === 0 && lineage.keyed.has(step.root)
                  ? id(segment as string)
                  : (segment as string)
              )
            : undefined;
      });
      return target;
    };
    // An embed's changes land on its seed counterpart, found through the
    // captured position of the same item before any text changes.
    const targets = new Map<Step, unknown>();
    for (const step of steps) {
      const embedded = step.path.some((segment) => typeof segment === "number");
      let target: unknown;
      if (!embedded) target = byKey(step);
      else {
        const [key, index, ...rest] = step.path;
        const counterpart = result.getMap(step.root).get(id(key as string));
        const f = alignments.get(`${step.root}\u0000${key}`);
        const item = step.target._item;
        if (rest.length || typeof index !== "number" || !f || !item)
          fail(`unexpected change at ${[step.root, ...step.path].join("/")}`);
        const c = f!.from.ids.indexOf(`${item!.id.client}:${item!.id.clock}`);
        const s = c < 0 ? -1 : f!.map[c];
        target = s < 0 ? undefined : childAtOffset(counterpart as Y.Text, s);
      }
      if (
        (target as object | undefined)?.constructor !== step.target.constructor
      )
        fail(
          `${[step.root, ...step.path].join("/")} is missing from the export`
        );
      targets.set(step, target);
    }

    // Parents first: an entity lands before changes inside it.
    steps.sort((a, b) => a.path.length - b.path.length);
    result.transact(() => {
      for (const step of steps) {
        const target = targets.get(step);
        if (step.kind === "text") {
          const f = alignments.get(`${step.root}\u0000${step.path[0]}`);
          if (step.path.length !== 1 || !f)
            fail(`unexpected text at ${[step.root, ...step.path].join("/")}`);
          (target as Y.Text).applyDelta(
            landDelta(
              f!,
              step.delta,
              (value) => copy(value),
              [step.root, ...step.path].join("/")
            ),
            { sanitize: false }
          );
        } else if (step.kind === "array") {
          const array = target as Y.Array<unknown>;
          array.delete(0, array.length);
          array.insert(
            0,
            step.target.toArray().map((item) => copy(item))
          );
        } else {
          const map = target as Y.Map<unknown>;
          const entities = !step.path.length && lineage.keyed.has(step.root);
          const positioned = step.root === lineage.positions?.root;
          for (const [key, action] of step.keys) {
            const name = entities ? id(key) : key;
            if (action === "delete") map.delete(name);
            else if (positioned && !entities && key === lineage.positions?.key)
              positions.push({ map, value: step.target.get(key) });
            else
              map.set(name, copy(step.target.get(key), positioned && entities));
          }
        }
      }
      for (const { map, value } of positions)
        map.set(
          lineage.positions!.key,
          lineage.positions!.rewrite(value, edited, result, id)
        );
      lineage.settle?.(result);
    });

    // Every touched entity reads as in the latest state, but for what the
    // export wrote its own way in a text it kept.
    const replaced = new Set(
      steps.flatMap((step) =>
        step.kind === "map" && !step.path.length && lineage.keyed.has(step.root)
          ? step.keys
              .filter(([, action]) => action !== "delete")
              .map(([key]) => `${step.root}\u0000${key}`)
          : []
      )
    );
    for (const [name, keys] of touched)
      for (const key of keys) {
        const value = later.getMap(name).get(key);
        const next = result.getMap(name).get(id(key));
        if ((value === undefined) !== (next === undefined))
          fail(`${name} ${key} did not land`);
        if (!(value instanceof Y.Text)) continue;
        // Units the export wrote differently stay as they were on each side.
        const f = replaced.has(`${name}\u0000${key}`)
          ? undefined
          : alignments.get(`${name}\u0000${key}`);
        const fromApart = new Set<string>();
        const toApart = new Set<string>(f?.to.ids);
        f?.map.forEach((target, index) => {
          if (target < 0) fromApart.add(f.from.ids[index]);
          else toApart.delete(f.to.ids[target]);
        });
        if (
          outside(units(value), fromApart) !==
          outside(units(next as Y.Text), toApart)
        )
          fail(`${name} ${key} did not land`);
      }
    return {
      state: Y.encodeStateAsUpdate(result),
      ids: new Map([...ids, ...renamed]),
    };
  } finally {
    for (const doc of [before, edited, later, result]) doc.destroy();
  }
}

/** The embed at `offset` of a text. */
function childAtOffset(text: Y.Text, offset: number): unknown {
  for (const [start, item] of items(text))
    if (offset < start + item.length)
      return item.content.getContent()[offset - start];
  return undefined;
}

/**
 * Captured DOCX story, paragraph and comment ids to seed(export)'s: stories by
 * their place, nested ones through their tables and content controls,
 * paragraphs by their marks' positions, comments by the numeric id a save
 * writes for them.
 */
export function docxIds(
  captured: Uint8Array,
  seed: Uint8Array
): Map<string, string> {
  const from = open(DOCX_LINEAGE, captured);
  const to = open(DOCX_LINEAGE, seed);
  try {
    const ids = new Map<string, string>();
    const stories = [from.getMap("stories"), to.getMap("stories")] as const;
    const pair = (a: unknown, b: unknown) => {
      if (typeof a !== "string" || typeof b !== "string" || ids.has(a)) return;
      ids.set(a, b);
      const left = stories[0].get(a);
      const right = stories[1].get(b);
      if (!(left instanceof Y.Text) || !(right instanceof Y.Text)) return;
      let f: Alignment;
      try {
        f = align(units(left), units(right));
      } catch {
        return; // Changes inside it cannot land and fail the rebase.
      }
      const theirs = new Map(embeds(right));
      for (const [index, embed] of embeds(left)) {
        const other = theirs.get(f.map[index]);
        const kind = embed.get("_kind");
        if (!other || other.get("_kind") !== kind) continue;
        if (kind === "pilcrow") pair(embed.get("paraId"), other.get("paraId"));
        if (kind === "blockSdt") pair(embed.get("story"), other.get("story"));
        if (kind !== "table") continue;
        const rows = embed.get("rows") as
          | Array<{ cells?: Array<{ story?: string }> }>
          | undefined;
        const otherRows = other.get("rows") as typeof rows;
        rows?.forEach((row, r) =>
          row.cells?.forEach((cell, c) =>
            pair(cell.story, otherRows?.[r]?.cells?.[c]?.story)
          )
        );
      }
    };
    for (const key of stories[0].keys())
      if (/^(body|(hf|fn|en):[^:]+)$/.test(key) && stories[1].has(key))
        pair(key, key);
    const comments = [...from.getMap("comments").entries()]
      .filter(([, comment]) => comment instanceof Y.Map)
      .map(([key]) => key);
    for (const [key, saved] of commentOoxmlIds(comments))
      if (to.getMap("comments").has(String(saved))) ids.set(key, String(saved));
    return ids;
  } finally {
    from.destroy();
    to.destroy();
  }
}

/**
 * Story and id of each DOCX paragraph whose mark or id `clientID` wrote into
 * `state`: the paragraphs whose identity came from the later edits.
 */
export function docxWrittenParagraphs(
  state: Uint8Array,
  clientID: number
): Array<[string, string]> {
  const doc = open(DOCX_LINEAGE, state);
  try {
    const found: Array<[string, string]> = [];
    for (const [story, text] of doc.getMap("stories").entries())
      if (text instanceof Y.Text)
        for (const [, embed] of embeds(text))
          if (
            embed.get("_kind") === "pilcrow" &&
            (embed._item?.id.client === clientID ||
              embed._map.get("paraId")?.id.client === clientID)
          )
            found.push([story, String(embed.get("paraId"))]);
    return found;
  } finally {
    doc.destroy();
  }
}

/**
 * A PPTX slide or shape the later edits brought back (an Undo of a deletion
 * made before the capture) is bound to source XML the export dropped.
 */
export function assertPptxRestorations(
  seed: Uint8Array,
  state: Uint8Array
): void {
  const published = open(PPTX_LINEAGE, seed);
  const rebased = open(PPTX_LINEAGE, state);
  try {
    for (const [root, bound] of [
      ["pptx:slides", "sourcePartPath"],
      ["pptx:shapes", "sourceId"],
    ] as const)
      for (const [id, entity] of rebased.getMap(root).entries())
        if (
          !published.getMap(root).has(id) &&
          entity instanceof Y.Map &&
          entity.get(bound)
        )
          fail(`${root} ${id} is bound to source XML the export dropped`);
  } finally {
    published.destroy();
    rebased.destroy();
  }
}

interface DeckShape {
  id: string;
  kind: string;
  textStories: Array<{ id: string; paragraphs: Array<{ id: string }> }>;
  children: DeckShape[];
}
interface Deck {
  slides: Array<{ id: string; shapes: DeckShape[] }>;
  comments?: Array<{
    id: string;
    slideId: string;
    author: string;
    text: string;
    created?: string | null;
  }>;
}

/** Captured PPTX slide, shape, story, paragraph and comment ids to seed(export)'s, by their place in the deck. */
export function pptxIds(captured: Deck, exported: Deck): Map<string, string> {
  const ids = new Map<string, string>();
  if (captured.slides.length !== exported.slides.length)
    fail("the export changed the slide count");
  const shapes = (a: DeckShape[], b: DeckShape[]) => {
    if (a.length !== b.length) fail("the export changed a slide's shapes");
    a.forEach((shape, index) => {
      const other = b[index];
      if (
        shape.kind !== other.kind ||
        shape.textStories.length !== other.textStories.length
      )
        fail("the export changed a shape");
      ids.set(shape.id, other.id);
      shape.textStories.forEach((story, k) => {
        const theirs = other.textStories[k];
        if (story.paragraphs.length !== theirs.paragraphs.length)
          fail("the export changed a text's paragraphs");
        ids.set(story.id, theirs.id);
        story.paragraphs.forEach((paragraph, p) =>
          ids.set(paragraph.id, theirs.paragraphs[p].id)
        );
      });
      shapes(shape.children, other.children);
    });
  };
  captured.slides.forEach((slide, index) => {
    ids.set(slide.id, exported.slides[index].id);
    shapes(slide.shapes, exported.slides[index].shapes);
  });
  const key = (
    comment: NonNullable<Deck["comments"]>[number],
    id: (value: string) => string
  ) =>
    JSON.stringify([
      id(comment.slideId),
      comment.author,
      comment.text,
      comment.created ?? null,
    ]);
  const theirs = new Map<string, string[]>();
  for (const comment of exported.comments ?? [])
    theirs.set(
      key(comment, (value) => value),
      [...(theirs.get(key(comment, (value) => value)) ?? []), comment.id]
    );
  for (const comment of captured.comments ?? []) {
    const match = theirs.get(key(comment, (value) => ids.get(value) ?? value));
    if (match?.length === 1) ids.set(comment.id, match[0]);
  }
  return ids;
}
