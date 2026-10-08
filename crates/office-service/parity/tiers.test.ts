import { expect, test } from "bun:test";
import * as Y from "yjs";
import { documentContent, structure } from "./tiers";

function state(text: string, bold: boolean, order: string[]) {
  const doc = new Y.Doc();
  doc.clientID = 7;
  const body = new Y.Text();
  doc.getMap("stories").set("body", body);
  body.insert(0, text, { bold });
  doc.getArray("pptx:slide-order").insert(0, order);
  return Y.encodeStateAsUpdate(doc);
}

test("states with other text, attributes or array order differ in both tiers", () => {
  const base = state("hello", true, ["s1", "s2"]);
  for (const other of [state("HELLO", true, ["s1", "s2"]), state("hello", false, ["s1", "s2"]), state("hello", true, ["s2", "s1"])]) {
    expect(structure(other)).not.toBe(structure(base));
    expect(documentContent(other)).not.toBe(documentContent(base));
  }
  expect(structure(state("hello", true, ["s1", "s2"]))).toBe(structure(base));
});

test("the same document written by another client is document-equal only", () => {
  const other = new Y.Doc();
  other.clientID = 8;
  const body = new Y.Text();
  other.getMap("stories").set("body", body);
  body.insert(0, "hello", { bold: true });
  other.getArray("pptx:slide-order").insert(0, ["s1", "s2"]);
  const [a, b] = [state("hello", true, ["s1", "s2"]), Y.encodeStateAsUpdate(other)];
  expect(structure(b)).not.toBe(structure(a));
  expect(documentContent(b)).toBe(documentContent(a));
});
