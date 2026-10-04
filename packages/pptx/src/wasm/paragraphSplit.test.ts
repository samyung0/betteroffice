import { beforeAll, describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import type { PresentationHandle, StorySnapshot, TextBoxPrimitive } from '../index';
import { initWasm, openPresentation } from '../index';

const root = resolve(import.meta.dir, '../../../..');
// The demo deck's first slide title: "Office files," then "without the office.".
const STORY = 'story:slide:0:256:shape:2:0';
const SLIDE = 'ppt/slides/slide1.xml';
// The title's first paragraph becomes a level-1 bullet with spacing and a
// hanging indent, so a split shows whether both halves keep all of it.
const PLAIN_PPR = '<a:pPr algn="l" lvl="0"></a:pPr><a:r><a:rPr lang="en-US" sz="4400" b="1"><a:solidFill><a:srgbClr val="101828"/>';
const LIST_PPR = '<a:pPr marL="457200" lvl="1" indent="-228600" algn="ctr"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="1200"/></a:spcBef><a:buFont typeface="Arial"/><a:buChar char="•"/></a:pPr>';
// "Office| files,"
const SPLIT = 'Office'.length;
let deck: Uint8Array;
let font: Uint8Array;

beforeAll(async () => {
  const [wasm, demo, liberation] = await Promise.all([
    readFile(resolve(import.meta.dir, 'generated/pptx_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
    readFile(resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')),
  ]);
  await initWasm(wasm);
  font = liberation;
  const zip = await JSZip.loadAsync(demo);
  const slide = await zip.file(SLIDE)!.async('string');
  expect(slide.split(PLAIN_PPR)).toHaveLength(2);
  zip.file(SLIDE, slide.replace(PLAIN_PPR, PLAIN_PPR.replace('<a:pPr algn="l" lvl="0"></a:pPr>', LIST_PPR)));
  deck = await zip.generateAsync({ type: 'uint8array' });
});

const open = (bytes = deck, clientId?: number) =>
  openPresentation(bytes, { clientId, fonts: [{ family: 'Arial', bytes: font }] });

const plain = (story: StorySnapshot) =>
  story.paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join('')).join('\n');

const properties = ({ alignment, level, bulletJson }: StorySnapshot['paragraphs'][number]) => ({ alignment, level, bulletJson });

async function slideXml(bytes: Uint8Array): Promise<string> {
  return (await JSZip.loadAsync(bytes)).file(SLIDE)!.async('string');
}

/** The title's `a:pPr` elements, in paragraph order. */
async function titleProperties(bytes: Uint8Array): Promise<string[]> {
  const xml = await slideXml(bytes);
  const title = xml.slice(xml.indexOf('name="Title"'), xml.indexOf('name="Subtitle"'));
  return [...title.matchAll(/<a:p>(<a:pPr[^>]*?(?:\/>|>.*?<\/a:pPr>))/g)].map((match) => match[1]);
}

/** The title as laid out: the editor's projection of the story. */
function titleLayout(handle: PresentationHandle) {
  const box = handle.layoutSlide(0).primitives.find(
    (primitive): primitive is TextBoxPrimitive => primitive.kind === 'textBox' && primitive.storyId === STORY
  )!;
  return { paragraphs: box.paragraphs, lines: box.lines.map(({ x, y, width, height, start, end }) => ({ x, y, width, height, start, end })) };
}

describe('splitting a paragraph', () => {
  test('keeps the id on the first half and every property on both, in the snapshot, the layout and the file', async () => {
    const handle = open();
    try {
      const before = handle.story(STORY);
      const original = before.paragraphs[0];
      expect(original.bulletJson).toContain('•');
      handle.insertParagraphBreak(STORY, SPLIT);

      const after = handle.story(STORY);
      expect(plain(after)).toBe('Office\n files,\nwithout the office.');
      const [first, second, third] = after.paragraphs;
      expect(first).toEqual({ ...original, runs: first.runs });
      expect(second.id).not.toBe(original.id);
      expect(second.propertiesFrom).toBe(original.id);
      expect(properties(second)).toEqual(properties(original));
      expect(third).toEqual(before.paragraphs[1]);

      const saved = handle.save();
      const [firstXml, secondXml, thirdXml] = await titleProperties(saved);
      expect(firstXml).toBe(secondXml);
      for (const part of ['marL="457200"', 'indent="-228600"', 'lvl="1"', 'algn="ctr"', '<a:spcPct val="90000"/>', '<a:spcPts val="1200"/>', '<a:buChar char="•"/>']) {
        expect(secondXml).toContain(part);
      }
      expect(thirdXml).toBe('<a:pPr algn="l" lvl="0"/>');
      const reopened = open(saved);
      try {
        expect(titleLayout(handle)).toEqual(titleLayout(reopened));
      } finally {
        reopened.dispose();
      }
    } finally {
      handle.dispose();
    }
  });

  test('Enter then Backspace at the same spot restores the paragraph, its id and the saved XML exactly', async () => {
    const handle = open();
    try {
      const before = handle.story(STORY);
      const source = await slideXml(handle.save());
      // As the editor's Enter and Backspace.
      handle.replaceText(STORY, SPLIT, SPLIT, '\n');
      handle.replaceText(STORY, SPLIT, SPLIT + 1, '');
      expect(handle.story(STORY)).toEqual(before);
      expect(await slideXml(handle.save())).toBe(source);
    } finally {
      handle.dispose();
    }
  });

  test('undoes and redoes as one step', async () => {
    const handle = open();
    try {
      const before = handle.story(STORY);
      const source = await slideXml(handle.save());
      handle.replaceText(STORY, SPLIT, SPLIT, '\n');
      const after = handle.story(STORY);

      expect(handle.undo().applied).toBe(true);
      expect(handle.story(STORY)).toEqual(before);
      expect(handle.canUndo()).toBe(false);
      expect(await slideXml(handle.save())).toBe(source);
      expect(handle.redo().applied).toBe(true);
      expect(handle.story(STORY)).toEqual(after);
    } finally {
      handle.dispose();
    }
  });

  test("lands a peer's concurrent typing in the second half and converges", () => {
    const left = open(deck, 9301);
    const right = open(deck, 9302);
    try {
      right.applyUpdate(left.encodeStateAsUpdate());
      left.applyUpdate(right.encodeStateAsUpdate());
      const original = left.story(STORY).paragraphs[0];
      left.insertParagraphBreak(STORY, SPLIT);
      // "Office fil|es," on the peer that has not seen the split yet.
      right.insertText(STORY, 'Office fil'.length, 'XX');
      left.applyUpdate(right.encodeDiff(left.encodeStateVector()));
      right.applyUpdate(left.encodeDiff(right.encodeStateVector()));

      expect(right.story(STORY)).toEqual(left.story(STORY));
      const [first, second] = left.story(STORY).paragraphs;
      expect(plain(left.story(STORY))).toBe('Office\n filXXes,\nwithout the office.');
      expect(first.id).toBe(original.id);
      expect([second.propertiesFrom, properties(second)]).toEqual([original.id, properties(original)]);
    } finally {
      left.dispose();
      right.dispose();
    }
  });

  test('two peers splitting one paragraph at once converge, and both copies lay out and save with its properties', async () => {
    const left = open(deck, 9311);
    const right = open(deck, 9312);
    try {
      right.applyUpdate(left.encodeStateAsUpdate());
      left.applyUpdate(right.encodeStateAsUpdate());
      left.insertParagraphBreak(STORY, SPLIT);
      right.insertParagraphBreak(STORY, 'Office fil'.length);
      left.applyUpdate(right.encodeDiff(left.encodeStateVector()));
      right.applyUpdate(left.encodeDiff(right.encodeStateVector()));

      const merged = left.story(STORY);
      expect(right.story(STORY)).toEqual(merged);
      expect(plain(merged)).toBe('Office\n fil\nes,\nwithout the office.');
      const saved = left.save();
      const pPr = await titleProperties(saved);
      expect(pPr.slice(0, 3).every((xml) => xml === pPr[0] && xml.includes('marL="457200"'))).toBe(true);
      const reopened = open(saved);
      try {
        expect(titleLayout(left)).toEqual(titleLayout(reopened));
      } finally {
        reopened.dispose();
      }
    } finally {
      left.dispose();
      right.dispose();
    }
  });
});
