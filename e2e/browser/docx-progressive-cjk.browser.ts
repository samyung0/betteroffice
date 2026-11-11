import { expect, test, type Locator, type Page } from 'playwright/test';
import JSZip from 'jszip';
import { pinnedEvaluate } from './pinned-evaluate';

// Capy's DOCX editor paints in its resident engine worker, so whether a sliced
// face's outlines arrived is read from the pixels. That the sliced faces lay out
// and draw exactly as the whole faces do is proven by the unit tests
// (viewer.slices.test.ts, residentSlicedFonts.test.ts): this build can no longer
// load a DOCX's CJK faces whole.
const japanese = '鬱蒼とした森の奥に、齟齬と矛盾が潜む。';
const traditional = '繁體中文的排版測試';
const hangul = '한글';
const relationships = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const rels = (body: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;

// The built app names each slice file `<name>-<hash>.<ext>`.
const sliceIndex = /\/slices-[\w-]+\.json$/;
const sliceChunk = /\/(?:core|c\d{3})-[\w-]+\.zst$/;
const sliceSkeleton = /\/skeleton-[\w-]+\.zst$/;
const wholeCjkFace = /\/Noto(?:Sans|Serif)(?:JP|KR|SC|TC)-Regular\.otf-[\w-]+\.zst$/;

async function documentBytes() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', rels(`<Relationship Id="doc" Type="${relationships}/officeDocument" Target="word/document.xml"/>`));
  const run = (family: string, lang: string, text: string) =>
    `<w:p><w:r><w:rPr><w:rFonts w:ascii="${family}" w:hAnsi="${family}" w:eastAsia="${family}" w:cs="${family}"/><w:sz w:val="40"/><w:lang w:eastAsia="${lang}"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
  const blank = '<w:p><w:r><w:rPr><w:sz w:val="40"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r></w:p>';
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${blank.repeat(6)}${run('MS Gothic', 'ja-JP', japanese)}${run('Microsoft JhengHei', 'zh-TW', traditional)}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

/** Dark pixels in the canvas's screenshot. */
async function ink(canvas: Locator): Promise<number> {
  const png = (await canvas.screenshot()).toString('base64');
  return pinnedEvaluate(canvas.page(), async (data) => {
    const image = await createImageBitmap(new Blob([Uint8Array.from(atob(data), (char) => char.charCodeAt(0))], { type: 'image/png' }));
    const context = new OffscreenCanvas(image.width, image.height).getContext('2d')!;
    context.drawImage(image, 0, 0);
    const pixels = context.getImageData(0, 0, image.width, image.height).data;
    let dark = 0;
    for (let index = 0; index < pixels.length; index += 4) {
      if (Math.min(pixels[index]!, pixels[index + 1]!, pixels[index + 2]!) < 128) dark += 1;
    }
    return dark;
  }, png);
}

async function open(page: Page) {
  const paths: string[] = [];
  const errors: string[] = [];
  const warnings: string[] = [];
  let lastRequest = Date.now();
  page.on('request', (request) => {
    lastRequest = Date.now();
    paths.push(new URL(request.url()).pathname);
  });
  page.on('pageerror', (error) => errors.push(error.message));
  page.on('console', (message) => {
    if (/font.*(fail|still loading)|fontRegistry.*falling back|outlines failed/i.test(message.text())) warnings.push(message.text());
  });
  await page.goto('/?format=docx');
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'progressive-cjk.docx', mimeType: 'application/octet-stream', buffer: await documentBytes(),
  });
  const canvas = page.locator('.canvas-page canvas').first();
  await expect(canvas).toBeVisible({ timeout: 45_000 });
  const mirror = page.locator('.canvas-page-mirror').first();
  await expect(mirror).toContainText(japanese, { timeout: 45_000 });
  await expect(mirror).toContainText(traditional);
  const idle = () => expect.poll(() => Date.now() - lastRequest, { timeout: 40_000 }).toBeGreaterThanOrEqual(3_000);
  await idle();
  return { canvas, mirror, paths, errors, warnings, idle };
}

const distinct = (paths: string[], pattern: RegExp) => new Set(paths.filter((path) => pattern.test(path)));

test('a CJK document opens with its sliced faces, fetches only the chunks it draws and paints them', async ({ page }) => {
  test.setTimeout(120_000);
  const { canvas, paths, errors, warnings } = await open(page);
  // Japanese and Traditional Chinese at least; whole faces never.
  expect(distinct(paths, sliceIndex).size).toBeGreaterThanOrEqual(2);
  expect(paths.filter((path) => wholeCjkFace.test(path))).toEqual([]);
  // The two lines draw 28 characters: a handful of each face's ~100 chunks.
  expect(distinct(paths, sliceChunk).size).toBeGreaterThan(0);
  expect(distinct(paths, sliceChunk).size).toBeLessThan(40);
  // The engine worker takes the page's decoded files: no skeleton or chunk is fetched twice.
  const slices = paths.filter((path) => sliceChunk.test(path) || sliceSkeleton.test(path));
  expect(slices.length).toBe(new Set(slices).size);
  // The engine's own decoder decodes the slices; the fonts package's stays unloaded.
  expect(paths.filter((path) => /\/fontpack(-[\w-]+)?\.wasm$/.test(path))).toEqual([]);
  await expect.poll(() => ink(canvas), { timeout: 30_000 }).toBeGreaterThan(0);
  expect(errors).toEqual([]);
  expect(warnings).toEqual([]);
});

test('typing a script no face on the page covers loads that face\'s slices and paints it', async ({ page }) => {
  test.setTimeout(120_000);
  const { canvas, mirror, paths, errors, warnings, idle } = await open(page);
  const indexes = distinct(paths, sliceIndex).size;
  const chunks = distinct(paths, sliceChunk).size;
  const before = await ink(canvas);
  await canvas.click({ position: { x: 200, y: 200 } });
  await page.keyboard.insertText(hangul);
  await expect(mirror).toContainText(hangul);
  await expect.poll(() => distinct(paths, sliceIndex).size, { timeout: 30_000 }).toBeGreaterThan(indexes);
  await idle();
  expect(distinct(paths, sliceChunk).size).toBeGreaterThan(chunks);
  expect(paths.filter((path) => wholeCjkFace.test(path))).toEqual([]);
  // Two 20 pt Hangul glyphs add about 300 dark pixels; the caret alone about 55.
  await expect.poll(() => ink(canvas), { timeout: 30_000 }).toBeGreaterThan(before + 150);
  expect(errors).toEqual([]);
  expect(warnings).toEqual([]);
});
