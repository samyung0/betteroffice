import { expect, test, type Page } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { inputReady, slideCanvas } from './pptx-async-helpers';

// Capy's slide text input; it has no "Text selected" status, so the copied text shows what is selected.
const input = (page: Page) => page.getByTestId('pptx-text-input');
const clipboard = (page: Page) => page.evaluate(() => navigator.clipboard.readText());

async function fixture() {
  const zip = await JSZip.loadAsync(await readFile(new URL('../../crates/pptx-render/tests/fixtures/table-basic.pptx', import.meta.url)));
  const path = 'ppt/slides/slide1.xml';
  const box = `<p:sp><p:nvSpPr><p:cNvPr id="10" name="Box"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="2743200"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"/><a:t>Whole box</a:t></a:r></a:p></p:txBody></p:sp>`;
  zip.file(path, (await zip.file(path)!.async('string')).replace(/ (gridSpan|rowSpan|hMerge|vMerge)="[^"]*"/gu, '')
    .replace('</p:spTree>', `${box}</p:spTree>`));
  return { name: 'select-all.pptx', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: await zip.generateAsync({ type: 'nodebuffer' }) };
}

test.beforeEach(async ({ page, context }) => {
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  await page.goto('/?format=pptx');
  await page.locator('input[type=file]').last().setInputFiles(await fixture());
  await inputReady(page);
});

// Upstream's table-cell case is not kept: Capy's editor enters a cell only through its API and
// leaves Tab in a cell to the browser, so no pointer or keyboard path reaches a cell's text here.
async function click(page: Page, x: number, y: number, clickCount = 1) {
  const canvas = (await slideCanvas(page).boundingBox())!;
  await page.mouse.click(canvas.x + canvas.width * x, canvas.y + canvas.height * y, { clickCount });
}

// In Capy's editor, as in PowerPoint, a first click selects the shape; a double click edits its text.
async function caretInBox(page: Page) {
  await click(page, 0.15, 0.42, 2);
  await expect(input(page)).toBeFocused();
}

test('pptx select all: Ctrl or Meta A selects the whole text box', async ({ page }) => {
  await caretInBox(page);
  await input(page).press('ControlOrMeta+a');
  await input(page).press('ControlOrMeta+c');
  await expect.poll(() => clipboard(page)).toBe('Whole box');
});

test('pptx select all: typing replaces the selected text box text', async ({ page }) => {
  await caretInBox(page);
  await input(page).press('ControlOrMeta+a');
  await page.keyboard.type('Z');
  await input(page).press('ControlOrMeta+a');
  await input(page).press('ControlOrMeta+c');
  await expect.poll(() => clipboard(page)).toBe('Z');
});

test('pptx select all: text typed just before is included across paragraphs', async ({ page }) => {
  await caretInBox(page);
  await input(page).press('ControlOrMeta+End');
  await page.keyboard.press('Enter');
  await page.keyboard.type('Next');
  await input(page).press('ControlOrMeta+a');
  await input(page).press('ControlOrMeta+c');
  await expect.poll(async () => (await clipboard(page)).replace(/\r\n?/gu, '\n')).toBe('Whole box\nNext');
});

test('pptx select all: nothing selected keeps the page text unselected', async ({ page }) => {
  await click(page, 0.5, 0.95);
  await page.keyboard.press('ControlOrMeta+a');
  expect(await page.evaluate(() => getSelection()?.toString() ?? '')).toBe('');
});
