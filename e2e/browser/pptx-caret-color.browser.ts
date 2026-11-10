import { expect, test, type Page } from 'playwright/test';
import JSZip from 'jszip';
import { inputReady, slideCanvas } from './pptx-async-helpers';

const DARK_HOST = 'rgb(237, 237, 237)';

async function deck(background: string | null, color: string) {
  const zip = new JSZip();
  const p = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const rels = (body: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`);
  zip.file('_rels/.rels', rels(`<Relationship Id="rId1" Type="${r}/officeDocument" Target="ppt/presentation.xml"/>`));
  zip.file('ppt/presentation.xml', `<p:presentation xmlns:p="${p}" xmlns:r="${r}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', rels(`<Relationship Id="rId1" Type="${r}/slide" Target="slides/slide1.xml"/>`));
  const bg = background ? `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${background}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>` : '';
  zip.file('ppt/slides/slide1.xml', `<p:sld xmlns:p="${p}" xmlns:a="${a}"><p:cSld>${bg}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text 2"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Caret text</a:t></a:r><a:endParaRPr sz="2400"/></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  return { name: 'caret.pptx', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', buffer: await zip.generateAsync({ type: 'nodebuffer' }) };
}

async function edit(page: Page, background: string | null, color: string, host: boolean) {
  await page.goto('/?format=pptx');
  if (host) await page.addStyleTag({ content: `body { color: ${DARK_HOST}; }` });
  await page.locator('input[type=file]').last().setInputFiles(await deck(background, color));
  await inputReady(page);
  const canvas = (await slideCanvas(page).boundingBox())!;
  // A first click selects the text box; a double click edits its text.
  await page.mouse.dblclick(canvas.x + canvas.width * 0.15, canvas.y + canvas.height * 0.22);
  const input = page.getByTestId('pptx-text-input');
  await expect(input).toBeFocused();
  await input.press('End');
}

/** Capy paints the caret on the selection canvas over the slide: the colour of its most opaque pixel. */
const caretColor = (page: Page) => page.locator('[data-testid="pptx-slide-canvas"] + canvas').evaluate((canvas: HTMLCanvasElement) => {
  const { data } = canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height);
  let best = 0;
  let color = '';
  for (let index = 0; index < data.length; index += 4) {
    if (data[index + 3]! <= best) continue;
    best = data[index + 3]!;
    color = `rgb(${data[index]}, ${data[index + 1]}, ${data[index + 2]})`;
  }
  return best >= 128 ? color : null;
});

for (const host of [false, true]) {
  const suffix = host ? ' inside a dark-themed host' : '';
  test(`pptx caret takes the text colour on a white slide${suffix}`, async ({ page }) => {
    await edit(page, null, 'C00000', host);
    await expect.poll(() => caretColor(page)).toBe('rgb(192, 0, 0)');
  });

  test(`pptx caret stays visible with white text on a dark slide${suffix}`, async ({ page }) => {
    await edit(page, '0B1F3A', 'FFFFFF', host);
    await expect.poll(() => caretColor(page)).toBe('rgb(255, 255, 255)');
  });
}
