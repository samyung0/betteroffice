import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { buildResidentRegionLayoutRequest } from '../editor/computeLayout';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { unzipContainer } from '../docx/wasm';
import type { LayoutBlock, Run as LayoutRun, ShapeBlock } from '../layout/pagination/types';
import type { DisplayList, GlyphRunPrimitive } from '../layout/render/displayList';
import type { Document } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { documentToYrs } from './documentToYrs';
import { createYrsSession, type YrsSession } from './index';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const FONTS = '<w:rFonts w:ascii="Cambria" w:hAnsi="Cambria" w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>';

function fontPackage(latin: string | undefined, fonts: string, prefix = '', body?: string): Uint8Array<ArrayBuffer> {
  const parts = new Map([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      (latin !== undefined ? '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' : '') + '</Types>'],
    ['_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
    ['word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      (latin !== undefined ? `<Relationship Id="theme" Type="${R}/theme" Target="theme/theme1.xml"/>` : '') + '</Relationships>'],
    ['word/document.xml', `<w:document xmlns:w="${W}"><w:body>${body ?? `<w:p>${prefix}<w:r><w:rPr>${fonts}</w:rPr><w:t>Title</w:t></w:r></w:p>`}<w:sectPr/></w:body></w:document>`],
  ]);
  if (latin !== undefined) parts.set('word/theme/theme1.xml', `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="Synthetic"><a:majorFont><a:latin typeface="${latin}"/></a:majorFont></a:fontScheme></a:themeElements></a:theme>`);
  return new Uint8Array(rezipPartsToArrayBuffer(new Map([...parts].map(([name, xml]) => [name, toBytes(xml)]))));
}

function textBox(kind: 'drawingml' | 'vml', fonts: string): string {
  const content = `<w:p><w:r><w:rPr>${fonts}</w:rPr><w:t>Box</w:t></w:r></w:p>`;
  if (kind === 'vml') {
    return `<w:r><w:pict xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word"><v:shape id="Text Box 1" o:spid="_x0000_s1026" style="position:absolute;margin-left:36pt;margin-top:72pt;width:144pt;height:72pt;z-index:1;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;v-text-anchor:top" filled="f" stroked="f"><v:textbox inset="0,0,0,0"><w:txbxContent>${content}</w:txbxContent></v:textbox><w10:wrap type="none"/></v:shape></w:pict></w:r>`;
  }
  return `<w:r><w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="457200"/><wp:docPr id="41" name="Text Box 41"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent>${content}</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz"/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}

function mixedSlotsPackage(fonts: string): Uint8Array<ArrayBuffer> {
  const body = `<w:p>${textBox('drawingml', fonts)}</w:p><w:p>${textBox('vml', fonts)}</w:p>` +
    '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
    `<w:r><w:rPr>${fonts}</w:rPr><w:t>Title</w:t></w:r></w:p>`;
  const parts = new Map(Object.entries(unzipContainer(fontPackage('Calibri', fonts, '', body))));
  const types = new TextDecoder().decode(parts.get('[Content_Types].xml'));
  parts.set('[Content_Types].xml', toBytes(types.replace('</Types>', '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>')));
  const rels = new TextDecoder().decode(parts.get('word/_rels/document.xml.rels'));
  parts.set('word/_rels/document.xml.rels', toBytes(rels.replace('</Relationships>', `<Relationship Id="numbering" Type="${R}/numbering" Target="numbering.xml"/></Relationships>`)));
  parts.set('word/numbering.xml', toBytes(`<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:rPr>${fonts}</w:rPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`));
  parts.set('word/theme/theme1.xml', toBytes('<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="Synthetic"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>'));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

beforeAll(() => preloadEditWasm(new Uint8Array(
  readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))
)));

function layoutRuns(blocks: readonly LayoutBlock[]): LayoutRun[] {
  return blocks.flatMap(block => {
    const runs = block.kind === 'paragraph' ? block.runs
      : block.kind === 'shape' ? (block.innerText ?? []).flatMap(paragraph => paragraph.runs)
        : block.kind === 'textBox' ? block.content.flatMap(paragraph => paragraph.runs) : [];
    return runs.flatMap(run => run.kind === 'image' && run.inlineShape
      ? layoutRuns([{ ...(run.inlineShape as ShapeBlock), kind: 'shape' }]) : [run]);
  });
}

function assertThemeFontPaint(session: YrsSession, document: Document, texts: string[]): void {
  const themeFont = session.registerFont(new Uint8Array(readFileSync(
    resolve(import.meta.dir, '../../../fonts/assets/Carlito-Regular.ttf')
  )));
  const explicitFont = session.registerFont(new Uint8Array(readFileSync(
    resolve(import.meta.dir, '../../../fonts/assets/Caladea-Regular.ttf')
  )));
  expect(themeFont).not.toBe(explicitFont);
  const request = buildResidentRegionLayoutRequest(document, 24, {});
  request.measurement = {
    fontChains: { 'calibri|0|0': [themeFont], 'cambria|0|0': [explicitFont] },
    defaults: { fontSize: 11, fontFamily: 'Calibri' },
    compat: { noLeading: false, doNotExpandShiftReturn: false },
    authoritativeShaping: true,
  };
  const requirements = JSON.parse(session.layoutFontRequirementsJson(JSON.stringify(request))) as Array<{ family: string }>;
  expect(requirements.some(font => font.family === 'Calibri')).toBe(true);
  expect(requirements.some(font => font.family === 'Cambria')).toBe(false);
  const layout = JSON.parse(session.layoutDocumentWithRegionsJson(JSON.stringify(request)));
  const display = JSON.parse(session.buildDisplayListJson(JSON.stringify({
    ...layout, fontChains: request.measurement.fontChains,
  }))) as DisplayList;
  const painted = display.pages.flatMap(page => page.primitives).filter(
    primitive => (primitive.kind === 'glyphRun' || primitive.kind === 'text') && primitive.text.length > 0
  );
  expect(painted.map(primitive => primitive.kind)).toEqual(painted.map(() => 'glyphRun'));
  const glyphs = painted.filter((primitive): primitive is GlyphRunPrimitive => primitive.kind === 'glyphRun');
  const paintedText = glyphs.map(run => run.text).join('');
  for (const text of texts) expect(paintedText).toContain(text);
  expect(paintedText.length).toBe(texts.join('').length);
  expect(glyphs.map(run => run.fontId)).toEqual(glyphs.map(() => themeFont));
}

for (const path of ['native', 'typescript'] as const) {
  for (const [slot, fonts, ascii, hAnsi] of [
    ['hAnsi', '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Arial" w:hAnsiTheme="minorHAnsi"/>', 'Times New Roman', 'Calibri'],
    ['ascii', '<w:rFonts w:ascii="Times New Roman" w:asciiTheme="minorHAnsi" w:hAnsi="Arial"/>', 'Calibri', 'Arial'],
  ] as const) {
    test(`${path} mixed ${slot} theme fonts stay in their own run and numbering slots`, async () => {
      const bytes = mixedSlotsPackage(fonts);
      const document = await parseDocx(bytes.buffer, { preloadFonts: false });
      const source = await createYrsSession();
      const session = await createYrsSession();
      try {
        if (path === 'native') source.seedFromDocx(bytes);
        else documentToYrs(source, document);
        session.openDocx(bytes, false);
        session.loadState(source.encodeState());
        const original = session.encodeState();
        const request = buildResidentRegionLayoutRequest(document, 24, {});
        const blocks = session.yrsBlocksForStory('body', request.renderEnv) as LayoutBlock[];
        const runs = layoutRuns(blocks);
        const text = runs.filter(run => run.kind === 'text' && ['Title', 'Box'].includes(run.text));
        expect(text).toHaveLength(3);
        for (const run of text) expect(run).toMatchObject({
          fontFamily: ascii, fontSlots: { ascii, hAnsi },
        });
        const paragraph = blocks.find(block => block.kind === 'paragraph' &&
          block.runs.some(run => run.kind === 'text' && run.text === 'Title'));
        expect(paragraph).toMatchObject({ attrs: { listMarker: '1.', listMarkerFontFamily: ascii } });
        expect(session.storySegments('body').find(segment => segment.kind === 'text')?.attributes.fontFamily)
          .toMatchObject({ ascii: 'Times New Roman', hAnsi: 'Arial' });
        expect(session.encodeState()).toEqual(original);
      } finally {
        source.destroy();
        session.destroy();
      }
    });
  }

  for (const [authored, fonts, seeded] of [
    ['theme-only', '<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>', 'Times New Roman'],
    ['paired', FONTS, 'Cambria'],
  ] as const) {
    for (const context of ['unknown theme', 'document theme'] as const) {
      test(`${path} source-less ${authored} fonts resolve with ${context} without changing the model`, async () => {
        const bytes = fontPackage('Times New Roman', fonts, textBox('drawingml', fonts) + textBox('vml', fonts));
        const document = await parseDocx(bytes.buffer, { preloadFonts: false });
        const source = await createYrsSession();
        const session = await createYrsSession();
        try {
          if (path === 'native') source.seedFromDocx(bytes);
          else documentToYrs(source, document);
          session.loadState(source.encodeState());
          const env = context === 'document theme'
            ? buildResidentRegionLayoutRequest(document, 24, {}).renderEnv
            : undefined;
          for (const candidate of path === 'typescript' ? [source, session] : [session]) {
            const original = candidate.encodeState();
            const blocks = candidate.yrsBlocksForStory('body', env) as LayoutBlock[];
            const runs = layoutRuns(blocks);
            const text = runs.filter(run => run.kind === 'text' && ['Title', 'Box'].includes(run.text));
            expect(text).toHaveLength(3);
            expect(text.filter(run => run.kind === 'text' && run.text === 'Title')).toHaveLength(1);
            expect(text.filter(run => run.kind === 'text' && run.text === 'Box')).toHaveLength(2);
            const effective = context === 'document theme' ? 'Times New Roman' : seeded;
            for (const run of text) expect(run).toMatchObject({
              fontFamily: effective, fontSlots: { ascii: effective, hAnsi: effective },
            });
            expect(candidate.storySegments('body').find(segment => segment.kind === 'text')?.attributes.fontFamily)
              .toMatchObject({ ascii: seeded, hAnsi: seeded });
            expect(candidate.encodeState()).toEqual(original);
          }
        } finally {
          source.destroy();
          session.destroy();
        }
      });
    }
  }

  for (const [name, latin, fonts, effective] of [
    ['paired', 'Calibri', FONTS, 'Calibri'],
    ['theme-only', 'Calibri', '<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>', 'Calibri'],
    ['explicit-only', 'Calibri', '<w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/>', 'Cambria'],
    ['default major', undefined, FONTS, 'Calibri Light'],
    ['default minor', undefined, FONTS.replaceAll('majorHAnsi', 'minorHAnsi'), 'Calibri'],
    ['empty theme slot', '', FONTS, 'Cambria'],
    ['unknown reference', 'Calibri', FONTS.replaceAll('majorHAnsi', 'unknown'), 'Cambria'],
  ] as const) {
    test(`${path} ${name} fonts resolve only in layout`, async () => {
      const bytes = fontPackage(latin, fonts);
      const document = await parseDocx(bytes.buffer, { preloadFonts: false });
      const session = await createYrsSession();
      try {
        if (path === 'native') session.seedFromDocx(bytes);
        else documentToYrs(session, document);
        const original = session.encodeState();
        const segments = session.storySegments('body');
        const authored = name === 'theme-only' ? 'Calibri' : 'Cambria';
        expect(segments.find(segment => segment.kind === 'text')?.attributes.fontFamily).toMatchObject({
          ascii: authored, hAnsi: authored,
        });
        const request = buildResidentRegionLayoutRequest(document, 24, {});
        const blocks = session.yrsBlocksForStory('body', request.renderEnv) as LayoutBlock[];
        const runs = blocks.flatMap(block => block.kind === 'paragraph' ? block.runs : []);
        expect(runs).toHaveLength(1);
        expect(runs[0]).toMatchObject({
          kind: 'text', text: 'Title', fontFamily: effective,
          fontSlots: { ascii: effective, hAnsi: effective },
        });
        if (name === 'paired' || name === 'theme-only') assertThemeFontPaint(session, document, ['Title']);
        expect(session.encodeState()).toEqual(original);
      } finally {
        session.destroy();
      }
    });
  }

  for (const kind of ['drawingml', 'vml'] as const) {
    for (const [authored, fonts] of [
      ['paired', FONTS],
      ['theme-only', '<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>'],
    ] as const) {
      test(`${path} ${kind} ${authored} text box fonts resolve only in layout`, async () => {
        const bytes = fontPackage('Calibri', fonts, textBox(kind, fonts));
        const document = await parseDocx(bytes.buffer, { preloadFonts: false });
        const session = await createYrsSession();
        try {
          if (path === 'native') session.seedFromDocx(bytes);
          else documentToYrs(session, document);
          const original = session.encodeState();
          const request = buildResidentRegionLayoutRequest(document, 24, {});
          const blocks = session.yrsBlocksForStory('body', request.renderEnv) as LayoutBlock[];
          const runs = layoutRuns(blocks);
          const labels = runs.filter(run => run.kind === 'text' && run.text === 'Box');
          expect(labels).toHaveLength(1);
          expect(labels[0]).toMatchObject({
            fontFamily: 'Calibri', fontSlots: { ascii: 'Calibri', hAnsi: 'Calibri' },
          });
          assertThemeFontPaint(session, document, ['Box', 'Title']);
          expect(session.encodeState()).toEqual(original);
        } finally {
          session.destroy();
        }
      });
    }
  }
}
