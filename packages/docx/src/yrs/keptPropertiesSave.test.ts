import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { unzipContainer } from '../wasm/opc';
import type { Document } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';

// What Word reads that the editor's model does not hold has to come back on
// save as the source wrote it. The oracle is the source XML.

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

const STYLES =
  `<w:styles xmlns:w="${W}"><w:docDefaults><w:pPrDefault><w:pPr>` +
  '<w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>';

function fixture(body: string): Uint8Array<ArrayBuffer> {
  const parts = new Map<string, Uint8Array>();
  parts.set(
    '[Content_Types].xml',
    toBytes(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        `<Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/>` +
        `<Override PartName="/word/styles.xml" ContentType="${OFFICE}.styles+xml"/></Types>`
    )
  );
  parts.set(
    '_rels/.rels',
    toBytes(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    )
  );
  parts.set(
    'word/_rels/document.xml.rels',
    toBytes(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/></Relationships>`
    )
  );
  parts.set(
    'word/document.xml',
    toBytes(`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr/></w:body></w:document>`)
  );
  parts.set('word/styles.xml', toBytes(STYLES));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

/** As Capy opens a file: the native seed, with the base materialized from the source. */
async function open(bytes: Uint8Array, clientId = 7) {
  const session = await createYrsSession({ clientId });
  session.openDocx(bytes, true);
  const base = session.materializeDocx()!;
  return { session, base };
}

async function savedDocumentXml(session: YrsSession, base: Document): Promise<string> {
  const bytes = new Uint8Array(await repackDocx(yrsToDocument(session, base)));
  const files = unzipContainer(bytes) as Record<string, Uint8Array>;
  return new TextDecoder().decode(files['word/document.xml']);
}

/** Each saved `w:p`'s pPr, in document order ('' when it has none). */
function pPrs(xml: string): string[] {
  return [...xml.matchAll(/<w:p(?: [^>]*)?>(.*?)<\/w:p>|<w:p(?: [^>]*)?\/>/g)].map(
    (match) => match[1]?.match(/^<w:pPr>.*?<\/w:pPr>/)?.[0] ?? ''
  );
}

function at(session: YrsSession, story: string, index: number, offset = 0) {
  return { story, paraId: session.paragraphs(story)[index]!.paraId, offset };
}

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))
  )
);

describe('pPr children the model has no field for', () => {
  // Every CT_PPrBase child the model lacks, in schema order between modeled ones.
  const KEPT =
    '<w:pPr><w:suppressAutoHyphens/><w:kinsoku w:val="0"/><w:wordWrap w:val="0"/>' +
    '<w:overflowPunct w:val="0"/><w:topLinePunct/><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/>' +
    '<w:adjustRightInd w:val="0"/><w:snapToGrid w:val="0"/><w:spacing w:after="0"/><w:ind w:left="100"/>' +
    '<w:mirrorIndents/><w:suppressOverlap/><w:jc w:val="both"/><w:textDirection w:val="tbRl"/>' +
    '<w:textAlignment w:val="center"/><w:textboxTightWrap w:val="allLines"/><w:outlineLvl w:val="8"/>' +
    '<w:divId w:val="123"/><w:cnfStyle w:val="100000000000"/></w:pPr>';
  const FRAME =
    '<w:pPr><w:framePr w:hAnchor="text" w:vAnchor="text" w:wrap="around" w:dropCap="drop" w:lines="3" ' +
    'w:hSpace="72"/></w:pPr>';
  const CELL = '<w:pPr><w:kinsoku w:val="0"/><w:cnfStyle w:val="101000000000"/></w:pPr>';
  const bytes = fixture(
    `<w:p>${KEPT}<w:r><w:t>Hello world</w:t></w:r></w:p>` +
      `<w:p>${FRAME}<w:r><w:t>D</w:t></w:r></w:p>` +
      '<w:p><w:r><w:t>plain</w:t></w:r></w:p>' +
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>' +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p>${CELL}<w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
      '<w:p/>'
  );

  it('come back in schema order on an untouched paragraph, a framed one and a cell', async () => {
    const { session, base } = await open(bytes);
    session.insertText(at(session, 'body', 2), 'X');
    const cell = session.storyIds().find((id) => id.startsWith('body:t'))!;
    session.insertText(at(session, cell, 0), 'z');
    expect(pPrs(await savedDocumentXml(session, base))).toEqual([KEPT, FRAME, '', CELL, '']);
    session.destroy();
  });

  it('stay on a paragraph the editor changed and on both halves of a split', async () => {
    const { session, base } = await open(bytes);
    const { paraId } = session.paragraphs('body')[0]!;
    session.setParagraphAttrs(
      { story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 0 } },
      { alignment: 'center' }
    );
    session.insertText(at(session, 'body', 0, 5), 'X');
    session.splitParagraph(at(session, 'body', 0, 3));
    const centred = KEPT.replace('<w:jc w:val="both"/>', '<w:jc w:val="center"/>');
    expect(pPrs(await savedDocumentXml(session, base)).slice(0, 3)).toEqual([centred, centred, FRAME]);
    session.destroy();
  });
});
