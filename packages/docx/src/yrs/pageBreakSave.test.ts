import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { unzipContainer } from '../docx/wasm';
import { preloadEditWasm } from '../wasm/edit';
import { documentToYrs } from './documentToYrs';
import { createYrsSession, type YrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument';

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${OFFICE_DOC}.wordprocessingml.styles+xml"/></Types>`;

const PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPkg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`;

const SECTION = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;

function fixture(body: string): Uint8Array {
  const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}${SECTION}</w:body></w:document>`;
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes(CONTENT_TYPES));
  parts.set('_rels/.rels', toBytes(PACKAGE_RELS));
  parts.set('word/document.xml', toBytes(document));
  parts.set('word/styles.xml', toBytes(STYLES));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

/** The saved `w:body` up to its `w:sectPr`. */
function savedBody(buffer: ArrayBuffer): string {
  const xml = new TextDecoder().decode(
    unzipContainer(new Uint8Array(buffer))['word/document.xml'] as Uint8Array
  );
  const start = xml.indexOf('<w:body>') + '<w:body>'.length;
  const end = xml.indexOf('<w:sectPr');
  return xml.slice(start, end === -1 ? xml.indexOf('</w:body>') : end);
}

const PAGE_BREAK = '<w:r><w:br w:type="page"/></w:r>';
let clientId = 68100;

/** `[through the engine, straight through the serializer]`. */
async function save(
  body: string,
  edit?: (session: YrsSession) => void
): Promise<[string, string]> {
  const bytes = fixture(body);
  const parsed = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });
  const control = savedBody(await repackDocx(parsed));

  const session = await createYrsSession({ clientId: (clientId += 1) });
  try {
    session.seedFromDocx(bytes);
    edit?.(session);
    return [savedBody(await repackDocx(yrsToDocument(session, parsed))), control];
  } finally {
    session.destroy();
  }
}

describe('page break save projection', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  const shapes: Array<[string, string, number]> = [
    [
      'a break between two runs',
      `<w:p><w:r><w:t>HEAD</w:t></w:r>${PAGE_BREAK}<w:r><w:t>TARGET</w:t></w:r></w:p>`,
      1,
    ],
    [
      'a break closing a paragraph',
      `<w:p><w:r><w:t>HEAD</w:t></w:r>${PAGE_BREAK}</w:p><w:p><w:r><w:t>TAIL</w:t></w:r></w:p>`,
      1,
    ],
    ['a break opening a paragraph', `<w:p>${PAGE_BREAK}<w:r><w:t>TARGET</w:t></w:r></w:p>`, 1],
    [
      'two breaks in one paragraph',
      `<w:p><w:r><w:t>A</w:t></w:r>${PAGE_BREAK}<w:r><w:t>B</w:t></w:r>${PAGE_BREAK}<w:r><w:t>C</w:t></w:r></w:p>`,
      2,
    ],
    [
      'a break owning its paragraph',
      `<w:p><w:r><w:t>HEAD</w:t></w:r></w:p><w:p>${PAGE_BREAK}</w:p><w:p><w:r><w:t>TAIL</w:t></w:r></w:p>`,
      1,
    ],
    [
      'a column break',
      `<w:p><w:r><w:t>HEAD</w:t></w:r><w:r><w:br w:type="column"/></w:r><w:r><w:t>TARGET</w:t></w:r></w:p>`,
      0,
    ],
  ];

  for (const [name, body, pageBreaks] of shapes) {
    it(`saves ${name} as the serializer would`, async () => {
      const [through, control] = await save(body);
      expect(through).toBe(control);
      expect(through.match(/<w:br w:type="page"\/>/g) ?? []).toHaveLength(pageBreaks);
    });
  }

  it('saves a break through a paragraph whose formatting changed', async () => {
    const [through, control] = await save(
      `<w:p><w:r><w:t>HEAD</w:t></w:r></w:p><w:p>${PAGE_BREAK}<w:r><w:t>TARGET</w:t></w:r></w:p>`,
      (session) => {
        const last = session.paragraphs('body').at(-1);
        if (last) session.setParagraphAttr(last.paraId, 'widowControl', false);
      }
    );
    expect(through).toContain('<w:br w:type="page"/>');
    expect(through).toContain('<w:widowControl w:val="0"/>');
    expect(control).toContain('<w:br w:type="page"/>');
  });

  it('records the same breaks through the projector and the engine', async () => {
    const bytes = fixture(
      `<w:p><w:r><w:t>HEAD</w:t></w:r>${PAGE_BREAK}<w:r><w:t>TARGET</w:t></w:r></w:p>`
    );
    const parsed = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });
    // The byte seeder writes under the fixed seed client 0.
    const projected = await createYrsSession({ clientId: 0 });
    const engine = await createYrsSession({ clientId: 68001 });
    try {
      documentToYrs(projected, parsed);
      engine.seedFromDocx(bytes);
      expect(engine.storySegments('body')).toEqual(projected.storySegments('body'));
      expect(engine.encodeState()).toEqual(projected.encodeState());
      expect(engine.paragraphs('body')[0]?.properties._originalRunBoundaries).toEqual([
        { text: 'HEAD', marksKey: '' },
        { text: '', breaks: [{ offset: 0, type: 'page' }] },
        { text: 'TARGET', marksKey: '' },
      ]);
    } finally {
      projected.destroy();
      engine.destroy();
    }
  });

  it('seeds leading, tracked and cell breaks through the projector as the engine does', async () => {
    const tracked = `<w:ins w:id="9" w:author="A">${PAGE_BREAK}</w:ins>`;
    const paragraphs =
      `<w:p>${PAGE_BREAK}${tracked}<w:r><w:t>LEAD</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>END</w:t><w:br w:type="column"/></w:r></w:p><w:p>${PAGE_BREAK}</w:p>` +
      `<w:p><w:bookmarkStart w:id="5" w:name="b"/>${PAGE_BREAK}<w:r><w:t>BM</w:t></w:r>` +
      `<w:bookmarkEnd w:id="5"/></w:p><w:p>${PAGE_BREAK}<w:fldSimple w:instr=" SEQ x "/></w:p>`;
    const cell =
      `<w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>` +
      `<w:tr><w:tc>${paragraphs}</w:tc></w:tr></w:tbl>`;
    const bytes = fixture(`${paragraphs}${cell}<w:p/>`);
    const parsed = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });
    const projected = await createYrsSession({ clientId: 0 });
    const engine = await createYrsSession({ clientId: 68002 });
    try {
      documentToYrs(projected, parsed);
      engine.seedFromDocx(bytes);
      for (const story of ['body', 'body:t0:r0c0']) {
        expect(engine.storySegments(story)).toEqual(projected.storySegments(story));
        expect(
          engine
            .storySegments(story)
            .filter((segment) => segment.kind === 'embed')
            .map((segment) => [
              segment.embedKind,
              segment.payload.leading === true,
              !!segment.attributes.ins,
            ])
        ).toEqual([
          ['pageBreak', true, false],
          ['pageBreak', true, true],
          ['columnBreak', false, false],
          ['pageBreak', false, false],
          ['pageBreak', true, false],
          ['pageBreak', true, false],
          ['field', false, false],
          ...(story === 'body' ? [['table', false, false]] : []),
        ]);
      }
    } finally {
      projected.destroy();
      engine.destroy();
    }
  });

  describe('saves the break units the session holds', () => {
    const units = (session: YrsSession): string =>
      session
        .storySegments('body')
        .map((segment) =>
          segment.kind === 'text'
            ? segment.text
            : segment.kind === 'pilcrow'
              ? '¶'
              : `[${segment.embedKind}${segment.payload.leading ? '^' : ''}]`
        )
        .join('');
    /** The edited body's units, and the units of its save seeded again. */
    async function roundTrip(
      body: string,
      edit: (session: YrsSession) => void
    ): Promise<[string, string]> {
      const bytes = fixture(body);
      const parsed = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });
      const session = await createYrsSession({ clientId: (clientId += 1) });
      const reseeded = await createYrsSession({ clientId: (clientId += 1) });
      try {
        session.seedFromDocx(bytes);
        edit(session);
        reseeded.seedFromDocx(new Uint8Array(await repackDocx(yrsToDocument(session, parsed))));
        return [units(session), units(reseeded)];
      } finally {
        session.destroy();
        reseeded.destroy();
      }
    }
    const pageBreakParagraph = `<w:p><w:r><w:t>HEAD</w:t></w:r></w:p><w:p>${PAGE_BREAK}</w:p><w:p><w:r><w:t>TAIL</w:t></w:r></w:p>`;
    const paraId = (session: YrsSession, index: number) =>
      session.paragraphs('body')[index]!.paraId;
    const cases: Array<[string, string, (session: YrsSession) => void]> = [
      [
        'a break inserted between paragraphs',
        `<w:p><w:r><w:t>HEAD</w:t></w:r></w:p><w:p><w:r><w:t>TAIL</w:t></w:r></w:p>`,
        (session) =>
          session.insertPageBreak({ story: 'body', paraId: paraId(session, 1), offset: 0 }),
      ],
      [
        'a deleted break',
        pageBreakParagraph,
        (session) =>
          session.deleteRange({
            story: 'body',
            start: { paraId: paraId(session, 2), offset: 0 },
            end: { paraId: paraId(session, 2), offset: 1 },
          }),
      ],
      [
        'text typed into the paragraph a break owns',
        pageBreakParagraph,
        (session) =>
          session.insertText({ story: 'body', paraId: paraId(session, 1), offset: 0 }, 'typed'),
      ],
      [
        'a paragraph merged into the paragraph a break owns',
        pageBreakParagraph,
        (session) => session.mergeParagraphs('body', paraId(session, 0), 'forward'),
      ],
      [
        'an edit to a paragraph a break closes',
        `<w:p><w:r><w:t>HEAD</w:t></w:r>${PAGE_BREAK}<w:r><w:t>TARGET</w:t></w:r></w:p><w:p><w:r><w:t>TAIL</w:t></w:r></w:p>`,
        (session) =>
          session.insertText({ story: 'body', paraId: paraId(session, 0), offset: 1 }, 'x'),
      ],
      [
        'an edit to a paragraph a break opens',
        `<w:p><w:r><w:t>HEAD</w:t></w:r></w:p><w:p>${PAGE_BREAK}<w:r><w:t>TARGET</w:t></w:r></w:p>`,
        (session) =>
          session.insertText({ story: 'body', paraId: paraId(session, 1), offset: 6 }, ' typed'),
      ],
      [
        'an unedited column break opening a paragraph',
        `<w:p><w:r><w:t>HEAD</w:t></w:r></w:p><w:p><w:r><w:br w:type="column"/><w:t>TARGET</w:t></w:r></w:p>`,
        () => {},
      ],
    ];
    for (const [name, body, edit] of cases) {
      it(name, async () => {
        const [edited, saved] = await roundTrip(body, edit);
        expect(saved).toBe(edited);
      });
    }
  });

  it('does not open a paragraph the source never had', async () => {
    const [through, control] = await save(
      `<w:p><w:r><w:t>HEAD</w:t></w:r>${PAGE_BREAK}<w:r><w:t>TARGET</w:t></w:r></w:p>`
    );
    expect(through.match(/<w:p[ >]/g) ?? []).toHaveLength(
      (control.match(/<w:p[ >]/g) ?? []).length
    );
  });
});
