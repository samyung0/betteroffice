import { beforeAll, expect, test } from "bun:test";
import { createYrsSession, type YrsSession } from "../packages/docx/src/yrs";
import { rezipContainer } from "../packages/docx/src/wasm/opc";
import { exportOffice, seedOffice } from "./office-checkpoint";

const fixed = { seed: "0".repeat(64), now: "2026-09-29T00:00:00.000Z" };
const W =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const OFFICE = "application/vnd.openxmlformats-officedocument.wordprocessingml";
const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

function docx(body: string): Uint8Array {
  const parts: Record<string, string> = {
    "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="${OFFICE}.comments+xml"/></Types>`,
    "_rels/.rels":
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    "word/_rels/document.xml.rels":
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>',
    "word/document.xml": `<w:document ${W}><w:body><w:p w14:paraId="11111111">${body}</w:p><w:p w14:paraId="22222222"><w:r><w:t>tail</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
    "word/comments.xml": `<w:comments ${W}><w:comment w:id="1" w:author="Reviewer" w:date="2026-09-01T00:00:00Z"><w:p><w:r><w:t>earlier</w:t></w:r></w:p></w:comment></w:comments>`,
  };
  return rezipContainer(
    Object.fromEntries(
      Object.entries(parts).map(([path, xml]) => [
        path,
        new TextEncoder().encode(xml),
      ])
    )
  );
}

// Seeding loads the engines the sessions below use directly.
beforeAll(() => seedOffice("docx", docx(run("x"))));
let clientId = 9400;
async function open(bytes: Uint8Array): Promise<YrsSession> {
  const session = await createYrsSession({ clientId: (clientId += 1) });
  session.openDocx(bytes, true);
  return session;
}

async function publish(
  bytes: Uint8Array,
  session: YrsSession
): Promise<Uint8Array> {
  const state = session.encodeState();
  const baseSha256 = new Bun.CryptoHasher("sha256").update(bytes).digest("hex");
  return exportOffice(
    bytes,
    { format: "docx", schemaVersion: 1, baseSha256, state },
    fixed
  );
}

const note = [
  {
    type: "paragraph",
    content: [{ type: "run", content: [{ type: "text", text: "new" }] }],
  },
];
/** Anchors of the comment the test adds; comment 1 comes with the source. */
const range = (session: YrsSession) =>
  session.resolveComment(
    session.listComments().find((comment) => comment.id !== "1")!.id
  );

// "alpha" and one story unit lead " beta", which the new comment covers.
const alpha = run("alpha");
test.each([
  [
    "an earlier comment's reference mark",
    `<w:commentRangeStart w:id="1"/>${alpha}<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>`,
  ],
  ["a line break", `${alpha}<w:r><w:br/></w:r>`],
  ["a symbol", `${alpha}<w:r><w:sym w:font="Wingdings" w:char="F0E0"/></w:r>`],
  [
    "a field",
    `${alpha}<w:fldSimple w:instr=" DATE "><w:r><w:t>2026-09-29</w:t></w:r></w:fldSimple>`,
  ],
])(
  "a comment after %s keeps its range across publications",
  async (_, lead) => {
    let bytes = docx(`${lead}${run(" beta gamma")}`);
    let session = await open(bytes);
    session.addComment(
      [
        {
          story: "body",
          start: { paraId: "11111111", offset: 6 },
          end: { paraId: "11111111", offset: 11 },
        },
      ],
      "Reviewer",
      "2026-09-29T00:00:00Z",
      note
    );
    const expected = [{ story: "body", start: 6, end: 11 }];
    expect(range(session)).toEqual(expected);
    for (let publication = 0; publication < 2; publication += 1) {
      bytes = await publish(bytes, session);
      session.destroy();
      session = await open(bytes);
      expect(range(session)).toEqual(expected);
      session.insertText({ story: "body", paraId: "22222222", offset: 0 }, "x");
    }
    session.destroy();
  }
);
