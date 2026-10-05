import { test, expect } from "bun:test";
import { exportOffice, seedOffice } from "./office-checkpoint";
import { createYrsSession } from "../packages/docx/src/yrs";
import { rezipContainer, unzipContainer } from "../packages/docx/src/wasm/opc";

const fixed = { seed: "0".repeat(64), now: "2026-10-05T00:00:00.000Z" };
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const OFFICE = "application/vnd.openxmlformats-officedocument.wordprocessingml";
const encode = (xml: string) => new TextEncoder().encode(xml);

function docx(body: string): Uint8Array {
  return rezipContainer({
    "[Content_Types].xml": encode(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${OFFICE}.styles+xml"/></Types>`
    ),
    "_rels/.rels": encode(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    "word/_rels/document.xml.rels": encode(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/></Relationships>`
    ),
    "word/document.xml": encode(
      `<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr/></w:body></w:document>`
    ),
    "word/styles.xml": encode(
      `<w:styles xmlns:w="${W}"><w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style></w:styles>`
    ),
  });
}

const documentXml = (bytes: Uint8Array) =>
  new TextDecoder().decode(unzipContainer(bytes)["word/document.xml"]);

test("DOCX paragraph edits and run font hints survive two publications", async () => {
  const bytes = docx(
    '<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>' +
      '<w:p><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>“中文”</w:t></w:r></w:p>'
  );
  const seed = await seedOffice("docx", bytes);
  const session = await createYrsSession({ clientId: 9301 });
  let state: Uint8Array;
  try {
    session.openDocx(bytes, false);
    session.loadState(seed.state);
    const { paraId } = session.paragraphs("body")[0]!;
    session.setParagraphAttrs(
      { story: "body", start: { paraId, offset: 0 }, end: { paraId, offset: 0 } },
      {
        indentLeft: 1234,
        lineSpacing: 480,
        lineSpacingRule: "auto",
        tabs: [{ position: 3000, alignment: "left" }],
      }
    );
    state = session.encodeState();
  } finally {
    session.destroy();
  }
  const published = await exportOffice(bytes, { ...seed, state }, fixed);
  const republished = await exportOffice(
    published,
    await seedOffice("docx", published),
    fixed
  );
  for (const saved of [published, republished]) {
    const xml = documentXml(saved);
    expect(xml).toContain(
      '<w:pPr><w:pStyle w:val="Title"/><w:tabs><w:tab w:val="left" w:pos="3000"/></w:tabs><w:spacing w:line="480" w:lineRule="auto"/><w:ind w:left="1234"/></w:pPr>'
    );
    expect(xml).toContain('<w:rFonts w:hint="eastAsia"/>');
  }
});
