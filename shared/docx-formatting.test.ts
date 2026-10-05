import { test, expect } from "bun:test";
import { createHash } from "node:crypto";
import { exportOffice, seedOffice } from "./office-checkpoint";
import { parseDocx } from "../packages/docx/src/docx";
import { createYrsSession, documentToYrs, type YrsSession } from "../packages/docx/src/yrs";
import { rezipContainer, unzipContainer } from "../packages/docx/src/wasm/opc";

const fixed = { seed: "0".repeat(64), now: "2026-10-05T00:00:00.000Z" };
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const OFFICE = "application/vnd.openxmlformats-officedocument.wordprocessingml";
const encode = (xml: string) => new TextEncoder().encode(xml);

const NUMBERING = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;

function docx(body: string): Uint8Array {
  return rezipContainer({
    "[Content_Types].xml": encode(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${OFFICE}.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="${OFFICE}.numbering+xml"/></Types>`
    ),
    "_rels/.rels": encode(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    "word/_rels/document.xml.rels": encode(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/><Relationship Id="numbering" Type="${R}/numbering" Target="numbering.xml"/></Relationships>`
    ),
    "word/numbering.xml": encode(NUMBERING),
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

// A first line set on a numbered paragraph, zero included, wins over its list
// level's (2026-10-06); a changed hash means the rule moved and DOCX seeds with it.
test("a zero first line on a numbered paragraph seeds as itself in both seeders", async () => {
  const numbered = (ind: string) =>
    `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>${ind}</w:pPr><w:r><w:t>Item</w:t></w:r></w:p>`;
  const bytes = docx(
    [
      "",
      '<w:ind w:firstLine="0"/>',
      '<w:ind w:hanging="0"/>',
      '<w:ind w:firstLine="240"/>',
      '<w:ind w:left="1440" w:firstLine="0"/>',
    ]
      .map(numbered)
      .join("")
  );
  const seed = await seedOffice("docx", bytes);
  expect(createHash("sha256").update(seed.state).digest("hex")).toBe(
    "69c97e8f51eeec42126dd0e481fda525c904e9d53d61a3543f6b4d3c3d26003f"
  );
  const indents = (session: YrsSession) =>
    session.paragraphs("body").map(({ properties }) => [
      properties.indentLeft,
      properties.indentFirstLine,
      properties.hangingIndent ?? false,
    ]);
  const expected = [
    [720, -360, true],
    [720, 0, false],
    [720, 0, true],
    [720, 240, false],
    [1440, 0, false],
  ];
  const native = await createYrsSession({ clientId: 9302 });
  const projected = await createYrsSession({ clientId: 9303 });
  try {
    native.openDocx(bytes, false);
    native.loadState(seed.state);
    documentToYrs(projected, await parseDocx(bytes.slice().buffer, { preloadFonts: false }));
    expect(indents(native)).toEqual(expected);
    expect(indents(projected)).toEqual(expected);
  } finally {
    native.destroy();
    projected.destroy();
  }
});

// pPr children the model has no field for ride the seed's source formatting
// (2026-10-06); a changed hash means that seed moved.
test("unmodeled pPr children seed into the source formatting in both seeders", async () => {
  const KEPT =
    '<w:pPr><w:kinsoku w:val="0"/><w:wordWrap w:val="0"/><w:overflowPunct w:val="0"/><w:autoSpaceDE w:val="0"/>' +
    '<w:adjustRightInd w:val="0"/><w:textAlignment w:val="baseline"/></w:pPr>';
  const CELL = '<w:pPr><w:cnfStyle w:firstRow="1" w:val="100000000000"/></w:pPr>';
  const bytes = docx(
    `<w:p>${KEPT}<w:r><w:t>中文段落</w:t></w:r></w:p>` +
      '<w:p><w:pPr><w:framePr w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/></w:pPr><w:r><w:t>D</w:t></w:r></w:p>' +
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>' +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p>${CELL}<w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>`
  );
  const seed = await seedOffice("docx", bytes);
  expect(createHash("sha256").update(seed.state).digest("hex")).toBe(
    "bb7c8b97ecdac25570f815056924f7114d10a90d75bfcbb74c31cd7505677781"
  );
  const kept = (session: YrsSession) =>
    session.storyIds().flatMap((story) =>
      session.paragraphs(story).map(({ properties }) => {
        const original = properties._originalFormatting as Record<string, unknown> | null | undefined;
        return [original?.extraChildren ?? null, (original?.frame as { dropCap?: string } | undefined)?.dropCap ?? null];
      })
    );
  const native = await createYrsSession({ clientId: 9304 });
  const projected = await createYrsSession({ clientId: 9305 });
  try {
    native.openDocx(bytes, false);
    native.loadState(seed.state);
    documentToYrs(projected, await parseDocx(bytes.slice().buffer, { preloadFonts: false }));
    expect(kept(projected)).toEqual(kept(native));
    expect(kept(native)).toContainEqual([
      { kinsoku: { val: "0" }, wordWrap: { val: "0" }, overflowPunct: { val: "0" }, adjustRightInd: { val: "0" }, textAlignment: { val: "baseline" } },
      null,
    ]);
  } finally {
    native.destroy();
    projected.destroy();
  }
  const xml = documentXml(await exportOffice(bytes, seed, fixed));
  expect(xml).toContain(KEPT);
  expect(xml).toContain(CELL);
});
