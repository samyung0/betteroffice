import { expect, test } from "bun:test";
import { exportOffice, seedOffice } from "./office-checkpoint";
import { createYrsSession, type YrsSession } from "../packages/docx/src/yrs";
import { rezipContainer, unzipContainer } from "../packages/docx/src/wasm/opc";

// Edits published the way the collaboration service publishes them (Capy's
// path, through TS yrsToDocument): the base package, its seed and the editor's
// state through exportOffice.

const fixed = { seed: "0".repeat(64), now: "2026-10-13T00:00:00.000Z" };
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const OFFICE = "application/vnd.openxmlformats-officedocument";
const NS = `xmlns:w="${W}" xmlns:r="${R}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"`;
const encode = (xml: string) => new TextEncoder().encode(xml);
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

interface Part {
  path: string;
  xml: string;
  contentType: string;
  /** Relationship from document.xml; omitted for parts reached otherwise. */
  rel?: { id: string; type: string };
}

function docx(body: string, parts: Part[] = [], sectPr = ""): Uint8Array {
  const overrides = parts
    .map((part) => `<Override PartName="/${part.path}" ContentType="${part.contentType}"/>`)
    .join("");
  const rels = parts
    .filter((part) => part.rel)
    .map(
      (part) =>
        `<Relationship Id="${part.rel!.id}" Type="${part.rel!.type}" Target="${part.path.slice("word/".length)}"/>`
    )
    .join("");
  return rezipContainer({
    "[Content_Types].xml": encode(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.wordprocessingml.document.main+xml"/>${overrides}</Types>`
    ),
    "_rels/.rels": encode(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    "word/_rels/document.xml.rels": encode(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`
    ),
    "word/document.xml": encode(
      `<w:document ${NS}><w:body>${body}<w:sectPr>${sectPr}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`
    ),
    ...Object.fromEntries(parts.map((part) => [part.path, encode(part.xml)])),
  });
}

/** Opens `bytes` as the editor does, runs `edit`, and publishes twice. */
async function publishTwice(
  bytes: Uint8Array,
  edit: (session: YrsSession) => void
): Promise<Uint8Array[]> {
  const seed = await seedOffice("docx", bytes);
  const session = await createYrsSession({ clientId: 9901 });
  let state: Uint8Array;
  try {
    session.openDocx(bytes, false);
    session.loadState(seed.state);
    edit(session);
    state = session.encodeState();
  } finally {
    session.destroy();
  }
  const published = await exportOffice(bytes, { ...seed, state }, fixed);
  const republished = await exportOffice(published, await seedOffice("docx", published), fixed);
  return [published, republished];
}

const text = (bytes: Uint8Array, path: string) => {
  const part = unzipContainer(bytes)[path];
  return part ? new TextDecoder().decode(part) : undefined;
};

/** Types `typed` at `offset` of the story's first paragraph. */
const typeAt = (story: string, offset: number, typed: string) => (session: YrsSession) => {
  const [{ paraId }] = session.paragraphs(story);
  session.insertText({ story, paraId, offset }, typed);
};

test("text typed into a header and a footer is saved into their parts", async () => {
  const bytes = docx(
    `<w:p w14:paraId="10000001">${run("Body")}</w:p>`,
    [
      {
        path: "word/header1.xml",
        xml: `<w:hdr ${NS}><w:p w14:paraId="20000001">${run("Unit 3")}</w:p></w:hdr>`,
        contentType: `${OFFICE}.wordprocessingml.header+xml`,
        rel: { id: "rIdH", type: `${R}/header` },
      },
      {
        path: "word/footer1.xml",
        xml: `<w:ftr ${NS}><w:p w14:paraId="30000001">${run("Draft")}</w:p></w:ftr>`,
        contentType: `${OFFICE}.wordprocessingml.footer+xml`,
        rel: { id: "rIdF", type: `${R}/footer` },
      },
    ],
    '<w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/>'
  );
  const saves = await publishTwice(bytes, (session) => {
    typeAt("hf:rIdH", 6, ": Forces")(session);
    typeAt("hf:rIdF", 0, "Second ")(session);
  });
  for (const saved of saves) {
    expect(text(saved, "word/header1.xml")).toContain("Unit 3: Forces");
    expect(text(saved, "word/footer1.xml")).toContain("Second Draft");
    expect(text(saved, "word/document.xml")).toContain(
      '<w:headerReference w:type="default" r:id="rIdH"/>'
    );
  }
});

const DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const FRAME = `<w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="3657600" cy="914400"/><wp:docPr id="1" name="Diagram 1"/><a:graphic xmlns:a="${A}"><a:graphicData uri="${DGM}"><dgm:relIds xmlns:dgm="${DGM}" r:dm="rIdDm" r:lo="rIdLo" r:qs="rIdQs" r:cs="rIdCs"/></a:graphicData></a:graphic></wp:inline></w:drawing>`;
const point = (id: string, label: string) =>
  `<dgm:pt modelId="${id}"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:p><a:r><a:t>${label}</a:t></a:r></a:p></dgm:t></dgm:pt>`;
const diagram = (name: string, xml: string, kind: string, id: string, rel: string): Part => ({
  path: `word/diagrams/${name}.xml`,
  xml,
  contentType: `${OFFICE}.drawingml.${kind}+xml`,
  rel: { id, type: `${R}/${rel}` },
});
const DIAGRAM_PARTS: Part[] = [
  diagram(
    "data1",
    `<dgm:dataModel xmlns:dgm="${DGM}" xmlns:a="${A}"><dgm:ptLst><dgm:pt modelId="0" type="doc"><dgm:prSet/><dgm:spPr/></dgm:pt>${point("1", "Plan")}${point("2", "Test")}</dgm:ptLst><dgm:cxnLst><dgm:cxn modelId="3" srcId="0" destId="1" srcOrd="0" destOrd="0"/><dgm:cxn modelId="4" srcId="0" destId="2" srcOrd="1" destOrd="0"/></dgm:cxnLst><dgm:bg/><dgm:whole/></dgm:dataModel>`,
    "diagramData",
    "rIdDm",
    "diagramData"
  ),
  diagram(
    "layout1",
    `<dgm:layoutDef xmlns:dgm="${DGM}" uniqueId="urn:microsoft.com/office/officeart/2005/8/layout/process1"><dgm:title val=""/><dgm:layoutNode name="Name0"/></dgm:layoutDef>`,
    "diagramLayout",
    "rIdLo",
    "diagramLayout"
  ),
  diagram(
    "quickStyle1",
    `<dgm:styleDef xmlns:dgm="${DGM}" uniqueId="urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1"/>`,
    "diagramStyle",
    "rIdQs",
    "diagramQuickStyle"
  ),
  diagram(
    "colors1",
    `<dgm:colorsDef xmlns:dgm="${DGM}" uniqueId="urn:microsoft.com/office/officeart/2005/8/colors/accent1_2"/>`,
    "diagramColors",
    "rIdCs",
    "diagramColors"
  ),
];

test("a SmartArt frame and its parts are published as authored after an edit to its paragraph", async () => {
  const bytes = docx(`<w:p w14:paraId="10000001">${run("Steps: ")}<w:r>${FRAME}</w:r>${run(" done.")}</w:p>`, DIAGRAM_PARTS);
  for (const saved of await publishTwice(bytes, typeAt("body", 0, "Two "))) {
    const xml = text(saved, "word/document.xml")!;
    expect(xml).toContain("Two Steps: ");
    expect(xml).toContain(FRAME);
    const rels = text(saved, "word/_rels/document.xml.rels")!;
    for (const part of DIAGRAM_PARTS) {
      expect(text(saved, part.path)).toBe(part.xml);
      expect(rels).toContain(`Id="${part.rel!.id}"`);
    }
  }
});

const RUBY = `<w:ruby><w:rubyPr><w:hps w:val="10"/><w:hpsRaise w:val="18"/><w:hpsBaseText w:val="21"/><w:lid w:val="ja-JP"/></w:rubyPr><w:rt><w:r><w:t>かんじ</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>漢字</w:t></w:r></w:rubyBase></w:ruby>`;

test("a phonetic guide is published with its guide, base text and run formatting, bold over it too", async () => {
  const bytes = docx(
    `<w:p w14:paraId="10000001">${run("読み: ")}<w:r><w:rPr><w:color w:val="C00000"/></w:rPr>${RUBY}</w:r>${run(" です")}</w:p>`
  );
  for (const saved of await publishTwice(bytes, typeAt("body", 0, "Lesson "))) {
    const xml = text(saved, "word/document.xml")!;
    expect(xml).toContain("Lesson 読み: ");
    expect(xml).toContain(`<w:r><w:rPr><w:color w:val="C00000"/></w:rPr>${RUBY}</w:r>`);
  }
  const bold = await publishTwice(bytes, (session) => {
    const [{ paraId }] = session.paragraphs("body");
    // The guide is one unit, after "読み: ".
    session.toggleMark(
      { story: "body", start: { paraId, offset: 4 }, end: { paraId, offset: 5 } },
      { type: "bold" }
    );
  });
  for (const saved of bold)
    expect(text(saved, "word/document.xml")).toContain(
      `<w:r><w:rPr><w:b/><w:bCs/><w:color w:val="C00000"/></w:rPr>${RUBY}</w:r>`
    );
});
