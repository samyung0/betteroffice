import { expect, test } from "bun:test";
import { exportOffice, seedOffice } from "./office-checkpoint";
import { createYrsSession, type YrsSession } from "../packages/docx/src/yrs";
import { rezipContainer, unzipContainer } from "../packages/docx/src/wasm/opc";

// Edits saved the way the collaboration service publishes them: the base
// package, its seed and the editor's state through exportOffice.

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
