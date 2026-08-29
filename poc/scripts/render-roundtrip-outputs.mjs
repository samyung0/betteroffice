#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";

import {
  FileBlob,
  PresentationFile,
  SpreadsheetFile,
} from "@oai/artifact-tool";

const root = path.resolve(import.meta.dirname, "../..");
const output = path.join(root, "poc/output");
const preview = path.join(output, "independent-preview");

async function writeBlob(filePath, blob) {
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
}

await fs.mkdir(preview, { recursive: true });

const workbook = await SpreadsheetFile.importXlsx(
  await FileBlob.load(path.join(output, "feature-rich.edited.xlsx"))
);
for (const sheetName of ["Summary", "Source Data"]) {
  const image = await workbook.render({
    sheetName,
    autoCrop: "all",
    scale: 1,
    format: "png",
  });
  await writeBlob(
    path.join(
      preview,
      `edited-xlsx-${sheetName.toLowerCase().replaceAll(" ", "-")}.png`
    ),
    image
  );
}

const presentation = await PresentationFile.importPptx(
  await FileBlob.load(path.join(output, "feature-rich.edited.pptx"))
);
for (const [index, slide] of presentation.slides.items.entries()) {
  const image = await presentation.export({ slide, format: "png", scale: 1 });
  await writeBlob(
    path.join(preview, `edited-pptx-slide-${index + 1}.png`),
    image
  );
}
