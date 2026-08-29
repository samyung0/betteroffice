#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import {
  Presentation,
  PresentationFile,
  SpreadsheetFile,
  Workbook,
} from "@oai/artifact-tool";

const [format, outputPath, previewDirectory] = process.argv.slice(2);

if (!format || !outputPath || !previewDirectory) {
  throw new Error(
    "usage: generate_artifact_fixture.mjs <xlsx|pptx> OUTPUT PREVIEW_DIRECTORY"
  );
}

async function writeBlob(filePath, blob) {
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
}

async function buildWorkbook() {
  const workbook = Workbook.create();
  const summary = workbook.worksheets.add("Summary");
  const source = workbook.worksheets.add("Source Data");

  source.getRange("A1:D6").values = [
    ["Month", "Viewer ms", "Editor ms", "Files"],
    ["Jan", 240, 810, 12],
    ["Feb", 225, 790, 18],
    ["Mar", 210, 760, 24],
    ["Apr", 198, 735, 30],
    ["May", 185, 710, 36],
  ];
  source.freezePanes.freezeRows(1);
  source.showGridLines = false;
  source.getRange("A1:D1").format = {
    fill: "#0F172A",
    font: { bold: true, color: "#FFFFFF" },
    horizontalAlignment: "center",
  };
  source.getRange("A1:D6").format.borders = {
    preset: "all",
    style: "thin",
    color: "#CBD5E1",
  };
  source.getRange("A1:A6").format.columnWidth = 14;
  source.getRange("B1:D6").format.columnWidth = 13;

  summary.showGridLines = false;
  summary.getRange("A1:H2").merge();
  summary.getRange("A1").values = [["BetterOffice browser round-trip fixture"]];
  summary.getRange("A1:H2").format = {
    fill: "#0F172A",
    font: { bold: true, color: "#FFFFFF", size: 20 },
    horizontalAlignment: "left",
    verticalAlignment: "center",
  };
  summary.getRange("A4:D4").values = [
    ["Metric", "Current", "Target", "Status"],
  ];
  summary.getRange("A5:C7").values = [
    ["Viewer load (ms)", 185, 200],
    ["Editor load (ms)", 710, 750],
    ["Fixture count", 36, 40],
  ];
  summary.getRange("D5").formulas = [['=IF(B5<=C5,"PASS","CHECK")']];
  summary.getRange("D5:D6").fillDown();
  summary.getRange("D7").formulas = [['=IF(B7>=C7,"PASS","CHECK")']];
  summary.getRange("A4:D4").format = {
    fill: "#2563EB",
    font: { bold: true, color: "#FFFFFF" },
    horizontalAlignment: "center",
  };
  summary.getRange("A4:D7").format.borders = {
    preset: "all",
    style: "thin",
    color: "#CBD5E1",
  };
  summary.getRange("A4:A7").format.columnWidth = 24;
  summary.getRange("B4:D7").format.columnWidth = 13;
  summary.getRange("B5:C6").format.numberFormat = '0 "ms"';

  summary.getRange("A9:D9").merge();
  summary.getRange("A9").values = [["EVO_EDIT_MARKER_XLSX"]];
  summary.getRange("A9:D9").format = {
    fill: "#DBEAFE",
    font: { bold: true, color: "#1E3A8A" },
    horizontalAlignment: "center",
  };

  summary.getRange("F4:H10").formulas = [
    ["='Source Data'!A1", "='Source Data'!B1", "='Source Data'!C1"],
    ["='Source Data'!A2", "='Source Data'!B2", "='Source Data'!C2"],
    ["='Source Data'!A3", "='Source Data'!B3", "='Source Data'!C3"],
    ["='Source Data'!A4", "='Source Data'!B4", "='Source Data'!C4"],
    ["='Source Data'!A5", "='Source Data'!B5", "='Source Data'!C5"],
    ["='Source Data'!A6", "='Source Data'!B6", "='Source Data'!C6"],
    ["", "", ""],
  ];
  const chart = summary.charts.add("line", summary.getRange("F4:H9"));
  chart.title = "Viewer stays lighter than editor";
  chart.hasLegend = true;
  chart.xAxis = { axisType: "textAxis" };
  chart.yAxis = { numberFormatCode: '0 "ms"' };
  chart.setPosition("F4", "M19");
  summary.freezePanes.freezeRows(2);

  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(outputPath);

  for (const sheetName of ["Summary", "Source Data"]) {
    const preview = await workbook.render({
      sheetName,
      autoCrop: "all",
      scale: 1,
      format: "png",
    });
    await writeBlob(
      path.join(
        previewDirectory,
        `${sheetName.toLowerCase().replaceAll(" ", "-")}.png`
      ),
      preview
    );
  }
}

function addTextBox(slide, text, position, style = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    position,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  shape.text = text;
  shape.text.style = style;
  return shape;
}

async function buildPresentation() {
  const presentation = Presentation.create({
    slideSize: { width: 1280, height: 720 },
  });

  const titleSlide = presentation.slides.add();
  titleSlide.background.fill = "slate-50";
  addTextBox(
    titleSlide,
    "BETTEROFFICE · BROWSER PROOF",
    { left: 72, top: 64, width: 500, height: 32 },
    { fontSize: 13, bold: true, color: "blue-600" }
  );
  addTextBox(
    titleSlide,
    "View first.\nLoad editing only on demand.",
    { left: 72, top: 150, width: 620, height: 190 },
    { fontSize: 44, bold: true, color: "slate-950" }
  );
  addTextBox(
    titleSlide,
    "EVO_EDIT_MARKER_PPTX",
    { left: 72, top: 380, width: 420, height: 44 },
    { fontSize: 18, bold: true, color: "blue-700" }
  );

  const stages = [
    ["1", "Inspect", "Identify format and cheap metadata"],
    ["2", "View", "Render with the smallest format-specific engine"],
    ["3", "Edit", "Replace the iframe with the full editor runtime"],
  ];
  stages.forEach(([number, label, detail], index) => {
    const left = 720 + index * 166;
    titleSlide.shapes.add({
      geometry: "roundRect",
      position: { left, top: 220, width: 148, height: 178 },
      fill: index === 2 ? "blue-600" : "white",
      line: {
        style: "solid",
        fill: index === 2 ? "blue-600" : "slate-200",
        width: 1,
      },
      borderRadius: "rounded-xl",
      shadow: "shadow-sm",
    });
    addTextBox(
      titleSlide,
      number,
      { left: left + 16, top: 238, width: 28, height: 28 },
      { fontSize: 13, bold: true, color: index === 2 ? "white" : "blue-600" }
    );
    addTextBox(
      titleSlide,
      label,
      { left: left + 16, top: 274, width: 116, height: 28 },
      { fontSize: 18, bold: true, color: index === 2 ? "white" : "slate-950" }
    );
    addTextBox(
      titleSlide,
      detail,
      { left: left + 16, top: 316, width: 116, height: 64 },
      { fontSize: 11, color: index === 2 ? "blue-100" : "slate-600" }
    );
  });

  const chartSlide = presentation.slides.add();
  chartSlide.background.fill = "white";
  addTextBox(
    chartSlide,
    "Runtime budgets are format-specific",
    { left: 72, top: 58, width: 700, height: 54 },
    { fontSize: 32, bold: true, color: "slate-950" }
  );
  addTextBox(
    chartSlide,
    "Synthetic values make the chart deterministic; the POC replaces them with measured results.",
    { left: 72, top: 118, width: 840, height: 34 },
    { fontSize: 15, color: "slate-600" }
  );
  chartSlide.charts.add("bar", {
    position: { left: 72, top: 190, width: 760, height: 430 },
    categories: ["DOCX", "XLSX", "PPTX"],
    series: [
      { name: "Viewer", values: [2.4, 1.2, 1.4], fill: "blue-500" },
      { name: "Editor", values: [8.1, 4.7, 3.9], fill: "slate-700" },
    ],
    hasLegend: true,
    dataLabels: { showValue: true, position: "outEnd" },
    xAxis: { title: { text: "Relative memory units" } },
    yAxis: {
      majorGridlines: { style: "solid", fill: "slate-200", width: 1 },
    },
  });
  addTextBox(
    chartSlide,
    "Preservation sentinels",
    { left: 890, top: 205, width: 300, height: 40 },
    { fontSize: 20, bold: true, color: "slate-950" }
  );
  addTextBox(
    chartSlide,
    "• Two slides\n• Editable text boxes\n• Native chart object\n• Theme colors\n• Unicode: 中文 · ✓",
    { left: 890, top: 265, width: 300, height: 220 },
    { fontSize: 17, color: "slate-700" }
  );

  await fs.mkdir(previewDirectory, { recursive: true });
  for (const [index, slide] of presentation.slides.items.entries()) {
    const stem = `slide-${String(index + 1).padStart(2, "0")}`;
    const png = await presentation.export({ slide, format: "png", scale: 1 });
    await writeBlob(path.join(previewDirectory, `${stem}.png`), png);
    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(
      path.join(previewDirectory, `${stem}.layout.json`),
      await layout.text()
    );
  }
  const montage = await presentation.export({
    format: "webp",
    montage: true,
    scale: 1,
  });
  await writeBlob(path.join(previewDirectory, "montage.webp"), montage);
  const output = await PresentationFile.exportPptx(presentation);
  await output.save(outputPath);
}

await fs.mkdir(path.dirname(outputPath), { recursive: true });
await fs.mkdir(previewDirectory, { recursive: true });

if (format === "xlsx") {
  await buildWorkbook();
} else if (format === "pptx") {
  await buildPresentation();
} else {
  throw new Error(`unsupported format: ${format}`);
}
