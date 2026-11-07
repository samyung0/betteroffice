import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';

import { initWasm, openWorkbook } from '../index';
import type { WorkbookHandle } from '../index';

const WASM = resolve(import.meta.dir, './generated/xlsx_wasm_bg.wasm');

const CELLS = [
  { a1: 'A1', row: 0, col: 0, stored: '_xlfn.SEQUENCE(2,2)', shown: 'SEQUENCE(2,2)' },
  { a1: 'D1', row: 0, col: 3, stored: '_xlfn.LET(_xlpm.x,2,_xlpm.y,3,_xlpm.x*_xlpm.y)', shown: 'LET(x,2,y,3,x*y)' },
  { a1: 'D2', row: 1, col: 3, stored: '_xlfn.REDUCE(0,{1,2},_xlfn.LAMBDA(_xlpm.a,_xlpm.v,_xlpm.a+_xlpm.v))', shown: 'REDUCE(0,{1,2},LAMBDA(a,v,a+v))' },
  { a1: 'D3', row: 2, col: 3, stored: 'IF(TRUE,"_xlfn.SEQUENCE(2,2)","")', shown: 'IF(TRUE,"_xlfn.SEQUENCE(2,2)","")' },
];

const ESCAPES: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' };
const UNESCAPES: Record<string, string> = Object.fromEntries(Object.entries(ESCAPES).map(([text, entity]) => [entity, text]));

function escapeXml(text: string): string {
  return text.replace(/[&<>"]/g, (character) => ESCAPES[character]);
}

function unescapeXml(text: string): string {
  return text.replace(/&(?:amp|lt|gt|quot);/g, (entity) => UNESCAPES[entity]);
}

async function workbookBytes(cells: string): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="Pick">_xlfn.SEQUENCE(3)</definedName></definedNames></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${cells}</sheetData></worksheet>`);
  return zip.generateAsync({ type: 'uint8array' });
}

function excelSaved(): Promise<Uint8Array> {
  const rows = [0, 1, 2].map((row) => {
    const cells = CELLS.filter((cell) => cell.row === row)
      .map((cell) => `<c r="${cell.a1}"><f>${escapeXml(cell.stored)}</f><v>1</v></c>`)
      .join('');
    return `<row r="${row + 1}">${cells}</row>`;
  });
  return workbookBytes(rows.join(''));
}

async function formulaTexts(saved: Uint8Array): Promise<string[]> {
  const sheet = await (await JSZip.loadAsync(saved)).file('xl/worksheets/sheet1.xml')!.async('string');
  return [...sheet.matchAll(/<f(?: [^>]*)?>([^<]*)<\/f>/g)].map((match) => unescapeXml(match[1]));
}

// Capy's handle shows formulas through cell() and rangeCells(); upstream's
// readCells/exportStructured agent reads stay deleted (2026-10-17).
function expectShown(handle: WorkbookHandle): void {
  for (const cell of CELLS) {
    expect(handle.cell(0, cell.row, cell.col).input).toBe(`=${cell.shown}`);
    expect(handle.rangeCells(0, cell.a1)[0][0].input).toBe(`=${cell.shown}`);
  }
}

describe('formula text without storage prefixes', () => {
  beforeAll(() => initWasm(new Uint8Array(readFileSync(WASM))));

  it('shows Excel-saved formulas without prefixes and saves them unchanged', async () => {
    const source = await excelSaved();
    const handle = openWorkbook(source);
    try {
      expectShown(handle);
      expect(handle.save()).toEqual(source);
      handle.editCells(0, CELLS.map((cell) => ({ row: cell.row, col: cell.col, input: handle.cell(0, cell.row, cell.col).input })));
      expect(handle.save()).toEqual(source);
      handle.editCells(0, [{ row: 4, col: 0, input: '5' }]);
      expect(await formulaTexts(handle.save())).toEqual(CELLS.map((cell) => cell.stored));
    } finally {
      handle.dispose();
    }
  });

  it('shows formulas BetterOffice saved with prefixes the same way', async () => {
    const blank = openWorkbook(await workbookBytes(''));
    const saved = (() => {
      try {
        blank.editCells(0, CELLS.map((cell) => ({ row: cell.row, col: cell.col, input: `=${cell.shown}` })));
        return blank.save();
      } finally {
        blank.dispose();
      }
    })();
    expect(await formulaTexts(saved)).toEqual(CELLS.map((cell) => cell.stored));
    const reopened = openWorkbook(saved);
    try {
      expectShown(reopened);
    } finally {
      reopened.dispose();
    }
  });
});
