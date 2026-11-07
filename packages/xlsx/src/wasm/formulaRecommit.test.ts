import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';

import { initWasm, openWorkbook } from './loader';

const WASM = resolve(import.meta.dir, './generated/xlsx_wasm_bg.wasm');
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const RELS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const STYLES = `<styleSheet xmlns="${MAIN}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>`;

const CELLS = [
  { a1: 'A1', col: 0, xml: '<v>21</v>', shown: '21', formula: null },
  { a1: 'B1', col: 1, xml: '<f>[1]Sheet1!A1</f><v>5</v>', shown: '=[1]Sheet1!A1', formula: '[1]Sheet1!A1' },
  { a1: 'C1', col: 2, xml: '<f>SUM([1]Prices!$B$2:$B$4)*2</f><v>12</v>', shown: '=SUM([1]Prices!$B$2:$B$4)*2', formula: 'SUM([1]Prices!$B$2:$B$4)*2' },
  { a1: 'D1', col: 3, xml: '<f>=A1*2</f><v>0</v>', shown: '=A1*2', formula: 'A1*2' },
  { a1: 'E1', col: 4, xml: '<f>A1*2</f><v>42</v>', shown: '=A1*2', formula: 'A1*2', style: 1 },
];

async function book(calcPr = ''): Promise<Uint8Array> {
  const cells = CELLS.map((cell) => `<c r="${cell.a1}"${cell.style ? ` s="${cell.style}"` : ''}>${cell.xml}</c>`).join('');
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="${PACKAGE_RELS}"><Relationship Id="rId1" Type="${RELS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('xl/workbook.xml', `<workbook xmlns="${MAIN}" xmlns:r="${RELS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>${calcPr}</workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="${PACKAGE_RELS}"><Relationship Id="rId1" Type="${RELS}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${RELS}/styles" Target="styles.xml"/></Relationships>`);
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="${MAIN}"><sheetData><row r="1">${cells}</row></sheetData></worksheet>`);
  zip.file('xl/styles.xml', STYLES);
  return zip.generateAsync({ type: 'uint8array' });
}

async function savedCaches(bytes: Uint8Array): Promise<Record<string, string | null>> {
  const sheet = await (await JSZip.loadAsync(bytes)).file('xl/worksheets/sheet1.xml')!.async('string');
  return Object.fromEntries([...sheet.matchAll(/<c r="([A-Z]+\d+)"[^>]*>(.*?)<\/c>/g)]
    .map((cell) => [cell[1], /<v>(.*?)<\/v>/.exec(cell[2])?.[1] ?? null]));
}

const FORMULAS = CELLS.filter((cell) => cell.formula !== null && !cell.style);

// Capy's handle reads and edits through cell() and editCells(); upstream's
// readCells/applyEdits agent API stays deleted (2026-10-17).
describe('re-committing shown formulas', () => {
  beforeAll(() => initWasm(new Uint8Array(readFileSync(WASM))));

  it('shows one equals sign and saves byte for byte after re-committing every shown formula', async () => {
    const source = await book();
    const handle = openWorkbook(source);
    try {
      for (const cell of CELLS) {
        expect(handle.cell(0, 0, cell.col)).toMatchObject({ input: cell.shown, isFormula: cell.formula !== null && !cell.style });
      }
      handle.editCells(0, FORMULAS.map((cell) => ({ row: 0, col: cell.col, input: handle.cell(0, 0, cell.col).input })));
      expect(new Uint8Array(handle.save())).toEqual(new Uint8Array(source));
    } finally {
      handle.dispose();
    }
  });

  it('takes a typed external reference as a formula and keeps cached external values', async () => {
    const handle = openWorkbook(await book('<calcPr fullCalcOnLoad="1"/>'));
    try {
      handle.editCells(0, [{ row: 1, col: 0, input: '=[1]Sheet1!B2*2' }]);
      expect(handle.cell(0, 1, 0)).toMatchObject({ input: '=[1]Sheet1!B2*2', isFormula: true });
      expect(await savedCaches(handle.save())).toMatchObject({ B1: '5', D1: '42' });
    } finally {
      handle.dispose();
    }
  });

  it('enters a formula typed into a Text-formatted cell as text, as Excel does', async () => {
    const handle = openWorkbook(await book());
    try {
      handle.editCells(0, [{ row: 0, col: 4, input: handle.cell(0, 0, 4).input }]);
      expect(handle.cell(0, 0, 4).isFormula).toBe(false);
    } finally {
      handle.dispose();
    }
  });
});
