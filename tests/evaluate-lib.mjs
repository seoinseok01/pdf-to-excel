// Convert a PDF and score the result against the workbook it was printed from.
import ExcelJS from 'exceljs';
import XLSX from 'xlsx';
import { convertPdf } from '../src/node.mjs';
import { compareWorkbooks } from './compare-lib.mjs';
import { compareFormatting } from './format-lib.mjs';

export async function evaluate(pdf, answer, options = {}) {
  const t0 = performance.now();
  const { workbook, grids } = await convertPdf(pdf, options);
  const ms = performance.now() - t0;
  const buffer = Buffer.from(await workbook.xlsx.writeBuffer());
  const orig = XLSX.readFile(answer);
  const rows = compareWorkbooks(orig, XLSX.read(buffer));
  const sum = (k) => rows.reduce((s, r) => s + r[k], 0);
  const origX = new ExcelJS.Workbook();
  const convX = new ExcelJS.Workbook();
  await origX.xlsx.readFile(answer);
  await convX.xlsx.load(buffer);
  const f = compareFormatting(origX, convX);
  return {
    workbook,
    quality: workbook.quality,
    pages: grids.length,
    ms,
    sheets: workbook.worksheets.length,
    originalSheets: orig.SheetNames.length,
    exact: sum('exact') / sum('cells'),
    merges: sum('merges') ? sum('mergesKept') / sum('merges') : 1,
    numbers: sum('numbers') ? sum('numbersKept') / sum('numbers') : 1,
    bold: f.pairs ? f.bold / f.pairs : 1,
    borders: f.borderSides ? f.borderStyleKept / f.borderSides : 1,
  };
}
