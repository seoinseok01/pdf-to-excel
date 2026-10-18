// Render an .xlsx back to HTML to eyeball conversion results.
// Usage: node scripts/preview.mjs out/file.xlsx  -> out/file.html
import ExcelJS from 'exceljs';
import { writeFile } from 'node:fs/promises';
import { renderSheet } from '../src/preview.mjs';

const input = process.argv[2];
const wb = new ExcelJS.Workbook();
await wb.xlsx.readFile(input);

let html = `<!doctype html><meta charset="utf-8"><title>${input}</title>
<style>body{font-family:sans-serif;margin:16px;background:#fff}
.sheet{border-collapse:collapse;table-layout:fixed;margin-bottom:40px}
.sheet td{font-size:11px;padding:1px 3px;overflow:hidden;white-space:pre;vertical-align:middle}
.sheet td.num{outline:1px dashed #0a0}h2{font-size:14px}</style>`;
for (const ws of wb.worksheets) html += `<h2>${ws.name}</h2>${renderSheet(ws)}`;
const output = input.replace(/\.xlsx$/, '.html');
await writeFile(output, html);
console.log('->', output);
