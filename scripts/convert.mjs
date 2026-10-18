// Usage: node scripts/convert.mjs input.pdf [output.xlsx] [--split-lines]
import path from 'node:path';
import { convertPdf } from '../src/node.mjs';

const args = process.argv.slice(2).filter((a) => !a.startsWith('--'));
const splitLines = process.argv.includes('--split-lines');
const input = args[0];
const output = args[1] ?? path.join('out', path.basename(input).replace(/\.pdf$/i, '') + '.xlsx');
const { grids, groups, workbook } = await convertPdf(input, { splitLines });
await workbook.xlsx.writeFile(output);
workbook.worksheets.forEach((ws, i) => {
  const pages = groups[i].map((k) => grids[k].page).join(',');
  console.log(`${ws.name}: ${ws.rowCount} rows x ${ws.columnCount} cols (PDF pages ${pages})`);
});
console.log('->', output);
