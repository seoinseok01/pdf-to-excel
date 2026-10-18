// Compare a converted workbook against the original Excel file the PDF was printed from.
// Usage: node scripts/compare.mjs original.xlsx converted.xlsx [originalSheet...]
import XLSX from 'xlsx';
import { compareWorkbooks } from '../tests/compare-lib.mjs';

const [origPath, convPath, ...sheetNames] = process.argv.slice(2);
const orig = XLSX.readFile(origPath);
const conv = XLSX.readFile(convPath);
const pct = (a, b) => (b ? `${((100 * a) / b).toFixed(1)}%` : '-');

for (const r of compareWorkbooks(orig, conv, sheetNames.length ? sheetNames : undefined)) {
  console.log(`\n## ${r.name}  (${r.cells} non-empty cells)`);
  console.log(`text exact   ${r.exact}/${r.cells} = ${pct(r.exact, r.cells)}`);
  console.log(`text partial ${r.partial}  (inside a larger cell or split)`);
  console.log(`missing      ${r.missing.length}`);
  console.log(`layout order ${(100 * r.ordered).toFixed(1)}% of matched pairs keep row/col order`);
  console.log(`merges       ${r.mergesKept}/${r.merges} merged blocks with text reproduced = ${pct(r.mergesKept, r.merges)}`);
  console.log(`numbers      ${r.numbersKept}/${r.numbers} numeric cells stay numeric = ${pct(r.numbersKept, r.numbers)}`);
  if (r.missing.length) console.log('missing samples:', r.missing.slice(0, 12).map((m) => `${m.addr}=${JSON.stringify(m.text.slice(0, 30))}`).join(', '));
}
