// Usage: node scripts/compare-format.mjs original.xlsx converted.xlsx
import ExcelJS from 'exceljs';
import { compareFormatting } from '../tests/format-lib.mjs';

const [a, b] = process.argv.slice(2);
const orig = new ExcelJS.Workbook();
const conv = new ExcelJS.Workbook();
await orig.xlsx.readFile(a);
await conv.xlsx.readFile(b);
const s = compareFormatting(orig, conv);
const pct = (x, y) => (y ? `${((100 * x) / y).toFixed(1)}%` : '-');
console.log(`paired cells   ${s.pairs}`);
console.log(`bold same      ${pct(s.bold, s.pairs)}`);
console.log(`colour same    ${pct(s.color, s.pairs)}`);
console.log(`borders kept   ${pct(s.borderSidesKept, s.borderSides)} of ${s.borderSides} bordered sides, same style ${pct(s.borderStyleKept, s.borderSides)}`);
