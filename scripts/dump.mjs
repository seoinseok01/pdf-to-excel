// Print non-empty cells of an .xlsx row by row: A1=value | B1=value ...
import ExcelJS from 'exceljs';
const wb = new ExcelJS.Workbook();
await wb.xlsx.readFile(process.argv[2]);
const only = process.argv[3];
for (const ws of wb.worksheets) {
  if (only && ws.name !== only) continue;
  console.log(`## ${ws.name} merges=${(ws.model.merges ?? []).length}`);
  ws.eachRow((row, r) => {
    const parts = [];
    row.eachCell((cell) => {
      if (cell.isMerged && cell.master !== cell) return;
      const v = cell.value;
      if (v == null || v === '') return;
      const t = typeof v === 'object' ? (v.text ?? JSON.stringify(v)) : v;
      parts.push(`${cell.address}=${typeof v === 'number' ? '#' + v : JSON.stringify(String(t).slice(0, 40))}`);
    });
    if (parts.length) console.log(parts.join(' | '));
  });
}
