// Convert every fixture PDF and score it against its answer-key workbook.
// Usage: node scripts/eval-fixtures.mjs [--split-lines]
import { existsSync, readdirSync } from 'node:fs';
import { evaluate } from '../tests/evaluate-lib.mjs';

const options = { splitLines: process.argv.includes('--split-lines') };
const pct = (v) => `${(100 * v).toFixed(1).padStart(5)}%`;

console.log('program      file          pages  sheets  text   merges  numbers  bold   borders  time');
for (const program of ['excel', 'libreoffice']) {
  const dir = `fixtures/pdf/${program}`;
  if (!existsSync(dir)) continue;
  for (const file of readdirSync(dir).filter((f) => f.endsWith('.pdf')).sort()) {
    const name = file.replace(/\.pdf$/, '');
    const r = await evaluate(`${dir}/${file}`, `fixtures/xlsx/${name}.xlsx`, options);
    const sheets = `${r.sheets}/${r.originalSheets}`;
    console.log(`${program.padEnd(12)} ${name.padEnd(13)} ${String(r.pages).padStart(5)}  ${sheets.padStart(6)} ${pct(r.exact)} ${pct(r.merges)} ${pct(r.numbers)} ${pct(r.bold)} ${pct(r.borders)}  ${(r.ms / 1000).toFixed(1)}s`);
  }
}
