// Synthetic office documents (scripts/make-fixtures.mjs) printed to PDF by Excel and
// LibreOffice (scripts/print-fixtures.sh), scored against the workbooks they came from.
// Minimums are the scores reached so far; raise them when a change improves one.
import { existsSync } from 'node:fs';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { evaluate } from './evaluate-lib.mjs';

// Merges that a PDF cannot show (a title merged across empty cells with no border, or
// columns whose edge is never drawn) keep some merge scores below 100%.
const MIN = {
  budget: { exact: 1, merges: 0.857, numbers: 1, bold: 0.995, borders: 0.997 },
  form: { exact: 1, merges: 0.266, numbers: 1, bold: 1, borders: 0.947 },
  large: { exact: 1, merges: 1, numbers: 1, bold: 0.999, borders: 1 },
  'multi-sheet': { exact: 1, merges: 1, numbers: 1, bold: 1, borders: 1 },
};

for (const program of ['excel', 'libreoffice']) {
  for (const [name, min] of Object.entries(MIN)) {
    const pdf = `fixtures/pdf/${program}/${name}.pdf`;
    test(`${program} / ${name}`, { skip: existsSync(pdf) ? false : `${pdf} 없음` }, async () => {
      const r = await evaluate(pdf, `fixtures/xlsx/${name}.xlsx`);
      assert.equal(r.sheets, r.originalSheets, '시트 개수가 원본과 같아야 함');
      // These are clean bordered documents: the estimate must say so, and the score back it.
      assert.equal(r.quality.label, '99% 이상', '예상 정확도');
      assert.ok(r.exact >= 0.99, '예상 정확도 99% 이상이면 실제도 99% 이상');
      for (const [key, value] of Object.entries(min)) {
        assert.ok(r[key] >= value - 1e-9, `${key} ${(100 * r[key]).toFixed(1)}% < 기준 ${(100 * value).toFixed(1)}%`);
      }
    });
  }
}
