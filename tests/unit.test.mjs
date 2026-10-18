// Small checks of the conversion rules that need no sample files.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { autoGroups, estimateAccuracy, parseValue } from '../src/convert.mjs';
import { resolveBlocks } from '../src/table.mjs';

test('숫자로 바꾸는 값', () => {
  assert.deepEqual(parseValue('1,234'), { value: 1234, numFmt: '#,##0' });
  assert.deepEqual(parseValue('(1,234)'), { value: -1234, numFmt: '#,##0;(#,##0)' });
  assert.deepEqual(parseValue('△500'), { value: -500, numFmt: '#,##0;△#,##0' });
  assert.deepEqual(parseValue('△6,351'), { value: -6351, numFmt: '#,##0;△#,##0' });
  assert.deepEqual(parseValue('15,000,000원'), { value: 15000000, numFmt: '#,##0"원"' });
  assert.deepEqual(parseValue('△1.5%'), { value: -0.015, numFmt: '0.0%;△0.0%' });
  assert.deepEqual(parseValue('-1.2%'), { value: -0.012, numFmt: '0.0%' });
  assert.deepEqual(parseValue('12.5%'), { value: 0.125, numFmt: '0.0%' });
  assert.deepEqual(parseValue('3.14'), { value: 3.14, numFmt: undefined });
  assert.deepEqual(parseValue('2026'), { value: 2026, numFmt: undefined });
});

test('셀 전체가 날짜면 엑셀 날짜로, 보이는 모양은 그대로', () => {
  const cases = [
    ['2025-03-14', 'yyyy-mm-dd'],
    ['2025-3-4', 'yyyy-m-d'],
    ['2025/03/14', 'yyyy/mm/dd'],
    ['2025.03.14', 'yyyy\\.mm\\.dd'],
    ['2025. 3. 14.', 'yyyy\\. m\\. d\\.'],
    ['2025년 3월 14일', 'yyyy"년" m"월" d"일"'],
  ];
  for (const [text, numFmt] of cases) {
    const r = parseValue(text);
    assert.ok(r.value instanceof Date, text);
    assert.equal(r.numFmt, numFmt, text);
  }
  assert.equal(parseValue('2025-03-04').value.toISOString(), '2025-03-04T00:00:00.000Z');
});

test('날짜처럼 보여도 글자로 남기는 값', () => {
  for (const s of ['2025-02-30', '2025-13-01', '2025-03-14 오전', '2025년 3월 14일 까지', '25.3.14', '1.5.2']) {
    assert.equal(typeof parseValue(s).value, 'string', s);
  }
});

test('글자로 남겨야 하는 값', () => {
  for (const s of ['010-0000-0000', '007', '1.5/2.0', '약 1백만 원', '(1,234', '12345678901234567890', '▲500', '1,000만원']) {
    assert.equal(typeof parseValue(s).value, 'string', s);
  }
});

const grid = (xs, size = [842, 595]) => ({ xs, size });

test('열 경계가 같은 쪽은 한 시트로 묶음', () => {
  const xs = [25, 47, 120, 341, 790];
  assert.deepEqual(autoGroups([grid(xs), grid(xs), grid(xs)]), [[0, 1, 2]]);
});

test('한 쪽의 경계가 다른 쪽에 다 들어 있으면 한 시트로 묶음', () => {
  assert.deepEqual(autoGroups([grid([39, 80, 300, 556]), grid([39, 556])]), [[0, 1]]);
});

test('경계가 거의 안 겹치거나 쪽 크기가 다르면 시트를 나눔', () => {
  assert.deepEqual(autoGroups([grid([39, 80, 300, 556]), grid([39, 102, 233, 400, 500])]), [[0], [1]]);
  assert.deepEqual(autoGroups([grid([39, 556]), grid([39, 556], [595, 842])]), [[0], [1]]);
});

test('사방이 선으로 막힌 칸만 병합', () => {
  // 2x2 grid. Outer box drawn, no inner lines: one enclosed merged block.
  const T = 'thin';
  const box = resolveBlocks(
    [[T, T], [null, null], [T, T]],
    [[T, null, T], [T, null, T]],
  );
  assert.deepEqual(box.merges, [{ r0: 0, r1: 1, c0: 0, c1: 1, enclosed: true }]);

  // Same grid with the bottom line missing: open space, nothing is merged.
  const open = resolveBlocks(
    [[T, T], [null, null], [null, null]],
    [[T, null, T], [T, null, T]],
  );
  assert.deepEqual(open.merges, []);
});

test('예상 정확도: 테두리 칸에 든 표는 높음, 흩어진 양식은 보통', () => {
  const base = { chars: 1000, enclosedChars: 0, multiLineChars: 0, openChars: 0, scatteredChars: 0, textChars: 1000, lostChars: 0 };
  // 글자 대부분이 테두리 칸 안에 있는 표
  const table = estimateAccuracy({ ...base, enclosedChars: 950, multiLineChars: 700, openChars: 50, scatteredChars: 3 });
  assert.equal(table.level, 'high');
  assert.equal(table.label, '99% 이상');
  assert.ok(table.notes.some((n) => n.includes('여러 줄')));
  // 테두리 밖 글자가 많고 흩어진 양식
  const form = estimateAccuracy({ ...base, enclosedChars: 650, multiLineChars: 400, openChars: 350, scatteredChars: 20 });
  assert.equal(form.level, 'medium');
  assert.equal(form.label, '약 70~85%');
  assert.equal(estimateAccuracy({ ...base, enclosedChars: 990, openChars: 10, lostChars: 30 }).preserved, 0.97);
});
