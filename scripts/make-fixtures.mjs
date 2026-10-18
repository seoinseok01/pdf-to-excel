// Build synthetic Excel workbooks that look like typical public-office documents.
// They hold no real data, so unlike samples/ they can be committed; the PDFs made from
// them in Excel / LibreOffice (scripts/print-fixtures.sh) are the test inputs and these
// files are the answer key.
// Usage: node scripts/make-fixtures.mjs  -> fixtures/xlsx/*.xlsx
import ExcelJS from 'exceljs';

const OUT = 'fixtures/xlsx';
const thin = { style: 'thin' };
const box = { top: thin, left: thin, bottom: thin, right: thin };
const font = (extra = {}) => ({ name: '맑은 고딕', size: 10, ...extra });

function frame(ws, r0, c0, r1, c1) {
  for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) ws.getCell(r, c).border = box;
}

// Seeded random numbers so the fixtures are identical every run.
let seed = 42;
const rand = () => ((seed = (seed * 1103515245 + 12345) % 2147483648) / 2147483648);
const pick = (list) => list[Math.floor(rand() * list.length)];

// 1. Budget table: two-level merged header, numbers with commas, negative changes, percents,
//    subtotals, a unit note, repeated header rows on every printed page.
async function budget() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('예산');
  ws.columns = [{ width: 6 }, { width: 22 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 10 }, { width: 18 }];
  ws.mergeCells('A1:G1');
  Object.assign(ws.getCell('A1'), { value: '2026년도 세출예산 사업별 현황', font: font({ size: 16, bold: true }), alignment: { horizontal: 'center' } });
  Object.assign(ws.getCell('G2'), { value: '(단위: 천원)', font: font({ size: 9 }), alignment: { horizontal: 'right' } });
  const head = [['연번', '사업명', '예산액', '', '증감', '', '비고'], ['', '', '2026년', '2025년', '금액', '비율', '']];
  head.forEach((row, i) => row.forEach((v, c) => Object.assign(ws.getCell(3 + i, c + 1), { value: v || null, font: font({ bold: true }), alignment: { horizontal: 'center', vertical: 'middle' }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } } })));
  for (const m of ['A3:A4', 'B3:B4', 'C3:D3', 'E3:F3', 'G3:G4']) ws.mergeCells(m);
  const names = ['청년 일자리 지원', '노인 돌봄 서비스', '도로 보수', '하천 정비', '문화센터 운영', '도서관 확충', '방역 소독', '농업인 교육', '소상공인 지원', '전통시장 활성화', '공원 조성', '주차장 확충'];
  let r = 5;
  for (let i = 1; i <= 120; i++) {
    const now = Math.round(rand() * 900000 + 10000);
    const prev = Math.round(now * (0.7 + rand() * 0.6));
    const diff = now - prev;
    ws.getRow(r).values = [i, `${pick(names)} ${i}`, now, prev, diff, prev ? diff / prev : 0, i % 7 === 0 ? '신규' : null];
    ws.getCell(r, 3).numFmt = ws.getCell(r, 4).numFmt = '#,##0';
    ws.getCell(r, 5).numFmt = '#,##0;△#,##0';
    ws.getCell(r, 6).numFmt = '0.0%';
    ws.getRow(r).font = font();
    ws.getCell(r, 1).alignment = { horizontal: 'center' };
    r++;
  }
  ws.mergeCells(r, 1, r, 2);
  Object.assign(ws.getCell(r, 1), { value: '합계', font: font({ bold: true }), alignment: { horizontal: 'center' } });
  for (const c of [3, 4, 5]) {
    const col = String.fromCharCode(64 + c);
    Object.assign(ws.getCell(r, c), { value: { formula: `SUM(${col}5:${col}${r - 1})` }, numFmt: c === 5 ? '#,##0;△#,##0' : '#,##0', font: font({ bold: true }) });
  }
  frame(ws, 3, 1, r, 7);
  ws.pageSetup = { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: '3:4' };
  await wb.xlsx.writeFile(`${OUT}/budget.xlsx`);
}

// 2. Workbook with three differently laid out sheets (printed as one PDF).
async function multiSheet() {
  const wb = new ExcelJS.Workbook();
  const a = wb.addWorksheet('인원현황');
  a.columns = [{ width: 18 }, { width: 10 }, { width: 10 }, { width: 10 }];
  a.getRow(1).values = ['부서', '정원', '현원', '결원'];
  ['기획예산과', '총무과', '민원과', '세무과', '복지정책과', '환경과'].forEach((d, i) => {
    const q = 20 + i * 3;
    a.getRow(i + 2).values = [d, q, q - (i % 3), { formula: `B${i + 2}-C${i + 2}` }];
  });
  a.getRow(1).font = font({ bold: true });
  frame(a, 1, 1, 7, 4);

  const b = wb.addWorksheet('일정');
  b.columns = [{ width: 14 }, { width: 30 }, { width: 14 }];
  b.getRow(1).values = ['날짜', '행사', '장소'];
  const events = ['주민설명회', '안전점검', '체육대회', '봉사활동', '정책토론회'];
  events.forEach((e, i) => {
    b.getRow(i + 2).values = [new Date(Date.UTC(2026, 9, 5 + i * 4)), e, i % 2 ? '대회의실' : '시민회관'];
    b.getCell(i + 2, 1).numFmt = 'yyyy-mm-dd';
  });
  b.getRow(1).font = font({ bold: true });
  frame(b, 1, 1, 6, 3);

  const c = wb.addWorksheet('메모');
  c.columns = [{ width: 60 }];
  c.getCell('A1').value = '※ 위 자료는 2026. 10. 1. 기준입니다.';
  c.getCell('A2').value = '※ 문의: 기획예산과 (내선 1234)';
  await wb.xlsx.writeFile(`${OUT}/multi-sheet.xlsx`);
}

// 3. Application form: merged label cells, fills, a thick outer border, bold and colored text,
//    an empty signature area.
async function form() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('신청서');
  ws.columns = Array.from({ length: 8 }, () => ({ width: 11 }));
  ws.mergeCells('A1:H1');
  Object.assign(ws.getCell('A1'), { value: '보조금 지원 신청서', font: font({ size: 18, bold: true }), alignment: { horizontal: 'center' } });
  const rows = [
    ['신청인', '성명', '홍길동', '', '생년월일', '1990. 5. 17.', '', ''],
    ['', '주소', '서울특별시 중구 세종대로 110', '', '', '', '', ''],
    ['', '연락처', '02-123-4567', '', '휴대전화', '010-0000-0000', '', ''],
    ['사업', '사업명', '마을 공동체 텃밭 가꾸기', '', '', '', '', ''],
    ['', '기간', '2026-11-01', '', '~', '2027-03-31', '', ''],
    ['', '금액', 15000000, '', '자부담', 3000000, '', ''],
  ];
  rows.forEach((vals, i) => vals.forEach((v, c) => (ws.getCell(i + 3, c + 1).value = v === '' ? null : v)));
  for (const m of ['A3:A5', 'A6:A8', 'C3:D3', 'F3:H3', 'C4:H4', 'C5:D5', 'F5:H5', 'C6:H6', 'C7:D7', 'F7:H7', 'C8:D8', 'F8:H8']) ws.mergeCells(m);
  for (const addr of ['C8', 'F8']) ws.getCell(addr).numFmt = '#,##0"원"';
  for (let r = 3; r <= 8; r++) {
    for (let c = 1; c <= 8; c++) {
      const cell = ws.getCell(r, c);
      cell.border = box;
      cell.font = font();
      cell.alignment = { vertical: 'middle', horizontal: c <= 2 || c === 5 ? 'center' : 'left' };
      if (c <= 2 || c === 5) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
        cell.font = font({ bold: true, color: { argb: 'FFFFFFFF' } });
      }
    }
  }
  for (let c = 1; c <= 8; c++) {
    ws.getCell(3, c).border = { ...ws.getCell(3, c).border, top: { style: 'medium' } };
    ws.getCell(8, c).border = { ...ws.getCell(8, c).border, bottom: { style: 'medium' } };
  }
  ws.mergeCells('A10:H10');
  Object.assign(ws.getCell('A10'), { value: '위와 같이 보조금 지원을 신청합니다.', alignment: { horizontal: 'center' }, font: font({ size: 11 }) });
  ws.mergeCells('A12:H12');
  Object.assign(ws.getCell('A12'), { value: '※ 허위 기재 시 지원이 취소될 수 있습니다.', font: font({ color: { argb: 'FFC00000' }, bold: true }) });
  ws.pageSetup = { paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 1 };
  await wb.xlsx.writeFile(`${OUT}/form.xlsx`);
}

// 4. Large landscape table for speed tests: about 3000 rows over many pages.
async function large() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('대장');
  ws.columns = [{ width: 8 }, { width: 14 }, { width: 24 }, { width: 12 }, { width: 14 }, { width: 14 }, { width: 30 }];
  ws.getRow(1).values = ['번호', '등록일', '업체명', '구분', '금액', '담당자', '비고'];
  ws.getRow(1).font = font({ bold: true });
  const kinds = ['물품', '용역', '공사'];
  const people = ['김민수', '이서연', '박지훈', '최유진', '정다은'];
  for (let i = 1; i <= 3000; i++) {
    const row = ws.getRow(i + 1);
    row.values = [i, new Date(Date.UTC(2026, Math.floor(rand() * 9), 1 + Math.floor(rand() * 28))), `(주)업체${String(i).padStart(4, '0')}`, pick(kinds), Math.round(rand() * 50000000), pick(people), i % 11 === 0 ? '분할 납품' : null];
    row.getCell(2).numFmt = 'yyyy-mm-dd';
    row.getCell(5).numFmt = '#,##0';
  }
  frame(ws, 1, 1, 3001, 7);
  ws.pageSetup = { paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: '1:1' };
  await wb.xlsx.writeFile(`${OUT}/large.xlsx`);
}

await budget();
await multiSheet();
await form();
await large();
console.log('fixtures written to', OUT);
