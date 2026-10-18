// Turn extracted page grids into an .xlsx workbook.
import ExcelJS from 'exceljs';
import { extractPage } from './extract.mjs';
import { buildPageGrid, cellText, finalizeGrid } from './table.mjs';

const PT_PER_CHAR = 5.6; // rough Excel column-width unit in points

// Read every page's text and lines. onProgress(pageNumber, pageCount) is called after each page.
export async function extractPages(doc, onProgress) {
  const pages = [];
  for (let p = 1; p <= doc.numPages; p++) {
    const page = await extractPage(doc, p);
    page.number = p;
    pages.push(page);
    onProgress?.(p, doc.numPages);
  }
  return pages;
}

// Raw grids of the pages that hold any text. Options are those of buildPageGrid.
export function gridsFromPages(pages, options = {}) {
  const grids = [];
  for (const page of pages) {
    const grid = buildPageGrid(page, options);
    grid.page = page.number;
    grid.size = [Math.round(page.width), Math.round(page.height)];
    if (grid.cells.some((row) => row.some((cell) => cell.texts.length))) grids.push(grid);
  }
  return grids;
}

// Convert a whole document. groups lists, per sheet, the indices into grids it is built from;
// the UI changes groups (merge/split sheets) or options and calls buildWorkbook again.
export async function convertDocument(doc, onProgress, options = {}) {
  const pages = await extractPages(doc, onProgress);
  const grids = gridsFromPages(pages, options);
  const groups = autoGroups(grids);
  return { pages, grids, groups, workbook: buildWorkbook(grids, groups) };
}

// The workbook carries .quality (see estimateAccuracy) for the UI.
export function buildWorkbook(grids, groups) {
  const workbook = new ExcelJS.Workbook();
  const stats = { chars: 0, enclosedChars: 0, multiLineChars: 0, openChars: 0, scatteredChars: 0 };
  groups.forEach((group, i) => {
    const ws = workbook.addWorksheet(`Sheet${i + 1}`);
    const grid = finalizeGrid(stitch(group.map((k) => grids[k])));
    for (const key of Object.keys(stats)) stats[key] += grid.stats[key];
    writeSheet(ws, grid);
  });
  const used = groups.flat().map((k) => grids[k]);
  stats.textChars = used.reduce((n, g) => n + g.textChars, 0);
  stats.lostChars = used.reduce((n, g) => n + g.lostChars, 0);
  workbook.quality = estimateAccuracy(stats);
  return workbook;
}

// Without the original Excel file accuracy cannot be measured, only estimated. Two things are
// reported:
//  - preserved: share of the PDF's characters that ended up in a cell (counted exactly)
//  - level: how clear the table structure is, with the accuracy range measured on test
//    documents of that kind (tests/fixtures.test.mjs, samples/): documents whose text sits in
//    bordered cells scored 99.8-100%; forms with much free text scattered outside borders
//    scored about 70-85%.
export function estimateAccuracy(stats) {
  const share = (n) => (stats.chars ? n / stats.chars : 0);
  const preserved = stats.textChars ? 1 - stats.lostChars / stats.textChars : 1;
  const open = share(stats.openChars);
  const scattered = share(stats.scatteredChars);
  const clear = open < 0.25 && scattered < 0.01 && preserved > 0.99;
  const notes = [];
  if (open >= 0.25) notes.push('테두리 밖에 놓인 글자가 많아서, 칸 위치가 원본과 다를 수 있어요.');
  if (scattered >= 0.01) notes.push('한 줄에 글자가 여러 조각으로 흩어진 곳(띄어 쓴 제목, 날짜·서명 줄, 체크박스 등)이 있어요.');
  if (share(stats.multiLineChars) >= 0.2) notes.push("칸 안에 여러 줄이 든 칸이 많아요. 원본이 한 줄에 한 행씩 쓴 양식이면 '여러 줄 나누기'를 켜 보세요.");
  if (preserved < 0.999) notes.push('표 바깥에 있어 옮기지 못한 글자가 조금 있어요.');
  return {
    preserved,
    level: clear ? 'high' : 'medium',
    label: clear ? '99% 이상' : '약 70~85%',
    notes,
  };
}

// A PDF does not say which sheet a page came from. Pages of one sheet share its column
// edges, so consecutive pages of the same size are grouped when the edges of one page are
// (almost) all found on the other. Anything less certain stays separate for the user to merge.
export function autoGroups(grids) {
  const groups = [];
  grids.forEach((g, i) => {
    const last = groups[groups.length - 1];
    const prev = last && grids[last[last.length - 1]];
    if (prev && samePageSize(prev, g) && edgesContained(prev.xs, g.xs)) last.push(i);
    else groups.push([i]);
  });
  return groups;
}

const samePageSize = (a, b) => a.size[0] === b.size[0] && a.size[1] === b.size[1];

function edgesContained(a, b) {
  const [small, big] = a.length <= b.length ? [a, b] : [b, a];
  const found = small.filter((x) => big.some((y) => Math.abs(x - y) < SNAP)).length;
  return small.length >= 2 && found / small.length >= 0.9;
}

const SNAP = 1.5;

function sameColumns(a, b) {
  return a.length === b.length && a.every((x, i) => Math.abs(x - b[i]) < SNAP);
}

// Column edges shared by all pages: every page's edges, with near-equal ones joined.
function unionEdges(pages) {
  const all = pages.flatMap((p) => p.xs).sort((a, b) => a - b);
  const out = [];
  for (const x of all) {
    if (out.length && x - out[out.length - 1].max < SNAP) {
      const g = out[out.length - 1];
      g.sum += x;
      g.n++;
      g.max = x;
    } else out.push({ sum: x, n: 1, max: x });
  }
  return out.map((g) => g.sum / g.n);
}

// Re-express a page grid on wider column edges. A page column covering several union
// columns is split with no border between the parts, so enclosed cells merge back.
function onColumns(grid, xs) {
  if (sameColumns(grid.xs, xs)) return grid;
  const nearest = (x) => xs.reduce((best, y, i) => (Math.abs(y - x) < Math.abs(xs[best] - x) ? i : best), 0);
  const idx = grid.xs.map(nearest);
  const nc = xs.length - 1;
  const nr = grid.cells.length;
  const cells = Array.from({ length: nr }, () => Array.from({ length: nc }, () => ({ texts: [] })));
  const top = Array.from({ length: nr + 1 }, () => new Array(nc).fill(null));
  const left = Array.from({ length: nr }, () => new Array(nc + 1).fill(null));
  for (let c = 0; c < grid.xs.length - 1; c++) {
    const k = idx[c];
    const m = Math.max(k, idx[c + 1] - 1);
    for (let r = 0; r <= nr; r++) for (let j = k; j <= m; j++) top[r][j] = grid.top[r][c];
    for (let r = 0; r < nr; r++) {
      const src = grid.cells[r][c];
      left[r][k] = grid.left[r][c];
      left[r][m + 1] = grid.left[r][c + 1];
      for (let j = k; j <= m; j++) cells[r][j].fill = src.fill;
      if (src.url) cells[r][k].url = src.url;
      for (const t of src.texts) {
        let j = k;
        while (j < m && t.x0 + 0.5 >= xs[j + 1]) j++;
        cells[r][j].texts.push(t);
      }
    }
  }
  return { ...grid, xs, cells, top, left };
}

const rowKey = (grid, r) => grid.cells[r].map(cellText).join('\u0001');
const GAP_ROW_PT = 15;

// Concatenate page grids into one on shared column edges, dropping header rows that
// Excel repeats on each page.
function stitch(pages) {
  const xs = unionEdges(pages);
  const grids = pages.map((p) => onColumns(p, xs));
  const first = grids[0];
  const headerKeys = new Set();
  for (let r = 0; r < Math.min(first.cells.length, 6); r++) {
    const key = rowKey(first, r);
    if (key.replace(/\u0001/g, '')) headerKeys.add(key);
  }

  const nc = xs.length - 1;
  const out = { xs, cells: [], heights: [], left: [], top: [] };
  grids.forEach((page, i) => {
    let skip = 0;
    // A page laid out like the previous one continues its table; any other page starts
    // a new block below a blank row, the way it sat further down the original sheet.
    const continues = i > 0 && sameColumns(pages[i - 1].xs, pages[i].xs);
    if (continues) while (skip < page.cells.length && headerKeys.has(rowKey(page, skip))) skip++;
    if (i === 0) out.top.push(page.top[0]);
    else if (continues) {
      // The border between pages is whatever the previous page drew at its bottom:
      // a missing line there means the cell continues onto this page.
      out.top.pop();
      out.top.push(out.lastBottom.map((style, c) => (style && page.top[skip][c] ? style : null)));
    } else {
      out.cells.push(Array.from({ length: nc }, () => ({ texts: [] })));
      out.heights.push(GAP_ROW_PT);
      out.left.push(new Array(nc + 1).fill(null));
      out.top.push(page.top[0]);
    }
    for (let r = skip; r < page.cells.length; r++) {
      out.cells.push(page.cells[r]);
      out.heights.push(page.heights[r]);
      out.left.push(page.left[r]);
      out.top.push(page.top[r + 1]);
    }
    out.lastBottom = page.top[page.cells.length];
  });
  return out;
}

function writeSheet(ws, grid) {
  grid.xs.slice(0, -1).forEach((x, c) => {
    ws.getColumn(c + 1).width = Math.max(1, (grid.xs[c + 1] - x) / PT_PER_CHAR);
  });
  grid.heights.forEach((h, r) => {
    ws.getRow(r + 1).height = Math.max(6, h);
  });
  grid.cells.forEach((row, r) => row.forEach((_, c) => writeCell(ws, grid, r, c)));
  for (const m of grid.merges) ws.mergeCells(m.r0 + 1, m.c0 + 1, m.r1 + 1, m.c1 + 1);
}

function writeCell(ws, grid, r, c) {
  const xl = ws.getCell(r + 1, c + 1);
  const border = {};
  const set = (side, edge) => {
    if (!edge || edge.style === 'edge') return;
    border[side] = { style: edge.style };
    if (edge.color && edge.color !== '000000') border[side].color = { argb: 'FF' + edge.color };
  };
  set('top', grid.top[r][c]);
  set('bottom', grid.top[r + 1][c]);
  set('left', grid.left[r][c]);
  set('right', grid.left[r][c + 1]);
  if (Object.keys(border).length) xl.border = border;

  const o = grid.owner[r][c];
  const info = grid.values.get(`${o.r0},${o.c0}`);
  if (info.fill) xl.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + info.fill } };
  if (o.r0 !== r || o.c0 !== c || info.value == null) return;

  const font = { name: info.font ?? '맑은 고딕', size: Math.round(Math.max(6, Math.min(info.size, 28)) * 2) / 2 };
  const lines = info.lines ?? [];
  const colorOf = (hex) => (hex && hex !== '000000' ? { argb: 'FF' + hex } : undefined);
  const uniform = lines.every((l) => l.bold === lines[0].bold && l.color === lines[0].color);
  if (lines.length && uniform) {
    if (lines[0].bold) font.bold = true;
    if (colorOf(lines[0].color)) font.color = colorOf(lines[0].color);
  } else if (!lines.length && info.fill && isDark(info.fill)) {
    // No colour information: white bold text is the usual choice on a dark fill.
    Object.assign(font, { color: { argb: 'FFFFFFFF' }, bold: true });
  }

  const { value, numFmt } = parseValue(info.value);
  if (info.url) xl.value = { text: info.value, hyperlink: info.url };
  else if (!uniform) {
    // Lines styled differently (e.g. bold title line, plain details) become rich text.
    xl.value = {
      richText: lines.map((l, i) => ({
        text: l.text + (i < lines.length - 1 ? '\n' : ''),
        font: { ...font, ...(l.bold ? { bold: true } : {}), ...(colorOf(l.color) ? { color: colorOf(l.color) } : {}) },
      })),
    };
  } else xl.value = value;
  if (numFmt && uniform && !info.url) xl.numFmt = numFmt;
  xl.alignment = {
    horizontal: (typeof value === 'number' || value instanceof Date) && info.align !== 'center' ? 'right' : info.align,
    vertical: 'middle',
    wrapText: info.multiline || o.enclosed,
  };
  if (info.url) Object.assign(font, { color: { argb: 'FF0563C1' }, underline: true });
  xl.font = font;
}

// Dates written as a whole cell, kept looking the way they were printed.
// Each pattern captures year, month, day and gives the Excel format that prints it back.
const DATE_PATTERNS = [
  [/^(\d{4})-(\d{1,2})-(\d{1,2})$/, (m) => (m[2].length === 2 ? 'yyyy-mm-dd' : 'yyyy-m-d')],
  [/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/, (m) => (m[2].length === 2 ? 'yyyy/mm/dd' : 'yyyy/m/d')],
  // A bare "." in a number format is the decimal point, so dots are escaped to print as-is.
  [/^(\d{4})\.(\d{1,2})\.(\d{1,2})\.?$/, (m, s) => (m[2].length === 2 ? 'yyyy\\.mm\\.dd' : 'yyyy\\.m\\.d') + (s.endsWith('.') ? '\\.' : '')],
  [/^(\d{4})\. (\d{1,2})\. (\d{1,2})\.?$/, (m, s) => 'yyyy\\. m\\. d' + (s.endsWith('.') ? '\\.' : '')],
  [/^(\d{4})년 ?(\d{1,2})월 ?(\d{1,2})일$/, () => 'yyyy"년" m"월" d"일"'],
];

function parseDate(s) {
  for (const [re, format] of DATE_PATTERNS) {
    const m = s.match(re);
    if (!m) continue;
    const [y, mo, d] = [Number(m[1]), Number(m[2]), Number(m[3])];
    const date = new Date(Date.UTC(y, mo - 1, d));
    // Reject impossible dates such as 2026-02-30.
    if (y < 1900 || date.getUTCMonth() !== mo - 1 || date.getUTCDate() !== d) return null;
    return { value: date, numFmt: format(m, s) };
  }
  return null;
}

// Numbers as numbers: 1,234 / -1,234 / (1,234) / △1,234 / 12.5%. Codes like 010 or phone numbers stay text.
// Whole-cell dates become Excel dates.
export function parseValue(text) {
  const s = text.trim();
  const date = parseDate(s);
  if (date) return date;
  const pct = s.match(/^(-|△)?(\d+(?:\.\d+)?)%$/);
  if (pct) {
    const fmt = pct[2].includes('.') ? '0.0%' : '0%';
    const value = (Number(pct[2]) / 100) * (pct[1] ? -1 : 1);
    return { value, numFmt: pct[1] === '△' ? `${fmt};△${fmt}` : fmt };
  }
  // ▲ means "decrease" in some offices and "increase" in others, so it stays text.
  const m = s.match(/^(-|△|\()?\s*(\d{1,3}(?:,\d{3})+|\d+)(\.\d+)?(\))?\s*(원)?$/);
  if (!m) return { value: s };
  const [, sign, intPart, frac, close, unit] = m;
  if ((sign === '(') !== !!close) return { value: s };
  if (/^0\d/.test(intPart)) return { value: s };
  const digits = intPart.replace(/,/g, '');
  if (digits.length > 15) return { value: s };
  const negative = sign === '-' || sign === '△' || sign === '(';
  const n = Number(digits + (frac ?? '')) * (negative ? -1 : 1);
  const decimals = frac ? '.' + '0'.repeat(frac.length - 1) : '';
  // Keep the printed look: thousands separators, △ or (…) for negatives, a 원 suffix.
  const special = sign === '△' || sign === '(' || unit;
  if (!intPart.includes(',') && !special) return { value: n, numFmt: undefined };
  const pos = (intPart.includes(',') || special ? '#,##0' : '0') + decimals + (unit ? '"원"' : '');
  const numFmt = sign === '△' ? `${pos};△${pos}` : sign === '(' ? `${pos};(${pos})` : pos;
  return { value: n, numFmt };
}

function isDark(hex) {
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
  return 0.299 * r + 0.587 * g + 0.114 * b < 110;
}
