// Measure how closely a converted workbook matches the original Excel file the PDF was printed from.
// Both workbooks are SheetJS workbooks. For each original sheet it reports:
//  - exact:   original non-empty cells whose displayed text appears as one converted cell
//  - partial: text found only inside a larger converted cell or split over several
//  - ordered: share of matched pairs that keep their relative row/column order
//  - merges:  original merged blocks (with text) reproduced holding the same cells
//  - numbers: original numeric cells that came out numeric
import XLSX from 'xlsx';

const norm = (s) => String(s ?? '').replace(/\s+/g, '');

function cellsOf(ws) {
  const out = [];
  for (const [addr, cell] of Object.entries(ws)) {
    if (addr[0] === '!') continue;
    const text = cell.w ?? (cell.v != null ? String(cell.v) : '');
    if (!norm(text)) continue;
    const { r, c } = XLSX.utils.decode_cell(addr);
    out.push({ r, c, addr, text, key: norm(text), numeric: cell.t === 'n' });
  }
  return out;
}

const mergeAt = (ws) => new Map((ws['!merges'] ?? []).map((m) => [`${m.s.r},${m.s.c}`, m]));

export function compareWorkbooks(orig, conv, sheetNames = orig.SheetNames) {
  const convSheets = conv.SheetNames.map((name) => {
    const cells = cellsOf(conv.Sheets[name]);
    return { name, ws: conv.Sheets[name], cells, rows: Math.max(1, ...cells.map((c) => c.r + 1)) };
  });
  const pool = new Map();
  for (const s of convSheets) {
    for (const cell of s.cells) {
      if (!pool.has(cell.key)) pool.set(cell.key, []);
      pool.get(cell.key).push({ ...cell, sheet: s });
    }
  }
  const allConvText = convSheets.flatMap((s) => s.cells.map((c) => c.key));
  const rowTexts = convSheets.flatMap((s) => {
    const rows = new Map();
    for (const c of s.cells) {
      if (!rows.has(c.r)) rows.set(c.r, []);
      rows.get(c.r).push(c);
    }
    return [...rows.values()].map((row) => row.sort((a, b) => a.c - b.c).map((c) => c.key).join(''));
  });

  return sheetNames.map((name) => {
    const ws = orig.Sheets[name];
    const cells = cellsOf(ws);
    const used = new Set();
    const pairs = [];
    const missing = [];
    let partial = 0;
    const origRows = Math.max(1, ...cells.map((c) => c.r + 1));
    for (const cell of cells) {
      const candidates = (pool.get(cell.key) ?? []).filter((c) => !used.has(c));
      if (candidates.length) {
        // Repeated texts (e.g. the same note on many rows): take the copy at the most similar relative row.
        const cost = (h) => Math.abs(h.r / h.sheet.rows - cell.r / origRows) + Math.abs(h.c - cell.c) * 1e-3;
        const hit = candidates.reduce((best, h) => (cost(h) < cost(best) ? h : best));
        used.add(hit);
        pairs.push([cell, hit]);
      } else if (allConvText.some((t) => t.includes(cell.key)) || rowTexts.some((t) => t.includes(cell.key))) partial++;
      else missing.push(cell);
    }

    let ordered = 0;
    let comparisons = 0;
    for (let i = 0; i < pairs.length; i++) {
      for (let j = i + 1; j < pairs.length; j++) {
        const [a, ca] = pairs[i];
        const [b, cb] = pairs[j];
        if (ca.sheet !== cb.sheet) continue;
        comparisons++;
        if (Math.sign(a.r - b.r) === Math.sign(ca.r - cb.r) && Math.sign(a.c - b.c) === Math.sign(ca.c - cb.c)) ordered++;
      }
    }

    const om = mergeAt(ws);
    let merges = 0;
    let mergesKept = 0;
    let numbers = 0;
    let numbersKept = 0;
    for (const [cell, hit] of pairs) {
      const m = om.get(`${cell.r},${cell.c}`);
      if (m) {
        merges++;
        const cm = mergeAt(hit.sheet.ws).get(`${hit.r},${hit.c}`) ?? { s: { r: hit.r, c: hit.c }, e: { r: hit.r, c: hit.c } };
        if (sameShape(m, cm, pairs, hit.sheet)) mergesKept++;
      }
      if (cell.numeric) {
        numbers++;
        if (hit.numeric) numbersKept++;
      }
    }

    return {
      name,
      cells: cells.length,
      exact: pairs.length,
      partial,
      missing,
      ordered: comparisons ? ordered / comparisons : 1,
      merges,
      mergesKept,
      numbers,
      numbersKept,
    };
  });
}

// A merge is "the same shape" when it holds the same cells: every matched cell inside the
// original block is inside the converted one and vice versa. Row counts are ignored, since
// how many sheet rows a bordered block used is not visible in the PDF.
function sameShape(m, cm, pairs, sheet) {
  const inOrig = (cell) => cell.r >= m.s.r && cell.r <= m.e.r && cell.c >= m.s.c && cell.c <= m.e.c;
  const inConv = (cell) => cell.r >= cm.s.r && cell.r <= cm.e.r && cell.c >= cm.s.c && cell.c <= cm.e.c;
  if (m.e.c > m.s.c !== cm.e.c > cm.s.c) return false;
  return pairs.every(([o, h]) => h.sheet !== sheet || inOrig(o) === inConv(h) || o.c < m.s.c || o.c > m.e.c);
}
