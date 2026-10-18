// Compare cell formatting (bold, text colour, borders) of a converted workbook with the
// original .xlsx. Cells are paired by identical text (whitespace ignored, first unused
// match in reading order); merged-cell copies are skipped. Both are ExcelJS workbooks.

const norm = (s) => String(s ?? '').replace(/\s+/g, '');
const textOf = (v) => (v?.richText ? v.richText.map((r) => r.text).join('') : v?.text ?? v);

// 'bold', 'plain' or 'mixed' (rich text with both).
function boldOf(cell) {
  const runs = cell.value?.richText;
  if (!runs) return cell.font?.bold ? 'bold' : 'plain';
  const flags = runs.filter((r) => norm(r.text)).map((r) => !!(r.font?.bold ?? cell.font?.bold));
  return flags.every(Boolean) ? 'bold' : flags.some(Boolean) ? 'mixed' : 'plain';
}

// Explicit non-black RGB colours count as coloured; theme colours are treated as default.
function colored(cell) {
  const argb = cell.font?.color?.argb;
  return !!argb && !/^FF(000000|222222)$/i.test(argb);
}

function cells(ws) {
  const out = [];
  ws.eachRow((row) =>
    row.eachCell((cell) => {
      if (cell.isMerged && cell.master !== cell) return;
      const key = norm(textOf(cell.value));
      if (key) out.push({ cell, key });
    }),
  );
  return out;
}

export function compareFormatting(orig, conv) {
  const pool = new Map();
  for (const ws of conv.worksheets) {
    for (const c of cells(ws)) {
      if (!pool.has(c.key)) pool.set(c.key, []);
      pool.get(c.key).push(c.cell);
    }
  }
  const score = { pairs: 0, bold: 0, color: 0, borderSides: 0, borderSidesKept: 0, borderStyleKept: 0 };
  for (const ws of orig.worksheets) {
    for (const { cell, key } of cells(ws)) {
      const hit = pool.get(key)?.shift();
      if (!hit) continue;
      score.pairs++;
      if (boldOf(cell) === boldOf(hit)) score.bold++;
      if (colored(cell) === colored(hit)) score.color++;
      for (const side of ['top', 'left', 'bottom', 'right']) {
        const a = cell.border?.[side]?.style;
        if (!a) continue;
        score.borderSides++;
        const b = hit.border?.[side]?.style;
        if (b) score.borderSidesKept++;
        if (b === a) score.borderStyleKept++;
      }
    }
  }
  return score;
}
