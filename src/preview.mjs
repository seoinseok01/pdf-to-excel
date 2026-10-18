// Render an ExcelJS worksheet as an HTML table (merges, borders, fills, widths, alignment).
// Used by the web UI preview and by scripts/preview.mjs.

const PX_PER_WIDTH = 5.6 * 1.33; // Excel width unit -> CSS px (via points)
const PX_PER_PT = 1.33;

const esc = (s) => String(s).replace(/[&<>"]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[ch]);
const BORDER_PX = { thin: 1, medium: 2, thick: 3, dotted: 1 };
function borderCss(b) {
  if (!b?.style) return 'none';
  const color = b.color?.argb ? '#' + b.color.argb.slice(2) : b.style === 'dotted' ? '#8a8a8a' : '#3a3a3a';
  return `${BORDER_PX[b.style] ?? 1}px ${b.style === 'dotted' ? 'dotted' : 'solid'} ${color}`;
}

// Cell content as HTML; rich text keeps per-run bold and colour.
function cellHtml(cell) {
  const rich = cell.value?.richText;
  if (!rich) return esc(displayText(cell));
  return rich
    .map((run) => {
      const css = [];
      if (run.font?.bold) css.push('font-weight:600');
      if (run.font?.color?.argb) css.push(`color:#${run.font.color.argb.slice(2)}`);
      return css.length ? `<span style="${css.join(';')}">${esc(run.text)}</span>` : esc(run.text);
    })
    .join('');
}

export function displayText(cell) {
  const v = cell.value;
  if (v == null) return '';
  if (v instanceof Date) return formatDate(v, cell.numFmt ?? 'yyyy-mm-dd');
  if (typeof v === 'object') return v.text ?? v.richText?.map((t) => t.text).join('') ?? '';
  if (typeof v === 'number' && cell.numFmt) return formatNumber(v, cell.numFmt);
  return String(v);
}

// maxRows limits how many rows are drawn: building a table of thousands of rows freezes
// the page for seconds, and the downloaded file has them all anyway.
export function renderSheet(ws, maxRows = Infinity) {
  const rowCount = Math.min(ws.rowCount, maxRows);
  const covered = new Set();
  const spans = new Map();
  for (const range of ws.model.merges ?? []) {
    const [a, b] = range.split(':').map((ref) => ws.getCell(ref));
    if (a.row > rowCount) continue;
    spans.set(`${a.row},${a.col}`, { rs: Math.min(b.row, rowCount) - a.row + 1, cs: b.col - a.col + 1 });
    for (let r = a.row; r <= Math.min(b.row, rowCount); r++) {
      for (let c = a.col; c <= b.col; c++) if (r !== a.row || c !== a.col) covered.add(`${r},${c}`);
    }
  }

  const widths = [];
  for (let c = 1; c <= ws.columnCount; c++) widths.push(Math.max(4, Math.round((ws.getColumn(c).width ?? 8) * PX_PER_WIDTH)));
  const total = widths.reduce((a, b) => a + b, 0);

  let html = `<table class="sheet" style="width:${total}px"><colgroup>`;
  for (const w of widths) html += `<col style="width:${w}px">`;
  html += '</colgroup><tbody>';
  for (let r = 1; r <= rowCount; r++) {
    const row = ws.getRow(r);
    html += `<tr style="height:${Math.round((row.height ?? 15) * PX_PER_PT)}px">`;
    for (let c = 1; c <= ws.columnCount; c++) {
      if (covered.has(`${r},${c}`)) continue;
      const cell = ws.getCell(r, c);
      const span = spans.get(`${r},${c}`);
      // A merged block's right and bottom borders live on its last cell.
      const last = span ? ws.getCell(r + span.rs - 1, c + span.cs - 1) : cell;
      const b = cell.border ?? {};
      const lb = last.border ?? {};
      const style = [
        `border-top:${borderCss(b.top)}`,
        `border-left:${borderCss(b.left)}`,
        `border-bottom:${borderCss(lb.bottom)}`,
        `border-right:${borderCss(lb.right)}`,
      ];
      if (cell.fill?.fgColor?.argb) style.push(`background:#${cell.fill.fgColor.argb.slice(2)}`);
      if (cell.font?.color?.argb) style.push(`color:#${cell.font.color.argb.slice(2)}`);
      if (cell.font?.bold) style.push('font-weight:600');
      if (cell.font?.underline) style.push('text-decoration:underline');
      if (cell.font?.size) style.push(`font-size:${(cell.font.size * PX_PER_PT).toFixed(1)}px`);
      if (cell.alignment?.horizontal) style.push(`text-align:${cell.alignment.horizontal}`);
      if (cell.alignment?.wrapText) style.push('white-space:pre-wrap');
      const cls = typeof cell.value === 'number' ? ' class="num"' : '';
      const spanAttr = span ? ` rowspan="${span.rs}" colspan="${span.cs}"` : '';
      html += `<td${spanAttr}${cls} style="${style.join(';')}">${cellHtml(cell)}</td>`;
    }
    html += '</tr>';
  }
  return html + '</tbody></table>';
}

// Enough of Excel's date formats for the ones the converter writes.
function formatDate(date, numFmt) {
  const parts = { yyyy: date.getUTCFullYear(), mm: String(date.getUTCMonth() + 1).padStart(2, '0'), dd: String(date.getUTCDate()).padStart(2, '0'), m: date.getUTCMonth() + 1, d: date.getUTCDate() };
  return numFmt.replace(/"([^"]*)"|\\(.)|yyyy|mm|dd|m|d/g, (tok, quoted, escaped) => quoted ?? escaped ?? String(parts[tok]));
}

// Enough of Excel's number formats for the ones the converter writes:
// "pos;neg" sections, #,##0 / 0, decimals, %, and quoted or escaped literals.
function formatNumber(v, numFmt) {
  const [pos, neg] = numFmt.split(';');
  const section = v < 0 && neg ? neg : pos;
  const abs = v < 0 && neg ? -v : v;
  const scaled = section.includes('%') ? abs * 100 : abs;
  const decimals = (section.match(/0\.(0+)/)?.[1] ?? '').length;
  const digits = scaled.toLocaleString('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals, useGrouping: section.includes(',') });
  let done = false;
  return section.replace(/"([^"]*)"|\\(.)|[#0,.]+/g, (tok, quoted, escaped) => {
    if (quoted !== undefined) return quoted;
    if (escaped !== undefined) return escaped;
    if (done) return '';
    done = true;
    return digits;
  });
}
