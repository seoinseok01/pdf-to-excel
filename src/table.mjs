// Rebuild an Excel-like grid from extracted pages.
//
// A PDF printed from Excel draws every border on a column or row edge of the
// original sheet, so the distinct x positions of vertical lines are the sheet's
// column edges and the y positions of horizontal lines are its row edges.
// Cells whose shared edge has no line are merged, but only when the merged
// block is fully enclosed by lines, so free-floating text (titles, notes)
// is not swallowed into one giant cell.
//
// Pages are first turned into raw grids (texts placed in single cells), then
// stitched together by the caller, and only then are merges resolved, so a
// merged cell broken by a page break comes back as one cell.

const SNAP = 1.5; // positions closer than this are the same edge
const GAP = 3; // collinear segments closer than this are joined
const COVER = 0.6; // fraction of an edge that must be drawn to count as a border

// options.splitLines: give every text line inside a bordered cell its own row, the way
// many forms are typed (one sheet row per line). Off, a multi-line cell stays one cell.
export function buildPageGrid(page, options = {}) {
  // Edges of cell background fills separate cells too (colored label columns often have no lines).
  const pageArea = page.width * page.height;
  const fills = page.fills.filter((f) => f.color !== 'FFFFFF' && (f.x1 - f.x0) * (f.y1 - f.y0) < pageArea * 0.3);
  const hEdges = fills.flatMap((f) => [f.y0, f.y1].map((y) => ({ y, x0: f.x0, x1: f.x1, style: 'edge' })));
  const vEdges = fills.flatMap((f) => [f.x0, f.x1].map((x) => ({ x, y0: f.y0, y1: f.y1, style: 'edge' })));
  const lines = weighLines([...page.hLines, ...page.vLines]);
  const h = groupLines([...page.hLines.map(lines), ...hEdges], 'y', 'x0', 'x1');
  const v = groupLines([...page.vLines.map(lines), ...vEdges], 'x', 'y0', 'y1');
  const texts = joinRuns(page.texts.filter((t) => !isPageNumber(t, page)), v);

  const xs = v.map((g) => g.pos);
  extendToText(xs, texts, 'x0', 'x1');
  // A horizontal line that borders no whole cell (e.g. a hyperlink underline) is not a row edge.
  const rowLines = h.filter((g) => xs.some((x, c) => c < xs.length - 1 && coverage(g, x, xs[c + 1])));
  const ys = rowLines.map((g) => g.pos);
  extendToText(ys, texts, 'y0', 'y1');

  let grid = makeGrid(xs, ys, h, v);
  // Split bands that hold several lines of free-floating text into separate rows.
  const extra = virtualRows(grid, texts, options.splitLines);
  if (extra.length) {
    // Rows split inside a bordered cell get an undrawn separator, so each line becomes
    // its own (horizontally merged) cell instead of merging back into one.
    const allYs = [...ys];
    const hAll = [...h];
    for (const e of extra) {
      let y = allYs.find((p) => Math.abs(p - e.y) < SNAP);
      if (y === undefined) allYs.push((y = e.y));
      if (!e.inside) continue;
      let g = hAll.find((q) => q.pos === y);
      if (!g) hAll.push((g = { pos: y, spans: [] }));
      else if (h.includes(g)) hAll[hAll.indexOf(g)] = g = { pos: g.pos, spans: [...g.spans] };
      g.spans.push({ a: e.x0, b: e.x1, style: 'edge' });
      g.spans.sort((p, q) => p.a - q.a);
    }
    grid = makeGrid(xs, allYs.sort((a, b) => a - b), hAll, v);
  }

  placeTexts(grid, texts);
  placeFills(grid, page.fills, page);
  for (const link of page.links ?? []) {
    const r = rowAt(grid.ys, link.y);
    const c = rowAt(grid.xs, link.x);
    if (r >= 0 && c >= 0) grid.cells[r][c].url = link.url;
  }
  grid.heights = grid.ys.slice(1).map((y, r) => y - grid.ys[r]);
  return grid;
}

function isPageNumber(t, page) {
  return t.y0 > page.height * 0.92 && /^\d+\s*\/\s*\d+$/.test(t.str.trim());
}

// Cluster segments by their fixed coordinate and join overlapping spans.
function groupLines(lines, posKey, aKey, bKey) {
  const sorted = [...lines].sort((p, q) => p[posKey] - q[posKey]);
  const groups = [];
  for (const line of sorted) {
    const last = groups[groups.length - 1];
    if (last && line[posKey] - last.maxPos <= SNAP) {
      last.items.push(line);
      last.maxPos = line[posKey];
    } else groups.push({ items: [line], maxPos: line[posKey] });
  }
  return groups.map((g) => {
    const pos = g.items.reduce((s, l) => s + l[posKey], 0) / g.items.length;
    const spans = g.items.map((l) => ({ a: l[aKey], b: l[bKey], style: l.style, color: l.color })).sort((p, q) => p.a - q.a);
    const merged = [];
    for (const s of spans) {
      const last = merged[merged.length - 1];
      if (last && s.a <= last.b + GAP) {
        last.b = Math.max(last.b, s.b);
        if (stronger(last.style, s.style) !== last.style) Object.assign(last, { style: s.style, color: s.color });
      } else merged.push({ ...s });
    }
    return { pos, spans: merged };
  });
}

function extendToText(edges, texts, aKey, bKey) {
  if (!texts.length) return;
  const lo = Math.min(...texts.map((t) => t[aKey])) - 1;
  const hi = Math.max(...texts.map((t) => t[bKey])) + 1;
  if (!edges.length) edges.push(lo, hi);
  if (lo < edges[0] - SNAP) edges.unshift(lo);
  if (hi > edges[edges.length - 1] + SNAP) edges.push(hi);
}

// Line styles from weakest to strongest; 'edge' separates cells but is not drawn.
const RANK = { edge: 0, dotted: 1, thin: 2, medium: 3, thick: 4 };
const stronger = (a, b) => (RANK[b] > RANK[a] ? b : a);

// Excel draws thin, medium and thick borders at about 1, 2 and 3 times the same width,
// and the PDF is usually scaled, so thickness is judged against the page's commonest line.
function weighLines(lines) {
  const widths = lines.filter((l) => l.style === 'thin' && l.width > 0).map((l) => Math.round(l.width * 20) / 20);
  const counts = new Map();
  for (const w of widths) counts.set(w, (counts.get(w) ?? 0) + 1);
  const base = [...counts].sort((a, b) => b[1] - a[1] || a[0] - b[0])[0]?.[0];
  return (l) => {
    if (l.style !== 'thin' || !base || !l.width) return l;
    const ratio = l.width / base;
    return { ...l, style: ratio >= 2.6 ? 'thick' : ratio >= 1.7 ? 'medium' : 'thin' };
  };
}

// How a segment [a, b] on a line group is drawn: { style, color } or null.
// style is 'thin', 'medium', 'thick', 'dotted' or 'edge' (separates cells, not drawn).
function coverage(group, a, b) {
  if (!group) return null;
  let covered = 0;
  let best = { style: 'edge' };
  for (const s of group.spans) {
    const o = Math.min(b, s.b) - Math.max(a, s.a);
    if (o > 0) {
      covered += o;
      if (stronger(best.style, s.style) !== best.style) best = { style: s.style, color: s.color };
    }
  }
  return covered >= COVER * (b - a) ? best : null;
}

function rowAt(edges, p) {
  for (let i = 0; i < edges.length - 1; i++) if (p >= edges[i] - 0.01 && p < edges[i + 1]) return i;
  return -1;
}

function makeGrid(xs, ys, h, v) {
  const nr = ys.length - 1;
  const nc = xs.length - 1;
  const hAt = new Map(h.map((g) => [g.pos, g]));
  const vAt = new Map(v.map((g) => [g.pos, g]));
  // top[r][c]: border above row r (r = 0..nr). left[r][c]: border left of column c (c = 0..nc).
  const top = [];
  for (let r = 0; r <= nr; r++) {
    const g = hAt.get(ys[r]);
    top.push(Array.from({ length: nc }, (_, c) => coverage(g, xs[c], xs[c + 1])));
  }
  const left = [];
  for (let r = 0; r < nr; r++) {
    left.push(Array.from({ length: nc + 1 }, (_, c) => coverage(vAt.get(xs[c]), ys[r], ys[r + 1])));
  }
  const cells = [];
  for (let r = 0; r < nr; r++) cells.push(Array.from({ length: nc }, () => ({ texts: [] })));
  return { xs, ys, top, left, cells };
}

// Resolve merged blocks from borders. Returns owner[r][c] = { r0, r1, c0, c1, enclosed } and the merges.
export function resolveBlocks(top, left) {
  const nr = left.length;
  const nc = nr ? left[0].length - 1 : 0;
  const parent = Array.from({ length: nr * nc }, (_, k) => k);
  const find = (k) => {
    while (parent[k] !== k) k = parent[k] = parent[parent[k]];
    return k;
  };
  const union = (a, b) => {
    parent[find(b)] = find(a);
  };
  for (let r = 0; r < nr; r++) {
    for (let c = 0; c < nc; c++) {
      if (c + 1 < nc && !left[r][c + 1]) union(r * nc + c, r * nc + c + 1);
      if (r + 1 < nr && !top[r + 1][c]) union(r * nc + c, (r + 1) * nc + c);
    }
  }
  const comps = new Map();
  for (let k = 0; k < nr * nc; k++) {
    const root = find(k);
    if (!comps.has(root)) comps.set(root, []);
    comps.get(root).push(k);
  }

  const owner = Array.from({ length: nr }, () => new Array(nc));
  const merges = [];
  for (const members of comps.values()) {
    const rs = members.map((k) => Math.floor(k / nc));
    const cs = members.map((k) => k % nc);
    const r0 = Math.min(...rs);
    const r1 = Math.max(...rs);
    const c0 = Math.min(...cs);
    const c1 = Math.max(...cs);
    let enclosed = members.length === (r1 - r0 + 1) * (c1 - c0 + 1);
    for (let c = c0; enclosed && c <= c1; c++) enclosed = !!top[r0][c] && !!top[r1 + 1][c];
    for (let r = r0; enclosed && r <= r1; r++) enclosed = !!left[r][c0] && !!left[r][c1 + 1];
    if (enclosed) {
      const block = { r0, r1, c0, c1, enclosed: true };
      members.forEach((k, i) => (owner[rs[i]][cs[i]] = block));
      if (members.length > 1) merges.push(block);
    } else {
      members.forEach((k, i) => (owner[rs[i]][cs[i]] = { r0: rs[i], r1: rs[i], c0: cs[i], c1: cs[i], enclosed: false }));
    }
  }
  return { owner, merges };
}

// Text lines that sit in unenclosed space and share a grid row get their own rows.
// Returns { y, inside, x0, x1 } cut positions between text lines: always for text in open
// space, and with splitLines also inside bordered blocks (inside = true, x0..x1 = the block).
function virtualRows(grid, texts, splitLines) {
  const { owner } = resolveBlocks(grid.top, grid.left);
  const groups = new Map();
  for (const t of texts) {
    const r = rowAt(grid.ys, (t.y0 + t.y1) / 2);
    const c = rowAt(grid.xs, t.x0 + 0.5);
    if (r < 0 || c < 0) continue;
    const o = owner[r][c];
    if (o.enclosed && !splitLines) continue;
    const key = o.enclosed ? `block ${o.r0},${o.c0}` : `row ${r}`;
    if (!groups.has(key)) groups.set(key, { o, items: [] });
    groups.get(key).items.push(t);
  }
  const extra = [];
  for (const { o, items } of groups.values()) {
    const lines = textLines(items);
    for (let i = 1; i < lines.length; i++) {
      const gapTop = Math.max(...lines[i - 1].map((t) => t.y1));
      const gapBottom = Math.min(...lines[i].map((t) => t.y0));
      extra.push({ y: (gapTop + gapBottom) / 2, inside: o.enclosed, x0: grid.xs[o.c0], x1: grid.xs[o.c1 + 1] });
    }
  }
  return extra;
}

// Join items on one line that pdf.js split apart (separate words, font changes)
// unless a vertical border runs between them.
function joinRuns(texts, v) {
  const lines = textLines(texts);
  const out = [];
  for (const line of lines) {
    let run = null;
    for (const t of line) {
      const gap = run ? t.x0 - run.x1 : Infinity;
      const mid = (t.y0 + t.y1) / 2;
      const blocked = run && v.some((g) => g.pos > run.x1 - 1 && g.pos < t.x0 + 1 && g.spans.some((s) => s.a <= mid && s.b >= mid));
      if (run && gap < t.size * 0.8 && !blocked) {
        run.str += (gap > t.size * 0.2 ? ' ' : '') + t.str;
        run.x1 = Math.max(run.x1, t.x1);
        run.y0 = Math.min(run.y0, t.y0);
        run.y1 = Math.max(run.y1, t.y1);
        // A run is bold only if all of it is; its font is the one most of its text uses.
        run.bold = run.bold && t.bold;
        if (t.str.length > run.mainLength) Object.assign(run, { font: t.font ?? run.font, mainLength: t.str.length });
      } else {
        run = { ...t, mainLength: t.str.length };
        out.push(run);
      }
    }
  }
  return out;
}

// Group items into visual lines by vertical overlap, top to bottom, each sorted left to right.
function textLines(items) {
  const sorted = [...items].sort((a, b) => a.y0 + a.y1 - (b.y0 + b.y1));
  const lines = [];
  for (const t of sorted) {
    const mid = (t.y0 + t.y1) / 2;
    const line = lines[lines.length - 1];
    if (line && mid - line.mid < Math.max(t.size, line.size) * 0.5) line.items.push(t);
    else lines.push({ mid, size: t.size, items: [t] });
  }
  return lines.map((l) => l.items.sort((a, b) => a.x0 - b.x0));
}

const charCount = (str) => str.replace(/\s+/g, '').length;

function placeTexts(grid, texts) {
  // Characters that could not be placed in any cell count against text preservation.
  grid.textChars = texts.reduce((n, t) => n + charCount(t.str), 0);
  grid.lostChars = 0;
  for (const t of texts) {
    const r = rowAt(grid.ys, (t.y0 + t.y1) / 2);
    if (r < 0) {
      grid.lostChars += charCount(t.str);
      continue;
    }
    // The column holding most of the item; overflowing text belongs where it starts.
    let best = -1;
    let bestOverlap = 0;
    for (let c = 0; c < grid.xs.length - 1; c++) {
      const o = Math.min(t.x1, grid.xs[c + 1]) - Math.max(t.x0, grid.xs[c]);
      if (o > bestOverlap) {
        bestOverlap = o;
        best = c;
      }
    }
    if (best < 0 || t.x1 - t.x0 > (grid.xs[best + 1] - grid.xs[best]) * 1.5) best = rowAt(grid.xs, t.x0 + 0.5);
    if (best < 0) {
      grid.lostChars += charCount(t.str);
      continue;
    }
    grid.cells[r][best].texts.push(t);
  }
}

function placeFills(grid, fills, page) {
  const pageArea = page.width * page.height;
  grid.cells.forEach((row, r) =>
    row.forEach((cell, c) => {
      const cx = (grid.xs[c] + grid.xs[c + 1]) / 2;
      const cy = (grid.ys[r] + grid.ys[r + 1]) / 2;
      let best = null;
      for (const f of fills) {
        const area = (f.x1 - f.x0) * (f.y1 - f.y0);
        if (area > pageArea * 0.3 || f.color === 'FFFFFF') continue;
        if (cx < f.x0 || cx > f.x1 || cy < f.y0 || cy > f.y1) continue;
        if (!best || area < best.area) best = { area, color: f.color };
      }
      if (best) cell.fill = best.color;
    }),
  );
}

// Text of a single raw cell, used to recognise repeated header rows.
export function cellText(cell) {
  return textLines(cell.texts).map(joinLine).join('\n');
}

// Resolve merges on a (possibly stitched) grid and compute each block's value and look.
export function finalizeGrid(grid) {
  const { owner, merges } = resolveBlocks(grid.top, grid.left);
  grid.owner = owner;
  grid.merges = merges;
  grid.values = new Map(); // "r,c" of a block's top-left cell -> { value, align, size, multiline, fill, url }
  const blocks = new Map();
  owner.forEach((row, r) =>
    row.forEach((o, c) => {
      const key = `${o.r0},${o.c0}`;
      if (!blocks.has(key)) blocks.set(key, { o, texts: [], fill: null, url: null });
      const b = blocks.get(key);
      const cell = grid.cells[r][c];
      b.texts.push(...cell.texts.map((t) => ({ ...t, mid: r * 1e4 + (t.y0 + t.y1) / 2 })));
      b.fill ??= cell.fill;
      b.url ??= cell.url;
    }),
  );
  // How the text sits in the result, for the accuracy estimate (see estimateAccuracy).
  grid.stats = { chars: 0, enclosedChars: 0, multiLineChars: 0, openChars: 0, scatteredChars: 0 };
  const openPerRow = new Map(); // row -> characters of each unbordered text cell on it
  for (const [key, b] of blocks) {
    const info = { fill: b.fill, url: b.url };
    if (b.texts.length) {
      // Texts from stitched pages carry a row-ordered key so they stay in reading order.
      const ordered = b.texts.map((t) => ({ ...t, y0: t.mid - (t.y1 - t.y0) / 2, y1: t.mid + (t.y1 - t.y0) / 2 }));
      const lines = textLines(ordered);
      info.value = lines.map(joinLine).join('\n');
      // Per-line look, so a bold title line above plain detail lines can stay bold alone.
      info.lines = lines.map((line) => ({
        text: joinLine(line),
        bold: line.every((t) => t.bold),
        color: line[0].color,
      }));
      info.font = b.texts.find((t) => t.font)?.font;
      const chars = b.texts.reduce((n, t) => n + charCount(t.str), 0);
      grid.stats.chars += chars;
      if (!b.o.enclosed) {
        grid.stats.openChars += chars;
        if (!openPerRow.has(b.o.r0)) openPerRow.set(b.o.r0, []);
        openPerRow.get(b.o.r0).push(chars);
      } else {
        grid.stats.enclosedChars += chars;
        if (lines.length > 1) grid.stats.multiLineChars += chars;
      }
      const x0 = grid.xs[b.o.c0];
      const x1 = grid.xs[b.o.c1 + 1];
      // Every line must agree: a paragraph whose block happens to be balanced is still left-aligned.
      const lineAlign = (line) => {
        const leftGap = line[0].x0 - x0;
        const rightGap = x1 - line[line.length - 1].x1;
        if (leftGap > 2 && rightGap > 2 && Math.abs(leftGap - rightGap) < Math.max(3, (x1 - x0) * 0.12)) return 'center';
        if (rightGap < 3 && leftGap > 6) return 'right';
        return 'left';
      };
      const aligns = new Set(lines.map(lineAlign));
      info.align = b.o.enclosed && aligns.size === 1 ? [...aligns][0] : 'left';
      info.size = Math.max(...b.texts.map((t) => t.size));
      info.multiline = lines.length > 1;
    }
    grid.values.set(key, info);
  }
  // Free text broken into three or more pieces on one row (a letter-spaced title, a date
  // and signature line, checkbox labels) is where the layout is hardest to recover.
  for (const pieces of openPerRow.values()) {
    if (pieces.length >= 3) grid.stats.scatteredChars += pieces.reduce((a, b) => a + b, 0);
  }
  return grid;
}

function joinLine(items) {
  let out = '';
  let prev = null;
  for (const t of items) {
    if (prev && t.x0 - prev.x1 > t.size * 0.2 && !out.endsWith(' ') && !t.str.startsWith(' ')) out += ' ';
    out += t.str;
    prev = t;
  }
  return out.replace(/\s+/g, ' ').trim();
}
