// Extract positioned text and axis-aligned ruling lines from PDF pages with pdf.js.
// All output coordinates are in page space with the origin at the top-left (y grows downward).
import { getDocument, OPS, Util } from 'pdfjs-dist/legacy/build/pdf.mjs';

// Runs in the browser and in Node. assetBase is where pdf.js's cmaps/ and standard_fonts/
// folders are served from (needed for Korean CID fonts).
export function loadPdf(data, assetBase) {
  return getDocument({
    data,
    cMapUrl: assetBase + 'cmaps/',
    cMapPacked: true,
    standardFontDataUrl: assetBase + 'standard_fonts/',
    useSystemFonts: false,
    verbosity: 0,
  }).promise;
}

// Segments thinner than this are treated as lines even if drawn as filled rectangles.
const THIN = 2.5;

export async function extractPage(doc, pageNumber) {
  const page = await doc.getPage(pageNumber);
  const viewport = page.getViewport({ scale: 1 });
  const [texts, ops] = await Promise.all([page.getTextContent(), page.getOperatorList()]);

  const fontOf = fontLookup(page);
  const textItems = [];
  for (const item of texts.items) {
    if (!item.str || !item.str.trim()) continue;
    const font = fontOf(item.fontName);
    const tx = Util.transform(viewport.transform, item.transform);
    const fontHeight = Math.hypot(tx[2], tx[3]);
    const x0 = tx[4];
    const baseline = tx[5];
    textItems.push({
      str: fixBrokenGlyphs(item.str),
      x0,
      x1: x0 + item.width,
      y0: baseline - fontHeight * 0.85,
      y1: baseline + fontHeight * 0.15,
      baseline,
      size: fontHeight,
      bold: font.bold,
      font: font.family,
    });
  }

  const { hLines, vLines, fills, textClips } = collectGraphics(ops, viewport.transform);
  // Excel repeats a cell's text on every page its merged range touches and hides the
  // extra copies with clip paths; drop items that are mostly outside their clip.
  const visible = textItems.filter((t) => {
    // pdf.js may split one show-text op into several items: use the nearest op starting at or left of the item.
    let rec = null;
    for (const o of textClips) {
      if (Math.abs(o.y - t.baseline) < 1.5 && o.x <= t.x0 + 1.5 && (!rec || o.x > rec.x)) rec = o;
    }
    if (!rec) return true;
    t.color = rec.color;
    if (!rec.clip) return true;
    const w = Math.max(0, Math.min(t.x1, rec.clip.x1) - Math.max(t.x0, rec.clip.x0));
    const h = Math.max(0, Math.min(t.y1, rec.clip.y1) - Math.max(t.y0, rec.clip.y0));
    return w * h >= 0.5 * (t.x1 - t.x0) * (t.y1 - t.y0);
  });
  const links = [];
  for (const a of await page.getAnnotations()) {
    if (a.subtype !== 'Link' || !a.url) continue;
    const [x0, y0, x1, y1] = viewport.convertToViewportRectangle(a.rect);
    links.push({ url: a.url, x: (x0 + x1) / 2, y: (y0 + y1) / 2 });
  }
  return { width: viewport.width, height: viewport.height, texts: visible, hLines, vLines, fills, links };
}

function collectGraphics(opList, baseTransform) {
  const hLines = [];
  const vLines = [];
  const fills = [];
  let ctm = baseTransform.slice();
  let fillColor = '000000';
  let fillPattern = false;
  let strokeColor = '000000';
  let lineWidth = 1;
  let clip = null; // device-space bounding box of the current clip path, null = whole page
  let tlm = [1, 0, 0, 1, 0, 0];
  const textClips = [];
  const stack = [];
  let subpaths = []; // arrays of device-space points
  let current = null;

  const toDevice = (x, y) => Util.applyTransform([x, y], ctm);

  const addBox = (pts, stroked, filled) => {
    const xs = pts.map((p) => p[0]);
    const ys = pts.map((p) => p[1]);
    // Clip to the current clip path, so hidden decorations do not count as borders.
    const x0 = Math.max(Math.min(...xs), clip ? clip.x0 : -Infinity);
    const x1 = Math.min(Math.max(...xs), clip ? clip.x1 : Infinity);
    const y0 = Math.max(Math.min(...ys), clip ? clip.y0 : -Infinity);
    const y1 = Math.min(Math.max(...ys), clip ? clip.y1 : Infinity);
    if (x1 < x0 || y1 < y0) return;
    const w = x1 - x0;
    const h = y1 - y0;
    const style = filled && fillPattern ? 'dotted' : 'thin';
    // A thin filled box is a line as thick as the box; a stroke is as thick as the pen.
    const pen = lineWidth * Math.hypot(ctm[0], ctm[1]);
    const color = filled ? fillColor : strokeColor;
    if (h <= THIN && w > THIN) hLines.push({ y: (y0 + y1) / 2, x0, x1, style, width: filled ? h : pen, color });
    else if (w <= THIN && h > THIN) vLines.push({ x: (x0 + x1) / 2, y0, y1, style, width: filled ? w : pen, color });
    else if (w > THIN && h > THIN) {
      if (filled && !fillPattern) fills.push({ x0, y0, x1, y1, color: fillColor });
      if (stroked) {
        const line = { style: 'thin', width: pen, color: strokeColor };
        hLines.push({ ...line, y: y0, x0, x1 }, { ...line, y: y1, x0, x1 });
        vLines.push({ ...line, x: x0, y0, y1 }, { ...line, x: x1, y0, y1 });
      }
    }
  };

  const flush = (stroked, filled) => {
    for (const pts of subpaths) {
      if (pts.length < 2) continue;
      if (filled || pts.length > 2) {
        // A closed or filled shape: classify by its bounding box.
        if (filled) addBox(pts, stroked, filled);
        else for (let k = 1; k < pts.length; k++) addBox([pts[k - 1], pts[k]], true, false);
      } else addBox(pts, true, false);
    }
    subpaths = [];
    current = null;
  };

  const { fnArray, argsArray } = opList;
  for (let i = 0; i < fnArray.length; i++) {
    const fn = fnArray[i];
    const args = argsArray[i];
    switch (fn) {
      case OPS.save:
        stack.push({ ctm, fillColor, fillPattern, strokeColor, lineWidth, clip });
        break;
      case OPS.restore: {
        const st = stack.pop();
        if (st) ({ ctm, fillColor, fillPattern, strokeColor, lineWidth, clip } = st);
        break;
      }
      case OPS.setFillRGBColor:
        fillColor = toHex(args);
        fillPattern = false;
        break;
      case OPS.setStrokeRGBColor:
        strokeColor = toHex(args);
        break;
      case OPS.setLineWidth:
        lineWidth = args[0];
        break;
      case OPS.setFillColorN:
        fillPattern = args[0] === 'TilingPattern' || args[0] === 'Shading';
        break;
      case OPS.transform:
        ctm = Util.transform(ctm, args);
        break;
      case OPS.paintFormXObjectBegin:
        stack.push({ ctm, fillColor, fillPattern, strokeColor, lineWidth, clip });
        if (args[0]) ctm = Util.transform(ctm, args[0]);
        break;
      case OPS.paintFormXObjectEnd: {
        const st = stack.pop();
        if (st) ({ ctm, fillColor, fillPattern, strokeColor, lineWidth, clip } = st);
        break;
      }
      case OPS.constructPath: {
        const [subOps, coords] = args;
        let j = 0;
        for (const op of subOps) {
          if (op === OPS.moveTo) {
            current = [toDevice(coords[j], coords[j + 1])];
            subpaths.push(current);
            j += 2;
          } else if (op === OPS.lineTo) {
            if (!current) subpaths.push((current = []));
            current.push(toDevice(coords[j], coords[j + 1]));
            j += 2;
          } else if (op === OPS.rectangle) {
            const [x, y, w, h] = coords.slice(j, j + 4);
            subpaths.push([toDevice(x, y), toDevice(x + w, y), toDevice(x + w, y + h), toDevice(x, y + h), toDevice(x, y)]);
            current = null;
            j += 4;
          } else if (op === OPS.curveTo) {
            j += 6;
          } else if (op === OPS.curveTo2 || op === OPS.curveTo3) {
            j += 4;
          } else if (op === OPS.closePath) {
            if (current && current.length) current.push(current[0]);
          }
        }
        break;
      }
      case OPS.stroke:
      case OPS.closeStroke:
        flush(true, false);
        break;
      case OPS.fill:
      case OPS.eoFill:
        flush(false, true);
        break;
      case OPS.fillStroke:
      case OPS.eoFillStroke:
      case OPS.closeFillStroke:
      case OPS.closeEOFillStroke:
        flush(true, true);
        break;
      case OPS.clip:
      case OPS.eoClip: {
        const pts = subpaths.flat();
        if (pts.length) {
          const box = {
            x0: Math.min(...pts.map((p) => p[0])),
            y0: Math.min(...pts.map((p) => p[1])),
            x1: Math.max(...pts.map((p) => p[0])),
            y1: Math.max(...pts.map((p) => p[1])),
          };
          clip = clip
            ? { x0: Math.max(clip.x0, box.x0), y0: Math.max(clip.y0, box.y0), x1: Math.min(clip.x1, box.x1), y1: Math.min(clip.y1, box.y1) }
            : box;
        }
        break;
      }
      case OPS.beginText:
        tlm = [1, 0, 0, 1, 0, 0];
        break;
      case OPS.setTextMatrix:
        tlm = Array.from(args);
        break;
      case OPS.moveText:
      case OPS.setLeadingMoveText:
        tlm = Util.transform(tlm, [1, 0, 0, 1, args[0], args[1]]);
        break;
      case OPS.showText:
      case OPS.showSpacedText:
      case OPS.nextLineShowText:
      case OPS.nextLineSetSpacingShowText: {
        const [x, y] = Util.applyTransform([tlm[4], tlm[5]], ctm);
        textClips.push({ x, y, clip, color: fillColor });
        break;
      }
      case OPS.endPath:
        subpaths = [];
        current = null;
        break;
    }
  }
  return { hLines, vLines, fills, textClips };
}

function toHex(args) {
  if (typeof args[0] === 'string') return args[0].replace('#', '').toUpperCase();
  return [args[0], args[1], args[2]].map((v) => Math.round(v).toString(16).padStart(2, '0')).join('').toUpperCase();
}

// Some fonts in Excel exports map ASCII punctuation to CJK compatibility squares in their
// ToUnicode table (seen: "(주)" -> "㏙주㏚", "-" -> "㎿", "~" -> "㚕"). Map the known ones back.
const BROKEN_GLYPHS = { '㏙': '(', '㏚': ')', '㎿': '-', '㚕': '~' };
function fixBrokenGlyphs(str) {
  return str.replace(/[㏙㏚㎿㚕]/g, (ch) => BROKEN_GLYPHS[ch]);
}

// Bold flag and family of a pdf.js font, from its PostScript name ("AAAAAE+MalgunGothicBold").
function fontLookup(page) {
  const cache = new Map();
  return (fontName) => {
    if (!cache.has(fontName)) {
      let name = '';
      try {
        name = page.commonObjs.get(fontName)?.name ?? '';
      } catch {
        // Font not loaded (e.g. a standard font): fall back to plain.
      }
      const base = name.replace(/^[A-Z]{6}\+/, '');
      cache.set(fontName, {
        bold: /bold|black|heavy|semibold|demi/i.test(base),
        family: familyName(base),
      });
    }
    return cache.get(fontName);
  };
}

const FAMILIES = [
  [/malgun/i, '맑은 고딕'],
  [/gulim/i, '굴림'],
  [/dotum/i, '돋움'],
  [/batang/i, '바탕'],
  [/gungsuh/i, '궁서'],
  [/nanumgothic/i, '나눔고딕'],
  [/nanummyeongjo/i, '나눔명조'],
  [/arial/i, 'Arial'],
  [/calibri/i, 'Calibri'],
  [/times/i, 'Times New Roman'],
];

function familyName(base) {
  for (const [re, family] of FAMILIES) if (re.test(base)) return family;
  return null;
}
