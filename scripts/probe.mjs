// Dump text items and ruling-line stats of a PDF, for debugging the extractor.
import { extractPage } from '../src/extract.mjs';
import { openPdf } from '../src/node.mjs';

const doc = await openPdf(process.argv[2]);
console.log('pages', doc.numPages);
const maxPages = Number(process.argv[3] ?? 2);
for (let p = 1; p <= Math.min(doc.numPages, maxPages); p++) {
  const page = await extractPage(doc, p);
  console.log(`\n=== page ${p} size ${page.width.toFixed(0)}x${page.height.toFixed(0)} texts ${page.texts.length} h-lines ${page.hLines.length} v-lines ${page.vLines.length} rects ${page.fills.length}`);
  for (const t of page.texts.slice(0, 40)) {
    console.log(`  [${t.x0.toFixed(1)},${t.y0.toFixed(1)} - ${t.x1.toFixed(1)},${t.y1.toFixed(1)}] ${JSON.stringify(t.str)}`);
  }
  console.log('  h-lines sample', page.hLines.slice(0, 5));
  console.log('  v-lines sample', page.vLines.slice(0, 5));
}
