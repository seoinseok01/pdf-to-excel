// Node entry points: open PDFs from disk with pdf.js assets from node_modules.
import { readFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import path from 'node:path';
import { loadPdf } from './extract.mjs';
import { convertDocument } from './convert.mjs';

const pdfjsDir = path.dirname(createRequire(import.meta.url).resolve('pdfjs-dist/package.json'));

export async function openPdf(filePath) {
  return loadPdf(new Uint8Array(await readFile(filePath)), pdfjsDir + path.sep);
}

// Returns { grids, groups, workbook }; see convertDocument.
export async function convertPdf(filePath, options) {
  return convertDocument(await openPdf(filePath), undefined, options);
}
