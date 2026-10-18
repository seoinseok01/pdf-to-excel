import { cpSync } from 'node:fs';
import { createRequire } from 'node:module';
import path from 'node:path';
import { defineConfig } from 'vite';

// pdf.js needs its CMaps (Korean CID fonts) and standard fonts served next to the app.
const pdfjsDir = path.dirname(createRequire(import.meta.url).resolve('pdfjs-dist/package.json'));
for (const dir of ['cmaps', 'standard_fonts']) {
  cpSync(path.join(pdfjsDir, dir), path.join('public', 'pdfjs', dir), { recursive: true });
}

export default defineConfig({
  base: './',
  build: { outDir: 'dist', chunkSizeWarningLimit: 2000 },
});
