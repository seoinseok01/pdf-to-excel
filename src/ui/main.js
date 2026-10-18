// Web UI: converts PDFs entirely in the browser and offers the .xlsx for download.
import { GlobalWorkerOptions } from 'pdfjs-dist/legacy/build/pdf.mjs';
import workerUrl from 'pdfjs-dist/legacy/build/pdf.worker.mjs?url';
import { loadPdf } from '../extract.mjs';
import { autoGroups, buildWorkbook, convertDocument, gridsFromPages } from '../convert.mjs';
import { renderSheet } from '../preview.mjs';

GlobalWorkerOptions.workerSrc = workerUrl;
const assetBase = new URL('pdfjs/', document.baseURI).href;

const $ = (id) => document.getElementById(id);
const drop = $('drop');
const input = $('file');
const list = $('files');
const preview = $('preview');

const jobs = [];
let shown = null;

input.addEventListener('change', () => {
  addFiles(input.files);
  input.value = '';
});
drop.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' || e.key === ' ') {
    e.preventDefault();
    input.click();
  }
});
for (const type of ['dragenter', 'dragover']) {
  drop.addEventListener(type, (e) => {
    e.preventDefault();
    drop.classList.add('over');
  });
}
for (const type of ['dragleave', 'drop']) drop.addEventListener(type, () => drop.classList.remove('over'));
drop.addEventListener('drop', (e) => {
  e.preventDefault();
  addFiles(e.dataTransfer.files);
});
// Dropping a file outside the zone should not navigate away from the page.
window.addEventListener('dragover', (e) => e.preventDefault());
window.addEventListener('drop', (e) => e.preventDefault());

$('preview-download').addEventListener('click', () => shown && download(shown));

function addFiles(fileList) {
  for (const file of fileList) {
    const isPdf = file.type === 'application/pdf' || /\.pdf$/i.test(file.name);
    const job = { file, name: outputName(file.name), status: isPdf ? 'queued' : 'error', error: isPdf ? null : 'PDF 파일이 아니에요.' };
    jobs.push(job);
    job.el = document.createElement('li');
    list.append(job.el);
    renderJob(job);
    if (isPdf) run(job);
  }
}

// "보고서.xlsx.pdf" -> "보고서.xlsx", "보고서.pdf" -> "보고서.xlsx"
function outputName(name) {
  const base = name.replace(/\.pdf$/i, '').replace(/\.(xlsx|xlsm|xls)$/i, '');
  return base + '.xlsx';
}

async function run(job) {
  job.status = 'working';
  job.progress = [0, 0];
  renderJob(job);
  try {
    const data = new Uint8Array(await job.file.arrayBuffer());
    const doc = await loadPdf(data, assetBase);
    job.progress = [0, doc.numPages];
    renderJob(job);
    const result = await convertDocument(doc, (page, total) => {
      job.progress = [page, total];
      renderJob(job);
    });
    await doc.destroy();
    Object.assign(job, result, { history: [], selected: new Set(), options: {} });
    if (!job.workbook.worksheets.length) throw new Error('NO_TEXT');
    job.status = 'done';
    if (!shown) showPreview(job);
  } catch (err) {
    console.error(err);
    job.status = 'error';
    job.error =
      err.message === 'NO_TEXT'
        ? '글자를 찾지 못했어요. 스캔한 PDF는 아직 변환할 수 없어요.'
        : err.name === 'PasswordException'
          ? '암호가 걸린 PDF예요. 암호를 해제한 뒤 다시 올려 주세요.'
          : '변환하지 못했어요. 엑셀에서 만든 PDF가 맞는지 확인해 주세요.';
  }
  renderJob(job);
}

function summary(workbook) {
  const sheets = workbook.worksheets;
  const first = sheets[0];
  const size = `${first.rowCount}행 × ${first.columnCount}열`;
  const shape = sheets.length > 1 ? `시트 ${sheets.length}개 · 첫 시트 ${size}` : size;
  return `${shape} · 예상 정확도 ${workbook.quality.label}`;
}

const percent = (v) => `${(Math.floor(v * 1000) / 10).toFixed(1).replace(/\.0$/, '')}%`;

function renderQuality(job) {
  const q = job.workbook.quality;
  const el = $('quality');
  el.className = `quality ${q.level}`;
  el.innerHTML = `
    <div class="quality-main">
      <div class="quality-figure"><span class="quality-label">예상 정확도</span><strong></strong></div>
      <div class="quality-figure"><span class="quality-label">글자 보존</span><strong></strong></div>
    </div>
    <ul class="quality-notes"></ul>
    <details>
      <summary>예상 정확도는 어떻게 정하나요?</summary>
      <p>원본 엑셀 파일이 없으면 정확도를 직접 잴 수 없어서, 변환 결과가 어떤 모양인지를 보고 비슷한 시험 문서에서 실제로 잰 정확도를 보여 드려요.</p>
      <ul>
        <li><b>99% 이상</b>: 글자 대부분이 테두리 칸 안에 있는 표. 시험한 문서 9개에서 99.8~100%였어요.</li>
        <li><b>약 70~85%</b>: 테두리 밖에 글자가 많거나 흩어진 양식(지원서, 동의서 등). 시험한 양식에서 69~79%였어요. 다운로드한 뒤 내용을 한 번 확인해 주세요.</li>
      </ul>
      <p><b>글자 보존</b>은 PDF에 있는 글자 중 엑셀 칸에 옮겨진 비율로, 실제로 센 값이에요.</p>
    </details>`;
  const [level, kept] = el.querySelectorAll('.quality-figure strong');
  level.textContent = q.label;
  kept.textContent = percent(q.preserved);
  const notes = el.querySelector('.quality-notes');
  if (!q.notes.length) notes.remove();
  for (const n of q.notes) {
    const li = document.createElement('li');
    li.textContent = n;
    notes.append(li);
  }
}

function renderJob(job) {
  const el = job.el;
  el.className = `file ${job.status}`;
  el.classList.toggle('active', shown === job);
  let status = '';
  if (job.status === 'queued') status = '대기 중';
  else if (job.status === 'working') {
    const [page, total] = job.progress;
    status = total ? `변환 중 · ${page}/${total}쪽` : '여는 중';
  } else if (job.status === 'done') status = summary(job.workbook);
  else status = job.error;

  const pct = job.status === 'working' && job.progress[1] ? (100 * job.progress[0]) / job.progress[1] : job.status === 'done' ? 100 : 0;
  el.innerHTML = `
    <div class="file-main">
      <div class="file-name"></div>
      <div class="file-status"></div>
      <div class="bar"><span style="width:${pct}%"></span></div>
    </div>
    <div class="file-actions"></div>`;
  el.querySelector('.file-name').textContent = job.file.name;
  el.querySelector('.file-status').textContent = status;
  if (job.status === 'done') {
    const actions = el.querySelector('.file-actions');
    const view = button('미리보기', 'btn', () => showPreview(job));
    const save = button('다운로드', 'btn primary', () => download(job));
    actions.append(view, save);
  }
}

function button(label, cls, onClick) {
  const b = document.createElement('button');
  b.type = 'button';
  b.className = cls;
  b.textContent = label;
  b.addEventListener('click', onClick);
  return b;
}

// ---- Sheet merging and splitting ----

// "PDF 1~7쪽" for a run of pages, "PDF 1·3·4쪽" otherwise.
function pageLabel(job, group) {
  const pages = group.map((k) => job.grids[k].page);
  const run = pages.every((p, i) => i === 0 || p === pages[i - 1] + 1);
  return `PDF ${run && pages.length > 2 ? `${pages[0]}~${pages[pages.length - 1]}` : pages.join('·')}쪽`;
}

function applyGroups(job, groups, sheetIndex = 0) {
  job.history.push(job.groups);
  job.groups = groups;
  job.workbook = buildWorkbook(job.grids, groups);
  job.selected = new Set();
  renderJob(job);
  showPreview(job, Math.min(sheetIndex, groups.length - 1));
}

function mergeSelected(job) {
  const picked = [...job.selected].sort((a, b) => a - b);
  if (picked.length < 2) return;
  const merged = picked.flatMap((i) => job.groups[i]);
  const groups = job.groups.filter((_, i) => !job.selected.has(i));
  groups.splice(picked[0], 0, merged);
  applyGroups(job, groups, picked[0]);
}

function splitSelected(job) {
  const picked = [...job.selected].sort((a, b) => a - b);
  if (!picked.length) return;
  const groups = job.groups.flatMap((g, i) => (job.selected.has(i) ? g.map((k) => [k]) : [g]));
  applyGroups(job, groups, picked[0]);
}

function undo(job) {
  if (!job.history.length) return;
  job.groups = job.history.pop();
  job.workbook = buildWorkbook(job.grids, job.groups);
  job.selected = new Set();
  renderJob(job);
  showPreview(job);
}

function renderManager(job) {
  const panel = $('manage');
  // Nothing to merge or split in a one-page document.
  panel.hidden = job.grids.length < 2;
  if (panel.hidden) return;
  const list = $('sheet-list');
  list.innerHTML = '';
  job.workbook.worksheets.forEach((ws, i) => {
    const li = document.createElement('li');
    li.innerHTML = '<label><input type="checkbox" /><span class="sheet-name"></span><span class="meta"></span></label>';
    const box = li.querySelector('input');
    box.checked = job.selected.has(i);
    box.addEventListener('change', () => {
      if (box.checked) job.selected.add(i);
      else job.selected.delete(i);
      updateManagerButtons(job);
    });
    li.querySelector('.sheet-name').textContent = ws.name;
    li.querySelector('.meta').textContent = `${pageLabel(job, job.groups[i])} · ${ws.rowCount}행 × ${ws.columnCount}열`;
    list.append(li);
  });
  updateManagerButtons(job);
}

function updateManagerButtons(job) {
  const picked = [...job.selected];
  $('merge').disabled = picked.length < 2;
  $('split').disabled = !picked.some((i) => job.groups[i].length > 1);
  $('undo').disabled = !job.history.length;
  const auto = autoGroups(job.grids);
  $('reset').disabled = JSON.stringify(auto) === JSON.stringify(job.groups);
}

// Re-cut the already extracted pages with the new option; sheet grouping is kept.
$('split-lines').addEventListener('change', (e) => {
  if (!shown) return;
  const job = shown;
  job.options = { ...job.options, splitLines: e.target.checked };
  job.grids = gridsFromPages(job.pages, job.options);
  job.workbook = buildWorkbook(job.grids, job.groups);
  renderJob(job);
  showPreview(job, currentSheet);
});

$('merge').addEventListener('click', () => shown && mergeSelected(shown));
$('split').addEventListener('click', () => shown && splitSelected(shown));
$('undo').addEventListener('click', () => shown && undo(shown));
$('reset').addEventListener('click', () => shown && applyGroups(shown, autoGroups(shown.grids)));

let currentSheet = 0;

function showPreview(job, sheetIndex = 0) {
  shown = job;
  currentSheet = sheetIndex;
  $('split-lines').checked = !!job.options.splitLines;
  jobs.forEach((j) => j.el.classList.toggle('active', j === job));
  preview.hidden = false;
  $('preview-title').textContent = job.name;
  const tabs = $('tabs');
  tabs.innerHTML = '';
  const sheets = job.workbook.worksheets;
  tabs.hidden = sheets.length < 2;
  sheets.forEach((ws, i) => {
    const tab = button(ws.name, 'tab', () => showPreview(job, i));
    tab.setAttribute('role', 'tab');
    tab.setAttribute('aria-selected', String(i === sheetIndex));
    tabs.append(tab);
  });
  renderManager(job);
  renderQuality(job);
  drawSheet(sheets[sheetIndex], PREVIEW_ROWS);
  $('sheet').scrollTo(0, 0);
}

const PREVIEW_ROWS = 300;

function drawSheet(ws, maxRows) {
  $('sheet').innerHTML = renderSheet(ws, maxRows);
  const cut = ws.rowCount > maxRows;
  $('more').hidden = !cut;
  if (cut) {
    $('more-text').textContent = `미리보기는 처음 ${maxRows.toLocaleString()}행만 보여요. 전체 ${ws.rowCount.toLocaleString()}행은 다운로드한 파일에 모두 들어 있어요.`;
    $('more-button').onclick = () => drawSheet(ws, Infinity);
  }
}

async function download(job) {
  const buffer = await job.workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = job.name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
