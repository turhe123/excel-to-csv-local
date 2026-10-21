// app.js — Excel to CSV Converter. Everything runs on the device: the workbook is read with the File API,
// parsed by the vendored xlsx library (in a Web Worker for files over 5 MB) and turned into CSV here.
import { toast, downloadBlob, formatBytes, wireDropzone } from './site.js';
import './csv-core.js';

const { CsvCore } = self;

const WORKER_THRESHOLD = 5 * 1024 * 1024;
const PREVIEW_ROWS = 20;
const BIG_CELLS = 10_000_000; // rough point where a sheet's text reaches hundreds of MB in memory
const EXTENSIONS = ['xlsx', 'xlsm', 'xlsb', 'xltx', 'xltm', 'xls', 'ods', 'fods'];

const $ = (id) => document.getElementById(id);
const zone = $('dropzone');
const input = $('file-input');
const options = $('options');
const results = $('results');
const controls = {
  sheet: $('sheet'),
  delimiter: $('delimiter'),
  encoding: $('encoding'),
  dates: $('dates'),
  formulas: $('formulas'),
  lineEnding: $('line-ending'),
};

const state = { engine: null, file: null, baseName: '', sheets: [], notices: [], token: 0 };

const extOf = (name) => (name.match(/\.([^.]+)$/) || [, ''])[1].toLowerCase();

function el(tag, props = {}, children = []) {
  const node = Object.assign(document.createElement(tag), props);
  for (const child of [].concat(children)) if (child != null) node.append(child);
  return node;
}

// ---------- Engines: same API on the main thread and in the worker ----------

class LocalEngine {
  async load(buffer, ext) { this.wb = CsvCore.parse(buffer, ext); return CsvCore.describe(this.wb); }
  async build(name, opts, limit) { return CsvCore.build(this.wb, name, opts, limit); }
  dispose() { this.wb = null; }
}

class WorkerEngine {
  constructor() {
    this.worker = new Worker('worker.js');
    this.pending = new Map();
    this.seq = 0;
    this.worker.onmessage = ({ data }) => {
      const p = this.pending.get(data.id);
      if (!p) return;
      this.pending.delete(data.id);
      if (data.error) {
        const err = new Error(data.error);
        err.friendly = data.friendly;
        p.reject(err);
      } else p.resolve(data.result);
    };
    this.worker.onerror = (e) => {
      e.preventDefault();
      for (const p of this.pending.values()) p.reject(new Error('The background converter stopped unexpectedly. The file may be too large for this device.'));
      this.pending.clear();
    };
  }
  call(msg, transfer) {
    const id = ++this.seq;
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.worker.postMessage({ ...msg, id }, transfer || []);
    });
  }
  load(buffer, ext) { return this.call({ type: 'load', buffer, ext }, [buffer]); }
  build(name, opts, limit) { return this.call({ type: 'build', name, opts, limit }); }
  dispose() { this.worker.terminate(); }
}

// ---------- Rendering helpers ----------

function alertBox(kind, text) {
  return el('div', { className: `alert alert--${kind}`, role: kind === 'error' ? 'alert' : 'status', textContent: text });
}

function showOnly(...nodes) {
  results.replaceChildren(...nodes.filter(Boolean));
}

function progressView(message, pct) {
  const bar = el('span');
  const wrap = el('div', { className: `progress${pct == null ? ' progress--busy' : ''}` }, bar);
  wrap.setAttribute('role', 'progressbar');
  wrap.setAttribute('aria-label', message);
  if (pct != null) {
    bar.style.width = `${pct}%`;
    wrap.setAttribute('aria-valuenow', String(pct));
  }
  return el('div', { className: 'x2c-status' }, [el('p', { className: 'x2c-status__text', textContent: pct == null ? message : `${message} ${pct}%` }), wrap]);
}

// Floor to one decimal so 99.96% reads 99.9%, never a premature 100%.
const percent = (done, total) => (total ? Math.floor((done / total) * 1000) / 10 : 0);

function readFile(file, token) {
  if (file.size <= WORKER_THRESHOLD) return file.arrayBuffer();
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onprogress = (e) => {
      if (token === state.token && e.lengthComputable) showOnly(...noticeNodes(), progressView(`Reading ${file.name} (${formatBytes(file.size)})…`, percent(e.loaded, e.total)));
    };
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error('The file could not be read. It may have been moved or is still being written.'));
    reader.readAsArrayBuffer(file);
  });
}

function noticeNodes() {
  return state.notices.map(([kind, text]) => alertBox(kind, text));
}

function currentOpts() {
  return {
    fs: controls.delimiter.value,
    rs: controls.lineEnding.value === 'lf' ? '\n' : '\r\n',
    dates: controls.dates.value,
    formulas: controls.formulas.value,
  };
}

function safeName(text) {
  return text.replace(/[\\/:*?"<>|\u0000-\u001f]+/g, '_').replace(/\s+/g, ' ').trim() || 'Sheet';
}

function csvName(sheet) {
  return `${state.baseName}-${safeName(sheet)}.csv`;
}

function csvBlob(text) {
  const parts = controls.encoding.value === 'bom' ? ['\uFEFF', text] : [text];
  return new Blob(parts, { type: 'text/csv;charset=utf-8' });
}

// ---------- Main flow ----------

async function handleFiles(files) {
  const token = ++state.token;
  const valid = files.filter((f) => EXTENSIONS.includes(extOf(f.name)));
  const rejected = files.filter((f) => !valid.includes(f));
  state.notices = [];
  if (rejected.length) {
    const names = rejected.map((f) => f.name).join(', ');
    state.notices.push(['error', `${names} ${rejected.length > 1 ? 'are not spreadsheets' : 'is not a spreadsheet'}. Choose an XLSX, XLS, XLSM, XLSB or ODS file.`]);
  }
  if (valid.length > 1) state.notices.push(['info', `One workbook at a time: converting ${valid[0].name}. Drop the others one by one.`]);
  if (!valid.length) {
    resetWorkbook();
    showOnly(...noticeNodes());
    return;
  }
  await loadFile(valid[0], token);
}

function resetWorkbook() {
  if (state.engine) state.engine.dispose();
  state.engine = null;
  state.sheets = [];
  options.hidden = true;
  zone.classList.remove('dropzone--compact');
  $('dropzone-title').textContent = 'Drop an Excel file here';
}

async function loadFile(file, token) {
  resetWorkbook();
  state.file = file;
  state.baseName = safeName(file.name.replace(/\.[^.]+$/, ''));
  const big = file.size > WORKER_THRESHOLD;
  if (big) showOnly(...noticeNodes(), progressView(`Reading ${file.name} (${formatBytes(file.size)})…`, 0));
  try {
    const buffer = await readFile(file, token);
    if (token !== state.token) return;
    let engine = null;
    if (big && typeof Worker !== 'undefined') {
      try { engine = new WorkerEngine(); } catch { engine = null; }
    }
    engine = engine || new LocalEngine();
    if (big) showOnly(...noticeNodes(), progressView(`Parsing ${file.name} in the background, the page stays usable…`));
    const sheets = await engine.load(buffer, extOf(file.name));
    if (token !== state.token) { engine.dispose(); return; }
    state.engine = engine;
    state.sheets = sheets;
    controls.sheet.replaceChildren(...sheets.map((s) => el('option', { value: s.name, textContent: s.hidden ? `${s.name} (hidden)` : s.name })));
    const firstVisible = sheets.find((s) => !s.hidden) || sheets[0];
    if (firstVisible) controls.sheet.value = firstVisible.name;
    options.hidden = false;
    zone.classList.add('dropzone--compact');
    $('dropzone-title').textContent = 'Drop another Excel file here';
    toast(`${file.name} opened on this device, nothing was uploaded`);
    await render({ scroll: true });
  } catch (err) {
    if (token !== state.token) return;
    resetWorkbook();
    const message = err && err.friendly ? err.message : `This file could not be converted: ${(err && err.message) || err}`;
    showOnly(...noticeNodes(), alertBox('error', `${file.name}: ${message}`));
  }
}

async function render({ scroll = false } = {}) {
  if (!state.engine) return;
  const token = state.token;
  const name = controls.sheet.value;
  let preview;
  try {
    preview = await state.engine.build(name, currentOpts(), PREVIEW_ROWS);
  } catch (err) {
    if (token === state.token) showOnly(...noticeNodes(), alertBox('error', err.message));
    return;
  }
  if (token !== state.token || name !== controls.sheet.value) return;
  const card = resultCard(name, preview);
  const extra = [];
  if (preview.rows * preview.cols > BIG_CELLS || (state.file && state.file.size > 100 * 1024 * 1024)) {
    extra.push(alertBox('warn', `This sheet is very large (${preview.rows.toLocaleString()} rows × ${preview.cols.toLocaleString()} columns). Mobile Safari may run out of memory; a desktop browser is safer.`));
  }
  showOnly(...noticeNodes(), ...extra, card);
  if (scroll) card.scrollIntoView({ block: 'nearest' });
}

function previewTable(preview) {
  const [head = [], ...body] = preview.grid;
  const headRow = el('tr', {}, [el('th', { className: 'x2c-rownum', scope: 'col', textContent: '#' }), ...head.map((t) => el('th', { scope: 'col', textContent: t }))]);
  const bodyRows = body.map((row, i) => el('tr', {}, [
    el('th', { className: 'x2c-rownum', scope: 'row', textContent: String(preview.startRow + i + 2) }),
    ...row.map((t) => el('td', { textContent: t })),
  ]));
  headRow.firstChild.textContent = String(preview.startRow + 1);
  return el('table', { id: 'preview-table' }, [el('thead', {}, headRow), el('tbody', {}, bodyRows)]);
}

function resultCard(name, preview) {
  const sheetMeta = state.sheets.find((s) => s.name === name);
  const empty = preview.rows === 0;
  const shown = Math.min(PREVIEW_ROWS, preview.rows);
  const stats = empty
    ? 'This sheet is empty, there is nothing to convert.'
    : `${preview.rows.toLocaleString()} ${preview.rows === 1 ? 'row' : 'rows'} × ${preview.cols.toLocaleString()} ${preview.cols === 1 ? 'column' : 'columns'}${preview.rows > shown ? ` · preview shows the first ${shown}` : ''}`;
  const fileName = csvName(name);

  const download = el('button', { type: 'button', id: 'download', className: 'btn btn--primary', disabled: empty }, `Download ${fileName}`);
  download.addEventListener('click', () => downloadSheet(name, download));
  const copy = el('button', { type: 'button', id: 'copy', className: 'btn', disabled: empty }, 'Copy to clipboard');
  copy.addEventListener('click', () => copySheet(name, copy));
  const all = state.sheets.length > 1
    ? el('button', { type: 'button', id: 'download-all', className: 'btn btn--ghost' }, `Download every sheet (${state.sheets.length} CSV files)`)
    : null;
  if (all) all.addEventListener('click', () => downloadAll(all));

  const head = el('div', { className: 'x2c-card__head' }, [
    el('h3', { className: 'x2c-card__title' }, [document.createTextNode(name), sheetMeta && sheetMeta.hidden ? el('span', { className: 'chip', textContent: 'hidden sheet' }) : null]),
    el('p', { className: 'result__info', id: 'sheet-stats', textContent: stats }),
  ]);
  const parts = [head];
  if (!empty) {
    const wrap = el('div', { className: 'table-wrap x2c-table', tabIndex: 0 }, previewTable(preview));
    wrap.setAttribute('role', 'region');
    wrap.setAttribute('aria-label', `Preview of the first ${shown} rows of ${name}`);
    const pre = el('pre', { className: 'code x2c-csv', id: 'csv-preview', tabIndex: 0, textContent: preview.csv });
    pre.setAttribute('aria-label', 'CSV text preview');
    parts.push(wrap, el('p', { className: 'x2c-label', textContent: `CSV output, first ${shown} ${shown === 1 ? 'line' : 'lines'}` }), pre);
  }
  parts.push(el('div', { className: 'row x2c-actions' }, [download, copy, all]));
  const card = el('section', { className: 'x2c-card', id: 'result-card' }, parts);
  card.setAttribute('aria-label', `Result for sheet ${name}`);
  return card;
}

async function fullCsv(name) {
  const out = await state.engine.build(name, currentOpts(), 0);
  return out.csv;
}

async function busy(button, label, work) {
  const original = button.textContent;
  button.disabled = true;
  button.textContent = label;
  try {
    await work();
  } catch (err) {
    results.prepend(alertBox('error', err.message || String(err)));
  } finally {
    button.disabled = false;
    button.textContent = original;
  }
}

function downloadSheet(name, button) {
  return busy(button, 'Building CSV…', async () => {
    const text = await fullCsv(name);
    downloadBlob(csvBlob(text), csvName(name));
    toast(`Saved ${csvName(name)}`);
  });
}

function copySheet(name, button) {
  return busy(button, 'Copying…', async () => {
    const text = await fullCsv(name);
    try {
      await navigator.clipboard.writeText(text);
    } catch {
      const area = el('textarea', { value: text, readOnly: true });
      area.setAttribute('aria-hidden', 'true');
      area.style.cssText = 'position:fixed;left:-9999px;top:0';
      document.body.append(area);
      area.select();
      const ok = document.execCommand('copy');
      area.remove();
      if (!ok) throw new Error('The browser blocked clipboard access. Use Download instead.');
    }
    toast(`Copied ${name} as CSV`);
  });
}

function downloadAll(button) {
  return busy(button, 'Saving every sheet…', async () => {
    let saved = 0;
    for (const sheet of state.sheets) {
      const text = await fullCsv(sheet.name);
      downloadBlob(csvBlob(text), csvName(sheet.name));
      saved++;
      // Browsers drop downloads fired in the same instant; space them out.
      await new Promise((r) => setTimeout(r, 350));
    }
    toast(`Saved ${saved} CSV files`);
  });
}

wireDropzone(zone, input, handleFiles);
for (const control of Object.values(controls)) control.addEventListener('change', () => render());

// Called by the factory's screenshot script to show the tool in a realistic, fictional state.
window.__demo = async () => {
  const res = await fetch('./tests/fixtures/sample-workbook.xlsx');
  const blob = await res.blob();
  await handleFiles([new File([blob], 'harbor-lane-orders-q3.xlsx', { type: blob.type })]);
};

export { handleFiles };
