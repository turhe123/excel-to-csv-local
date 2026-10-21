// csv-core.js — workbook parsing and CSV building, shared by the page (imported by app.js) and worker.js
// (importScripts). Plain script with no import/export so it runs in both; it publishes self.CsvCore and
// uses the global XLSX from vendor/xlsx at call time.
(function () {
  'use strict';

  const ZIP_TYPES = ['xlsx', 'xlsm', 'xlsb', 'xltx', 'xltm', 'ods'];

  function friendly(message) {
    const err = new Error(message);
    err.friendly = true;
    return err;
  }

  // Look at the first bytes before handing the file to the parser: XLSX.read happily treats any text
  // (a PDF, a PNG renamed to .xls) as a one-column CSV, which would "convert" garbage.
  function sniff(bytes, ext) {
    const b = (i) => bytes[i];
    const zip = b(0) === 0x50 && b(1) === 0x4b && (b(2) === 3 || b(2) === 5) && (b(3) === 4 || b(3) === 6);
    const cfb = b(0) === 0xd0 && b(1) === 0xcf && b(2) === 0x11 && b(3) === 0xe0;
    const biff = b(0) === 0x09 && [0x00, 0x02, 0x04, 0x08].includes(b(1));
    let i = 0;
    if (b(0) === 0xef && b(1) === 0xbb && b(2) === 0xbf) i = 3;
    while (i < bytes.length && i < 64 && [0x20, 0x09, 0x0a, 0x0d].includes(bytes[i])) i++;
    const markup = bytes[i] === 0x3c; // '<': SpreadsheetML 2003 or an HTML table saved as .xls
    if (zip || cfb) return;
    if (ext === 'xls' && (biff || markup)) return;
    if (ext === 'fods' && markup) return;
    throw friendly('This file is not a spreadsheet. Its contents do not look like an Excel or OpenDocument workbook.');
  }

  function parse(buffer, ext) {
    const bytes = new Uint8Array(buffer);
    if (!bytes.length) throw friendly('This file is empty.');
    sniff(bytes.subarray(0, 64), ext);
    try {
      // Dates stay as serial numbers (cellDates off) so no time zone is ever applied; cellNF keeps the
      // number format so date cells can be recognised, and cellFormula keeps formula text.
      return XLSX.read(bytes, { type: 'array', dense: true, cellFormula: true, cellNF: true, cellDates: false, cellHTML: false, raw: false });
    } catch (e) {
      const msg = String((e && e.message) || e);
      if (/password|encrypt/i.test(msg)) {
        throw friendly('This file is password protected. Open it in your spreadsheet app, remove the password, save it and try again.');
      }
      if (/zip/i.test(msg) && ZIP_TYPES.includes(ext)) {
        throw friendly('This file is a ZIP archive but not a spreadsheet, or the workbook inside is damaged.');
      }
      throw friendly(`This file could not be read as a spreadsheet (${msg}).`);
    }
  }

  function cellAt(ws, r, c) {
    const data = ws['!data'];
    if (data) return data[r] ? data[r][c] : undefined;
    return ws[XLSX.utils.encode_cell({ r, c })];
  }

  const pad = (n, w = 2) => String(n).padStart(w, '0');

  function isoDate(cell, date1904) {
    const p = XLSX.SSF.parse_date_code(cell.v, { date1904 });
    if (!p) return null;
    const fmt = String(cell.z || '').toLowerCase().replace(/"[^"]*"|\[[^\]]*\]/g, '');
    const hasDate = /[dy]/.test(fmt) || (/m/.test(fmt) && !/[hs]/.test(fmt));
    const hasTime = /[hs]/.test(fmt);
    const date = `${pad(p.y, 4)}-${pad(p.m)}-${pad(p.d)}`;
    const time = `${pad(p.H)}:${pad(p.M)}:${pad(p.S)}`;
    if (hasDate && hasTime) return `${date} ${time}`;
    if (hasTime && !hasDate) return time;
    return date;
  }

  function cellText(cell, opts, date1904) {
    if (!cell) return '';
    if (opts.formulas === 'formulas' && cell.f) return `=${cell.f}`;
    if (cell.t === 'z' || cell.v == null) return cell.w != null ? String(cell.w) : '';
    if (opts.dates === 'iso' && cell.t === 'n' && cell.z && XLSX.SSF.is_date(cell.z)) {
      const iso = isoDate(cell, date1904);
      if (iso) return iso;
    }
    if (cell.t === 'n' && isGeneral(cell.z)) return rawNumber(cell.v);
    if (cell.w != null) return String(cell.w);
    if (cell.t === 'b') return cell.v ? 'TRUE' : 'FALSE';
    return String(cell.v);
  }

  const isGeneral = (z) => !z || String(z).toLowerCase() === 'general';

  // General-format numbers carry no display rule, so write the stored value in full. The formatted text
  // (cell.w) imitates Excel's narrow-column display and turns 9780306406157 into 9.78031E+12, which
  // silently destroys barcodes, phone numbers and account numbers. Excel keeps 15 significant digits, so
  // rounding there also hides binary noise such as 0.30000000000000004.
  function rawNumber(v) {
    if (typeof v !== 'number' || !Number.isFinite(v)) return String(v);
    return String(Number(v.toPrecision(15)));
  }

  // Bounds of the sheet with empty trailing rows and columns removed (formatting can stretch !ref far
  // beyond the data). Returns null for an empty sheet.
  function bounds(wb, ws, opts) {
    if (!ws || !ws['!ref']) return null;
    const range = XLSX.utils.decode_range(ws['!ref']);
    const date1904 = isDate1904(wb);
    let lastRow = -1;
    let lastCol = -1;
    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        if (cellText(cellAt(ws, r, c), opts, date1904) !== '') {
          lastRow = r;
          if (c > lastCol) lastCol = c;
        }
      }
    }
    if (lastRow < 0) return null;
    return { s: range.s, e: { r: lastRow, c: lastCol } };
  }

  function isDate1904(wb) {
    return Boolean(wb.Workbook && wb.Workbook.WBProps && wb.Workbook.WBProps.date1904);
  }

  function quote(text, fs) {
    if (text.includes(fs) || text.includes('"') || text.includes('\n') || text.includes('\r')) {
      return `"${text.replace(/"/g, '""')}"`;
    }
    return text;
  }

  // Builds the rows for a sheet. `limit` caps the number of rows (preview); grid is only kept then.
  function build(wb, name, opts, limit) {
    const ws = wb.Sheets[name];
    if (!ws) throw friendly(`The sheet “${name}” was not found in this workbook.`);
    const b = bounds(wb, ws, opts);
    if (!b) return { csv: '', grid: [], rows: 0, cols: 0, startRow: 0 };
    const date1904 = isDate1904(wb);
    const rows = b.e.r - b.s.r + 1;
    const cols = b.e.c - b.s.c + 1;
    const endRow = limit ? Math.min(b.e.r, b.s.r + limit - 1) : b.e.r;
    const lines = [];
    const grid = [];
    for (let r = b.s.r; r <= endRow; r++) {
      const row = [];
      for (let c = b.s.c; c <= b.e.c; c++) row.push(cellText(cellAt(ws, r, c), opts, date1904));
      if (limit) grid.push(row);
      lines.push(row.map((t) => quote(t, opts.fs)).join(opts.fs));
    }
    return { csv: lines.join(opts.rs), grid, rows, cols, startRow: b.s.r };
  }

  function describe(wb) {
    const meta = (wb.Workbook && wb.Workbook.Sheets) || [];
    return wb.SheetNames.map((name, i) => ({ name, hidden: Boolean(meta[i] && meta[i].Hidden) }));
  }

  self.CsvCore = { parse, describe, build };
})();
