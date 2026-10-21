// worker.js — parses big workbooks off the main thread (files over 5 MB) and keeps the workbook here,
// so building CSV for another sheet or delimiter never blocks the page either.
importScripts('vendor/xlsx/xlsx.full.min.js', 'csv-core.js');

let wb = null;

self.onmessage = ({ data }) => {
  const { id, type } = data;
  try {
    let result;
    if (type === 'load') {
      wb = self.CsvCore.parse(data.buffer, data.ext);
      result = self.CsvCore.describe(wb);
    } else if (type === 'build') {
      result = self.CsvCore.build(wb, data.name, data.opts, data.limit);
    }
    self.postMessage({ id, result });
  } catch (e) {
    self.postMessage({ id, error: String((e && e.message) || e), friendly: Boolean(e && e.friendly) });
  }
};
