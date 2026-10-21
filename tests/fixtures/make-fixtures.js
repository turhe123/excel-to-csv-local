// make-fixtures.js — generates the small, fictional spreadsheet fixtures used by tests and __demo.
// Run with: node tests/fixtures/make-fixtures.js
const fs = require('fs');
const path = require('path');
const XLSX = require('../../vendor/xlsx/xlsx.full.min.js');

const out = (name) => path.join(__dirname, name);

// Excel serial for a calendar date, no time zone involved.
const serial = (y, m, d) => (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;

const customers = [
  ['Renée Fontaine', 'Brookhaven'], ['Milo Ashcombe', 'Fernridge'], ['Tessa Lindqvist', 'Oakmere'],
  ['Jasper Quill', 'Saltmarsh'], ['Nadia Orwell', 'Pinecrest'], ['Felix Arden', 'Larkfield'],
  ['Ivy Calloway', 'Westerby'], ['Owen Pembrook', 'Stonebridge'], ['Clara Wren', 'Millbrook'],
  ['Hugo Sterling', 'Ashgrove'], ['Lena Marlowe', 'Dunmore'], ['Theo Hartwell', 'Elmstead'],
  ['Maya Thorne', 'Coldwater'], ['Rufus Kettle', 'Harrowgate'], ['Iris Fenwick', 'Wrenford'],
  ['Silas Penrose', 'Brightwater'], ['Greta Holloway', 'Rookwood'], ['Ezra Blackwood', 'Glenhollow'],
  ['Poppy Ainsley', 'Kingsmere'], ['Rowan Easton', 'Thistledown'],
];
const products = [
  ['HLG-101', 'Driftwood candle holder', 18.5], ['HLG-102', 'Mug, large', 12], ['HLG-103', 'Linen tea towel', 9.75],
  ['HLG-104', 'Rope doorstop', 22], ['HLG-105', 'Sea glass coasters (set of 4)', 16.25], ['HLG-106', 'Canvas tote', 14],
];

function ordersSheet() {
  const aoa = [['Order ID', 'Order date', 'Customer', 'City', 'Item', 'Qty', 'Unit price', 'Total']];
  for (let i = 0; i < 24; i++) {
    const [name, city] = customers[i % customers.length];
    const [, item, price] = products[(i * 5) % products.length];
    aoa.push([`HL-${1001 + i}`, serial(2026, 7, 1 + i), name, city, item, 1 + (i % 4), price, null]);
  }
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  for (let r = 1; r < aoa.length; r++) {
    const date = ws[XLSX.utils.encode_cell({ r, c: 1 })];
    date.z = 'mmm d, yyyy';
    ws[XLSX.utils.encode_cell({ r, c: 6 })].z = '0.00';
    const qty = aoa[r][5];
    const price = aoa[r][6];
    ws[XLSX.utils.encode_cell({ r, c: 7 })] = { t: 'n', v: Math.round(qty * price * 100) / 100, f: `F${r + 1}*G${r + 1}`, z: '0.00' };
  }
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length - 1, c: 7 } });
  return ws;
}

function customersSheet() {
  const aoa = [['Customer ID', 'Name', 'City', 'Joined', 'Address']];
  customers.forEach(([name, city], i) => {
    aoa.push([`C-${201 + i}`, name, city, serial(2025, 1 + (i % 12), 3 + i), i === 0 ? '12 Pier Road, Unit 4' : `${10 + i * 3} Harbor Lane`]);
  });
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  for (let r = 1; r < aoa.length; r++) ws[XLSX.utils.encode_cell({ r, c: 3 })].z = 'yyyy-mm-dd';
  return ws;
}

function productsSheet() {
  return XLSX.utils.aoa_to_sheet([['SKU', 'Product', 'Price'], ...products]);
}

function sampleWorkbook() {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ordersSheet(), 'Orders');
  XLSX.utils.book_append_sheet(wb, customersSheet(), 'Customers');
  XLSX.utils.book_append_sheet(wb, productsSheet(), 'Products');
  return wb;
}

function edgeWorkbook() {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ['Label', 'Value'],
    ['Merged title', null],
    ['Note', 'Line one\nLine two'],
    ['Quote', 'She said "hi"'],
    ['Leading equals', '=not a formula'],
    ['Sum', null],
    ['Arrival', serial(2026, 3, 9) + 0.5],
  ]);
  ws['!merges'] = [{ s: { r: 1, c: 0 }, e: { r: 1, c: 1 } }];
  ws.B6 = { t: 'n', v: 5, f: '2+3' };
  ws.B7.z = 'yyyy-mm-dd hh:mm';
  // Trailing empty rows that only carry an empty value (as formatting leftovers do).
  ws.A9 = { t: 's', v: '' };
  ws.A10 = { t: 'z' };
  ws['!ref'] = 'A1:C10';
  XLSX.utils.book_append_sheet(wb, ws, 'Visible');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['secret'], ['value']]), 'Archive');
  wb.Workbook = { Sheets: [{ Hidden: 0 }, { Hidden: 1 }] };
  return wb;
}

function lockedWorkbook() {
  // Same container layout as an encrypted Office file: a CFB with EncryptionInfo + EncryptedPackage.
  const cfb = XLSX.CFB.utils.cfb_new();
  XLSX.CFB.utils.cfb_add(cfb, '/EncryptionInfo', Buffer.from([4, 0, 4, 0, 0x40, 0, 0, 0]));
  XLSX.CFB.utils.cfb_add(cfb, '/EncryptedPackage', Buffer.alloc(64));
  return XLSX.CFB.write(cfb, { type: 'buffer' });
}

const save = (wb, name, bookType) => fs.writeFileSync(out(name), XLSX.write(wb, { type: 'buffer', bookType, compression: true }));
save(sampleWorkbook(), 'sample-workbook.xlsx', 'xlsx');
save(sampleWorkbook(), 'sample-workbook.ods', 'ods');
save(sampleWorkbook(), 'sample-workbook.xls', 'biff8');
save(edgeWorkbook(), 'edge-cases.xlsx', 'xlsx');
fs.writeFileSync(out('locked.xlsx'), lockedWorkbook());
console.log('fixtures written');
