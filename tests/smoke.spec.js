// smoke.spec.js — tool-specific checks for Excel to CSV Converter, on the fictional fixtures in
// tests/fixtures/ (regenerate them with `node tests/fixtures/make-fixtures.js`).
const { test, expect } = require('@playwright/test');
const fs = require('fs');
const path = require('path');

const fixture = (name) => path.join(__dirname, 'fixtures', name);
const ORDERS_HEADER = ['Order ID', 'Order date', 'Customer', 'City', 'Item', 'Qty', 'Unit price', 'Total'];

async function open(page, name = 'sample-workbook.xlsx') {
  await page.goto('./');
  await page.setInputFiles('#file-input', fixture(name));
  await expect(page.locator('#result-card')).toBeVisible();
}

async function download(page, selector = '#download') {
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click(selector)]);
  const bytes = fs.readFileSync(await dl.path());
  return { name: dl.suggestedFilename(), bytes, text: bytes.toString('utf8').replace(/^\uFEFF/, '') };
}

test.describe('Excel to CSV Converter', () => {
  test('loading sample-workbook.xlsx lists its three sheet names in a selector', async ({ page }) => {
    await open(page);
    await expect(page.locator('#options')).toBeVisible();
    await expect(page.locator('#sheet option')).toHaveText(['Orders', 'Customers', 'Products']);
    await expect(page.locator('label[for="sheet"]')).toHaveText('Sheet');
  });

  test('choosing the Orders sheet shows a preview of the first 20 rows', async ({ page }) => {
    await open(page);
    await page.selectOption('#sheet', 'Customers');
    await expect(page.locator('#result-card h3')).toHaveText('Customers');
    await page.selectOption('#sheet', 'Orders');
    await expect(page.locator('#result-card h3')).toHaveText('Orders');
    await expect(page.locator('#preview-table tr')).toHaveCount(20);
    await expect(page.locator('#preview-table thead th')).toHaveText(['1', ...ORDERS_HEADER]);
    await expect(page.locator('#sheet-stats')).toContainText('25 rows × 8 columns');
    await expect(page.locator('#sheet-stats')).toContainText('first 20');
    // Accented name and a value containing a comma survive into the preview, the comma value quoted.
    await expect(page.locator('#preview-table')).toContainText('Renée Fontaine');
    await expect(page.locator('#csv-preview')).toContainText('"Mug, large"');
    await expect(page.locator('#csv-preview')).toContainText('HL-1001,2026-07-01,Renée Fontaine');
  });

  test('downloading produces a .csv file whose first line matches the sheet header row', async ({ page }) => {
    await open(page);
    await expect(page.locator('#download')).toHaveText('Download sample-workbook-Orders.csv');
    const file = await download(page);
    expect(file.name).toBe('sample-workbook-Orders.csv');
    expect(file.bytes.subarray(0, 3)).toEqual(Buffer.from([0xef, 0xbb, 0xbf])); // UTF-8 BOM by default
    const lines = file.text.split('\r\n');
    expect(lines[0]).toBe(ORDERS_HEADER.join(','));
    expect(lines).toHaveLength(25); // header + 24 orders, no trailing empty rows
    expect(file.text).toContain('Renée Fontaine');
    expect(file.text).toContain('"Mug, large"');
  });

  test('Enter on the focused Download button downloads', async ({ page }) => {
    await open(page);
    await page.focus('#download');
    const [dl] = await Promise.all([page.waitForEvent('download'), page.keyboard.press('Enter')]);
    expect(dl.suggestedFilename()).toMatch(/\.csv$/);
  });

  test('switching the delimiter to semicolon changes the preview and the downloaded file', async ({ page }) => {
    await open(page);
    await expect(page.locator('#csv-preview')).toContainText('Order ID,Order date');
    await page.selectOption('#delimiter', ';');
    await expect(page.locator('#csv-preview')).toContainText('Order ID;Order date;Customer');
    await expect(page.locator('#csv-preview')).toContainText('Mug, large'); // no longer needs quotes
    await expect(page.locator('#csv-preview')).not.toContainText('"Mug, large"');
    const file = await download(page);
    expect(file.text.split('\r\n')[0]).toBe(ORDERS_HEADER.join(';'));
    expect(file.text).toContain(';Mug, large;');
  });

  test('encoding, line ending, dates and formulas options change the output', async ({ page }) => {
    await open(page);
    await page.selectOption('#encoding', 'plain');
    await page.selectOption('#line-ending', 'lf');
    await page.selectOption('#dates', 'text');
    await page.selectOption('#formulas', 'formulas');
    await expect(page.locator('#csv-preview')).toContainText('HL-1001,"Jul 1, 2026"');
    const file = await download(page);
    expect(file.bytes[0]).not.toBe(0xef);
    expect(file.text).not.toContain('\r');
    const second = file.text.split('\n')[1];
    expect(second).toBe('HL-1001,"Jul 1, 2026",Renée Fontaine,Brookhaven,Driftwood candle holder,1,18.50,=F2*G2');
  });

  test('Download every sheet saves one CSV per sheet', async ({ page }) => {
    await open(page);
    const names = [];
    page.on('download', (d) => names.push(d.suggestedFilename()));
    await page.click('#download-all');
    await expect.poll(() => names.length, { timeout: 10_000 }).toBe(3);
    expect(names).toEqual(['sample-workbook-Orders.csv', 'sample-workbook-Customers.csv', 'sample-workbook-Products.csv']);
  });

  test('copy to clipboard copies the CSV', async ({ page, context }) => {
    await context.grantPermissions(['clipboard-read', 'clipboard-write']);
    await open(page);
    await page.selectOption('#sheet', 'Products');
    await expect(page.locator('#result-card h3')).toHaveText('Products');
    await page.click('#copy');
    await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toContain('SKU,Product,Price');
  });

  test('edge cases: quoting, merged cells, hidden sheet, formulas left alone, long numbers kept, trailing rows trimmed', async ({ page }) => {
    await open(page, 'edge-cases.xlsx');
    await expect(page.locator('#sheet option')).toHaveText(['Visible', 'Archive (hidden)']);
    const file = await download(page);
    expect(file.name).toBe('edge-cases-Visible.csv');
    expect(file.text).toBe([
      'Label,Value',
      'Merged title,',
      'Note,"Line one\nLine two"',
      'Quote,"She said ""hi"""',
      'Leading equals,=not a formula',
      'Sum,5',
      'Arrival,2026-03-09 12:00:00',
      'Barcode,9780306406157',
      'Phone,447123456789',
      'Ratio,0.3',
    ].join('\r\n'));
    // General-format numbers keep every digit in the preview too, never scientific notation.
    await expect(page.locator('#csv-preview')).toContainText('9780306406157');
    await expect(page.locator('#csv-preview')).not.toContainText('E+');
    await page.selectOption('#sheet', 'Archive');
    await expect(page.locator('#result-card .chip')).toHaveText('hidden sheet');
  });

  test('XLS and ODS workbooks convert too', async ({ page }) => {
    for (const name of ['sample-workbook.xls', 'sample-workbook.ods']) {
      await open(page, name);
      await expect(page.locator('#sheet option')).toHaveText(['Orders', 'Customers', 'Products']);
      await expect(page.locator('#csv-preview')).toContainText(ORDERS_HEADER.join(','));
    }
  });

  test('a non-spreadsheet file shows a friendly error and no download', async ({ page }) => {
    await page.goto('./');
    let downloads = 0;
    page.on('download', () => downloads++);
    await page.setInputFiles('#file-input', fixture('sample.pdf'));
    await expect(page.locator('#results .alert--error')).toContainText('sample.pdf is not a spreadsheet');
    // A PNG renamed to .xlsx is caught by its contents, not just its name.
    await page.setInputFiles('#file-input', { name: 'orders.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: fs.readFileSync(fixture('sample.png')) });
    await expect(page.locator('#results .alert--error')).toContainText('not a spreadsheet');
    await expect(page.locator('#download')).toHaveCount(0);
    await expect(page.locator('#options')).toBeHidden();
    expect(downloads).toBe(0);
  });

  test('password protected files get a clear message', async ({ page }) => {
    await page.goto('./');
    await page.setInputFiles('#file-input', fixture('locked.xlsx'));
    await expect(page.locator('#results .alert--error')).toContainText('password protected');
    await expect(page.locator('#download')).toHaveCount(0);
  });

  test('wrong files mixed with a workbook are rejected while the workbook still converts', async ({ page }) => {
    await page.goto('./');
    await page.evaluate(async () => {
      const buf = await (await fetch('./tests/fixtures/sample-workbook.xlsx')).arrayBuffer();
      const { handleFiles } = await import('./app.js');
      await handleFiles([new File(['hello'], 'notes.txt', { type: 'text/plain' }), new File([buf], 'orders-q3.xlsx')]);
    });
    await expect(page.locator('#results .alert--error')).toContainText('notes.txt is not a spreadsheet');
    await expect(page.locator('#download')).toHaveText('Download orders-q3-Orders.csv');
  });

  test('files over 5 MB are parsed in a Web Worker', async ({ page }) => {
    await page.goto('./');
    const workers = [];
    page.on('worker', (w) => workers.push(w.url()));
    await page.evaluate(async () => {
      const buf = await (await fetch('./tests/fixtures/sample-workbook.xlsx')).arrayBuffer();
      // Pad the ZIP with a large trailing comment-like block: readers find the directory from the end.
      const padded = new Uint8Array(6 * 1024 * 1024 + buf.byteLength);
      padded.set(new Uint8Array(buf), 0);
      const { handleFiles } = await import('./app.js');
      await handleFiles([new File([padded], 'big-orders.xlsx')]);
    });
    await expect(page.locator('#result-card h3')).toHaveText('Orders', { timeout: 30_000 });
    expect(workers.some((u) => u.endsWith('/worker.js'))).toBeTruthy();
    const file = await download(page);
    expect(file.text.split('\r\n')[0]).toBe(ORDERS_HEADER.join(','));
  });

  test('works at 390px wide', async ({ page }) => {
    await page.setViewportSize({ width: 390, height: 844 });
    await open(page);
    await expect(page.locator('#download')).toBeVisible();
    const widths = await page.evaluate(() => ({ doc: document.documentElement.scrollWidth, win: window.innerWidth }));
    expect(widths.doc).toBeLessThanOrEqual(widths.win + 1);
    const box = await page.locator('#result-card').boundingBox();
    expect(box.x).toBeGreaterThanOrEqual(0);
    expect(box.x + box.width).toBeLessThanOrEqual(391);
    const file = await download(page);
    expect(file.name).toBe('sample-workbook-Orders.csv');
  });

  test('__demo shows a finished result', async ({ page }) => {
    await page.goto('./');
    await page.evaluate(() => window.__demo());
    await expect(page.locator('#preview-table tbody tr')).toHaveCount(19);
    await expect(page.locator('#download')).toHaveText('Download harbor-lane-orders-q3-Orders.csv');
  });
});
