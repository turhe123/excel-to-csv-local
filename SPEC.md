# Excel to CSV Converter — specification

**Tagline:** Convert XLSX and XLS sheets to clean CSV in your browser, pick the sheet and delimiter, nothing uploaded.

**Category:** Data & Files

**Free alternative to:** Convertio Excel to CSV (uploads your spreadsheet to its servers, free plan capped at 100 MB and 10 conversions a day, ads, sign-up for more)

**Search phrases:** excel to csv, xlsx to csv, convert xls to csv online, export excel sheet as csv, excel to csv utf-8, excel to csv semicolon

**Allowed libraries:** xlsx

**Published at:** https://turhe123.github.io/excel-to-csv-local/

## Acceptance checks (each must have a test)

- loading sample-workbook.xlsx lists its three sheet names in a selector
- choosing the Orders sheet shows a preview of the first 20 rows
- downloading produces a .csv file whose first line matches the sheet header row
- switching the delimiter to semicolon changes the preview and the downloaded file
- a non-spreadsheet file shows a friendly error and no download
- works at 390px wide

## Spec

Search demand: people type excel to csv, xlsx to csv and convert xls to csv online thousands of times a month, usually because an import tool, database or shop platform wants CSV. The incumbents such as Convertio, Zamzar and CloudConvert upload the workbook to a server, cap free size and daily conversions, show ads and often push a sign-up. Spreadsheets frequently contain customer lists, payroll or invoices, so the upload is a real worry, and a common follow-up complaint is broken accents because the output is not UTF-8.

Our version does one job: turn one sheet of an XLSX, XLS, XLSM or ODS file into a CSV file. Features: drop or choose a file, see the sheet names, choose a sheet, choose the delimiter (comma, semicolon, tab, pipe), choose the line ending, choose whether to add a UTF-8 byte order mark so Excel and accented text open correctly, choose whether dates are written as ISO dates or as displayed text, and choose whether formulas export their computed values. A Download every sheet option saves one CSV per sheet.

UI: a large drop zone above the fold with a Choose file button. After loading, a compact options row appears with sheet selector, delimiter, encoding and date format, then a scrollable preview table of the first 20 rows with row and column counts, then a primary Download CSV button and a secondary Copy to clipboard button. The download is named after the source file with the sheet name appended, for example orders-q3-Orders.csv. Empty state explains in one line that files stay on the device. Error state names the problem, for example this file is password protected or is not a spreadsheet. Keyboard: Enter on the focused Download button downloads, and the drop zone is a labelled file input so it is fully keyboard reachable.

Technical approach: read the file with the File API into an ArrayBuffer, parse with the allowed xlsx library using cellDates and raw false options as selected, then use its sheet to CSV utility with the chosen field separator. Prefix the byte order mark when requested and build a Blob with a text/csv type for download through an object URL that is revoked afterwards. For the all-sheets option, generate one CSV string per sheet with the xlsx library and offer sequential downloads of each file. Run parsing in a Web Worker when the file exceeds 5 MB so the page stays responsive, and show a progress message. Preview only slices the first rows so memory stays modest.

Edge cases: very large sheets over 200 MB of cells should warn about memory on mobile Safari merged cells export only the top-left value cells beginning with equals, plus, minus or at are left untouched because this is a converter not a spreadsheet values containing the delimiter, quotes or line breaks are quoted correctly empty trailing rows are trimmed hidden sheets are listed with a hidden label password protected files produce a clear message time zones are not applied to date serials, so a date stays the date shown in Excel.

Demo data: a fictional workbook from a made up shop called Harbor Lane Goods with sheets named Orders, Customers and Products, containing twenty rows of invented names, cities, order totals and ISO dates, including one name with an accented letter and one cell with a comma to prove quoting.
