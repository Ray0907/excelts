# excelts

A TypeScript xlsx reader and writer. It is an independent rewrite, not a fork of exceljs and not affiliated with it. It ships an exceljs-compatible API (`excelts/compat`) so common code changes only its import.

Status: pre-release (0.x). Written for speed and a small install: one runtime dependency (`fflate`); native `node:zlib` is used when available.

## Use

```ts
// lean API: one sheet, rows as arrays
import { writeXlsx, readXlsx } from 'excelts';
const bytes = writeXlsx({ name: 'Sheet1', columns: [{ header: 'name' }, { header: 'amount', format: '#,##0.00' }], rows: [['a', 1.5], ['b', 2]] });
for (const sheet of readXlsx(bytes).sheets) for (const row of sheet.rows()) console.log(row);

// exceljs-compatible API
import { Workbook } from 'excelts/compat';
const wb = new Workbook();
wb.addWorksheet('Report').addRow(['hello', 42]);
const buffer = await wb.xlsx.writeBuffer();
```

## What works

Workbook, worksheets, `addRow`/`addRows`/`getRow`/`getCell`, `columns`, column widths, row heights, merged cells, frozen panes, auto filter, page setup, fonts, fills, borders, alignment, number formats, formulas with cached results, rich text, hyperlinks, data validation, notes, images, defined names, CSV, `writeBuffer`/`writeFile`/`write(stream)`, `load`/`readFile`, and the streaming `WorkbookWriter`.

## What does not

Throws `NotSupportedError`: tables, pivot tables, conditional formatting, sheet protection, the streaming reader (`WorkbookReader`), charts, and the internal `.model` APIs. Loading a file that contains an unsupported part and saving it again throws unless you pass `{ ignoreUnsupported: true }` (exceljs drops such parts silently).

## Differences from exceljs

- Control characters in text are kept as `_xHHHH_` escapes; exceljs removes them.
- `cell.style` returns a write-through view; `cell.style === other.style` is false, and changes deeper than one level (`cell.font.color.argb = ...`) throw.
- `NaN`, infinities and invalid dates are skipped rather than written as invalid XML.
- A malformed shared-strings or styles part is reported when first used.
- Dates from 1900-01-01 to 1900-02-28 follow Excel's 1900 leap-year numbering.

## Reading untrusted files

The reader enforces limits (entry and total size, compression ratio, entry count, rows, columns, strings, styles) and throws `XlsxLimitError`. Override them with `readXlsx(bytes, { limits })`.

## Verification

`npm test`, `npm run security` (attack fixtures and a mutation fuzzer), `npm run validate:xsd` (needs `xmllint`, run `sh xsd/fetch.sh` once), `npm run verify:text`. Output has not been opened in Microsoft Excel, Numbers or LibreOffice; passing the ECMA-376 schema does not guarantee that Excel will not offer a repair.

## License

MIT
