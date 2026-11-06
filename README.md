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

## Benchmarks

Synthetic, reproducible here: 100,000 rows x 10 columns (strings, numbers, dates, booleans), Node 24.13, 10 cores, one fresh process per run, 5 runs, median, output verified against the input. Machine load was about 1.7, not idle.

| Write | time | peak RSS | file |
|---|---|---|---|
| excelts (`writeXlsx`) | 304 ms | 248 MB | 6.3 MB |
| excelts/compat `WorkbookWriter` style, in memory | 341 ms | 362 MB | 6.3 MB |
| excelts/compat streaming | 318 ms | 261 MB | 6.3 MB |
| exceljs streaming | 1246 ms | 260 MB | 7.3 MB |
| exceljs in memory | 2820 ms | 1292 MB | 6.2 MB |

| Read | time | peak RSS |
|---|---|---|
| excelts (`readXlsx`) | 139 ms | 169 MB |
| excelts/compat in memory | 333 ms | 379 MB |
| exceljs streaming reader | 921 ms | 355 MB |
| exceljs in memory | 1338 ms | 700 MB |

Reproduce: `npm run bench:inputs`, `npm run bench`, `npm run bench:read` (they need `npm ci` and a build, `npm run build`).

Real files: on a private corpus of about 2,300 spreadsheets (98% under 100 KB) written by Excel, LibreOffice and many libraries, compared with exceljs streaming, median per file:

| | same process (warm) | new process per file (cold) |
|---|---|---|
| read | 9.1x | 3.7x |
| write | 9.9x | 3.4x |

The cold gap is mostly module loading (about 6 ms against 69 ms), not reading or writing. The corpus and its harness are not in this repository, so these two rows cannot be reproduced from here.

Where it is not faster: very long strings (1 to 30k characters) write only about 1.3x faster than exceljs streaming, because deflate dominates; wide sheets use about as much memory as exceljs streaming; and on small files a native reader (Rust calamine) is faster than this one. Sizes above 10 MB were measured on one real file only.

## Reading untrusted files

The reader enforces limits (entry and total size, compression ratio, entry count, rows, columns, strings, styles) and throws `XlsxLimitError`. Override them with `readXlsx(bytes, { limits })`.

## Verification

`npm test`, `npm run security` (attack fixtures and a mutation fuzzer), `npm run validate:xsd` (needs `xmllint`, run `sh xsd/fetch.sh` once), `npm run verify:text`. Output has not been opened in Microsoft Excel, Numbers or LibreOffice; passing the ECMA-376 schema does not guarantee that Excel will not offer a repair.

## License

MIT
