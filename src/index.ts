import { strToU8 } from 'fflate';
import { dateToSerial1900 } from './serial.js';
import { SharedStrings } from './sharedStrings.js';
import { Styles } from './styles.js';
import { XML_DECLARATION, escapeXml } from './xml.js';
import { concat, defaultCompressor, levelFor, writeZip, type Compressor } from './zip.js';

/** An error value such as `#DIV/0!`; the same shape exceljs uses. */
export interface CellError {
  error: string;
}
export type CellValue = string | number | boolean | Date | CellError | null | undefined;
export interface Column {
  header: string;
  /** Excel number format code, applied to every data cell of the column. */
  format?: string;
}
export interface SheetInput {
  name: string;
  columns: Column[];
  rows: Iterable<readonly CellValue[]>;
}

const ERROR_CODES: ReadonlySet<string> = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA']);
const MAX_COLUMNS = 16384; // XFD
const MAX_ROWS = 1048576;
const CHUNK_CHARS = 65536;
const SHEET_NAME_ILLEGAL = /[\[\]:*?/\\]/;

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';

/** Column letters, computed once per column index and shared by every row. */
class ColumnLetters {
  private readonly letters: string[] = [];

  at(index: number): string {
    if (index >= MAX_COLUMNS) throw new RangeError(`column ${index + 1} exceeds the ${MAX_COLUMNS}-column limit (XFD)`);
    for (let i = this.letters.length; i <= index; i++) {
      let s = '';
      for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
      this.letters.push(s);
    }
    return this.letters[index]!;
  }
}

function checkSheetName(name: string): void {
  if (name.length < 1 || name.length > 31 || SHEET_NAME_ILLEGAL.test(name) || name.startsWith("'") || name.endsWith("'")) {
    throw new Error(`invalid sheet name ${JSON.stringify(name)}: must be 1-31 chars without [ ] : * ? / \\ and not start or end with an apostrophe`);
  }
}

/** @internal tests pass `fallbackCompressor` to exercise the non-Node path. */
export function writeXlsx(sheet: SheetInput, compressor: Compressor = defaultCompressor): Uint8Array {
  checkSheetName(sheet.name);

  const strings = new SharedStrings();
  const styles = new Styles();
  const letters = new ColumnLetters();
  const numberXf = sheet.columns.map((c) => styles.xfForFormat(c.format));
  const dateXf = sheet.columns.map((c) => (c.format ? styles.xfForFormat(c.format) : Styles.DATE_XF));

  const sheetChunks: Uint8Array[] = [];
  let buf = XML_DECLARATION + `<worksheet xmlns="${NS_MAIN}"><sheetData>`;
  let rowNum = 0;

  const writeRow = (cells: readonly CellValue[], isHeader: boolean): void => {
    rowNum++;
    if (rowNum > MAX_ROWS) throw new RangeError(`row ${rowNum} exceeds the ${MAX_ROWS}-row limit`);
    buf += `<row r="${rowNum}">`;
    for (let j = 0; j < cells.length; j++) {
      const v = cells[j];
      const ref = letters.at(j) + rowNum;
      const nx = isHeader ? Styles.DEFAULT_XF : (numberXf[j] ?? Styles.DEFAULT_XF);
      switch (typeof v) {
        case 'string':
          buf += `<c r="${ref}" t="s"${nx ? ` s="${nx}"` : ''}><v>${strings.add(v)}</v></c>`;
          break;
        case 'number':
          // NaN and ±Infinity have no XLSX representation; a written "NaN" makes Excel report a corrupt file.
          if (Number.isFinite(v)) buf += `<c r="${ref}"${nx ? ` s="${nx}"` : ''}><v>${v}</v></c>`;
          break;
        case 'boolean':
          buf += `<c r="${ref}" t="b"${nx ? ` s="${nx}"` : ''}><v>${v ? 1 : 0}</v></c>`;
          break;
        case 'object':
          if (v instanceof Date) {
            const ms = v.getTime();
            if (Number.isFinite(ms)) {
              buf += `<c r="${ref}" s="${dateXf[j] ?? Styles.DATE_XF}"><v>${dateToSerial1900(ms)}</v></c>`;
            }
            break;
          }
          if (v === null) break;
          if (typeof (v as CellError).error === 'string') {
            const code = (v as CellError).error;
            if (!ERROR_CODES.has(code)) throw new TypeError(`unsupported error value ${JSON.stringify(code)} at ${ref}`);
            buf += `<c r="${ref}" t="e"${nx ? ` s="${nx}"` : ''}><v>${code}</v></c>`;
            break;
          }
          throw new TypeError(`unsupported cell value at ${ref}`);
        case 'undefined':
          break;
        default:
          throw new TypeError(`unsupported cell value of type ${typeof v} at ${ref}`);
      }
    }
    buf += '</row>';
    if (buf.length >= CHUNK_CHARS) {
      sheetChunks.push(strToU8(buf));
      buf = '';
    }
  };

  writeRow(sheet.columns.map((c) => c.header), true);
  for (const r of sheet.rows) writeRow(r, false);

  // No <dimension>: it is optional and needs the final extent, which is unknown while rows stream out.
  buf += '</sheetData></worksheet>';
  sheetChunks.push(strToU8(buf));

  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      XML_DECLARATION +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
        '</Types>',
    ),
    '_rels/.rels': strToU8(
      XML_DECLARATION +
        `<Relationships xmlns="${NS_PKG_REL}">` +
        `<Relationship Id="rId1" Type="${NS_REL}/officeDocument" Target="xl/workbook.xml"/>` +
        '</Relationships>',
    ),
    'xl/workbook.xml': strToU8(
      XML_DECLARATION +
        `<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}"><sheets><sheet name="${escapeXml(sheet.name)}" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      XML_DECLARATION +
        `<Relationships xmlns="${NS_PKG_REL}">` +
        `<Relationship Id="rId1" Type="${NS_REL}/worksheet" Target="worksheets/sheet1.xml"/>` +
        `<Relationship Id="rId2" Type="${NS_REL}/styles" Target="styles.xml"/>` +
        `<Relationship Id="rId3" Type="${NS_REL}/sharedStrings" Target="sharedStrings.xml"/>` +
        '</Relationships>',
    ),
    'xl/styles.xml': strToU8(styles.toXml()),
    'xl/sharedStrings.xml': strings.toBytes(),
  };
  return writeZip(
    [{ name: 'xl/worksheets/sheet1.xml', data: concat(sheetChunks) }, ...Object.entries(files).map(([name, data]) => ({ name, data }))],
    compressor,
    levelFor(compressor),
  );
}

export { readXlsx, type ReadSheet, type ReadWorkbook } from './read.js';
export { XlsxLimitError, type XlsxLimits, type ReadOptions } from './limits.js';
