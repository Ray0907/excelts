import ExcelJS from 'exceljs';
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { readXlsx, writeXlsx, type CellError, type CellValue } from '../src/index.js';
import { workbookOf } from './handBuilt.js';

const CODES = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA'];
const rows = (bytes: Uint8Array): CellValue[][] => [...readXlsx(bytes).sheets[0]!.rows()];

describe('error cells', () => {
  it('reads t="e" as { error }', () => {
    const cells = CODES.map((c, i) => `<c r="${String.fromCharCode(65 + i)}1" t="e"><f>1/0</f><v>${c}</v></c>`).join('');
    expect(rows(workbookOf(`<row r="1">${cells}</row>`))[0]).toEqual(CODES.map((error) => ({ error })));
  });

  it('writes t="e" and reads all 8 codes back', () => {
    const bytes = writeXlsx({ name: 'e', columns: [{ header: 'h' }], rows: [CODES.map((error) => ({ error }))] });
    const sheet = strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']!);
    expect(sheet).toContain('<c r="A2" t="e"><v>#NULL!</v></c>');
    expect(sheet).toContain('<c r="B2" t="e"><v>#DIV/0!</v></c>');
    expect(rows(bytes)[1]).toEqual(CODES.map((error) => ({ error })));
  });

  it('an error is not a string and not a Date', () => {
    const v = rows(workbookOf('<row r="1"><c r="A1" t="e"><v>#N/A</v></c></row>'))[0]![0] as CellError;
    expect(typeof v).toBe('object');
    expect(v.error).toBe('#N/A');
  });

  it('exceljs reads what we write', async () => {
    const bytes = writeXlsx({ name: 'e', columns: [{ header: 'h' }], rows: [[{ error: '#DIV/0!' }]] });
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(bytes as unknown as Buffer);
    expect(wb.worksheets[0]!.getCell('A2').value).toEqual({ error: '#DIV/0!' });
  });

  it.each(['#FOO', '', '#div/0!', '#N/A ', 'N/A'])('rejects the unknown error text %j', (error) => {
    expect(() => writeXlsx({ name: 'e', columns: [{ header: 'h' }], rows: [[{ error }]] })).toThrow(TypeError);
  });

  it('rejects a non-string error', () => {
    expect(() => writeXlsx({ name: 'e', columns: [{ header: 'h' }], rows: [[{ error: 7 } as unknown as CellValue]] })).toThrow(TypeError);
  });
});
