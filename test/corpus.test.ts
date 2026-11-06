import { existsSync, readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { readXlsx, type CellValue } from '../src/index.js';

// Expected values are hard-coded from what the oracle's openpyxl and calamine adapters reported (oracle/out/*.ndjson), not computed here.
const RAW = new URL('../corpus/raw/', import.meta.url);
const have = existsSync(RAW);

function cell(file: string, sheet: number, ref: string): CellValue {
  const m = /^([A-Z]+)(\d+)$/.exec(ref)!;
  const col = [...m[1]!].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const rows = [...readXlsx(new Uint8Array(readFileSync(new URL(file, RAW)))).sheets[sheet]!.rows()];
  return rows[Number(m[2]) - 1]?.[col];
}
const ms = (v: CellValue): number => (v as Date).getTime();

describe.skipIf(!have)('real-world files', () => {
  it('calamine tests_date.xlsx: [hh]:mm:ss is elapsed time since 1899-12-30', () => {
    const f = 'calamine/tests_date.xlsx';
    expect(ms(cell(f, 0, 'A3'))).toBe(-2208242990000);
    expect(ms(cell(f, 0, 'A2'))).toBe(1609545600000);
  });

  it('calamine tests_date_1904.xlsx: a duration ignores date1904', () => {
    expect(ms(cell('calamine/tests_date_1904.xlsx', 0, 'A3'))).toBe(-2208242990000);
  });

  it('calamine tests_date_east_asian.xlsx: 0.5 in the time formats 20, 46, [h]:mm:ss and hh:mm:ss', () => {
    const f = 'calamine/tests_date_east_asian.xlsx';
    for (const ref of ['B22', 'B24', 'B26', 'B27']) expect(ms(cell(f, 0, ref)), ref).toBe(-2209118400000);
    expect(ms(cell(f, 0, 'C24'))).toBe(1705287466667);
    expect(ms(cell(f, 0, 'A22'))).toBe(1705276800000);
    expect(cell(f, 0, 'B28')).toBe(0.5);
  });

  it('calamine tests_errors.xlsx: error cells', () => {
    const f = 'calamine/tests_errors.xlsx';
    const got = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7'].map((r) => cell(f, 0, r));
    expect(got).toEqual(['#DIV/0!', '#NAME?', '#VALUE!', '#NULL!', '#REF!', '#NUM!', '#N/A'].map((error) => ({ error })));
  });

  it('rust_xlsxwriter tests_input_escapes09.xlsx: _x001B_ is ESC', () => {
    expect(cell('rust_xlsxwriter/tests_input_escapes09.xlsx', 0, 'A1')).toBe('Data\u001b[32m1');
  });

  it('exceljs-spec shared_string_with_escape.xlsx: the escaped escape', () => {
    expect(cell('exceljs-spec/shared_string_with_escape.xlsx', 0, 'A1')).toBe('_x000D_');
  });

  it('calamine tests_inlineStr_cdata.xlsx: CDATA in inline strings', () => {
    const f = 'calamine/tests_inlineStr_cdata.xlsx';
    expect(['A1', 'B1', 'A2', 'B2'].map((r) => cell(f, 0, r))).toEqual(['Hello CDATA', 'World', '12345', 'NormalText']);
  });

  it('closedxml files: CRLF in text becomes LF', () => {
    expect(cell('closedxml/ClosedXML.Tests_Resource_Other_RichText_kanji-with-new-line-input.xlsx', 0, 'B3')).toBe('\n漢字\n');
    expect(cell('closedxml/ClosedXML.Tests_Resource_Other_StyleReferenceFiles_RowAndColumnStyles_inputfile.xlsx', 0, 'F4')).toBe('Items\ntotal');
    expect(cell('closedxml/ClosedXML.Tests_Resource_TryToLoad_TableHeadersWithLineBreaks.xlsx', 0, 'B1')).toBe('Room/\narea');
  });
});
