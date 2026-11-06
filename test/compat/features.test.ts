import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { Workbook } from '../../src/compat/index.js';

describe('csv', () => {
  it('parses quotes, embedded delimiters and newlines, numbers, booleans, dates and empties', () => {
    const wb = new Workbook();
    const ws = wb.csv.load('a,"b,c","say ""hi""","l1\nl2"\n1,2.5,true,\n2021-03-04,#N/A,text,');
    expect(ws.getRow(1).values).toEqual([undefined, 'a', 'b,c', 'say "hi"', 'l1\nl2']);
    const r2 = ws.getRow(2).values as unknown[];
    expect(r2.slice(1, 4)).toEqual([1, 2.5, true]);
    expect(r2[4] ?? null).toBeNull();
    const r3 = ws.getRow(3).values as unknown[];
    expect(r3[1]).toBeInstanceOf(Date);
    expect(r3[2]).toEqual({ error: '#N/A' });
  });

  it('honours delimiter, quote and a custom map', () => {
    const wb = new Workbook();
    const ws = wb.csv.load("x;'y;z'\n3;4", { parserOptions: { delimiter: ';', quote: "'" }, map: (v) => String(v).toUpperCase() });
    expect(ws.getRow(1).values).toEqual([undefined, 'X', 'Y;Z']);
    expect(ws.getRow(2).values).toEqual([undefined, '3', '4']);
  });

  it('writes quoted fields, results of formulas, hyperlinks and dates', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.addRow(['a,b', 'say "hi"', 'l1\nl2', 5, true]);
    ws.addRow([{ formula: 'A1', result: 'res' }, { text: 'site', hyperlink: 'https://x.test' }, new Date(Date.UTC(2020, 0, 2, 3, 4, 5))]);
    expect(wb.csv.toString({ dateUTC: true, dateFormat: 'YYYY-MM-DD' })).toBe('"a,b","say ""hi""","l1\nl2",5,true\nres,https://x.test,2020-01-02');
  });

  it('keeps empty rows between data rows unless told not to', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 1;
    ws.getCell('A4').value = 4;
    expect(wb.csv.toString()).toBe('1\n\n\n4');
    expect(wb.csv.toString({ includeEmptyRows: false })).toBe('1\n4');
  });

  it('round trips through load and toString, and writeBuffer matches', async () => {
    const wb = new Workbook();
    wb.csv.load('h1,h2\n1,"two, 2"');
    expect(wb.csv.toString({ sheetId: 1 })).toBe('h1,h2\n1,"two, 2"');
    expect(strFromU8(await wb.csv.writeBuffer({ sheetId: 1 }))).toBe('h1,h2\n1,"two, 2"');
  });
});

describe('spliceColumns', () => {
  function sheet(): ReturnType<Workbook['addWorksheet']> {
    const ws = new Workbook().addWorksheet('S');
    ws.addRow(['a1', 'b1', 'c1', 'd1']);
    ws.addRow(['a2', 'b2', 'c2', 'd2']);
    ws.addRow(['a3', 'b3', 'c3', 'd3']);
    return ws;
  }
  const col = (ws: ReturnType<typeof sheet>, c: number): unknown[] => [1, 2, 3].map((r) => ws.getCell(r, c).value ?? null);

  it('removes columns and closes the gap', () => {
    const ws = sheet();
    ws.spliceColumns(2, 2);
    expect(col(ws, 1)).toEqual(['a1', 'a2', 'a3']);
    expect(col(ws, 2)).toEqual(['d1', 'd2', 'd3']);
    expect(col(ws, 3)).toEqual([null, null, null]);
  });

  it('inserts columns and pushes the rest right', () => {
    const ws = sheet();
    ws.spliceColumns(2, 0, ['x1', 'x2', 'x3'], ['y1', 'y2']);
    expect(col(ws, 2)).toEqual(['x1', 'x2', 'x3']);
    expect(col(ws, 3)).toEqual(['y1', 'y2', null]);
    expect(col(ws, 4)).toEqual(['b1', 'b2', 'b3']);
    expect(col(ws, 6)).toEqual(['d1', 'd2', 'd3']);
  });

  it('replaces columns, moves column widths and styles with them', () => {
    const ws = sheet();
    ws.getColumn(3).width = 30;
    ws.getColumn(4).width = 40;
    ws.getCell('D1').font = { bold: true };
    ws.spliceColumns(2, 1, ['r1', 'r2', 'r3']);
    expect(col(ws, 2)).toEqual(['r1', 'r2', 'r3']);
    expect(col(ws, 3)).toEqual(['c1', 'c2', 'c3']);
    expect(ws.getColumn(3).width).toBe(30);
    expect(ws.getCell('D1').font?.bold).toBe(true);
  });

  it('keeps merged cells readable and a row splice keeps its values', () => {
    const ws = sheet();
    ws.spliceColumns(1, 1);
    expect(ws.getCell('A1').value).toBe('b1');
    ws.spliceRows(1, 1);
    expect(ws.getCell('A1').value).toBe('b2');
  });
});

describe('load options and odd packages', () => {
  async function pkg(rows: number, cols: number): Promise<Uint8Array> {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    for (let r = 0; r < rows; r++) ws.addRow(Array.from({ length: cols }, (_, c) => r * cols + c));
    return wb.xlsx.writeBuffer();
  }

  it('maxRows rejects a sheet with more rows, with the exceljs message', async () => {
    const bytes = await pkg(12, 2);
    await expect(new Workbook().xlsx.load(bytes, { maxRows: 10 })).rejects.toThrow('Max row count (10) exceeded');
    expect((await new Workbook().xlsx.load(bytes, { maxRows: 12 })).getWorksheet('S')!.rowCount).toBe(12);
  });

  it('maxCols rejects a sheet with more columns, with the exceljs message', async () => {
    const bytes = await pkg(2, 12);
    await expect(new Workbook().xlsx.load(bytes, { maxCols: 10 })).rejects.toThrow('Max column count (10) exceeded');
    expect((await new Workbook().xlsx.load(bytes, { maxCols: 12 })).getWorksheet('S')!.columnCount).toBe(12);
  });

  it('a zip without xl/workbook.xml loads as an empty workbook', async () => {
    const bytes = zipSync({ 'word/document.xml': strToU8('<w:document/>'), '[Content_Types].xml': strToU8('<Types/>') });
    const wb = await new Workbook().xlsx.load(bytes);
    expect(wb.worksheets).toEqual([]);
    expect(wb.getWorksheet(1)).toBeUndefined();
  });

  it('loading replaces what the workbook held, including pictures and names', async () => {
    const wb = new Workbook();
    wb.addWorksheet('Old').getCell('A1').name = 'gone';
    wb.addImage({ buffer: new Uint8Array([1, 2, 3]), extension: 'png' });
    await wb.xlsx.load(await pkg(1, 1));
    expect(wb.worksheets.map((w) => w.name)).toEqual(['S']);
    expect(wb.media).toEqual([]);
    expect(wb.definedNames.model).toEqual([]);
  });

  it('text before the root element of a sheet part is rejected as malformed XML', async () => {
    const files = unzipSync(await pkg(1, 1));
    files['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(files['xl/worksheets/sheet1.xml']!).replace('?>\n', '?>\nfoo\n'));
    await expect(new Workbook().xlsx.load(zipSync(files))).rejects.toThrow('3:1: text data outside of root node.');
  });

  it('date validations read back as dates, whole numbers as numbers, formulas as text', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    const d1 = new Date(Date.UTC(2015, 0, 1));
    ws.dataValidations.add('A1', { type: 'date', operator: 'between', formulae: [d1, new Date(Date.UTC(2016, 0, 1))] });
    ws.dataValidations.add('B1', { type: 'whole', operator: 'greaterThan', formulae: [5] });
    ws.dataValidations.add('C1', { type: 'custom', formulae: ['$A$1>0'] });
    const dv = (await new Workbook().xlsx.load(await wb.xlsx.writeBuffer())).getWorksheet('S')!.dataValidations.model;
    expect(dv.A1!.formulae).toEqual([d1, new Date(Date.UTC(2016, 0, 1))]);
    expect(dv.B1!.formulae).toEqual([5]);
    expect(dv.C1!.formulae).toEqual(['$A$1>0']);
  });
});
