import { describe, expect, it } from 'vitest';
import { Workbook, ValueType } from '../../src/compat/index.js';

const sheet = () => new Workbook().addWorksheet('S');

describe('workbook and worksheet', () => {
  it('addWorksheet / getWorksheet / eachSheet / removeWorksheet', () => {
    const wb = new Workbook();
    const a = wb.addWorksheet('A');
    const b = wb.addWorksheet('B');
    expect(wb.getWorksheet('B')).toBe(b);
    expect(wb.getWorksheet(1)).toBe(a);
    expect(wb.getWorksheet()).toBe(a);
    const seen: string[] = [];
    wb.eachSheet((ws, id) => seen.push(`${ws.name}:${id}`));
    expect(seen).toEqual(['A:1', 'B:2']);
    wb.removeWorksheet('A');
    expect(wb.worksheets.map((w) => w.name)).toEqual(['B']);
  });
  it('validates sheet names', () => {
    const wb = new Workbook();
    wb.addWorksheet('Dup');
    expect(() => wb.addWorksheet('dup')).toThrow(/already exists/);
    expect(() => wb.addWorksheet('a:b')).toThrow(/cannot include/);
    expect(() => wb.addWorksheet("'x")).toThrow(/quotation/);
    expect(() => wb.addWorksheet('')).toThrow(/empty/);
  });
  it('columns create header row and keyed rows', () => {
    const ws = sheet();
    ws.columns = [{ header: 'Id', key: 'id', width: 8 }, { header: 'Name', key: 'name' }];
    ws.addRow({ id: 1, name: 'x' });
    expect(ws.getRow(1).values).toEqual([, 'Id', 'Name']);
    expect(ws.getRow(2).values).toEqual([, 1, 'x']);
    expect(ws.getColumn('id').width).toBe(8);
    expect(ws.getColumn(2)).toBe(ws.columns![1]);
  });
});

describe('rows and cells', () => {
  it('addRow array shapes: own index 0 means 0-based, otherwise 1-based sparse', () => {
    const ws = sheet();
    ws.addRow([1, 'a']);
    ws.addRow([, 2, 'b']);
    expect(ws.getRow(1).values).toEqual([, 1, 'a']);
    expect(ws.getRow(2).values).toEqual([, 2, 'b']);
  });
  it('getCell creates and has identity; findCell does not create', () => {
    const ws = sheet();
    expect(ws.findCell('B2')).toBeUndefined();
    const c = ws.getCell('B2');
    expect(ws.getCell(2, 2)).toBe(c);
    expect(c.address).toBe('B2');
    expect(c.value).toBeNull();
    expect(c.type).toBe(ValueType.Null);
    expect(ws.getRow(2).cellCount).toBe(2);
  });
  it('classifies value types', () => {
    const ws = sheet();
    const t = (v: unknown) => {
      ws.getCell('A1').value = v;
      return ws.getCell('A1').type;
    };
    expect(t(1)).toBe(ValueType.Number);
    expect(t('s')).toBe(ValueType.String);
    expect(t(true)).toBe(ValueType.Boolean);
    expect(t(new Date())).toBe(ValueType.Date);
    expect(t({ formula: 'A2' })).toBe(ValueType.Formula);
    expect(t({ text: 'x', hyperlink: 'http://a' })).toBe(ValueType.Hyperlink);
    expect(t({ error: '#N/A' })).toBe(ValueType.Error);
    expect(t({ richText: [{ text: 'a' }] })).toBe(ValueType.RichText);
    expect(t({ a: 1 })).toBe(ValueType.String);
    expect(ws.getCell('A1').value).toEqual({ a: 1 });
  });
  it('eachRow skips rows without values; eachCell skips empty cells', () => {
    const ws = sheet();
    ws.getCell('C3').value = 1;
    ws.getCell('A1');
    const rows: number[] = [];
    ws.eachRow((r, n) => rows.push(n));
    expect(rows).toEqual([3]);
    const cols: number[] = [];
    ws.getRow(3).eachCell((c, n) => cols.push(n));
    expect(cols).toEqual([3]);
    const all: number[] = [];
    ws.getRow(3).eachCell({ includeEmpty: true }, (c, n) => all.push(n));
    expect(all).toEqual([1, 2, 3]);
  });
  it('getRows returns undefined below 1; rowCount ignores holes at the end', () => {
    const ws = sheet();
    expect(ws.getRows(1, 0)).toBeUndefined();
    ws.getRow(5);
    expect(ws.rowCount).toBe(5);
  });
  it('formula cell exposes formula, result and shared-formula translation', () => {
    const ws = sheet();
    ws.getCell('A1').value = { formula: 'B1+1', result: 3 };
    ws.getCell('A2').value = { sharedFormula: 'A1', result: 4 };
    expect(ws.getCell('A1').formula).toBe('B1+1');
    expect(ws.getCell('A1').result).toBe(3);
    expect(ws.getCell('A2').formula).toBe('B2+1');
  });
});

describe('merges', () => {
  it('slaves report the master value, writes go to the master, overlap throws', () => {
    const ws = sheet();
    ws.getCell('A1').value = 5;
    ws.mergeCells('A1:B2');
    expect(ws.getCell('B2').value).toBe(5);
    expect(ws.getCell('B2').type).toBe(ValueType.Merge);
    expect(ws.getCell('B2').master).toBe(ws.getCell('A1'));
    ws.getCell('B2').value = 9;
    expect(ws.getCell('A1').value).toBe(9);
    expect(() => ws.mergeCells('B2:C3')).toThrow(/already merged/);
    ws.unMergeCells('A1');
    expect(ws.getCell('B2').type).toBe(ValueType.Null);
    expect(ws.getCell('B2').isMerged).toBe(false);
  });
});

describe('row operations', () => {
  it('spliceRows, insertRow, duplicateRow keep values and styles with the rows', () => {
    const ws = sheet();
    ws.addRow([1]);
    ws.addRow([2]);
    ws.addRow([3]);
    ws.getCell('A3').font = { bold: true };
    ws.insertRow(2, ['new']);
    expect(ws.getSheetValues().slice(1).map((r) => r[1])).toEqual([1, 'new', 2, 3]);
    expect(ws.getCell('A4').font).toEqual({ bold: true });
    ws.spliceRows(1, 1);
    expect(ws.getRow(1).values).toEqual([, 'new']);
    ws.duplicateRow(1, 2, true);
    expect(ws.rowCount).toBe(5);
  });
  it('row.splice shifts cells', () => {
    const ws = sheet();
    const r = ws.addRow(['a', 'b', 'c']);
    r.splice(2, 1, 'x', 'y');
    expect(r.values).toEqual([, 'a', 'x', 'y', 'c']);
  });
});
