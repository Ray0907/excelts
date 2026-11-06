import { describe, expect, it } from 'vitest';
import { NotSupportedError, Workbook } from '../../src/compat/index.js';

const sheet = () => new Workbook().addWorksheet('S');

describe('write-through style objects', () => {
  it('cell.font.bold = true rewrites only that cell', () => {
    const ws = sheet();
    ws.getCell('A1').font = { name: 'Arial', size: 10 };
    ws.getCell('A2').font = { name: 'Arial', size: 10 };
    ws.getCell('A1').font!.bold = true;
    expect(ws.getCell('A1').font).toEqual({ name: 'Arial', size: 10, bold: true });
    expect(ws.getCell('A2').font).toEqual({ name: 'Arial', size: 10 });
  });
  it('one level: font.color = {...} and alignment.vertical work', () => {
    const ws = sheet();
    const c = ws.getCell('A1');
    c.font = { bold: true };
    c.font!.color = { argb: 'FFFF0000' };
    c.alignment = { horizontal: 'left' };
    c.alignment!.vertical = 'middle';
    expect(c.font).toEqual({ bold: true, color: { argb: 'FFFF0000' } });
    expect(c.alignment).toEqual({ horizontal: 'left', vertical: 'middle' });
  });
  it('deeper mutation throws NotSupportedError naming the feature', () => {
    const ws = sheet();
    const c = ws.getCell('A1');
    c.font = { color: { argb: 'FF000000' } };
    expect(() => {
      (c.font!.color as { argb: string }).argb = 'FFFFFFFF';
    }).toThrow(NotSupportedError);
    expect(c.font).toEqual({ color: { argb: 'FF000000' } });
  });
  it('unset components read as undefined; style assignment replaces the whole style', () => {
    const ws = sheet();
    const c = ws.getCell('A1');
    expect(c.font).toBeUndefined();
    expect(c.alignment).toBeUndefined();
    c.style = { numFmt: '0.00', fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00FF00' } } };
    expect(c.numFmt).toBe('0.00');
    c.style = { font: { italic: true } };
    expect(c.numFmt).toBeUndefined();
    expect(c.style).toEqual({ font: { italic: true } });
  });
  it('invalid alignment values are dropped as exceljs does', () => {
    const c = sheet().getCell('A1');
    c.alignment = { textRotation: -91, indent: -1 };
    expect(c.alignment).toBeUndefined();
  });
});

describe('row and column style inheritance', () => {
  it('column style reaches later cells first; retroactive only for column.font = ...', () => {
    const ws = sheet();
    ws.columns = [{ key: 'a', style: { numFmt: '0.0' } }, { key: 'b' }];
    ws.addRow({ a: 1, b: 2 });
    expect(ws.getCell('A1').numFmt).toBe('0.0');
    expect(ws.getCell('B1').numFmt).toBeUndefined();
    ws.getColumn('b').style = { numFmt: '0%' };
    expect(ws.getCell('B1').numFmt).toBeUndefined();
    ws.addRow({ a: 3, b: 4 });
    expect(ws.getCell('B2').numFmt).toBe('0%');
    ws.getColumn('b').font = { bold: true };
    expect(ws.getCell('B1').font).toEqual({ bold: true });
  });
  it('row style applies to cells created afterwards; row.font reaches existing ones', () => {
    const ws = sheet();
    const r = ws.getRow(1);
    r.getCell(1).value = 'x';
    r.font = { italic: true };
    r.getCell(2).value = 'y';
    expect(r.getCell(1).font).toEqual({ italic: true });
    expect(r.getCell(2).font).toEqual({ italic: true });
  });
});
