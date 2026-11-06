import { describe, expect, it } from 'vitest';
import { readXlsx, XlsxLimitError } from '../src/index.js';
import { workbookOf } from './handBuilt.js';

const rowsOf = (sheetData: string, limits = {}) => [...readXlsx(workbookOf(sheetData), { limits }).sheets[0]!.rows()];

describe('rows(): gaps between and before numbered rows', () => {
  it('yields row N as the Nth array, with an empty array for every skipped row', () => {
    const rows = rowsOf('<row r="3"><c r="A3"><v>7</v></c></row><row r="6"><c r="B6"><v>8</v></c></row>');
    expect(rows).toEqual([[], [], [7], [], [], [undefined, 8]]);
  });

  it('gives each gap row its own mutable array', () => {
    const rows = rowsOf('<row r="4"><c r="A4"><v>1</v></c></row>');
    expect(new Set(rows.slice(0, 3)).size).toBe(3);
    rows[0]!.push(99);
    expect(rows[1]).toEqual([]);
  });

  it('is a plain iterator: a fresh result per step, early exit, and re-reading restarts', () => {
    const wb = readXlsx(workbookOf('<row r="5"><c r="A5"><v>1</v></c></row>'));
    const it = wb.sheets[0]!.rows();
    expect(it[Symbol.iterator]()).toBe(it);
    const a = it.next();
    const b = it.next();
    expect(a).not.toBe(b);
    expect(a).toEqual({ value: [], done: false });
    for (const _ of wb.sheets[0]!.rows()) break;
    expect([...wb.sheets[0]!.rows()]).toHaveLength(5);
    expect(it.return?.()).toEqual({ value: undefined, done: true });
    expect(it.next().done).toBe(true);
  });

  it('charges a long gap to maxGapRows before yielding any of it', () => {
    const wb = readXlsx(workbookOf('<row r="1000"/>'), { limits: { maxGapRows: 998 } });
    const it = wb.sheets[0]!.rows();
    expect(() => it.next()).toThrow(XlsxLimitError);
  });

  it('reads a sheet that ends with the last Excel row without allocating a row object per gap row being a requirement of the caller', () => {
    const rows = rowsOf('<row r="2"><c r="A2"><v>1</v></c></row><row r="1048576"/>');
    expect(rows).toHaveLength(1048576);
    expect(rows[1]).toEqual([1]);
    expect(rows[1048575]).toEqual([]);
  });
});
