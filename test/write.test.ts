import { describe, expect, it } from 'vitest';
import ExcelJS from 'exceljs';
import { unzipSync, strFromU8 } from 'fflate';
import { writeXlsx, type CellValue } from '../src/index.js';
import { fallbackCompressor } from '../src/zip.js';
// @ts-expect-error untyped bench module
import { COLUMNS, row } from './helpers/data.mjs';

async function load(bytes: Uint8Array) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(Buffer.from(bytes) as unknown as ArrayBuffer);
  return wb.worksheets[0]!;
}

function sheetOf(columns: { header: string; format?: string }[], rows: Iterable<readonly CellValue[]>, name = 'S') {
  return writeXlsx({ name, columns, rows });
}

describe('writeXlsx round trip', () => {
  it('strings', async () => {
    const values = ['a & <b> "q"', 'héllo 世界 😀', '  lead', 'trail  ', 'line1\nline2', ''];
    const ws = await load(sheetOf([{ header: 'h' }], values.map((v) => [v])));
    expect(ws.getCell('A1').value).toBe('h');
    values.forEach((v, i) => {
      const got = ws.getCell(i + 2, 1).value;
      expect(got ?? '').toBe(v);
    });
  });

  it('numbers', async () => {
    const values = [0, 1, 42, -7, 3.14159, -0.5, 1e21, 1e-7];
    const ws = await load(sheetOf([{ header: 'n' }], values.map((v) => [v])));
    values.forEach((v, i) => expect(ws.getCell(i + 2, 1).value).toBe(v));
  });

  it('booleans', async () => {
    const ws = await load(sheetOf([{ header: 'b' }], [[true], [false]]));
    expect(ws.getCell('A2').value).toBe(true);
    expect(ws.getCell('A3').value).toBe(false);
  });

  it('dates keep their ms and take the column format', async () => {
    const d = new Date(Date.UTC(2021, 5, 15, 13, 45, 30, 250));
    const fmt = 'yyyy-mm-dd hh:mm:ss';
    const ws = await load(sheetOf([{ header: 'd', format: fmt }], [[d]]));
    const cell = ws.getCell('A2');
    expect((cell.value as Date).getTime()).toBe(d.getTime());
    expect(cell.numFmt).toBe(fmt);
  });

  it('dates in a column with no format use built-in format 22', async () => {
    const d = new Date(Date.UTC(2021, 5, 15, 13, 45));
    const ws = await load(sheetOf([{ header: 'd' }], [[d]]));
    const cell = ws.getCell('A2');
    expect(cell.value).toBeInstanceOf(Date);
    expect((cell.value as Date).getTime()).toBe(d.getTime());
  });

  it('applies a column format to numbers', async () => {
    const ws = await load(sheetOf([{ header: 'p', format: '0.00%' }, { header: 'q' }], [[0.5, 1]]));
    expect(ws.getCell('A2').numFmt).toBe('0.00%');
    expect(ws.getCell('B2').numFmt).toBeUndefined();
  });

  it('null and undefined leave gaps', async () => {
    const ws = await load(sheetOf([{ header: 'a' }, { header: 'b' }, { header: 'c' }], [[1, null, 3], [undefined, 'x', undefined]]));
    expect(ws.getCell('A2').value).toBe(1);
    expect(ws.getCell('B2').value).toBeNull();
    expect(ws.getCell('C2').value).toBe(3);
    expect(ws.getCell('A3').value).toBeNull();
    expect(ws.getCell('B3').value).toBe('x');
    expect(ws.getCell('C3').value).toBeNull();
  });

  it('empty rows yields header only', async () => {
    const ws = await load(sheetOf([{ header: 'a' }, { header: 'b' }], []));
    expect(ws.rowCount).toBe(1);
    expect(ws.getCell('A1').value).toBe('a');
    expect(ws.getCell('B1').value).toBe('b');
  });

  it('non-finite numbers and invalid dates become empty cells', async () => {
    const ws = await load(sheetOf([{ header: 'a' }], [[NaN], [Infinity], [-Infinity], [new Date(NaN)], [5]]));
    for (const r of [2, 3, 4, 5]) expect(ws.getCell(r, 1).value).toBeNull();
    expect(ws.getCell('A6').value).toBe(5);
  });

  it('accepts a generator of rows', async () => {
    function* gen() {
      for (let i = 0; i < 5; i++) yield [i, `s${i}`] as const;
    }
    const ws = await load(sheetOf([{ header: 'i' }, { header: 's' }], gen()));
    expect(ws.rowCount).toBe(6);
    expect(ws.getCell('B6').value).toBe('s4');
  });

  it('allows rows wider than columns, with general column letters', async () => {
    const wide = Array.from({ length: 30 }, (_, i) => i);
    const ws = await load(sheetOf([{ header: 'a' }], [wide]));
    expect(ws.getCell('AA2').value).toBe(26);
    expect(ws.getCell('AD2').value).toBe(29);
  });

  it('keeps XML-illegal control characters as _xHHHH_ escapes', async () => {
    const ws = await load(sheetOf([{ header: 'h' }], [['a\u0000b\u0008c\td']]));
    expect(ws.getCell('A2').value).toBe('a\u0000b\u0008c\td');
  });
});

describe('shared strings', () => {
  it('dedupes repeated strings', () => {
    const bytes = sheetOf([{ header: 'h' }], [['x'], ['x'], ['x'], ['y']]);
    const xml = strFromU8(unzipSync(bytes)['xl/sharedStrings.xml']!);
    const count = Number(/ count="(\d+)"/.exec(xml)![1]);
    const unique = Number(/uniqueCount="(\d+)"/.exec(xml)![1]);
    expect(unique).toBeLessThan(count);
    expect(unique).toBe(3);
    expect(count).toBe(5);
  });
});

describe('validation', () => {
  it.each(['', 'a'.repeat(32), 'a[b', 'a]b', 'a:b', 'a*b', 'a?b', 'a/b', 'a\\b'])('rejects sheet name %j', (name) => {
    expect(() => sheetOf([{ header: 'h' }], [], name)).toThrow();
  });

  it('accepts a 31-char sheet name', () => {
    expect(() => sheetOf([{ header: 'h' }], [], 'a'.repeat(31))).not.toThrow();
  });

  it('throws RangeError for a 33000-char string', () => {
    expect(() => sheetOf([{ header: 'h' }], [['x'.repeat(33000)]])).toThrow(RangeError);
  });

  it('throws RangeError past column XFD', () => {
    const wide = new Array(16385).fill(1);
    expect(() => sheetOf([{ header: 'h' }], [wide])).toThrow(RangeError);
  });
});

describe('bench-shaped dataset', () => {
  it('opens with exceljs for 1000 mixed rows', async () => {
    const rows = Array.from({ length: 1000 }, (_, i) => row(i) as CellValue[]);
    const ws = await load(sheetOf(COLUMNS, rows, 'Data'));
    expect(ws.rowCount).toBe(1001);
    expect(ws.name).toBe('Data');
    const src = row(999) as CellValue[];
    expect(ws.getCell(1001, 1).value).toBe(src[0]);
    expect(ws.getCell(1001, 2).value).toBe(src[1]);
    expect((ws.getCell(1001, 5).value as Date).getTime()).toBe((src[4] as Date).getTime());
    expect(ws.getCell(1001, 6).value).toBe(src[5]);
    expect(ws.getCell(1001, 10).value).toBe(src[9]);
  });
});

describe('zip container', () => {
  const columns = [{ header: 'id' }, { header: 'name' }, { header: 'when', format: 'yyyy-mm-dd' }];
  const rows = Array.from({ length: 500 }, (_, i) => [i, `name ${i % 17} é`, new Date(Date.UTC(2020, 0, 1 + i))] as CellValue[]);

  it('fallback compressor output opens with exceljs', async () => {
    const ws = await load(writeXlsx({ name: 'S', columns, rows }, fallbackCompressor));
    expect(ws.getCell('B2').value).toBe('name 0 é');
    expect(ws.getCell(501, 1).value).toBe(499);
  });

  it('is byte-identical across calls with the same input', () => {
    const a = writeXlsx({ name: 'S', columns, rows });
    const b = writeXlsx({ name: 'S', columns, rows });
    expect(Buffer.from(a).equals(Buffer.from(b))).toBe(true);
  });
});
