import { describe, expect, it } from 'vitest';
import { readXlsx } from '../src/index.js';
import { classifyFormat, XF } from '../src/readStyles.js';
import { stylesOf, workbookOf } from './handBuilt.js';

const { NONE, DATE, TIME } = XF;

const CODES: [string, number][] = [
  ['General', NONE],
  ['0', NONE],
  ['0.00', NONE],
  ['#,##0.00', NONE],
  ['0.00%', NONE],
  ['0.00E+00', NONE],
  ['@', NONE],
  ['[Red]0.00', NONE],
  ['[Red][<0]0.00;0', NONE],
  ['#,##0 "days"', NONE],
  ['"d y h s" 0', NONE],
  ['0\\d', NONE],
  ['"Total" 0', NONE],
  ['[$-409]0.0', NONE],
  ['m/d/yy', DATE],
  ['yyyy-mm-dd', DATE],
  ['yyyy\\-mm\\-dd', DATE],
  ['mmm-yy', DATE],
  ['d-mmm', DATE],
  ['m', DATE],
  ['mm', DATE],
  ['mmm', DATE],
  ['mmmm', DATE],
  ['yyyy-mm-dd hh:mm', DATE],
  ['yyyy-mm-dd\\Thh:mm:ss', DATE],
  ['[$-F800]dddd, mmmm dd, yyyy', DATE],
  ['mmm h:mm', DATE],
  ['"Q"q yyyy', DATE],
  ['d "at" h:mm', DATE],
  ['h:mm', TIME],
  ['hh:mm:ss', TIME],
  ['h:mm AM/PM', TIME],
  ['h:mm:ss am/pm', TIME],
  ['[h]:mm:ss', TIME],
  ['[hh]:mm:ss', TIME],
  ['[mm]:ss', TIME],
  ['[ss]', TIME],
  ['[m]', TIME],
  ['mm:ss', TIME],
  ['mm:ss.0', TIME],
  ['mmss.0', TIME],
  ['[$-F400]h:mm:ss AM/PM', TIME],
  ['"Total" h:mm', TIME],
  ['h "h" m "m"', TIME],
  ['h\\h mm\\m', TIME],
  ['[Red]h:mm', TIME],
  ['[h]"h"', TIME],
  ['s', TIME],
  ['0.0 "d"', NONE],
  ['h:mm;@', TIME],
];

describe('classifyFormat', () => {
  it.each(CODES)('%j', (code, kind) => {
    expect(classifyFormat(code)).toBe(kind);
  });
});

const BUILTIN: [number, number][] = [
  [0, NONE], [1, NONE], [2, NONE], [9, NONE], [10, NONE], [12, NONE], [37, NONE], [49, NONE],
  [14, DATE], [15, DATE], [16, DATE], [17, DATE], [22, DATE],
  [18, TIME], [19, TIME], [20, TIME], [21, TIME], [45, TIME], [46, TIME], [47, TIME],
];

describe('built-in number format ids', () => {
  for (const date1904 of [false, true]) {
    it(`classifies every id the same in a ${date1904 ? 1904 : 1900} workbook`, () => {
      const cells = BUILTIN.map((_, i) => `<row r="${i + 1}"><c r="A${i + 1}" s="${i}"><v>10.5</v></c></row>`).join('');
      const wb = readXlsx(workbookOf(cells, { styles: stylesOf(BUILTIN.map(([id]) => id)), date1904 }));
      const got = [...wb.sheets[0]!.rows()].map((r) => r[0]);
      BUILTIN.forEach(([id, kind], i) => {
        if (kind === NONE) expect(got[i], `id ${id}`).toBe(10.5);
        else expect(got[i], `id ${id}`).toBeInstanceOf(Date);
      });
    });
  }
});

describe('time-only cells', () => {
  const read = (serial: string, format: number | string, date1904: boolean) => {
    const wb = readXlsx(workbookOf(`<row r="1"><c r="A1" s="0"><v>${serial}</v></c></row>`, { styles: stylesOf([format]), date1904 }));
    return [...wb.sheets[0]!.rows()][0]![0];
  };
  const base = Date.UTC(1899, 11, 30);

  it.each([false, true])('is 1899-12-30 plus the serial, date1904=%s', (date1904) => {
    expect((read('10.5', '[h]:mm:ss', date1904) as Date).getTime()).toBe(base + 10.5 * 86400000);
    expect((read('0.25', 'h:mm', date1904) as Date).toISOString()).toBe('1899-12-30T06:00:00.000Z');
    expect((read('0.5', 20, date1904) as Date).toISOString()).toBe('1899-12-30T12:00:00.000Z');
    expect((read('10.6320601851852', '[hh]:mm:ss', date1904) as Date).getTime()).toBe(-2208242990000);
  });

  it('ignores the 1900 leap-year correction below serial 60', () => {
    expect((read('1.5', '[h]:mm', false) as Date).toISOString()).toBe('1899-12-31T12:00:00.000Z');
    expect((read('1.5', 'm/d/yy h:mm', false) as Date).toISOString()).toBe('1900-01-01T12:00:00.000Z');
  });

  it('1900 and 1904 workbooks give the same Date', () => {
    expect(read('100.125', '[h]:mm:ss', true)).toEqual(read('100.125', '[h]:mm:ss', false));
  });

  it('keeps the calendar behavior of dates', () => {
    expect((read('44197', 'yyyy-mm-dd', false) as Date).toISOString()).toBe('2021-01-01T00:00:00.000Z');
    expect((read('42735', 'yyyy-mm-dd', true) as Date).toISOString()).toBe('2021-01-01T00:00:00.000Z');
    expect((read('59', 'yyyy-mm-dd', false) as Date).toISOString()).toBe('1900-02-28T00:00:00.000Z');
  });

  it('takes the slow number path too (exponent, long mantissa)', () => {
    expect((read('1.5E0', 'h:mm', false) as Date).toISOString()).toBe('1899-12-31T12:00:00.000Z');
    expect((read('0.123456789012345678', 'h:mm:ss', true) as Date).getTime()).toBe(Math.round((0.123456789012345678) * 86400000) + base);
  });
});
