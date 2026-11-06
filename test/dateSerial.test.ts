import { describe, expect, it } from 'vitest';
import { readXlsx, writeXlsx } from '../src/index.js';
import { dateToSerial1900 } from '../src/serial.js';

// Excel numbers days from 1900-01-01 = 1 and counts a nonexistent 1900-02-29, so 1900-03-01 = 61 and earlier dates sit one lower than the linear count.
describe('dateToSerial1900', () => {
  it.each([
    ['1900-01-01', 1],
    ['1900-02-28', 59],
    ['1900-03-01', 61],
    ['1970-01-01', 25569],
    ['1899-12-30', 0],
    ['1899-12-31', 1],
    ['2024-01-01', 45292],
  ])('%s is serial %i', (iso, serial) => {
    expect(dateToSerial1900(Date.parse(iso + 'T00:00:00Z'))).toBe(serial);
  });
  it('leaves times of day and durations below 1900-01-01 unshifted', () => {
    expect(dateToSerial1900(Date.parse('1899-12-30T12:00:00Z'))).toBe(0.5);
    expect(dateToSerial1900(Date.parse('1899-12-31T02:31:12Z'))).toBeCloseTo(1.105, 6);
  });
  it('keeps the time of day before 1900-03-01', () => {
    expect(dateToSerial1900(Date.parse('1900-02-28T12:00:00Z'))).toBe(59.5);
  });
});

describe('dates before 1900-03-01 round trip through the writer', () => {
  it('reads back the same instant', () => {
    const dates = ['1900-01-01', '1900-01-31', '1900-02-28', '1900-03-01'].map((s) => new Date(s + 'T00:00:00Z'));
    const bytes = writeXlsx({ name: 'S', columns: [{ header: 'd' }], rows: dates.map((d) => [d]) });
    const rows = [...readXlsx(bytes).sheets[0]!.rows()].slice(1).map((r) => (r[0] as Date).toISOString().slice(0, 10));
    expect(rows).toEqual(['1900-01-01', '1900-01-31', '1900-02-28', '1900-03-01']);
  });
});
