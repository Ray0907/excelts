import { describe, expect, it } from 'vitest';
import { serialToDate } from '../src/readSheet.js';

const iso = (d: Date) => d.toISOString();

// Oracle adjudication (oracle/adjudications.tsv): ECMA-376 18.17.4.1 compat bases put serial 0 at 1899-12-31 (1900) and 1904-01-01 (1904).
describe('serialToDate epochs', () => {
  it('1900 compat base: serial 0 is 1899-12-31, serial 1 is 1900-01-01', () => {
    expect(iso(serialToDate(0, false))).toBe('1899-12-31T00:00:00.000Z');
    expect(iso(serialToDate(0.5, false))).toBe('1899-12-31T12:00:00.000Z');
    expect(iso(serialToDate(1, false))).toBe('1900-01-01T00:00:00.000Z');
    expect(iso(serialToDate(61, false))).toBe('1900-03-01T00:00:00.000Z');
    expect(iso(serialToDate(25569, false))).toBe('1970-01-01T00:00:00.000Z');
  });
  it('1904 base: serial 0 is 1904-01-01', () => {
    expect(iso(serialToDate(0, true))).toBe('1904-01-01T00:00:00.000Z');
    expect(iso(serialToDate(24107, true))).toBe('1970-01-01T00:00:00.000Z');
  });
  it('negative serials are undefined by the spec; they stay continuous with serial 0', () => {
    expect(iso(serialToDate(-1, false))).toBe('1899-12-30T00:00:00.000Z');
  });
});
