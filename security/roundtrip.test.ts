import { describe, expect, it } from 'vitest';
import * as fc from 'fast-check';
import { readXlsx, writeXlsx, type CellValue } from '../src/index.js';
import { sheetArb } from './arbitraries.js';
import { normGrid } from './normalize.js';

const RUNS = Number(process.env.FC_RUNS ?? 40);
const SEED = process.env.FC_SEED === undefined ? undefined : Number(process.env.FC_SEED);
const params = { numRuns: RUNS, seed: SEED, endOnFailure: false } as const;

const readGrid = (bytes: Uint8Array): { names: string[]; rows: CellValue[][] } => {
  const wb = readXlsx(bytes);
  return { names: wb.sheets.map((s) => s.name), rows: [...wb.sheets[0]!.rows()] };
};

describe('write then read round trip', () => {
  it('returns every sheet cell unchanged', { timeout: 120_000 }, () => {
    fc.assert(
      fc.property(sheetArb, ({ columns, rows }) => {
        const cols = columns.map(({ header, format }) => ({ header, format }));
        const bytes = writeXlsx({ name: 'Data', columns: cols, rows });
        const got = readGrid(bytes);
        expect(got.names).toEqual(['Data']);
        expect(normGrid(got.rows)).toEqual(normGrid([cols.map((c) => c.header), ...rows]));
      }),
      params,
    );
  });

  it('reads negative zero back as zero', () => {
    const got = readGrid(writeXlsx({ name: 'Z', columns: [{ header: 'n' }], rows: [[-0], [0]] }));
    expect(got.rows[1]![0]).toBe(0);
    expect(got.rows[2]![0]).toBe(0);
  });

  it('keeps the longest cell Excel allows', () => {
    const s = 'x'.repeat(32767);
    expect(readGrid(writeXlsx({ name: 'L', columns: [{ header: 'n' }], rows: [[s]] })).rows[1]![0]).toBe(s);
  });

  it('keeps a thousand rows of width one', () => {
    const rows = Array.from({ length: 1000 }, (_, i) => [i]);
    expect(normGrid(readGrid(writeXlsx({ name: 'R', columns: [{ header: 'n' }], rows })).rows)).toHaveLength(1001);
  });
});
