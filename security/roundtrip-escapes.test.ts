import { describe, expect, it } from 'vitest';
import * as fc from 'fast-check';
import { readXlsx, writeXlsx, type CellValue } from '../src/index.js';
import { textArb } from './arbitraries.js';
import { normGrid } from './normalize.js';

// Skipped until the reader and writer land the behaviors below, then flip `describe.skip` to `describe`:
//   - an error cell is written and read back as { error: '#DIV/0!' }
//   - `_xHHHH_` escapes: a control character (including NUL) is written escaped and read back unescaped; text that already looks like `_x0041_` round-trips literally
//   - control characters below 0x20 are preserved, not stripped
//   - CR and CRLF in a string are preserved, not normalized to LF
describe.skip('write then read round trip: escapes, control characters, CRLF, error cells', () => {
  const RUNS = Number(process.env.FC_RUNS ?? 40);
  const roundTrip = (rows: CellValue[][]): unknown[][] => {
    const bytes = writeXlsx({ name: 'E', columns: [{ header: 'h' }, { header: 'g' }], rows });
    return normGrid([...readXlsx(bytes).sheets[0]!.rows()].slice(1));
  };

  it('preserves arbitrary strings including control characters, CR/LF and _x escapes', () => {
    const tricky = fc.oneof(textArb(false), fc.constantFrom('_x0041_', '_x005F_x0041_', '_x', 'a_x000D_b', '\r\n', '\r', 'a\r\nb', '\u0000', '\u0001\u001f', '_xZZZZ_', '_x00_'));
    fc.assert(
      fc.property(fc.array(fc.tuple(tricky, tricky), { maxLength: 20 }), (rows) => {
        expect(roundTrip(rows)).toEqual(normGrid(rows));
      }),
      { numRuns: RUNS },
    );
  });

  it('preserves error cells', () => {
    const errors = fc.constantFrom('#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A');
    fc.assert(
      fc.property(fc.array(fc.tuple(errors, errors), { maxLength: 20 }), (codes) => {
        const rows = codes.map(([a, b]) => [{ error: a }, { error: b }]);
        expect(roundTrip(rows)).toEqual(rows);
      }),
      { numRuns: RUNS },
    );
  });

  it('does not decode _xHHHH_ in text that was written literally', () => {
    expect(roundTrip([['_x0041_', 'x']])).toEqual([['_x0041_', 'x']]);
  });
});
