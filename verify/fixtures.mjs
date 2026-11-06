import { row } from '../test/helpers/data.mjs';
import { COLUMNS } from '../test/helpers/data.mjs';

const T = Date.UTC(2021, 5, 15, 12, 30, 0);
const edgeCols = [{ header: 'text' }, { header: 'num' }, { header: 'flag' }, { header: 'when', format: 'yyyy-mm-dd hh:mm:ss' }, { header: 'plain-date' }, { header: 'pct', format: '0.00%' }];
const edgeRows = [
  ['a & <b> "q" \'s\'', 0, true, new Date(T), new Date(T), 0.5],
  ['  leading and trailing  ', -1.5, false, new Date(T + 1000), new Date(T + 86400000), 0.125],
  ['line1\nline2\ttab', 1e21, true, new Date(Date.UTC(1999, 11, 31)), new Date(Date.UTC(2100, 0, 1)), 1],
  ['日本語 ünïcødé 🎉', 3.141592653589793, false, new Date(Date.UTC(1970, 0, 1)), new Date(Date.UTC(1970, 0, 2)), 0],
  ['repeat', 42, true, new Date(T), new Date(T), 0.5],
  ['repeat', Number.MAX_SAFE_INTEGER, false, null, undefined, null],
  [null, undefined, null, null, null, null],
  ['wide', 1, true, new Date(T), new Date(T), 0.1, 'extra1', 2, false],
  [Number.NaN, Number.POSITIVE_INFINITY, 'x', 'y', 'z', 'w'],
];

export const FIXTURES = [
  { name: 'edge', columns: edgeCols, rows: edgeRows },
  { name: 'bench1k', columns: COLUMNS, rows: Array.from({ length: 1000 }, (_, i) => row(i)) },
  { name: 'bench100k', columns: COLUMNS, rows: Array.from({ length: 100000 }, (_, i) => row(i)), big: true },
];

// Canonical expected grid: header row then data, NaN/Infinity/null/undefined become null, Dates become epoch ms.
export function expectedGrid(f) {
  const norm = (v) => (v instanceof Date ? v.getTime() : typeof v === 'number' && !Number.isFinite(v) ? null : v ?? null);
  return [f.columns.map((c) => c.header), ...f.rows.map((r) => r.map(norm))];
}
