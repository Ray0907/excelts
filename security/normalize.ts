import type { CellValue } from '../src/index.js';

export type Norm = string | number | boolean | null | { error: string };

/** The same canonical form verify/fixtures.mjs expectedGrid uses: NaN/Infinity/undefined become null, a Date becomes its epoch ms. */
export function normCell(v: CellValue): Norm {
  if (v instanceof Date) return v.getTime();
  if (typeof v === 'number') return !Number.isFinite(v) ? null : v === 0 ? 0 : v; // -0 becomes 0: a spreadsheet cell has no negative zero
  return v ?? null;
}

/** Trailing empty cells and trailing empty rows are not part of the contract, so both sides drop them. */
export function normGrid(grid: readonly (readonly CellValue[])[]): Norm[][] {
  const out = grid.map((row) => {
    const cells = Array.from(row, normCell);
    while (cells.length && cells[cells.length - 1] === null) cells.pop();
    return cells;
  });
  while (out.length && out[out.length - 1]!.length === 0) out.pop();
  return out;
}
