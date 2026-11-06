import type { CellValue } from './index.js';

/**
 * Rows from the sheet scanner, which reports a run of skipped row numbers as one number. Handing those out here, from a counter,
 * costs a few ns a row; resuming the scanner's generator once per skipped row (a sheet formatted down to row 1,048,576 has about a million) costs about 4x that.
 * Each skipped row is still its own fresh, mutable empty array.
 */
export class RowIterator implements IterableIterator<CellValue[]> {
  private gapLeft = 0;

  constructor(private readonly runs: Iterator<CellValue[] | number>) {}

  next(): IteratorResult<CellValue[]> {
    if (this.gapLeft > 0) {
      this.gapLeft--;
      return { value: [], done: false };
    }
    const r = this.runs.next();
    if (r.done) return r as IteratorReturnResult<undefined> as IteratorResult<CellValue[]>;
    const v = r.value;
    if (typeof v === 'number') {
      this.gapLeft = v - 1;
      return { value: [], done: false };
    }
    return r as IteratorYieldResult<CellValue[]>;
  }

  return(): IteratorResult<CellValue[]> {
    this.gapLeft = 0;
    this.runs.return?.();
    return { value: undefined, done: true };
  }

  [Symbol.iterator](): this {
    return this;
  }
}
