import * as fc from 'fast-check';
import type { CellValue, Column } from '../src/index.js';

/** Excel's 1900 date system starts at 1900-03-01 for this purpose: earlier serials sit in the Lotus leap-year bug zone. */
const MIN_DATE_MS = Date.UTC(1900, 2, 1);
const MAX_DATE_MS = Date.UTC(9999, 11, 31, 23, 59, 59, 999);

const fromCodePoints = (cps: number[]): string => String.fromCodePoint(...cps);

/** Drops a dangling high surrogate left by cutting at a UTF-16 unit boundary. */
const cutUnits = (s: string, n: number): string => {
  const cut = s.slice(0, n);
  const last = cut.charCodeAt(cut.length - 1);
  return last >= 0xd800 && last <= 0xdbff ? cut.slice(0, -1) : cut;
};

/**
 * Lone surrogates are not valid text and cannot survive UTF-8, so they are never generated (documented, not tested).
 * U+FFFE and U+FFFF are XML non-characters and are left out too.
 * `plain` keeps the main round trip clear of the behaviors being changed: no \r, no control character but \t and \n, no `_x`.
 */
const codePoint = (plain: boolean): fc.Arbitrary<number> =>
  fc.oneof(
    { weight: 6, arbitrary: fc.integer({ min: 0x20, max: 0x7e }) },
    { weight: 2, arbitrary: fc.constantFrom(0x09, 0x0a, 0x20, 0x26, 0x3c, 0x3e, 0x22, 0x27) },
    { weight: 1, arbitrary: fc.integer({ min: 0x7f, max: 0x9f }) },
    { weight: 3, arbitrary: fc.integer({ min: 0xa0, max: 0xd7ff }) },
    { weight: 1, arbitrary: fc.integer({ min: 0xe000, max: 0xfffd }) },
    { weight: 2, arbitrary: fc.integer({ min: 0x10000, max: 0x10ffff }) },
    ...(plain ? [] : [{ weight: 3, arbitrary: fc.oneof(fc.integer({ min: 0x00, max: 0x1f }), fc.constant(0x0d)) }]),
  );

const shortText = (plain: boolean): fc.Arbitrary<string> => {
  const s = fc.array(codePoint(plain), { maxLength: 30 }).map(fromCodePoints);
  return plain ? s.map((t) => t.replace(/_x/g, '_y')) : s;
};

const longText = (plain: boolean): fc.Arbitrary<string> =>
  fc
    .tuple(shortText(plain).filter((s) => s.length > 0), fc.integer({ min: 1000, max: 32767 }))
    .map(([unit, n]) => cutUnits(unit.repeat(Math.ceil(n / unit.length)), n));

export const textArb = (plain = true): fc.Arbitrary<string> =>
  fc.oneof({ weight: 40, arbitrary: shortText(plain) }, { weight: 1, arbitrary: longText(plain) }, { weight: 2, arbitrary: fc.constantFrom('', ' ', '  x  ', '\n', '\t', 'a\nb', '=1+1', "'quoted", '0', '1e3', 'TRUE', '#N/A') });

export const doubleArb: fc.Arbitrary<number> = fc.oneof(
  { weight: 6, arbitrary: fc.double({ noNaN: false }) },
  { weight: 3, arbitrary: fc.integer({ min: -1000000, max: 1000000 }) },
  { weight: 2, arbitrary: fc.double({ min: -1e6, max: 1e6, noNaN: true }) },
  { weight: 1, arbitrary: fc.constantFrom(0, -0, Number.MIN_VALUE, -Number.MIN_VALUE, Number.MAX_VALUE, -Number.MAX_VALUE, Number.EPSILON, 1e21, 1e-7, 0.1 + 0.2, Number.MAX_SAFE_INTEGER, NaN, Infinity, -Infinity) },
);

export const dateArb: fc.Arbitrary<Date> = fc.integer({ min: MIN_DATE_MS, max: MAX_DATE_MS }).map((ms) => new Date(ms));

/** A column's format decides how its cells read back: number formats turn a Date into a serial number, date formats turn a number into a Date. */
export type Kind = 'general' | 'number' | 'date';
export interface GenColumn extends Column {
  kind: Kind;
}

const FORMATS: Record<Kind, (string | undefined)[]> = {
  general: [undefined],
  number: ['0.00', '0.00%', '#,##0', '@'],
  date: ['yyyy-mm-dd hh:mm:ss', 'yyyy-mm-dd'],
};

export const columnArb: fc.Arbitrary<GenColumn> = fc
  .tuple(fc.constantFrom<Kind>('general', 'general', 'number', 'date'), textArb(true), fc.nat())
  .map(([kind, header, pick]) => ({ kind, header, format: FORMATS[kind][pick % FORMATS[kind].length] }));

const common = [textArb(true), fc.boolean(), fc.constantFrom(null, undefined)] as const;
const cellFor = (kind: Kind): fc.Arbitrary<CellValue> =>
  kind === 'general'
    ? fc.oneof({ weight: 3, arbitrary: textArb(true) }, { weight: 4, arbitrary: doubleArb }, { weight: 1, arbitrary: fc.boolean() }, { weight: 2, arbitrary: dateArb }, { weight: 2, arbitrary: fc.constantFrom(null, undefined) })
    : kind === 'number'
      ? fc.oneof({ weight: 1, arbitrary: common[0] }, { weight: 5, arbitrary: doubleArb }, { weight: 1, arbitrary: common[1] }, { weight: 2, arbitrary: common[2] })
      : fc.oneof({ weight: 1, arbitrary: common[0] }, { weight: 5, arbitrary: dateArb }, { weight: 1, arbitrary: common[1] }, { weight: 2, arbitrary: common[2] });

export interface GenSheet {
  columns: GenColumn[];
  rows: CellValue[][];
}

/** Mostly small sheets, now and then 197 columns or 300 rows. Rows run up to three cells past the declared columns. */
export const sheetArb: fc.Arbitrary<GenSheet> = fc
  .tuple(fc.oneof({ weight: 9, arbitrary: fc.integer({ min: 1, max: 8 }) }, { weight: 1, arbitrary: fc.integer({ min: 100, max: 197 }) }), fc.oneof({ weight: 9, arbitrary: fc.integer({ min: 0, max: 12 }) }, { weight: 1, arbitrary: fc.integer({ min: 100, max: 300 }) }))
  .chain(([nCols, nRows]) =>
    fc.array(columnArb, { minLength: nCols, maxLength: nCols }).chain((columns) => {
      const kindAt = (j: number): Kind => columns[j]?.kind ?? 'general';
      const row = fc
        .tuple(fc.integer({ min: 0, max: nCols + 3 }), fc.tuple(...Array.from({ length: nCols + 3 }, (_, j) => cellFor(kindAt(j)))))
        .map(([n, cells]) => cells.slice(0, n));
      return fc.array(row, { minLength: nRows, maxLength: nRows }).map((rows) => ({ columns, rows }));
    }),
  );

export const bytesToCells = (cps: number[]): string => fromCodePoints(cps);
