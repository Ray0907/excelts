import { describe, expect, it } from 'vitest';
import { scanRows, type SheetContext } from '../src/readSheet.js';

const ctx: SheetContext = { sharedStrings: ['zero', 'one', 'two', 'three', 'four', 'five', 'six', 'seven'], xfKind: [0, 1], date1904: false };

const sheet = (cells: string[]): string => `<sheetData><row r="1">${cells.join('')}</row></sheetData>`;
const first = (cells: string[]) => [...scanRows(new TextEncoder().encode(sheet(cells)), ctx)][0]!;
const one = (cell: string) => first([cell])[0];

function xorshift(seed: number): () => number {
  let s = seed >>> 0 || 1;
  return () => {
    s ^= s << 13;
    s >>>= 0;
    s ^= s >>> 17;
    s ^= s << 5;
    s >>>= 0;
    return s;
  };
}

function randomDecimal(r: () => number): string {
  const digits = (n: number): string => {
    let d = '';
    for (let i = 0; i < n; i++) d += String(r() % 10);
    return d;
  };
  const neg = r() % 3 === 0 ? '-' : '';
  switch (r() % 9) {
    case 0: return neg + digits(1 + (r() % 17));
    case 1: return neg + digits(1 + (r() % 8)) + '.' + digits(1 + (r() % 15));
    case 2: return neg + '.' + digits(1 + (r() % 17));
    case 3: return neg + digits(1 + (r() % 8)) + '.';
    case 4: return neg + digits(1 + (r() % 6)) + '.' + digits(1 + (r() % 6)) + '0'.repeat(1 + (r() % 12));
    case 5: return neg + '0'.repeat(1 + (r() % 4)) + digits(1 + (r() % 4)) + '.' + digits(r() % 6);
    case 6: return neg + '0.' + '0'.repeat(r() % 20) + digits(1 + (r() % 17));
    case 7: return neg + digits(1 + (r() % 25));
    default: return neg + '0'.repeat(r() % 3) + '.' + '0'.repeat(r() % 3);
  }
}

describe('number cells read through the in-place path', () => {
  it('matches Number() bit for bit on 200k random decimal strings', () => {
    const r = xorshift(12345);
    const fixed = ['-0', '0', '-0.0', '0.', '007.50', '.5', '-.5', '5.', '-5.', '9007199254740993', '0.1234567890123456789', '1E-5', '1e5', '+5', '-', '.', '-.', '5.5.5', '1 2', 'NaN', 'INF'];
    for (let batch = 0; batch < 200; batch++) {
      const strs: string[] = batch === 0 ? [...fixed] : [];
      while (strs.length < 1000) strs.push(randomDecimal(r));
      const cells = strs.map((s) => `<c><v>${s}</v></c>`);
      const row = first(cells);
      for (let i = 0; i < strs.length; i++) {
        const want = Number(strs[i]);
        const got = row[i];
        if (Number.isNaN(want)) {
          expect(Number.isNaN(got as number), strs[i]).toBe(true);
        } else {
          expect(Object.is(got, want), `${strs[i]} -> ${String(got)} vs ${want}`).toBe(true);
        }
      }
    }
  });
});

describe('cell shapes around the fast path', () => {
  it('plain numbers and -0', () => {
    expect(one('<c r="A1"><v>12</v></c>')).toBe(12);
    expect(Object.is(one('<c r="A1"><v>-0</v></c>'), -0)).toBe(true);
    expect(one('<c r="A1" t="n"><v>1.5</v></c>')).toBe(1.5);
  });
  it('<f> before <v> still yields the cached value', () => {
    expect(one('<c r="A1"><f>1+1</f><v>2</v></c>')).toBe(2);
    expect(one('<c r="A1" t="s"><f>x</f><v>3</v></c>')).toBe('three');
  });
  it('<v/> and <v></v> are empty', () => {
    expect(one('<c r="A1"><v/></c>')).toBeUndefined();
    expect(one('<c r="A1"><v></v></c>')).toBeUndefined();
    expect(one('<c r="A1" t="s"><v></v></c>')).toBeUndefined();
  });
  it('whitespace inside <v> is trimmed by Number()', () => {
    expect(one('<c r="A1"><v> 12 </v></c>')).toBe(12);
    expect(one('<c r="A1" t="s"><v> 2 </v></c>')).toBe('two');
  });
  it('exponents, big integers and 17 significant digits fall back exactly', () => {
    expect(one('<c r="A1"><v>1E-5</v></c>')).toBe(1e-5);
    expect(one('<c r="A1"><v>1.5e3</v></c>')).toBe(1500);
    expect(one('<c r="A1"><v>9007199254740993</v></c>')).toBe(9007199254740992);
    expect(one('<c r="A1"><v>0.1234567890123456789</v></c>')).toBe(0.1234567890123456789);
  });
  it('shared string index with leading zeros, and out-of-range', () => {
    expect(one('<c r="A1" t="s"><v>007</v></c>')).toBe('seven');
    expect(one('<c r="A1" t="s"><v>0</v></c>')).toBe('zero');
    expect(one('<c r="A1" t="s"><v>99</v></c>')).toBeUndefined();
    expect(one('<c r="A1" t="s"><v>1x</v></c>')).toBeUndefined();
  });
  it('date styles still convert numbers', () => {
    expect(one('<c r="A1" s="1"><v>25570</v></c>')).toEqual(new Date(Date.UTC(1970, 0, 2)));
  });
  it('cells without a fast shape and their neighbours are unaffected', () => {
    const row = first(['<c r="A1"/>', '<c r="B1" t="inlineStr"><is><t>hi</t></is></c>', '<c r="C1" t="b"><v>1</v></c>', '<c r="D1" t="str"><v>a&amp;b</v></c>', '<c r="E1"><v>4</v></c>']);
    expect(row).toEqual([undefined, 'hi', true, 'a&b', 4]);
  });
  it('non-numeric content before the closing tag falls back', () => {
    expect(one('<c r="A1"><v>12<x/></v></c>')).toBeNaN();
  });
});
