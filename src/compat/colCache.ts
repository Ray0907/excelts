/** A1-style address helpers, same results as exceljs's col-cache. */
const MAX_COL = 16384;
const letters: string[] = [''];
const addrRe = /^[A-Z]+\d+$/;

export function n2l(n: number): string {
  if (n < 1 || n > MAX_COL) throw new Error(`${n} is out of bounds. Excel supports columns from 1 to ${MAX_COL}`);
  let s = letters[n];
  if (s === undefined) {
    s = '';
    for (let k = n; k > 0; k = Math.floor((k - 1) / 26)) s = String.fromCharCode(65 + ((k - 1) % 26)) + s;
    letters[n] = s;
  }
  return s;
}

export function l2n(l: string): number {
  const u = l.toUpperCase();
  let n = 0;
  for (let i = 0; i < u.length; i++) {
    const c = u.charCodeAt(i) - 64;
    if (c < 1 || c > 26) throw new Error(`Out of bounds. Invalid column letter: ${l}`);
    n = n * 26 + c;
  }
  if (n < 1 || n > MAX_COL) throw new Error(`Out of bounds. Invalid column letter: ${l}`);
  return n;
}

export interface Addr {
  address: string;
  col: number;
  row: number;
  $col$row: string;
}

export function encodeAddress(row: number, col: number): string {
  return n2l(col) + row;
}

const decodeRe = /^[$]?([A-Za-z]{1,3})[$]?(\d+)$/;
export function decodeAddress(address: string): Addr {
  const m = decodeRe.exec(address);
  if (m) {
    const col = l2n(m[1]!);
    const row = parseInt(m[2]!, 10);
    return { address: n2l(col) + row, col, row, $col$row: `$${n2l(col)}$${row}` };
  }
  // column-only ("A") or row-only ("5") references, as in "A:A" and "2:2"
  const c = /^[$]?([A-Za-z]{1,3})$/.exec(address);
  if (c) {
    const col = l2n(c[1]!);
    return { address: n2l(col), col, row: 0, $col$row: `$${n2l(col)}` } as Addr;
  }
  const r = /^[$]?(\d+)$/.exec(address);
  if (r) return { address: r[1]!, col: 0, row: parseInt(r[1]!, 10), $col$row: `$${r[1]}` } as Addr;
  throw new Error(`Invalid Address: ${address}`);
}

export function validateAddress(address: string): void {
  if (!addrRe.test(address.toUpperCase().replace(/\$/g, ''))) throw new Error(`Invalid Address: ${address}`);
}

export interface Decoded {
  top?: number;
  left?: number;
  bottom?: number;
  right?: number;
  row?: number;
  col?: number;
  sheetName?: string;
  address?: string;
  dimensions?: string;
}

/** "Sheet1!A1", "'My Sheet'!A1:B2", "A1:B2" or "A1". */
export function decodeEx(value: string): Decoded {
  const m = /^(?:'((?:[^']|'')+)'|([^'!][^!]*))!(.+)$/.exec(value);
  let sheetName: string | undefined;
  let rest = value;
  if (m) {
    sheetName = (m[1] ?? m[2])!.replace(/''/g, "'");
    rest = m[3]!;
  }
  const parts = rest.split(':');
  if (parts.length === 2) {
    const tl = decodeAddress(parts[0]!);
    const br = decodeAddress(parts[1]!);
    return { top: Math.min(tl.row, br.row), left: Math.min(tl.col, br.col), bottom: Math.max(tl.row, br.row), right: Math.max(tl.col, br.col), sheetName, dimensions: rest };
  }
  const a = decodeAddress(rest);
  return { row: a.row, col: a.col, address: a.address, sheetName };
}
