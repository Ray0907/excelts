import { decodeEx, n2l, type Decoded } from './colCache.js';
import type { Worksheet } from './worksheet.js';

interface Loc { sheetName: string; address: string; row: number; col: number }
interface Rect { sheetName: string; top: number; left: number; bottom: number; right: number }
/** A name kept as the file had it: scoped to a sheet, a formula or constant, whole rows/columns, or carrying other attributes. */
export interface RawName { name: string; scope?: Worksheet; attrs: Record<string, string>; text: string }

/** Ranges up to this many cells are tracked cell by cell (so names can be added and taken off single cells); larger ones stay rectangles. */
const EXPAND_MAX = 1024;
/** Cells one workbook may track individually; a bound for hostile files with thousands of mid-sized ranges. */
const BUDGET = 1_048_576;
const MAX_ROW = 1_048_576;
const MAX_COL = 16_384;

const key = (s: string, r: number, c: number): string => `${s}\u0000${r}\u0000${c}`;
const inside = (r: Rect, s: string, row: number, col: number): boolean => r.sheetName === s && row >= r.top && row <= r.bottom && col >= r.left && col <= r.right;

export function sheetRef(name: string): string {
  const plain = /^[A-Za-z][A-Za-z0-9]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) && !/^[rRcC]\d*$/.test(name);
  return plain ? name : `'${name.replace(/'/g, "''")}'`;
}

/** Splits a defined name's formula into its comma-separated parts, leaving commas inside quoted sheet names alone. */
function parts(text: string): string[] {
  const out: string[] = [];
  let cur = '';
  let q = false;
  for (const ch of text) {
    if (ch === "'") q = !q;
    if (ch === ',' && !q) {
      out.push(cur);
      cur = '';
    } else cur += ch;
  }
  out.push(cur);
  return out.map((p) => p.trim()).filter(Boolean);
}

const REF = /^(?:'(?:[^']|'')+'|[^'!]+)!\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$/;

/** Names of cells and ranges. Stored per name as the cells they cover (small ranges) or rectangles (large ones), never as a matrix of the sheet. */
export class DefinedNames {
  private cells = new Map<string, Map<string, Loc>>();
  private big = new Map<string, Rect[]>();
  private spent = 0;
  /** @internal names kept verbatim, written back unchanged */
  raw: RawName[] = [];

  add(locStr: string, name: string): void {
    this.addEx(decodeEx(locStr), name);
  }

  addEx(l: Decoded, name: string): void {
    const sheetName = l.sheetName ?? '';
    const top = l.top ?? l.row!;
    const left = l.left ?? l.col!;
    const bottom = l.bottom ?? top;
    const right = l.right ?? left;
    if (!(top >= 1 && left >= 1 && bottom <= MAX_ROW && right <= MAX_COL)) throw new RangeError(`defined name ${name}: location out of bounds`);
    const area = (bottom - top + 1) * (right - left + 1);
    if (area > EXPAND_MAX || this.spent + area > BUDGET) {
      const list = this.big.get(name) ?? [];
      if (!list.some((r) => r.sheetName === sheetName && r.top === top && r.left === left && r.bottom === bottom && r.right === right)) list.push({ sheetName, top, left, bottom, right });
      this.big.set(name, list);
      if (!this.cells.has(name)) this.cells.set(name, new Map());
      return;
    }
    let m = this.cells.get(name);
    if (!m) this.cells.set(name, (m = new Map()));
    for (let r = top; r <= bottom; r++) {
      for (let c = left; c <= right; c++) {
        const k = key(sheetName, r, c);
        if (!m.has(k)) {
          m.set(k, { sheetName, row: r, col: c, address: n2l(c) + r });
          this.spent++;
        }
      }
    }
  }

  remove(locStr: string, name: string): void {
    this.removeEx(decodeEx(locStr), name);
  }

  removeEx(l: Decoded, name: string): void {
    const m = this.cells.get(name);
    if (!m) return;
    const s = l.sheetName ?? '';
    const top = l.top ?? l.row!, left = l.left ?? l.col!, bottom = l.bottom ?? top, right = l.right ?? left;
    if ((bottom - top + 1) * (right - left + 1) > EXPAND_MAX) {
      for (const [k, v] of m) if (v.sheetName === s && v.row >= top && v.row <= bottom && v.col >= left && v.col <= right) m.delete(k);
    } else {
      for (let r = top; r <= bottom; r++) for (let c = left; c <= right; c++) m.delete(key(s, r, c));
    }
    const list = this.big.get(name);
    if (list) this.big.set(name, list.filter((r) => !(r.sheetName === s && r.top >= top && r.bottom <= bottom && r.left >= left && r.right <= right)));
  }

  removeAllNames(l: Decoded): void {
    for (const name of this.cells.keys()) this.removeEx(l, name);
  }

  getNames(addressStr: string): string[] {
    return this.getNamesEx(decodeEx(addressStr));
  }

  getNamesEx(l: Decoded): string[] {
    const s = l.sheetName ?? '';
    const out: string[] = [];
    for (const [name, m] of this.cells) {
      if (m.has(key(s, l.row!, l.col!)) || this.big.get(name)?.some((r) => inside(r, s, l.row!, l.col!))) out.push(name);
    }
    return out;
  }

  forEach(cb: (name: string, cell: Loc) => void): void {
    for (const [name, m] of this.cells) for (const loc of m.values()) cb(name, loc);
  }

  /** The ranges of a name: tracked cells joined into rectangles, then the large ranges. */
  getRanges(name: string): { name: string; ranges: string[] } {
    const m = this.cells.get(name);
    const ranges: string[] = [];
    if (m) {
      const todo = new Set(m.keys());
      const order = [...m.values()].sort((a, b) => (a.sheetName < b.sheetName ? -1 : a.sheetName > b.sheetName ? 1 : a.row - b.row || a.col - b.col));
      const live = (s: string, r: number, c: number): boolean => todo.has(key(s, r, c));
      for (const cell of order) {
        if (!todo.delete(key(cell.sheetName, cell.row, cell.col))) continue;
        const s = cell.sheetName;
        let top = cell.row, bottom = cell.row, left = cell.col, right = cell.col;
        for (let y = top - 1; live(s, y, cell.col); y--) { todo.delete(key(s, y, cell.col)); top = y; }
        for (let y = bottom + 1; live(s, y, cell.col); y++) { todo.delete(key(s, y, cell.col)); bottom = y; }
        const column = (x: number): boolean => {
          for (let y = top; y <= bottom; y++) if (!live(s, y, x)) return false;
          for (let y = top; y <= bottom; y++) todo.delete(key(s, y, x));
          return true;
        };
        for (let x = left - 1; column(x); x--) left = x;
        for (let x = right + 1; column(x); x++) right = x;
        ranges.push(this.text({ sheetName: s, top, left, bottom, right }));
      }
    }
    for (const r of this.big.get(name) ?? []) ranges.push(this.text(r));
    return { name, ranges };
  }

  private text(r: Rect): string {
    const a = `$${n2l(r.left)}$${r.top}`;
    const body = r.top === r.bottom && r.left === r.right ? a : `${a}:$${n2l(r.right)}$${r.bottom}`;
    return r.sheetName ? `${sheetRef(r.sheetName)}!${body}` : body;
  }

  private shift(sheet: string, axis: 'row' | 'col', start: number, del: number, ins: number): void {
    const move = (v: number): number | undefined => (v < start ? v : v < start + del ? undefined : v + ins - del);
    for (const [name, m] of this.cells) {
      const next = new Map<string, Loc>();
      for (const loc of m.values()) {
        if (loc.sheetName === sheet) {
          const v = move(loc[axis]);
          if (v === undefined) continue;
          const row = axis === 'row' ? v : loc.row, col = axis === 'col' ? v : loc.col;
          next.set(key(sheet, row, col), { sheetName: sheet, row, col, address: n2l(col) + row });
        } else next.set(key(loc.sheetName, loc.row, loc.col), loc);
      }
      this.cells.set(name, next);
    }
    for (const [name, list] of this.big) {
      const out: Rect[] = [];
      for (const r of list) {
        if (r.sheetName !== sheet) {
          out.push(r);
          continue;
        }
        const lo = axis === 'row' ? 'top' : 'left', hi = axis === 'row' ? 'bottom' : 'right';
        // the first and last surviving line of the rectangle
        let a = r[lo];
        let b = r[hi];
        if (a >= start && a < start + del) a = start + del;
        if (b >= start && b < start + del) b = start - 1;
        if (a > b) continue;
        out.push({ ...r, [lo]: move(a), [hi]: move(b) } as Rect);
      }
      this.big.set(name, out);
    }
  }

  spliceRows(sheetName: string, start: number, numDelete: number, numInsert: number): void {
    this.shift(sheetName, 'row', start, numDelete, numInsert);
  }

  spliceColumns(sheetName: string, start: number, numDelete: number, numInsert: number): void {
    this.shift(sheetName, 'col', start, numDelete, numInsert);
  }

  /** `[{ name, ranges }]` for the names that cover something. */
  get model(): { name: string; ranges: string[] }[] {
    return [...this.cells.keys()].map((n) => this.getRanges(n)).filter((d) => d.ranges.length);
  }
  set model(value: { name: string; ranges: string[] }[]) {
    this.cells = new Map();
    this.big = new Map();
    this.spent = 0;
    for (const d of value) for (const r of d.ranges) if (REF.test(r)) this.add(r, d.name);
  }

  /** Takes a `definedName` element of a file: plain cell ranges join the model, everything else is kept to be written back as it was. */
  load(name: string, text: string, scope: Worksheet | undefined, attrs: Record<string, string>): void {
    const ps = parts(text);
    if (!scope && !Object.keys(attrs).length && ps.length && ps.every((p) => REF.test(p)) && !name.startsWith('_xlnm.')) {
      try {
        for (const p of ps) this.add(p, name);
        return;
      } catch {
        // a reference outside the sheet grid: keep the name as it was
      }
    }
    this.raw.push({ name, scope, attrs, text });
  }
}
