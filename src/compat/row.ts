import { Cell } from './cell.js';
import { l2n } from './colCache.js';
import { bookOf } from './fallbackBook.js';
import { NotSupportedError } from './errors.js';
import { componentView, styleView } from './styleView.js';
import type { Style, StyleKey } from './styleTypes.js';
import type { Note } from './notes.js';
import { rowSplice } from './splice.js';
import { MergeSlave, storeValue, type Raw } from './values.js';
import type { Worksheet } from './worksheet.js';

const MAX_COLS = 16384;

function isObj(x: unknown): boolean {
  return typeof x === 'object' && x !== null && !(x instanceof Date);
}

/** Owns the data of a row: `_v` holds the raw values (index = column - 1, `undefined` = no cell), `_s` the explicit style ids. */
export class Row {
  /** @internal */ _v: Raw[] = [];
  /** @internal */ _s: (number | undefined)[] | null = null;
  /** @internal */ _cells: (Cell | undefined)[] | null = null;
  /** @internal notes by 1-based column; null until the first note */
  _n: Map<number, Note> | null = null;
  /** @internal */ _xf = 0;
  /** @internal */ _hidden = false;
  /** @internal */ _outline = 0;
  height?: number;

  constructor(
    readonly worksheet: Worksheet,
    readonly number: number,
  ) {}

  // ---- cells ----
  /** @internal the style a not-yet-styled cell at 0-based column j takes: the row's over the column's. */
  _inheritedXf(j: number): number {
    const ws = this.worksheet;
    return bookOf(ws).merge(this._xf, ws._colXf ? ws._colXf(j + 1) : 0);
  }
  /** @internal */
  _xfAt(j: number): number {
    return this._s?.[j] ?? this._inheritedXf(j);
  }
  /** @internal */
  _setXfAt(j: number, xf: number): void {
    (this._s ??= [])[j] = xf;
  }

  findCell(col: number): Cell | undefined {
    const j = col - 1;
    if (this._v[j] === undefined) return undefined;
    return this.getCell(col);
  }
  getCell(col: number | string): Cell {
    const n = typeof col === 'string' ? (this.worksheet.getColumnKey(col)?.number ?? l2n(col)) : col;
    const j = n - 1;
    const cells = (this._cells ??= []);
    const hit = cells[j];
    if (hit) return hit;
    if (!(n >= 1 && n <= MAX_COLS)) throw new Error(`Out of bounds: column ${n}`);
    if (this._v[j] === undefined) {
      this._v[j] = null;
      const xf = this._inheritedXf(j);
      if (xf) this._setXfAt(j, xf);
    }
    return (cells[j] = new Cell(this, n));
  }

  eachCell(options: { includeEmpty?: boolean } | ((c: Cell, col: number) => void), cb?: (c: Cell, col: number) => void): void {
    const f = cb ?? (options as (c: Cell, col: number) => void);
    const o = cb ? (options as { includeEmpty?: boolean }) : null;
    const v = this._v;
    if (o && o.includeEmpty) {
      for (let i = 1; i <= v.length; i++) f(this.getCell(i), i);
    } else {
      for (let j = 0; j < v.length; j++) {
        const x = v[j];
        if (x !== null && x !== undefined) f(this.getCell(j + 1), j + 1);
      }
    }
  }

  // ---- values ----
  get values(): Raw[] {
    const out: Raw[] = [];
    const v = this._v;
    for (let j = 0; j < v.length; j++) {
      const x = v[j];
      if (x === null || x === undefined) continue;
      out[j + 1] = x instanceof MergeSlave ? this.getCell(j + 1).value : x;
    }
    return out;
  }
  set values(value: Raw[] | Record<string, Raw> | null | undefined) {
    this._cells = null;
    this._s = null;
    if (!value) {
      this._v = [];
    } else if (Array.isArray(value)) {
      const off = Object.prototype.hasOwnProperty.call(value, '0') ? 0 : 1;
      const n = value.length - off;
      const out: Raw[] = new Array(n > 0 ? n : 0);
      let end = 0;
      for (let i = 0; i < n; i++) {
        const x = value[i + off];
        if (x === undefined) continue;
        out[i] = isObj(x) ? storeValue(x) : x;
        end = i + 1;
      }
      out.length = end;
      this._v = out;
    } else {
      this._v = [];
      this.worksheet.eachColumnKey((column, key) => {
        const x = (value as Record<string, Raw>)[key];
        if (x !== undefined) this.getCell(column.number).value = x;
      });
    }
  }
  get hasValues(): boolean {
    for (const x of this._v) if (x !== null && x !== undefined) return true;
    return false;
  }
  get cellCount(): number { return this._v.length; }
  get actualCellCount(): number {
    let c = 0;
    this.eachCell(() => c++);
    return c;
  }
  get dimensions(): { min: number; max: number } | null {
    let min = 0, max = 0;
    const v = this._v;
    for (let j = 0; j < v.length; j++) {
      if (v[j] !== null && v[j] !== undefined) {
        if (!min) min = j + 1;
        max = j + 1;
      }
    }
    return min ? { min, max } : null;
  }

  // ---- style ----
  private get book() { return bookOf(this.worksheet); }
  get style(): Style {
    const book = this.book;
    return styleView(() => book.get(this._xf), (k, v) => this.apply(k, v));
  }
  set style(s: Style | undefined) {
    this._freeze();
    this._xf = this.book.intern(s);
  }
  /** @internal gives every existing cell its own style id, so a later change of the row or column style only reaches cells created afterwards. */
  _freeze(): void {
    for (let j = 0; j < this._v.length; j++) if (this._v[j] !== undefined && this._s?.[j] === undefined) this._setXfAt(j, this._inheritedXf(j));
  }
  private apply(k: StyleKey, v: unknown): void {
    this._xf = this.book.withComponent(this._xf, k, v);
    for (let j = 0; j < this._v.length; j++) {
      if (this._v[j] !== undefined) this._setXfAt(j, this.book.withComponent(this._xfAt(j), k, v));
    }
  }
  private comp<T extends object>(k: StyleKey): T | undefined {
    return componentView(this.book.get(this._xf)[k] as T | undefined, k, (n) => this.apply(k, n));
  }
  get numFmt(): string | undefined { return this.book.get(this._xf).numFmt; }
  set numFmt(v: string | undefined) { this.apply('numFmt', v); }
  get font() { return this.comp<NonNullable<Style['font']>>('font'); }
  set font(v: Style['font']) { this.apply('font', v); }
  get alignment() { return this.comp<NonNullable<Style['alignment']>>('alignment'); }
  set alignment(v: Style['alignment']) { this.apply('alignment', v); }
  get border() { return this.comp<NonNullable<Style['border']>>('border'); }
  set border(v: Style['border']) { this.apply('border', v); }
  get fill() { return this.comp<NonNullable<Style['fill']>>('fill'); }
  set fill(v: Style['fill']) { this.apply('fill', v); }
  get protection() { return this.comp<NonNullable<Style['protection']>>('protection'); }
  set protection(v: Style['protection']) { this.apply('protection', v); }

  get hidden(): boolean { return this._hidden; }
  set hidden(v: boolean) { this._hidden = !!v; }
  get outlineLevel(): number { return this._outline; }
  set outlineLevel(v: number) { this._outline = v || 0; }
  get collapsed(): boolean { return !!(this._outline && this._outline >= this.worksheet.properties.outlineLevelRow); }

  addPageBreak(lft?: number, rght?: number): void {
    const left = Math.max(0, (lft ?? 0) - 1) || 0;
    const right = Math.max(0, (rght ?? 0) - 1) || 16838;
    const pb: { id: number; max: number; man: number; min?: number } = { id: this.number, max: right, man: 1 };
    if (left) pb.min = left;
    this.worksheet.rowBreaks.push(pb);
  }
  /** A no-op in the document model; the append-only writer's sheets serialise and release the row. */
  commit(): void {
    this.worksheet._commitRow?.(this);
  }
  destroy(): void {
    this._v = [];
    this._s = null;
    this._cells = null;
  }
  splice(start: number, count: number, ...inserts: Raw[]): void { rowSplice(this, start, count, inserts); }
  get model(): never { throw new NotSupportedError('row.model (internal model)'); }
  set model(_v: unknown) { throw new NotSupportedError('row.model (internal model)'); }
}

