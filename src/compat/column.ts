import { n2l } from './colCache.js';
import { bookOf } from './fallbackBook.js';
import { ValueType } from './enums.js';
import { componentView, styleView } from './styleView.js';
import type { Style, StyleKey } from './styleTypes.js';
import type { Raw } from './values.js';
import type { Worksheet } from './worksheet.js';

export const DEFAULT_COLUMN_WIDTH = 9;

export interface ColumnDefn {
  header?: string | string[];
  key?: string;
  width?: number;
  style?: Style;
  hidden?: boolean;
  outlineLevel?: number;
}

export class Column {
  /** @internal */ _xf = 0;
  /** @internal */ _key?: string;
  /** @internal */ _header?: string | string[];
  /** @internal */ _hidden = false;
  /** @internal */ _outline = 0;
  width?: number;

  constructor(
    readonly worksheet: Worksheet,
    readonly number: number,
    defn?: ColumnDefn | false,
  ) {
    if (defn !== false) this.defn = defn;
  }

  get letter(): string { return n2l(this.number); }
  get isCustomWidth(): boolean { return this.width !== undefined && this.width !== DEFAULT_COLUMN_WIDTH; }
  get defn(): ColumnDefn {
    return { header: this._header, key: this.key, width: this.width, style: this.style, hidden: this.hidden, outlineLevel: this.outlineLevel };
  }
  set defn(v: ColumnDefn | undefined | null) {
    if (v) {
      this.key = v.key;
      this.width = v.width !== undefined ? v.width : DEFAULT_COLUMN_WIDTH;
      this.outlineLevel = v.outlineLevel ?? 0;
      this.style = v.style ?? {};
      this.header = v.header;
      this._hidden = !!v.hidden;
    } else {
      this._header = undefined;
      this.key = undefined;
      this.width = undefined;
      this.style = {};
      this.outlineLevel = 0;
    }
  }
  get headers(): (string | undefined)[] { return Array.isArray(this._header) ? this._header : [this._header]; }
  get header(): string | string[] | undefined { return this._header; }
  set header(v: string | string[] | undefined) {
    if (v !== undefined) {
      this._header = v;
      this.headers.forEach((text, i) => {
        this.worksheet.getCell(i + 1, this.number).value = text;
      });
    } else this._header = undefined;
  }
  get headerCount(): number { return this.headers.length; }
  get key(): string | undefined { return this._key; }
  set key(v: string | undefined) {
    const ws = this.worksheet;
    if (this._key && ws.getColumnKey(this._key) === this) ws.deleteColumnKey(this._key);
    this._key = v;
    if (v) ws.setColumnKey(v, this);
  }
  get hidden(): boolean { return this._hidden; }
  set hidden(v: boolean) { this._hidden = !!v; }
  get outlineLevel(): number { return this._outline; }
  set outlineLevel(v: number) { this._outline = v || 0; }
  get collapsed(): boolean { return !!(this._outline && this._outline >= this.worksheet.properties.outlineLevelCol); }
  get isDefault(): boolean { return !(this.isCustomWidth || this.hidden || this.outlineLevel || this._xf); }
  toString(): string { return JSON.stringify({ key: this.key, width: this.width, headers: this.headers.length ? this.headers : undefined }); }
  equivalentTo(o: Column): boolean {
    return this.width === o.width && this.hidden === o.hidden && this.outlineLevel === o.outlineLevel && this._xf === o._xf;
  }

  eachCell(options: { includeEmpty?: boolean } | ((c: import('./cell.js').Cell, row: number) => void) | null, cb?: (c: import('./cell.js').Cell, row: number) => void): void {
    const f = cb ?? (options as (c: import('./cell.js').Cell, row: number) => void);
    const o = cb ? (options as { includeEmpty?: boolean } | null) : undefined;
    this.worksheet.eachRow(o ?? {}, (row, n) => f(row.getCell(this.number), n));
  }
  get values(): Raw[] {
    const out: Raw[] = [];
    this.eachCell((c, n) => {
      if (c.type !== ValueType.Null) out[n] = c.value;
    });
    return out;
  }
  set values(v: Raw[]) {
    if (!v) return;
    const off = Object.prototype.hasOwnProperty.call(v, '0') ? 1 : 0;
    v.forEach((x, i) => {
      this.worksheet.getCell(i + off, this.number).value = x;
    });
  }

  // ---- style ----
  private get book() { return bookOf(this.worksheet); }
  get style(): Style {
    const book = this.book;
    return styleView(() => book.get(this._xf), (k, v) => this.apply(k, v));
  }
  set style(s: Style | undefined) {
    const xf = this.book.intern(s);
    if (xf !== this._xf) {
      this.worksheet._freezeColumn?.(this.number);
      this._xf = xf;
    }
  }
  private apply(k: StyleKey, v: unknown): void {
    this.worksheet._freezeColumn?.(this.number);
    this._xf = this.book.withComponent(this._xf, k, v);
    this.eachCell(null, (cell) => {
      (cell as unknown as Record<string, unknown>)[k] = v;
    });
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
}

