import type { Cell } from './cell.js';
import { Column, type ColumnDefn } from './column.js';
import { decodeAddress, l2n } from './colCache.js';
import { ValueType } from './enums.js';
import { NotSupportedError } from './errors.js';
import { Range } from './range.js';
import { Row } from './row.js';
import type { Raw } from './values.js';
import { duplicateRow, fillFormula, insertRow, insertRows, spliceColumns, spliceRows } from './splice.js';
import { bookOf } from './fallbackBook.js';
import { Image, type ImageRangeInput } from './images.js';
import { DataValidations } from './dataValidations.js';
import type { Workbook } from './workbook.js';

export interface WorksheetOptions {
  state?: 'visible' | 'hidden' | 'veryHidden';
  properties?: Partial<SheetProperties>;
  pageSetup?: Record<string, unknown>;
  headerFooter?: Record<string, unknown>;
  views?: SheetView[];
  autoFilter?: string | { from: string | { row: number; column: number }; to: string | { row: number; column: number } } | null;
  tabColor?: { argb?: string; theme?: number };
}
export interface SheetProperties {
  defaultRowHeight: number; dyDescent: number; outlineLevelCol: number; outlineLevelRow: number;
  tabColor?: { argb?: string; theme?: number; indexed?: number }; defaultColWidth?: number;
  outlineProperties?: { summaryBelow?: boolean; summaryRight?: boolean };
  [k: string]: unknown;
}
export interface SheetView {
  state?: 'normal' | 'frozen' | 'split'; style?: string; xSplit?: number; ySplit?: number; topLeftCell?: string; activeCell?: string;
  showGridLines?: boolean; showRowColHeaders?: boolean; rightToLeft?: boolean; zoomScale?: number; zoomScaleNormal?: number;
  [k: string]: unknown;
}

const NAME_ILLEGAL = /[*?:/\\[\]]/;

export class Worksheet {
  id: number;
  orderNo: number;
  state: 'visible' | 'hidden' | 'veryHidden';
  properties: SheetProperties;
  pageSetup: Record<string, any>;
  headerFooter: Record<string, any>;
  views: SheetView[];
  autoFilter: WorksheetOptions['autoFilter'];
  dataValidations = new DataValidations();
  rowBreaks: { id: number; max: number; man: number; min?: number }[] = [];
  /** @internal */ _rows: (Row | undefined)[] = [];
  /** @internal */ _columns: Column[] | null = null;
  /** @internal */ _keys = new Map<string, Column>();
  /** @internal */ _merges = new Map<string, Range>();
  /** @internal pictures and the background picture, in the order added */
  _media: Image[] = [];
  /** @internal what a loaded sheet carried that cannot be written back */
  _unsupported: string[] = [];
  /** @internal loaded hyperlinks are plain cell values; nothing else is kept */
  private _name: string | undefined;

  constructor(
    readonly workbook: Workbook,
    options: WorksheetOptions & { id: number; orderNo?: number; name?: string },
  ) {
    this.id = options.id;
    this.orderNo = options.orderNo ?? 0;
    this.name = options.name as string;
    this.state = options.state ?? 'visible';
    this.properties = { defaultRowHeight: 15, dyDescent: 55, outlineLevelCol: 0, outlineLevelRow: 0, ...options.properties };
    if (options.tabColor) this.properties.tabColor = options.tabColor;
    const ps = options.pageSetup;
    this.pageSetup = {
      margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
      orientation: 'portrait', horizontalDpi: 4294967295, verticalDpi: 4294967295,
      fitToPage: !!(ps && (ps.fitToWidth || ps.fitToHeight) && !ps.scale),
      pageOrder: 'downThenOver', blackAndWhite: false, draft: false, cellComments: 'None', errors: 'displayed', scale: 100,
      fitToWidth: 1, fitToHeight: 1, paperSize: undefined, showRowColHeaders: false, showGridLines: false,
      firstPageNumber: undefined, horizontalCentered: false, verticalCentered: false, rowBreaks: null, colBreaks: null,
      ...ps,
    };
    this.headerFooter = {
      differentFirst: false, differentOddEven: false, oddHeader: null, oddFooter: null, evenHeader: null, evenFooter: null,
      firstHeader: null, firstFooter: null, ...options.headerFooter,
    };
    this.views = options.views ?? [];
    this.autoFilter = options.autoFilter ?? null;
  }

  get name(): string { return this._name as string; }
  set name(name: string | undefined) {
    if (name === undefined) name = `sheet${this.id}`;
    if (this._name === name) return;
    if (typeof name !== 'string') throw new Error('The name has to be a string.');
    if (name === '') throw new Error("The name can't be empty.");
    if (name === 'History') throw new Error('The name "History" is protected. Please use a different name.');
    if (NAME_ILLEGAL.test(name)) throw new Error(`Worksheet name ${name} cannot include any of the following characters: * ? : \\ / [ ]`);
    if (/(^')|('$)/.test(name)) throw new Error(`The first or last character of worksheet name cannot be a single quotation mark: ${name}`);
    if (name.length > 31) {
      console.warn(`Worksheet name ${name} exceeds 31 chars. This will be truncated`);
      name = name.substring(0, 31);
    }
    if (this.workbook._worksheets.find((ws) => ws && ws !== this && ws.name.toLowerCase() === name!.toLowerCase())) {
      throw new Error(`Worksheet name already exists: ${name}`);
    }
    this._name = name;
  }

  destroy(): void { this.workbook.removeWorksheetEx(this); }
  /** @internal set by the append-only writer's sheets */
  _commitRow?(row: Row): void;

  // ---- columns ----
  get columns(): Column[] | null { return this._columns; }
  set columns(value: ColumnDefn[]) {
    const cols: Column[] = [];
    if (this._rows.length) for (let n = 1; n <= value.length; n++) this._freezeColumn(n);
    this._columns = cols;
    value.forEach((defn, i) => {
      const c = new Column(this, i + 1, false);
      cols.push(c);
      c.defn = defn;
    });
  }
  /** @internal */ _colXf(c: number): number { return this._columns?.[c - 1]?._xf ?? 0; }
  /** @internal gives the existing cells of a column their own style before the column's changes. */
  _freezeColumn(c: number): void {
    const xf = this._colXf(c);
    for (const r of this._rows) {
      if (!r) continue;
      const j = c - 1;
      if (r._v[j] !== undefined && r._s?.[j] === undefined) r._setXfAt(j, bookOf(this).merge(r._xf, xf));
    }
  }
  /** @internal column letters or a column key to a number */
  _colNumber(c: string): number { return this._keys.get(c)?.number ?? l2n(c); }
  getColumnKey(key: string): Column | undefined { return this._keys.get(key); }
  setColumnKey(key: string, value: Column): void { this._keys.set(key, value); }
  deleteColumnKey(key: string): void { this._keys.delete(key); }
  eachColumnKey(f: (c: Column, key: string) => void): void { this._keys.forEach(f); }
  getColumn(c: number | string): Column {
    if (typeof c === 'string') {
      const k = this._keys.get(c);
      if (k) return k;
      c = l2n(c);
    }
    const cols = (this._columns ??= []);
    for (let n = cols.length + 1; n <= c; n++) cols.push(new Column(this, n));
    return cols[c - 1]!;
  }
  get lastColumn(): Column { return this.getColumn(this.columnCount); }
  get columnCount(): number {
    let m = 0;
    for (const r of this._rows) if (r && r._v.length > m) m = r._v.length;
    return m;
  }
  get actualColumnCount(): number {
    const seen = new Set<number>();
    this.eachRow((r) => r.eachCell((_c, col) => seen.add(col)));
    return seen.size;
  }

  // ---- rows ----
  get _lastRowNumber(): number {
    let n = this._rows.length;
    while (n > 0 && this._rows[n - 1] === undefined) n--;
    return n;
  }
  get rowCount(): number { return this._lastRowNumber; }
  get lastRow(): Row | undefined { return this._rows.length ? this._rows[this._rows.length - 1] : undefined; }
  get actualRowCount(): number {
    let c = 0;
    this.eachRow(() => c++);
    return c;
  }
  findRow(r: number): Row | undefined { return this._rows[r - 1]; }
  findRows(start: number, length: number): (Row | undefined)[] { return this._rows.slice(start - 1, start - 1 + length); }
  getRow(r: number): Row {
    let row = this._rows[r - 1];
    if (!row) row = this._rows[r - 1] = new Row(this, r);
    return row;
  }
  getRows(start: number, length: number): Row[] | undefined {
    if (length < 1) return undefined;
    const out: Row[] = [];
    for (let i = start; i < start + length; i++) out.push(this.getRow(i));
    return out;
  }
  addRow(value?: Raw[] | Record<string, Raw> | null, style = 'n'): Row {
    const n = this._lastRowNumber + 1;
    const row = this.getRow(n);
    row.values = value;
    if (style.charCodeAt(0) === 105 /* i */) this._setStyleOption(n, style);
    return row;
  }
  addRows(values: Iterable<Raw[] | Record<string, Raw>>, style = 'n'): Row[] {
    const out: Row[] = [];
    for (const v of values) out.push(this.addRow(v, style));
    return out;
  }
  eachRow(options: { includeEmpty?: boolean } | ((r: Row, n: number) => void) | undefined, cb?: (r: Row, n: number) => void): void {
    const f = cb ?? (options as (r: Row, n: number) => void);
    const o = cb ? (options as { includeEmpty?: boolean } | undefined) : undefined;
    if (o && o.includeEmpty) {
      const n = this._rows.length;
      for (let i = 1; i <= n; i++) f(this.getRow(i), i);
    } else {
      for (const r of this._rows) if (r && r.hasValues) f(r, r.number);
    }
  }
  getSheetValues(): Raw[][] {
    const out: Raw[][] = [];
    for (const r of this._rows) if (r) out[r.number] = r.values;
    return out;
  }
  get dimensions(): Range {
    const d = new Range();
    for (const r of this._rows) {
      const rd = r?.dimensions;
      if (r && rd) d.expand(r.number, rd.min, r.number, rd.max);
    }
    return d;
  }
  /** @internal */
  _setStyleOption(pos: number, style = 'n'): void {
    if (style[0] === 'o' && this.findRow(pos + 1) !== undefined) this._copyStyle(pos + 1, pos, style[1] === '+');
    else if (style[0] === 'i' && this.findRow(pos - 1) !== undefined) this._copyStyle(pos - 1, pos, style[1] === '+');
  }
  private _copyStyle(src: number, dest: number, styleEmpty: boolean): void {
    const s = this.getRow(src);
    const d = this.getRow(dest);
    d._xf = s._xf;
    s.eachCell({ includeEmpty: styleEmpty }, (_cell, col) => {
      d.getCell(col)._xf = s.getCell(col)._xf;
    });
    d.height = s.height;
  }

  // ---- cells ----
  findCell(r: number | string, c?: number): Cell | undefined {
    const a = typeof r === 'string' ? decodeAddress(r) : { row: r, col: c! };
    return this._rows[a.row - 1]?.findCell(a.col);
  }
  getCell(r: number | string, c?: number): Cell {
    if (typeof r === 'string') {
      const a = decodeAddress(r);
      return this.getRow(a.row).getCell(a.col);
    }
    return this.getRow(r).getCell(c as number);
  }

  // ---- merges ----
  mergeCells(...cells: unknown[]): void { this.mergeInternal(new Range(cells), false); }
  mergeCellsWithoutStyle(...cells: unknown[]): void { this.mergeInternal(new Range(cells), true); }
  private mergeInternal(d: Range, ignoreStyle: boolean): void {
    for (const m of this._merges.values()) if (m.intersects(d)) throw new Error('Cannot merge already merged cells');
    const master = this.getCell(d.top, d.left);
    for (let i = d.top; i <= d.bottom; i++) {
      for (let j = d.left; j <= d.right; j++) if (i > d.top || j > d.left) this.getCell(i, j).merge(master, ignoreStyle);
    }
    this._merges.set(master.address, d);
  }
  private unMergeMaster(master: Cell): void {
    const m = this._merges.get(master.address);
    if (!m) return;
    for (let i = m.top; i <= m.bottom; i++) for (let j = m.left; j <= m.right; j++) this.getCell(i, j).unmerge();
    this._merges.delete(master.address);
  }
  get hasMerges(): boolean { return this._merges.size > 0; }
  unMergeCells(...cells: unknown[]): void {
    const d = new Range(cells);
    for (let i = d.top; i <= d.bottom; i++) {
      for (let j = d.left; j <= d.right; j++) {
        const cell = this.findCell(i, j);
        if (!cell) continue;
        if (cell.type === ValueType.Merge) this.unMergeMaster(cell.master);
        else if (this._merges.has(cell.address)) this.unMergeMaster(cell);
      }
    }
  }

  get tabColor(): SheetProperties['tabColor'] { return this.properties.tabColor; }
  set tabColor(v: SheetProperties['tabColor']) { this.properties.tabColor = v; }

  insertRow(pos: number, value: Raw, style = 'n'): Row { return insertRow(this, pos, value, style); }
  insertRows(pos: number, values: Raw[], style = 'n'): Row[] | undefined { return insertRows(this, pos, values, style); }
  spliceRows(start: number, count: number, ...inserts: Raw[]): void { spliceRows(this, start, count, inserts); }
  duplicateRow(rowNum: number, count: number, insert = false): void { duplicateRow(this, rowNum, count, insert); }
  spliceColumns(start: number, count: number, ...inserts: Raw[][]): void { spliceColumns(this, start, count, inserts); }
  fillFormula(range: string, formula: string, results?: unknown, shareType = 'shared'): void { fillFormula(this, range, formula, results, shareType); }
  /** @internal */
  _copyStyleRow(src: number, dest: number, styleEmpty: boolean): void { this._copyStyle(src, dest, styleEmpty); }

  // ---- Tier 3: refuse loudly ----
  // ---- pictures ----
  addImage(imageId: number, range: string | ImageRangeInput): void { this._media.push(new Image(this, 'image', imageId, range)); }
  getImages(): Image[] { return this._media.filter((m) => m.type === 'image'); }
  addBackgroundImage(imageId: number): void { this._media.push(new Image(this, 'background', imageId)); }
  getBackgroundImageId(): number | undefined { return this._media.find((m) => m.type === 'background')?.imageId; }

  protect(): never { throw new NotSupportedError('worksheet.protect (sheet protection)'); }
  unprotect(): never { throw new NotSupportedError('worksheet.unprotect (sheet protection)'); }
  addTable(): never { throw new NotSupportedError('worksheet.addTable (tables)'); }
  getTable(): never { throw new NotSupportedError('worksheet tables'); }
  addPivotTable(): never { throw new NotSupportedError('worksheet.addPivotTable (pivot tables)'); }
  addConditionalFormatting(): never { throw new NotSupportedError('worksheet.addConditionalFormatting'); }
  get model(): never { throw new NotSupportedError('worksheet.model (internal model)'); }
  set model(_v: unknown) { throw new NotSupportedError('worksheet.model (internal model)'); }
}
