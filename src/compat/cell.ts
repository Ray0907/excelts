import { encodeAddress, n2l } from './colCache.js';
import { bookOf } from './fallbackBook.js';
import { FormulaType, ValueType } from './enums.js';
import { NotSupportedError } from './errors.js';
import type { Row } from './row.js';
import { checkNote, type Note } from './notes.js';
import { componentView, styleView } from './styleView.js';
import type { Alignment, Border, Fill, Font, Protection, Style, StyleKey } from './styleTypes.js';
import { MergeSlave, publicType, slideFormula, storeValue, typeOf, valueText, type Raw } from './values.js';
import type { Workbook } from './workbook.js';
import type { Worksheet } from './worksheet.js';

/** A cell is a view: it holds its row and column, and every read or write goes to the row's value array. Identity is cached on the row. */
export class Cell {
  constructor(
    readonly _row: Row,
    readonly _col: number,
  ) {}

  get worksheet(): Worksheet { return this._row.worksheet; }
  get workbook(): Workbook { return this._row.worksheet.workbook; }
  get row(): number { return this._row.number; }
  get col(): number { return this._col; }
  get address(): string { return encodeAddress(this._row.number, this._col); }
  get $col$row(): string { return `$${n2l(this._col)}$${this._row.number}`; }
  get fullAddress(): { sheetName: string; address: string; row: number; col: number } {
    return { sheetName: this.worksheet.name, address: this.address, row: this.row, col: this.col };
  }

  // ---- value ----
  private get raw(): Raw { return this._row._v[this._col - 1]; }
  get value(): unknown {
    const r = this.raw;
    if (r === undefined) return null;
    if (r instanceof MergeSlave) return this.master.value;
    return r;
  }
  set value(v: unknown) {
    const r = this.raw;
    if (r instanceof MergeSlave) {
      this.master.value = v;
      return;
    }
    this._row._v[this._col - 1] = storeValue(v);
  }
  get type(): number { return publicType(this.raw); }
  get effectiveType(): number {
    const r = this.raw;
    if (r instanceof MergeSlave) return this.master.effectiveType;
    const t = typeOf(r);
    if (t !== ValueType.Formula) return t === ValueType.JSON ? ValueType.String : t;
    const res = (r as { result?: unknown }).result;
    if (res === null || res === undefined) return ValueType.Null;
    if (typeof res === 'string') return ValueType.String;
    if (typeof res === 'number') return ValueType.Number;
    if (res instanceof Date) return ValueType.Date;
    return ValueType.Null;
  }
  get text(): string {
    const r = this.raw;
    return r instanceof MergeSlave ? valueText(this.master.value) : valueText(r);
  }
  get html(): string { return this.text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
  toString(): string { return this.text; }
  get formula(): string | undefined {
    const r = this.raw as { formula?: string; sharedFormula?: string } | null;
    if (typeOf(r) !== ValueType.Formula) return undefined;
    if (r!.formula) return r!.formula;
    if (r!.sharedFormula) {
      const m = this.worksheet.findCell(r!.sharedFormula);
      const mf = m?.formula;
      return mf && m ? slideFormula(mf, m.address, this.address) : undefined;
    }
    return undefined;
  }
  get result(): unknown { return typeOf(this.raw) === ValueType.Formula ? (this.raw as { result?: unknown }).result : undefined; }
  get formulaType(): number {
    const r = this.raw as { formula?: string; sharedFormula?: string } | null;
    if (typeOf(r) !== ValueType.Formula) return FormulaType.None;
    return r!.formula ? FormulaType.Master : r!.sharedFormula ? FormulaType.Shared : FormulaType.None;
  }
  get isHyperlink(): boolean { return this.type === ValueType.Hyperlink; }
  get hyperlink(): string | undefined { return this.isHyperlink ? (this.raw as { hyperlink: string }).hyperlink : undefined; }

  // ---- merge ----
  get isMerged(): boolean { return this.raw instanceof MergeSlave || this._row.worksheet._merges.has(this.address); }
  get master(): Cell {
    const r = this.raw;
    return r instanceof MergeSlave ? this._row.worksheet.getCell(r.masterRow, r.masterCol) : this;
  }
  merge(master: Cell, ignoreStyle?: boolean): void {
    this._row._v[this._col - 1] = new MergeSlave(master.row, master.col);
    if (!ignoreStyle) this._xf = master._xf;
  }
  unmerge(): void {
    if (this.raw instanceof MergeSlave) {
      this._row._v[this._col - 1] = null;
      this._xf = this._row._inheritedXf(this._col - 1);
    }
  }
  isMergedTo(master: Cell): boolean {
    const r = this.raw;
    return r instanceof MergeSlave && r.masterRow === master.row && r.masterCol === master.col;
  }

  // ---- style ----
  get _xf(): number { return this._row._xfAt(this._col - 1); }
  set _xf(xf: number) { this._row._setXfAt(this._col - 1, xf); }
  private get book() { return bookOf(this._row.worksheet); }
  get style(): Style {
    const book = this.book;
    return styleView(() => book.get(this._xf), (k, v) => this.setComponent(k, v));
  }
  set style(s: Style | undefined) { this._xf = this.book.intern(s); }
  private setComponent(k: StyleKey, v: unknown): void { this._xf = this.book.withComponent(this._xf, k, v); }
  private component<T extends object>(k: StyleKey): T | undefined {
    const cur = this.book.get(this._xf)[k] as T | undefined;
    return componentView(cur, k, (n) => this.setComponent(k, n));
  }
  get numFmt(): string | undefined { return this.book.get(this._xf).numFmt; }
  set numFmt(v: string | undefined) { this.setComponent('numFmt', v); }
  get font(): Font | undefined { return this.component<Font>('font'); }
  set font(v: Font | undefined) { this.setComponent('font', v); }
  get alignment(): Alignment | undefined { return this.component<Alignment>('alignment'); }
  set alignment(v: Alignment | undefined) { this.setComponent('alignment', v); }
  get border(): Border | undefined { return this.component<Border>('border'); }
  set border(v: Border | undefined) { this.setComponent('border', v); }
  get fill(): Fill | undefined { return this.component<Fill>('fill'); }
  set fill(v: Fill | undefined) { this.setComponent('fill', v); }
  get protection(): Protection | undefined { return this.component<Protection>('protection'); }
  set protection(v: Protection | undefined) { this.setComponent('protection', v); }

  // ---- unsupported ----
  get note(): Note | undefined { return this._row._n?.get(this._col); }
  set note(v: Note | undefined | null) {
    if (v === undefined || v === null) {
      this._row._n?.delete(this._col);
      return;
    }
    (this._row._n ??= new Map()).set(this._col, checkNote(v));
  }
  get model(): never { throw new NotSupportedError('cell.model (internal model)'); }
  set model(_v: unknown) { throw new NotSupportedError('cell.model (internal model)'); }
  // ---- defined names ----
  get names(): string[] { return this.workbook.definedNames.getNamesEx(this.fullAddress); }
  set names(names: string[]) {
    this.removeAllNames();
    for (const n of names) this.addName(n);
  }
  get name(): string | undefined { return this.names[0]; }
  set name(name: string | undefined) { this.names = name === undefined || name === null ? [] : [name]; }
  addName(name: string): void { this.workbook.definedNames.addEx(this.fullAddress, name); }
  removeName(name: string): void { this.workbook.definedNames.removeEx(this.fullAddress, name); }
  removeAllNames(): void { this.workbook.definedNames.removeAllNames(this.fullAddress); }
  get dataValidation(): unknown { return this.worksheet.dataValidations.find(this.address); }
  set dataValidation(v: unknown) { this.worksheet.dataValidations.add(this.address, v as never); }
}
