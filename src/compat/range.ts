import { decodeAddress, decodeEx, encodeAddress, n2l } from './colCache.js';

interface RangeModel {
  top: number;
  left: number;
  bottom: number;
  right: number;
  sheetName?: string;
}

/** Same surface as exceljs's Range. */
export class Range {
  model!: RangeModel;
  constructor(...args: unknown[]) {
    this.decode(args);
  }

  private setTLBR(t: unknown, l: unknown, b?: unknown, r?: unknown, s?: unknown): void {
    if (typeof t === 'string') {
      const tl = decodeAddress(t);
      const br = decodeAddress(l as string);
      this.model = { top: Math.min(tl.row, br.row), left: Math.min(tl.col, br.col), bottom: Math.max(tl.row, br.row), right: Math.max(tl.col, br.col), sheetName: b as string | undefined };
    } else {
      const [t1, l1, b1, r1] = [t, l, b, r] as number[];
      this.model = { top: Math.min(t1!, b1!), left: Math.min(l1!, r1!), bottom: Math.max(t1!, b1!), right: Math.max(l1!, r1!), sheetName: s as string | undefined };
    }
  }

  private decode(a: unknown[]): void {
    switch (a.length) {
      case 5:
      case 4:
        this.setTLBR(a[0], a[1], a[2], a[3], a[4]);
        break;
      case 3:
      case 2:
        this.setTLBR(a[0], a[1], a[2]);
        break;
      case 1: {
        const v = a[0] as Range | unknown[] | RangeModel | string;
        if (v instanceof Range) this.model = { ...v.model };
        else if (Array.isArray(v)) this.decode(v);
        else if (typeof v === 'object' && v.top && v.left && v.bottom && v.right) this.model = { top: v.top, left: v.left, bottom: v.bottom, right: v.right, sheetName: v.sheetName };
        else {
          const d = decodeEx(v as string);
          this.model = d.top ? { top: d.top, left: d.left!, bottom: d.bottom!, right: d.right!, sheetName: d.sheetName } : { top: d.row!, left: d.col!, bottom: d.row!, right: d.col!, sheetName: d.sheetName };
        }
        break;
      }
      case 0:
        this.model = { top: 0, left: 0, bottom: 0, right: 0 };
        break;
      default:
        throw new Error(`Invalid number of arguments to _getDimensions() - ${a.length}`);
    }
  }

  get top(): number { return this.model.top || 1; }
  set top(v: number) { this.model.top = v; }
  get left(): number { return this.model.left || 1; }
  set left(v: number) { this.model.left = v; }
  get bottom(): number { return this.model.bottom || 1; }
  set bottom(v: number) { this.model.bottom = v; }
  get right(): number { return this.model.right || 1; }
  set right(v: number) { this.model.right = v; }
  get sheetName(): string | undefined { return this.model.sheetName; }
  set sheetName(v: string | undefined) { this.model.sheetName = v; }

  private get sheetPrefix(): string {
    const s = this.model.sheetName;
    return s ? (/^[a-zA-Z0-9]*$/.test(s) ? `${s}!` : `'${s}'!`) : '';
  }
  expand(top: number, left: number, bottom: number, right: number): void {
    if (!this.model.top || top < this.top) this.top = top;
    if (!this.model.left || left < this.left) this.left = left;
    if (!this.model.bottom || bottom > this.bottom) this.bottom = bottom;
    if (!this.model.right || right > this.right) this.right = right;
  }
  expandToAddress(a: string): void {
    const d = decodeEx(a);
    this.expand(d.row!, d.col!, d.row!, d.col!);
  }
  get tl(): string { return n2l(this.left) + this.top; }
  get $t$l(): string { return `$${n2l(this.left)}$${this.top}`; }
  get br(): string { return n2l(this.right) + this.bottom; }
  get $b$r(): string { return `$${n2l(this.right)}$${this.bottom}`; }
  get range(): string { return `${this.sheetPrefix + this.tl}:${this.br}`; }
  get $range(): string { return `${this.sheetPrefix + this.$t$l}:${this.$b$r}`; }
  get shortRange(): string { return this.count > 1 ? this.range : this.sheetPrefix + this.tl; }
  get $shortRange(): string { return this.count > 1 ? this.$range : this.sheetPrefix + this.$t$l; }
  get count(): number { return (1 + this.bottom - this.top) * (1 + this.right - this.left); }
  toString(): string { return this.range; }
  intersects(o: Range): boolean {
    if (o.sheetName && this.sheetName && o.sheetName !== this.sheetName) return false;
    return !(o.bottom < this.top || o.top > this.bottom || o.right < this.left || o.left > this.right);
  }
  contains(a: string): boolean { return this.containsEx(decodeEx(a)); }
  containsEx(a: { row?: number; col?: number; sheetName?: string }): boolean {
    if (a.sheetName && this.sheetName && a.sheetName !== this.sheetName) return false;
    return a.row! >= this.top && a.row! <= this.bottom && a.col! >= this.left && a.col! <= this.right;
  }
  forEachAddress(cb: (address: string, row: number, col: number) => void): void {
    for (let col = this.left; col <= this.right; col++) for (let row = this.top; row <= this.bottom; row++) cb(encodeAddress(row, col), row, col);
  }
}
