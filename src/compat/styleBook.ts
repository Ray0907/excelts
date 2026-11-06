import { XML_DECLARATION, escapeXml } from '../xml.js';
import { classifyFormat, XF } from '../readStyles.js';
import { alignmentXml, borderXml, BUILTIN_FORMATS, BUILTIN_IDS, fillXml, fontXml, normalizeAlignment, protectionXml } from './styleEmit.js';
import type { Alignment, Border, Fill, Font, Protection, Style, StyleKey } from './styleTypes.js';

const FIRST_CUSTOM = 164;
const DATE_FORMAT_ID = 14; // mm-dd-yy, what exceljs gives a date cell without a number format

/** One deduplicated component table. Entry 0 is the "absent" entry; `xml` is what the part carries at that index. */
class Table<T> {
  readonly items: (T | undefined)[] = [undefined];
  readonly xml: string[] = [];
  private readonly index = new Map<string, number>();
  constructor(private readonly render: (v: T) => string, base: string) {
    this.xml.push(base);
  }
  add(v: T): number {
    const k = JSON.stringify(v);
    let i = this.index.get(k);
    if (i === undefined) {
      i = this.items.length;
      this.index.set(k, i);
      this.items.push(freeze(JSON.parse(k) as T));
      this.xml.push(this.render(v));
    }
    return i;
  }
  reset(): void {
    this.items.length = 0;
    this.xml.length = 0;
    this.index.clear();
  }
  /** Loaded entries keep their parsed form and their original XML. */
  addLoaded(v: T | undefined, xml: string): void {
    const i = this.items.length;
    this.items.push(v === undefined ? undefined : freeze(v));
    this.xml.push(xml);
    if (v !== undefined) {
      const k = JSON.stringify(v);
      if (!this.index.has(k)) this.index.set(k, i);
    }
  }
}

function freeze<T>(o: T): T {
  if (o && typeof o === 'object') {
    for (const v of Object.values(o)) freeze(v);
    Object.freeze(o);
  }
  return o;
}

export interface Xf { n: number; f: number; fi: number; b: number; a: number; p: number }

export class StyleBook {
  readonly fonts = new Table<Font>(fontXml, '<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>');
  readonly fills = new Table<Fill>(fillXml, '<fill><patternFill patternType="none"/></fill>');
  readonly borders = new Table<Border>(borderXml, '<border><left/><right/><top/><bottom/><diagonal/></border>');
  readonly aligns = new Table<Alignment>(alignmentXml, '');
  readonly prots = new Table<Protection>(protectionXml, '');
  /** numFmtId -> code, for custom formats (id >= 164, or any id the file defined). */
  readonly customFormats = new Map<number, string>();
  private readonly formatIds = new Map<string, number>();
  private nextFormat = FIRST_CUSTOM;
  readonly xfs: Xf[] = [];
  private readonly xfIndex = new Map<string, number>();
  /** Raw parts of a loaded styles.xml that are carried over unchanged. */
  raw: { cellStyleXfs?: string; cellStyles?: string; dxfs?: string; tableStyles?: string; colors?: string } = {};
  private readonly snapshots: (Style | undefined)[] = [];
  private readonly mergeCache = new Map<number, number>();
  private readonly dateCache = new Map<number, number>();
  private readonly kindCache = new Map<number, number>();

  constructor() {
    this.fills.addLoaded({ type: 'pattern', pattern: 'gray125' }, '<fill><patternFill patternType="gray125"/></fill>');
    this.xfs.push({ n: 0, f: 0, fi: 0, b: 0, a: 0, p: 0 });
    this.xfIndex.set('0,0,0,0,0,0', 0);
  }

  /** Empties every table so a loaded styles part can be replayed in file order. */
  resetForLoad(): void {
    for (const t of [this.fonts, this.fills, this.borders, this.aligns, this.prots] as Table<unknown>[]) t.reset();
    this.aligns.addLoaded(undefined, '');
    this.prots.addLoaded(undefined, '');
    this.xfs.length = 0;
    this.xfIndex.clear();
    this.snapshots.length = 0;
  }

  // ---- number formats ----
  formatId(code: string | undefined): number {
    if (!code || code === 'General') return 0;
    const b = BUILTIN_IDS.get(code);
    if (b !== undefined) return b;
    let id = this.formatIds.get(code);
    if (id === undefined) {
      while (this.customFormats.has(this.nextFormat)) this.nextFormat++;
      id = this.nextFormat++;
      this.customFormats.set(id, code);
      this.formatIds.set(code, id);
    }
    return id;
  }
  defineFormat(id: number, code: string): void {
    this.customFormats.set(id, code);
    if (!this.formatIds.has(code)) this.formatIds.set(code, id);
  }
  formatCode(id: number): string | undefined {
    return id ? (this.customFormats.get(id) ?? BUILTIN_FORMATS[id]) : undefined;
  }

  // ---- xf interning ----
  addXf(x: Xf): number {
    const k = `${x.n},${x.f},${x.fi},${x.b},${x.a},${x.p}`;
    let i = this.xfIndex.get(k);
    if (i === undefined) {
      i = this.xfs.length;
      this.xfs.push({ ...x });
      this.xfIndex.set(k, i);
    }
    return i;
  }
  /** Loaded cellXfs keep their file index even when two entries are equal. */
  pushLoadedXf(x: Xf): void {
    const k = `${x.n},${x.f},${x.fi},${x.b},${x.a},${x.p}`;
    if (!this.xfIndex.has(k)) this.xfIndex.set(k, this.xfs.length);
    this.xfs.push({ ...x });
  }

  private part(name: StyleKey, v: unknown): number {
    if (v === undefined || v === null) return 0;
    switch (name) {
      case 'numFmt': return this.formatId(v as string);
      case 'font': return this.fonts.add(v as Font);
      case 'fill': return this.fills.add(v as Fill);
      case 'border': return this.borders.add(v as Border);
      case 'alignment': {
        const a = normalizeAlignment(v as Alignment);
        return a ? this.aligns.add(a) : 0;
      }
      case 'protection': return Object.keys(v as object).length ? this.prots.add(v as Protection) : 0;
    }
  }

  intern(s: Style | undefined): number {
    if (!s) return 0;
    return this.addXf({
      n: this.part('numFmt', s.numFmt), f: this.part('font', s.font), fi: this.part('fill', s.fill),
      b: this.part('border', s.border), a: this.part('alignment', s.alignment), p: this.part('protection', s.protection),
    });
  }

  /** Replaces one component of an xf, reusing the other ids. */
  withComponent(xf: number, name: StyleKey, v: unknown): number {
    const x = { ...this.xfs[xf]! };
    const id = this.part(name, v);
    switch (name) {
      case 'numFmt': x.n = id; break;
      case 'font': x.f = id; break;
      case 'fill': x.fi = id; break;
      case 'border': x.b = id; break;
      case 'alignment': x.a = id; break;
      case 'protection': x.p = id; break;
    }
    return this.addXf(x);
  }

  /** Per component, the row's wins and the column's fills the gaps (exceljs `_mergeStyle`). */
  merge(rowXf: number, colXf: number): number {
    if (!rowXf) return colXf;
    if (!colXf) return rowXf;
    const key = rowXf * 65536 + colXf;
    let r = this.mergeCache.get(key);
    if (r === undefined) {
      const a = this.xfs[rowXf]!, b = this.xfs[colXf]!;
      r = this.addXf({ n: a.n || b.n, f: a.f || b.f, fi: a.fi || b.fi, b: a.b || b.b, a: a.a || b.a, p: a.p || b.p });
      this.mergeCache.set(key, r);
    }
    return r;
  }

  /** The xf a date cell is written with: its own, with the default date format when it has none. */
  dateXf(xf: number): number {
    let r = this.dateCache.get(xf);
    if (r === undefined) {
      const x = this.xfs[xf]!;
      r = x.n ? xf : this.addXf({ ...x, n: DATE_FORMAT_ID });
      this.dateCache.set(xf, r);
    }
    return r;
  }

  /** XF.NONE / DATE / TIME for the number format of an xf. */
  kind(xf: number): number {
    let k = this.kindCache.get(xf);
    if (k === undefined) {
      const n = this.xfs[xf]?.n ?? 0;
      const code = this.formatCode(n);
      k = code ? classifyFormat(code) : XF.NONE;
      this.kindCache.set(xf, k);
    }
    return k;
  }

  /** The style an xf stands for; a shared object, never mutate it. */
  get(xf: number): Style {
    let s = this.snapshots[xf];
    if (!s) {
      const x = this.xfs[xf];
      if (!x) throw new RangeError(`unknown style id ${xf}`);
      const o: Style = {};
      const fmt = this.formatCode(x.n);
      if (fmt) o.numFmt = fmt;
      if (x.f) o.font = this.fonts.items[x.f]!;
      if (x.fi) o.fill = this.fills.items[x.fi]!;
      if (x.b) o.border = this.borders.items[x.b]!;
      if (x.a) o.alignment = this.aligns.items[x.a]!;
      if (x.p) o.protection = this.prots.items[x.p]!;
      s = this.snapshots[xf] = o;
    }
    return s;
  }

  toXml(): string {
    const fmts = [...this.customFormats.entries()].sort((a, b) => a[0] - b[0]);
    const numFmts = fmts.length ? `<numFmts count="${fmts.length}">${fmts.map(([id, c]) => `<numFmt numFmtId="${id}" formatCode="${escapeXml(c)}"/>`).join('')}</numFmts>` : '';
    const xfs = this.xfs
      .map((x) => {
        let a = `<xf numFmtId="${x.n}" fontId="${x.f}" fillId="${x.fi}" borderId="${x.b}" xfId="0"`;
        if (x.n) a += ' applyNumberFormat="1"';
        if (x.f) a += ' applyFont="1"';
        if (x.fi) a += ' applyFill="1"';
        if (x.b) a += ' applyBorder="1"';
        if (x.a) a += ' applyAlignment="1"';
        if (x.p) a += ' applyProtection="1"';
        const kids = (x.a ? this.aligns.xml[x.a]! : '') + (x.p ? this.prots.xml[x.p]! : '');
        return kids ? `${a}>${kids}</xf>` : `${a}/>`;
      })
      .join('');
    const r = this.raw;
    return (
      XML_DECLARATION +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      numFmts +
      `<fonts count="${this.fonts.xml.length}">${this.fonts.xml.join('')}</fonts>` +
      `<fills count="${this.fills.xml.length}">${this.fills.xml.join('')}</fills>` +
      `<borders count="${this.borders.xml.length}">${this.borders.xml.join('')}</borders>` +
      (r.cellStyleXfs ?? '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>') +
      `<cellXfs count="${this.xfs.length}">${xfs}</cellXfs>` +
      (r.cellStyles ?? '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>') +
      (r.dxfs ?? '') + (r.tableStyles ?? '') + (r.colors ?? '') +
      '</styleSheet>'
    );
  }
}
