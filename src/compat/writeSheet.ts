import { strToU8 } from 'fflate';
import type { SharedStrings } from '../sharedStrings.js';
import { escapeXml } from '../xml.js';
import { decodeAddress, n2l } from './colCache.js';
import { NS_MAIN, NS_REL } from './rels.js';
import { Range } from './range.js';
import { richTextXml } from './richText.js';
import type { Row } from './row.js';
import type { StyleBook } from './styleBook.js';
import { colorXml } from './styleEmit.js';
import { Extras as ExtrasImpl, type Extras } from './writeExtras.js';
import { flags, MergeSlave, slideFormula } from './values.js';
import type { NoteAt } from './notes.js';
import { sheetTail } from './writeSheetTail.js';
import { dateToSerial1900 } from '../serial.js';
import type { Worksheet } from './worksheet.js';

export { NS_MAIN, NS_REL };
const MS_PER_DAY = 86400000;
const UNIX_EPOCH_SERIAL = 25569;
const CHUNK_CHARS = 65536;
const ERRORS = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA']);

export interface WriteCtx {
  strings: SharedStrings;
  book: StyleBook;
  date1904: boolean;
  /** parts that sheets add beside themselves (pictures, notes); created on first use */
  extras?: Extras;
}
export interface HyperlinkOut { ref: string; target?: string; location?: string; tooltip?: string; display?: string }

/** Serialises one worksheet straight from its rows. Nothing is built per cell except the output text. */
export class SheetWriter {
  readonly chunks: Uint8Array[] = [];
  readonly links: HyperlinkOut[] = [];
  readonly notes: NoteAt[] = [];
  private readonly relList: { type: string; target: string; external: boolean }[] = [];
  private buf = '';
  private readonly colXf: number[] = [];
  private shared: Map<string, { si: number; range: Range; closed?: boolean; formula?: string }> | undefined;
  private maxRow = 0;
  private maxCol = 0;
  private sharedCount = 0;

  constructor(
    readonly ws: Worksheet,
    private readonly ctx: WriteCtx,
    private readonly stream = false,
  ) {
    const cols = ws._columns;
    if (cols) for (let i = 0; i < cols.length; i++) this.colXf[i] = cols[i]!._xf;
  }

  private flush(force = false): void {
    if (force || this.buf.length >= CHUNK_CHARS) {
      this.chunks.push(strToU8(this.buf));
      this.buf = '';
    }
  }

  /** Everything up to and including `<sheetData>`. */
  head(): void {
    const ws = this.ws;
    let maxRow = 0, maxCol = 0;
    if (!this.stream) {
      for (let i = 0; i < ws._rows.length; i++) {
        const r = ws._rows[i];
        if (r && (r._v.length || r.height || r._xf)) {
          maxRow = i + 1;
          if (r._v.length > maxCol) maxCol = r._v.length;
        }
      }
      if (flags.shared) this.prepareShared();
    }
    this.maxRow = maxRow;
    this.maxCol = maxCol;
    const { sheetPr, views, fmt, cols } = sheetTail.headParts(ws, this.ctx.book, maxRow);
    this.buf =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_REL}" ` +
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">' +
      sheetPr +
      (maxRow && maxCol ? `<dimension ref="A1:${n2l(maxCol)}${maxRow}"/>` : '') +
      views + fmt + cols + '<sheetData>';
  }

  private prepareShared(rows: Iterable<Row | undefined> = this.ws._rows): void {
    for (const r of rows) {
      if (!r) continue;
      for (let j = 0; j < r._v.length; j++) {
        const v = r._v[j];
        if (v === null || typeof v !== 'object' || v instanceof Date || v instanceof MergeSlave) continue;
        const f = v as { sharedFormula?: string; formula?: string; shareType?: string; ref?: string };
        if (f.shareType === 'shared' && f.formula && f.ref) {
          const addr = `${n2l(j + 1)}${r.number}`;
          const m = (this.shared ??= new Map());
          if (!m.has(addr)) m.set(addr, { si: this.sharedCount++, range: new Range(f.ref), formula: f.formula });
        }
      }
    }
    for (const r of rows) {
      if (!r) continue;
      for (let j = 0; j < r._v.length; j++) {
        const v = r._v[j];
        if (v === null || typeof v !== 'object' || v instanceof Date || v instanceof MergeSlave) continue;
        const sf = (v as { sharedFormula?: string }).sharedFormula;
        if (!sf) continue;
        const key = decodeAddress(sf).address;
        const m = (this.shared ??= new Map());
        const e = m.get(key);
        if (e && !e.closed) e.range.expand(r.number, j + 1, r.number, j + 1);
        else if (!e && (!this.stream || decodeAddress(key).row === r.number)) {
          const created = { si: this.sharedCount++, range: new Range(key, key) };
          m.set(key, created);
          created.range.expand(r.number, j + 1, r.number, j + 1);
        }
      }
    }
  }

  private formulaCell(ref: string, sAttr: string, v: { formula?: string; sharedFormula?: string; result?: unknown; shareType?: string; ref?: string }, addr: string): string {
    let f = '';
    if (v.sharedFormula) {
      const mk = decodeAddress(v.sharedFormula).address;
      const m = this.shared?.get(mk);
      if (m && !m.closed) f = `<f t="shared" si="${m.si}"/>`;
      else if (m?.formula) f = `<f>${escapeXml(slideFormula(m.formula, mk, addr))}</f>`;
    } else if (v.formula) {
      const text = escapeXml(v.formula);
      const m = this.shared?.get(addr);
      if (m) m.formula = v.formula;
      if (m) f = `<f t="shared" ref="${m.range.range}" si="${m.si}">${text}</f>`;
      else if (v.shareType === 'array' && v.ref) f = `<f t="array" ref="${escapeXml(v.ref)}">${text}</f>`;
      else f = `<f>${text}</f>`;
    }
    const res = v.result;
    let t = '';
    let val = '';
    if (typeof res === 'number') val = Number.isFinite(res) ? `<v>${res}</v>` : '';
    else if (typeof res === 'string') { t = ' t="str"'; val = `<v>${escapeXml(res)}</v>`; }
    else if (typeof res === 'boolean') { t = ' t="b"'; val = `<v>${res ? 1 : 0}</v>`; }
    else if (res instanceof Date) {
      const ms = res.getTime();
      if (Number.isFinite(ms)) val = `<v>${this.serial(ms)}</v>`;
    } else if (res && typeof res === 'object' && typeof (res as { error?: unknown }).error === 'string') {
      t = ' t="e"';
      val = `<v>${escapeXml((res as { error: string }).error)}</v>`;
    }
    return `<c r="${ref}"${sAttr}${t}>${f}${val}</c>`;
  }

  private serial(ms: number): number {
    return this.ctx.date1904 ? ms / MS_PER_DAY + UNIX_EPOCH_SERIAL - 1462 : dateToSerial1900(ms);
  }

  /** Streaming: writes a batch of rows that may share formulas with each other; shared formulas of earlier batches are written out in full. */
  rows(batch: Row[]): void {
    if (flags.shared) this.prepareShared(batch);
    for (const r of batch) this.row(r);
    if (this.shared) for (const e of this.shared.values()) e.closed = true;
  }

  row(row: Row): void {
    const book = this.ctx.book;
    const v = row._v;
    const rowXf = row._xf;
    const attrs =
      (rowXf ? ` s="${rowXf}" customFormat="1"` : '') +
      (row.height ? ` ht="${row.height}" customHeight="1"` : '') +
      (row._hidden ? ' hidden="1"' : '') +
      (row._outline ? ` outlineLevel="${row._outline}"` : '') +
      (row.collapsed ? ' collapsed="1"' : '');
    const rn = row.number;
    if (row._n) for (const [col, note] of row._n) this.notes.push({ row: rn, col, note });
    if (!v.length && !attrs) return;
    let out = `<row r="${rn}"${attrs}>`;
    const s = row._s;
    const ctx = this.ctx;
    for (let j = 0; j < v.length; j++) {
      const x = v[j];
      if (x === undefined) continue;
      let xf = s !== null ? s[j] : undefined;
      if (xf === undefined) xf = rowXf ? book.merge(rowXf, this.colXf[j] ?? 0) : (this.colXf[j] ?? 0);
      const ref = n2l(j + 1) + rn;
      const sa = xf ? ` s="${xf}"` : '';
      switch (typeof x) {
        case 'string':
          out += `<c r="${ref}" t="s"${sa}><v>${ctx.strings.add(x)}</v></c>`;
          break;
        case 'number':
          if (Number.isFinite(x)) out += `<c r="${ref}"${sa}><v>${x}</v></c>`;
          else if (xf) out += `<c r="${ref}"${sa}/>`;
          break;
        case 'boolean':
          out += `<c r="${ref}" t="b"${sa}><v>${x ? 1 : 0}</v></c>`;
          break;
        case 'object':
          if (x === null || x instanceof MergeSlave) {
            if (xf) out += `<c r="${ref}"${sa}/>`;
          } else if (x instanceof Date) {
            const ms = x.getTime();
            const dx = book.dateXf(xf);
            if (Number.isFinite(ms)) out += `<c r="${ref}" s="${dx}"><v>${this.serial(ms)}</v></c>`;
            else if (xf) out += `<c r="${ref}"${sa}/>`;
          } else out += this.objectCell(ref, sa, x as Record<string, unknown>, `${n2l(j + 1)}${rn}`);
          break;
        default:
          throw new TypeError(`unsupported cell value of type ${typeof x} at ${ref}`);
      }
    }
    out += '</row>';
    this.buf += out;
    this.flush();
  }

  private objectCell(ref: string, sa: string, x: Record<string, unknown>, addr: string): string {
    if (x.text && x.hyperlink) {
      const h = String(x.hyperlink);
      const link: HyperlinkOut = { ref, tooltip: x.tooltip ? String(x.tooltip) : undefined };
      if (h.startsWith('#')) link.location = h.slice(1);
      else link.target = h;
      this.links.push(link);
      if (x.formula) return this.formulaCell(ref, sa, { formula: String(x.formula), result: String(x.text) }, addr);
      return `<c r="${ref}" t="s"${sa}><v>${this.ctx.strings.add(String(x.text))}</v></c>`;
    }
    if (x.formula || x.sharedFormula) return this.formulaCell(ref, sa, x, addr);
    if (x.richText) return `<c r="${ref}" t="s"${sa}><v>${this.ctx.strings.addRich(richTextXml(x.richText as never))}</v></c>`;
    if (typeof x.error === 'string') {
      if (!ERRORS.has(x.error)) throw new TypeError(`unsupported error value ${JSON.stringify(x.error)} at ${ref}`);
      return `<c r="${ref}" t="e"${sa}><v>${x.error}</v></c>`;
    }
    return `<c r="${ref}" t="s"${sa}><v>${this.ctx.strings.add(JSON.stringify(x))}</v></c>`;
  }

  /** Writes every row and the tail; returns once `chunks` holds the whole part. */
  finish(): void {
    this.buf += '</sheetData>';
    this.buf += sheetTail.tail(this.ws, this.ctx.book);
    this.buf += this.linksXml();
    this.buf += sheetTail.printParts(this.ws);
    const ctx = this.ctx;
    this.buf += (ctx.extras ??= new ExtrasImpl()).sheet(this);
    this.buf += '</worksheet>';
    this.flush(true);
  }

  private linksXml(): string {
    if (!this.links.length) return '';
    return (
      '<hyperlinks>' +
      this.links
        .map((l) => {
          const a = l.target !== undefined ? ` r:id="${this.addRel('hyperlink', l.target, true)}"` : ` location="${escapeXml(l.location ?? '')}"`;
          return `<hyperlink ref="${l.ref}"${a}${l.tooltip ? ` tooltip="${escapeXml(l.tooltip)}"` : ''}/>`;
        })
        .join('') +
      '</hyperlinks>'
    );
  }

  /** Registers a relationship of this sheet and returns its id. */
  addRel(type: string, target: string, external = false): string {
    this.relList.push({ type, target, external });
    return `rId${this.relList.length}`;
  }

  /** The sheet's relationships part (hyperlinks, pictures, notes), or undefined when it has none. Valid once `finish` has run. */
  relsXml(): string | undefined {
    if (!this.relList.length) return undefined;
    return (
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      this.relList.map((l, i) => `<Relationship Id="rId${i + 1}" Type="${NS_REL}/${l.type}" Target="${escapeXml(l.target)}"${l.external ? ' TargetMode="External"' : ''}/>`).join('') +
      '</Relationships>'
    );
  }

  /** Whole-sheet write. */
  writeAll(): void {
    this.head();
    for (const r of this.ws._rows) if (r) this.row(r);
    this.finish();
  }
}

export { colorXml };
