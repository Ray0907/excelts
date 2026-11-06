import type { Workbook } from './workbook.js';
import type { Worksheet } from './worksheet.js';

export interface CsvOptions {
  sheetName?: string;
  sheetId?: number;
  encoding?: string;
  dateFormat?: string;
  dateFormats?: string[];
  dateUTC?: boolean;
  includeEmptyRows?: boolean;
  map?: (v: unknown) => unknown;
  formatterOptions?: { delimiter?: string; quote?: string | boolean; rowDelimiter?: string };
  parserOptions?: { delimiter?: string; quote?: string | boolean };
}

const SPECIAL: Record<string, unknown> = {
  true: true, false: false, '#N/A': { error: '#N/A' }, '#REF!': { error: '#REF!' }, '#NAME?': { error: '#NAME?' },
  '#DIV/0!': { error: '#DIV/0!' }, '#NULL!': { error: '#NULL!' }, '#VALUE!': { error: '#VALUE!' }, '#NUM!': { error: '#NUM!' },
};
const DEFAULT_DATE_FORMATS = ['YYYY-MM-DD[T]HH:mm:ssZ', 'YYYY-MM-DD[T]HH:mm:ss', 'MM-DD-YYYY', 'YYYY-MM-DD'];
const p2 = (n: number): string => String(n).padStart(2, '0');

/** The few dayjs tokens CSV dates use: YYYY MM DD HH mm ss Z and [literals]. */
function formatDate(d: Date, fmt: string | undefined, utc: boolean): string {
  const g = (l: 'FullYear' | 'Month' | 'Date' | 'Hours' | 'Minutes' | 'Seconds'): number => (utc ? d[`getUTC${l}`]() : d[`get${l}`]());
  const off = -d.getTimezoneOffset();
  const z = utc ? 'Z' : `${off >= 0 ? '+' : '-'}${p2(Math.floor(Math.abs(off) / 60))}:${p2(Math.abs(off) % 60)}`;
  const f = fmt ?? 'YYYY-MM-DD[T]HH:mm:ssZ';
  return f.replace(/\[([^\]]*)\]|YYYY|MM|DD|HH|mm|ss|Z/g, (m, lit: string | undefined) => {
    if (lit !== undefined) return lit;
    switch (m) {
      case 'YYYY': return String(g('FullYear')).padStart(4, '0');
      case 'MM': return p2(g('Month') + 1);
      case 'DD': return p2(g('Date'));
      case 'HH': return p2(g('Hours'));
      case 'mm': return p2(g('Minutes'));
      case 'ss': return p2(g('Seconds'));
      default: return z;
    }
  });
}

function parseDate(text: string, fmt: string): Date | null {
  const names: string[] = [];
  const re = fmt
    .replace(/[.*+?^${}()|\\]/g, (c) => (c === '{' || c === '}' ? `\\${c}` : `\\${c}`))
    .replace(/\\\[([^\]]*)\\\]|\[([^\]]*)\]|YYYY|MM|DD|HH|mm|ss|Z/g, (m, a: string | undefined, b: string | undefined) => {
      if (a !== undefined || b !== undefined) return (a ?? b)!;
      names.push(m);
      return m === 'YYYY' ? '(\\d{4})' : m === 'Z' ? '(Z|[+-]\\d\\d:?\\d\\d)' : '(\\d\\d)';
    });
  const m = new RegExp(`^${re}$`).exec(text);
  if (!m) return null;
  const v: Record<string, string> = {};
  names.forEach((n, i) => (v[n] = m[i + 1]!));
  const Y = +(v.YYYY ?? 1970), M = +(v.MM ?? 1) - 1, D = +(v.DD ?? 1), h = +(v.HH ?? 0), mi = +(v.mm ?? 0), s = +(v.ss ?? 0);
  if (M < 0 || M > 11 || D < 1 || D > 31 || h > 23 || mi > 59 || s > 59) return null;
  let d: Date;
  if (v.Z !== undefined) {
    const z = v.Z === 'Z' ? 0 : (v.Z[0] === '-' ? -1 : 1) * (+v.Z.slice(1, 3) * 60 + +v.Z.slice(-2));
    d = new Date(Date.UTC(Y, M, D, h, mi, s) - z * 60000);
  } else d = new Date(Y, M, D, h, mi, s);
  return Number.isNaN(d.getTime()) || (v.Z === undefined && (d.getMonth() !== M || d.getDate() !== D)) ? null : d;
}

function parseCsv(text: string, delimiter: string, quote: string | false): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQ = false;
  let i = 0;
  const n = text.length;
  let touched = false;
  while (i < n) {
    const c = text[i]!;
    if (inQ) {
      if (c === quote) {
        if (text[i + 1] === quote) { field += quote; i += 2; continue; }
        inQ = false;
      } else field += c;
      i++;
    } else if (quote && c === quote && field === '') { inQ = true; touched = true; i++; }
    else if (c === delimiter) { row.push(field); field = ''; touched = true; i++; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && text[i + 1] === '\n') i++;
      row.push(field); rows.push(row); row = []; field = ''; touched = false; i++;
    } else { field += c; touched = true; i++; }
  }
  if (touched || field !== '' || row.length) { row.push(field); rows.push(row); }
  return rows;
}

export class CsvIO {
  constructor(readonly _wb: Workbook) {}

  load(text: string, options: CsvOptions = {}): Worksheet {
    const ws = this._wb.addWorksheet(options.sheetName);
    const formats = options.dateFormats ?? DEFAULT_DATE_FORMATS;
    const map =
      options.map ??
      ((datum: unknown): unknown => {
        const s = datum as string;
        if (s === '') return null;
        const num = Number(s);
        if (!Number.isNaN(num) && num !== Infinity) return num;
        for (const f of formats) {
          const d = parseDate(s, f);
          if (d) return d;
        }
        return SPECIAL[s] !== undefined ? SPECIAL[s] : s;
      });
    const po = options.parserOptions ?? {};
    const quote = po.quote === false ? false : typeof po.quote === 'string' ? po.quote : '"';
    for (const r of parseCsv(text, po.delimiter ?? ',', quote)) ws.addRow(r.map(map));
    return ws;
  }
  async read(stream: AsyncIterable<Uint8Array | string>, options: CsvOptions = {}): Promise<Worksheet> {
    const parts: string[] = [];
    const dec = new TextDecoder(options.encoding && /^utf-?8$/i.test(options.encoding) ? 'utf-8' : (options.encoding ?? 'utf-8'));
    for await (const c of stream) parts.push(typeof c === 'string' ? c : dec.decode(c, { stream: true }));
    return this.load(parts.join(''), options);
  }
  async readFile(filename: string, options: CsvOptions = {}): Promise<Worksheet> {
    const fs = await import('node:fs/promises');
    let data: Uint8Array;
    try {
      data = await fs.readFile(filename);
    } catch (e) {
      if ((e as { code?: string }).code === 'ENOENT') throw new Error(`File not found: ${filename}`);
      throw e;
    }
    return this.load(new TextDecoder().decode(data), options);
  }

  /** The sheet as CSV text. */
  toString(options: CsvOptions = {}): string {
    const ws = this._wb.getWorksheet(options.sheetName ?? options.sheetId);
    const fo = options.formatterOptions ?? {};
    const delim = fo.delimiter ?? ',';
    const quote = fo.quote === false ? false : typeof fo.quote === 'string' ? fo.quote : '"';
    const rowDelim = fo.rowDelimiter ?? '\n';
    const map =
      options.map ??
      ((value: unknown): unknown => {
        if (value) {
          const v = value as { text?: string; hyperlink?: string; formula?: string; result?: unknown; error?: string };
          if (v.text || v.hyperlink) return v.hyperlink || v.text || '';
          if (v.formula || v.result) return v.result || '';
          if (value instanceof Date) return formatDate(value, options.dateFormat, !!options.dateUTC);
          if (v.error) return v.error;
          if (typeof value === 'object') return JSON.stringify(value);
        }
        return value;
      });
    const lines: string[] = [];
    const fmt = (v: unknown): string => {
      if (v === null || v === undefined) return '';
      const s = String(v);
      if (quote && (s.includes(delim) || s.includes(quote) || s.includes('\n') || s.includes('\r'))) return quote + s.split(quote).join(quote + quote) + quote;
      return s;
    };
    const includeEmpty = options.includeEmptyRows === undefined || options.includeEmptyRows;
    let last = 1;
    ws?.eachRow((row, n) => {
      if (includeEmpty) while (last++ < n - 1) lines.push('');
      const values = row.values.slice(1);
      const out: string[] = [];
      for (let j = 0; j < values.length; j++) out.push(fmt(map(values[j])));
      lines.push(out.join(delim));
      last = n;
    });
    return lines.length ? lines.join(rowDelim) : '';
  }
  async writeBuffer(options: CsvOptions = {}): Promise<Uint8Array> {
    return new TextEncoder().encode(this.toString(options));
  }
  async writeFile(filename: string, options: CsvOptions = {}): Promise<void> {
    const fs = await import('node:fs/promises');
    await fs.writeFile(filename, this.toString(options), { encoding: (options.encoding ?? 'utf8') as BufferEncoding });
  }
  async write(stream: { write(b: string, cb?: (e?: Error | null) => void): unknown; end(cb?: () => void): unknown }, options: CsvOptions = {}): Promise<void> {
    await new Promise<void>((resolve, reject) => stream.write(this.toString(options), (e) => (e ? reject(e) : stream.end(() => resolve()))));
  }
}
