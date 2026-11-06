import type { XlsxLimits } from '../limits.js';
import { XlsxLimitError } from '../limits.js';
import { XF } from '../readStyles.js';
import { serialToDate, timeToDate } from '../readSheet.js';
import { decodeText } from '../xmlRead.js';
import { Row } from './row.js';
import type { StyleBook } from './styleBook.js';
import { flags } from './values.js';
import type { Worksheet } from './worksheet.js';

export interface SheetLoadCtx {
  sst: readonly unknown[];
  book: StyleBook;
  date1904: boolean;
  limits: Readonly<XlsxLimits>;
  maxRows?: number;
  maxCols?: number;
}

const isEnd = (c: number): boolean => c === 32 || c === 62 || c === 47 || c === 10 || c === 13 || c === 9;

/** Raw value of `name="..."` inside a start tag's text. */
function at(tag: string, name: string): string | undefined {
  const i = tag.indexOf(` ${name}=`);
  if (i < 0) return undefined;
  const q = i + name.length + 2;
  const quote = tag.charCodeAt(q);
  if (quote !== 34 && quote !== 39) return undefined;
  const e = tag.indexOf(String.fromCharCode(quote), q + 1);
  return e < 0 ? undefined : tag.slice(q + 1, e);
}

function colOf(ref: string): number {
  let n = 0;
  for (let k = 0; k < ref.length; k++) {
    const c = ref.charCodeAt(k);
    if (c < 65 || c > 90) break;
    n = n * 26 + (c - 64);
  }
  return n;
}

function inlineText(body: string): string {
  let out = '';
  const is = body.indexOf('<is');
  const end = body.indexOf('</is>', is);
  const seg = body.slice(is, end < 0 ? body.length : end).replace(/<rPh[\s>][\s\S]*?<\/rPh>/g, '');
  const re = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(seg))) out += decodeText(m[1]!);
  return out;
}

/** Fills `ws` from the text between `<sheetData>` and `</sheetData>`; the rows' value arrays are built in place, so loading copies nothing. */
export function loadSheetData(ws: Worksheet, xml: string, from: number, to: number, ctx: SheetLoadCtx): void {
  const { sst, book, limits } = ctx;
  const maxRows = limits.maxRows;
  const maxCols = limits.maxColumns;
  const nxf = book.xfs.length;
  const shared = new Map<string, string>(); // si -> master address
  let rowNum = 0;
  let p = xml.indexOf('<row', from);
  while (p >= 0 && p < to) {
    if (!isEnd(xml.charCodeAt(p + 4))) {
      p = xml.indexOf('<row', p + 4);
      continue;
    }
    const gt = xml.indexOf('>', p);
    if (gt < 0) break;
    const rtag = xml.slice(p + 4, gt);
    const rAttr = at(rtag, 'r');
    rowNum = rAttr !== undefined ? parseInt(rAttr, 10) : rowNum + 1;
    if (!(rowNum >= 1) || rowNum > maxRows) throw new XlsxLimitError('maxRows', `row number ${rAttr ?? rowNum} passes ${maxRows}`);
    if (ctx.maxRows !== undefined && rowNum > ctx.maxRows) throw new Error(`Max row count (${ctx.maxRows}) exceeded`);
    const row = new Row(ws, rowNum);
    ws._rows[rowNum - 1] = row;
    const ht = at(rtag, 'ht');
    if (ht !== undefined && at(rtag, 'customHeight') !== '0') row.height = parseFloat(ht);
    if (at(rtag, 'hidden') === '1' || at(rtag, 'hidden') === 'true') row._hidden = true;
    const ol = at(rtag, 'outlineLevel');
    if (ol) row._outline = parseInt(ol, 10) || 0;
    const rs = at(rtag, 's');
    if (rs && (at(rtag, 'customFormat') === '1' || at(rtag, 'customFormat') === 'true')) {
      const x = parseInt(rs, 10);
      if (x > 0 && x < nxf) row._xf = x;
    }
    const selfClose = xml.charCodeAt(gt - 1) === 47;
    let next = gt + 1;
    if (!selfClose) {
      let rowEnd = xml.indexOf('</row>', gt);
      if (rowEnd < 0 || rowEnd > to) rowEnd = to;
      next = rowEnd + 6;
      const v = row._v;
      let s: number[] | null = null;
      let col = 0;
      let q = xml.indexOf('<c', gt);
      while (q >= 0 && q < rowEnd) {
        if (!isEnd(xml.charCodeAt(q + 2))) {
          q = xml.indexOf('<c', q + 2);
          continue;
        }
        const cg = xml.indexOf('>', q);
        const ctag = xml.slice(q + 2, cg);
        const ref = at(ctag, 'r');
        col = ref !== undefined ? colOf(ref) : col + 1;
        if (col < 1 || col > maxCols) throw new XlsxLimitError('maxColumns', `column ${col} passes ${maxCols}`);
        if (ctx.maxCols !== undefined && col > ctx.maxCols) throw new Error(`Max column count (${ctx.maxCols}) exceeded`);
        const st = at(ctag, 's');
        const xf = st ? parseInt(st, 10) : 0;
        const sx = xf > 0 && xf < nxf ? xf : 0;
        const type = at(ctag, 't');
        let val: unknown = null;
        let ce: number;
        if (xml.charCodeAt(cg - 1) === 47) {
          ce = cg + 1;
        } else {
          ce = xml.indexOf('</c>', cg);
          if (ce < 0 || ce > rowEnd) ce = rowEnd;
          val = cellValue(xml, cg + 1, ce, type, sx, ctx, shared, ref ?? `${col}`);
        }
        v[col - 1] = val;
        if (sx) (s ??= [])[col - 1] = sx;
        q = xml.indexOf('<c', ce);
      }
      if (s) row._s = s;
    }
    p = xml.indexOf('<row', next);
  }
}

function cellValue(xml: string, from: number, to: number, type: string | undefined, sx: number, ctx: SheetLoadCtx, shared: Map<string, string>, ref: string): unknown {
  const body = xml.slice(from, to);
  const fi = body.indexOf('<f');
  let ftext: string | undefined;
  let ftag = '';
  if (fi >= 0 && isEnd(body.charCodeAt(fi + 2))) {
    const fg = body.indexOf('>', fi);
    ftag = body.slice(fi + 2, fg);
    if (body.charCodeAt(fg - 1) !== 47) {
      const fe = body.indexOf('</f>', fg);
      ftext = decodeText(body.slice(fg + 1, fe < 0 ? body.length : fe));
    } else ftext = '';
  }
  let result: unknown;
  const vi = body.indexOf('<v');
  let raw: string | undefined;
  if (vi >= 0 && isEnd(body.charCodeAt(vi + 2))) {
    const vg = body.indexOf('>', vi);
    if (body.charCodeAt(vg - 1) !== 47) {
      const ve = body.indexOf('</v>', vg);
      raw = body.slice(vg + 1, ve < 0 ? body.length : ve);
    }
  }
  switch (type) {
    case 's': {
      if (raw === undefined || raw === '') break;
      const idx = +raw;
      result = ctx.sst[idx];
      break;
    }
    case 'str':
      if (raw !== undefined) result = decodeText(raw);
      break;
    case 'inlineStr':
      result = inlineText(body);
      break;
    case 'b':
      if (raw !== undefined) result = raw === '1' || raw === 'true';
      break;
    case 'e':
      if (raw !== undefined) result = { error: decodeText(raw) };
      break;
    case 'd':
      if (raw) {
        const ms = Date.parse(raw.includes('T') && !/Z|[+-]\d\d:?\d\d$/.test(raw) ? raw + 'Z' : raw);
        result = Number.isNaN(ms) ? raw : new Date(ms);
      }
      break;
    default: {
      if (raw === undefined || raw === '') break;
      const n = Number(raw);
      if (n !== n) break;
      const kind = sx ? ctx.book.kind(sx) : XF.NONE;
      result = kind === XF.NONE ? n : kind === XF.TIME ? timeToDate(n) : serialToDate(n, ctx.date1904);
    }
  }
  if (ftext === undefined) return result === undefined ? null : result;
  const f: Record<string, unknown> = {};
  const ft = at(ftag, 't');
  const si = at(ftag, 'si');
  if (ft === 'shared' && si !== undefined) {
    if (ftext !== '') {
      shared.set(si, ref);
      f.shareType = 'shared';
      const fref = at(ftag, 'ref');
      if (fref) f.ref = fref;
      f.formula = ftext;
    } else {
      const m = shared.get(si);
      if (m) f.sharedFormula = m;
      else f.formula = ftext;
    }
    flags.shared = true;
  } else if (ft === 'array') {
    f.shareType = 'array';
    const fref = at(ftag, 'ref');
    if (fref) f.ref = fref;
    f.formula = ftext;
  } else f.formula = ftext;
  if (result !== undefined) f.result = result;
  return f;
}

