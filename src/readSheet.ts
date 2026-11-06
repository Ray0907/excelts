import type { CellValue } from './index.js';
import { XF } from './readStyles.js';
import { ascii, viewOf, type ByteView } from './bytes.js';
import { rootPrefixOf, stripPrefix } from './dialect.js';
import { decodeEntities, decodeText, readText } from './xmlRead.js';
import { DEFAULT_LIMITS, XlsxLimitError, type XlsxLimits } from './limits.js';
import { RowIterator } from './rowIterator.js';

export interface SheetContext {
  sharedStrings: readonly string[];
  /** `XF.*` per cellXfs entry, see readStyles. */
  xfKind: ArrayLike<number>;
  date1904: boolean;
  /** Defaults apply when absent. */
  limits?: Readonly<XlsxLimits>;
}

const MS_PER_DAY = 86400000;
const UNIX_EPOCH_SERIAL_1900 = 25569;
const UNIX_EPOCH_SERIAL_1904 = 24107;
const LOTUS_LEAP_BUG_SERIAL = 60;

export function serialToDate(serial: number, date1904: boolean): Date {
  if (date1904) return new Date(Math.round((serial - UNIX_EPOCH_SERIAL_1904) * MS_PER_DAY));
  // Excel counts a nonexistent 1900-02-29, so serials before it sit one day later than the linear formula says.
  const days = serial < LOTUS_LEAP_BUG_SERIAL ? serial - UNIX_EPOCH_SERIAL_1900 + 1 : serial - UNIX_EPOCH_SERIAL_1900;
  return new Date(Math.round(days * MS_PER_DAY));
}

/** A time of day or elapsed duration is a plain offset from 1899-12-30, whatever the workbook's epoch and without the 1900 leap-year shift. */
export function timeToDate(serial: number): Date {
  return new Date(Math.round((serial - UNIX_EPOCH_SERIAL_1900) * MS_PER_DAY));
}

function numberCell(n: number, style: number, ctx: SheetContext, scan: Scan): CellValue {
  if (style < 0) return n;
  const kind = (scan.xf ??= ctx.xfKind)[style];
  return kind === undefined || kind === XF.NONE ? n : kind === XF.TIME ? timeToDate(n) : serialToDate(n, ctx.date1904);
}

/** `t="d"` ISO 8601; a value without a zone is UTC (JS would read it as local time). */
function isoToDate(s: string): CellValue {
  const hasTime = s.indexOf('T') >= 0;
  const zoned = s.endsWith('Z') || /[+-]\d\d:?\d\d$/.test(s.slice(10));
  const ms = Date.parse(hasTime && !zoned ? s + 'Z' : s);
  return Number.isNaN(ms) ? s : new Date(ms);
}

const T_NONE = 0; // absent or "n"
const T_SHARED = 1;
const T_STR = 2;
const T_INLINE = 3;
const T_BOOL = 4;
const T_ERROR = 5;
const T_DATE = 6;
const T_OTHER = 7;

const SHEET_DATA = ascii('<sheetData');
const SHEET_DATA_END = ascii('</sheetData>');
const ROW = ascii('<row');
const ROW_END = ascii('</row>');
const CELL = ascii('<c');
const CELL_END = ascii('</c>');
const V_OPEN = ascii('<v');
const V_END = ascii('</v>');
const IS = ascii('<is');
const IS_END = ascii('</is>');
const IS_T = ascii('<is><t>');
const T_IS_END = ascii('</t></is>');
const INLINE_STR = ascii('inlineStr');

// Bytes the inline-string scan stops at: '<' ends the text; '&', CR and '_' may start something that needs decoding.
const STOP = new Uint8Array(256);
STOP[60] = STOP[38] = STOP[13] = STOP[95] = 1;

const LT = 60;
const GT = 62;
const SLASH = 47;
const EQ = 61;
const DQUOTE = 34;
const SQUOTE = 39;

// Results of scanTag. Module-level so a tag costs no object or string; the scanner is synchronous, so nothing interleaves.
let tagHasR = false;
let tagCol = -1; // 0-based column from the letters of r="AB12"
let tagNum = 0; // digits of r: the row number for <row r="..">
let tagStyle = -1;
let tagType = T_NONE;

function matches(data: Uint8Array, pattern: Uint8Array, at: number): boolean {
  if (at < 0 || at + pattern.length > data.length) return false;
  for (let k = 0; k < pattern.length; k++) if (data[at + k] !== pattern[k]) return false;
  return true;
}

const NEAR = 48;

/** The last far search for one pattern: `hit` is the first occurrence at or after `from` (-1: none to the end of the part). */
interface Memo {
  from: number;
  hit: number;
}
const memo = (): Memo => ({ from: Infinity, hit: -1 });

/** Cell-level targets sit a few dozen bytes away, where a JS loop beats the per-call cost of a native search; farther hits fall through to it.
 * The far search is remembered per pattern: when the pattern is absent from here on (or its next occurrence lies beyond the cell being read),
 * every later cell asks the same question, and answering from the memo keeps a part of N such cells linear instead of N times the part's size. */
function find(view: ByteView, pattern: Uint8Array, from: number, m: Memo): number {
  const data = view.data;
  const first = pattern[0]!;
  const len = pattern.length;
  const stop = Math.min(from + NEAR, data.length - len + 1);
  for (let i = from; i < stop; i++) {
    if (data[i] !== first) continue;
    let k = 1;
    while (k < len && data[i + k] === pattern[k]) k++;
    if (k === len) return i;
  }
  const at = Math.max(stop, from);
  if (at >= m.from && (m.hit < 0 || m.hit >= at)) return m.hit;
  const hit = view.indexOfSeq(pattern, at);
  m.from = at;
  m.hit = hit;
  return hit;
}

/** Per-scan state: the search memos and the cell text ceiling. One per scanRows call, so interleaved sheet iterators cannot see each other's. */
interface Scan {
  maxChars: number;
  /** The context's tables, taken on first use: the context may load them lazily, and a getter call per cell would show on dense sheets. */
  sst: readonly string[] | undefined;
  xf: ArrayLike<number> | undefined;
  v: Memo;
  vEnd: Memo;
  is: Memo;
  isEnd: Memo;
  cellEnd: Memo;
  cell: Memo;
  row: Memo;
  rowEnd: Memo;
}

function capped(text: string, max: number): string {
  if (text.length > max) throw new XlsxLimitError('maxCellChars', `cell text of ${text.length} characters`);
  return text;
}

/** Walks the attributes of a start tag whose name ends at `from`, reading r, s and t by byte, and returns the index of the closing `>`.
 * Quoted values are skipped whole, so a `>` inside one (ph="a>b") does not end the tag. The first occurrence of an attribute wins. */
function scanTag(data: Uint8Array, from: number): number {
  tagHasR = false;
  tagCol = -1;
  tagNum = 0;
  tagStyle = -1;
  tagType = T_NONE;
  let seenS = false;
  let seenT = false;
  const n = data.length;
  let i = from;
  while (i < n) {
    let c = data[i]!;
    if (c === GT) return i;
    if (c <= 32 || c === SLASH) {
      i++;
      continue;
    }
    const nameStart = i;
    while (i < n) {
      c = data[i]!;
      if (c === EQ || c === GT || c === SLASH || c <= 32) break;
      i++;
    }
    const nameLen = i - nameStart;
    while (i < n && data[i]! <= 32) i++;
    if (data[i] !== EQ) continue;
    i++;
    while (i < n && data[i]! <= 32) i++;
    const quote = data[i];
    if (quote !== DQUOTE && quote !== SQUOTE) continue;
    const vs = i + 1;
    let ve = vs;
    while (ve < n && data[ve] !== quote) ve++;
    if (ve >= n) return n;
    i = ve + 1;
    if (nameLen !== 1) continue;
    const name = data[nameStart];
    if (name === 114 && !tagHasR) {
      tagHasR = true;
      let col = 0;
      let k = vs;
      for (; k < ve; k++) {
        const ch = data[k]!;
        if (ch < 65 || ch > 90) break;
        col = col * 26 + (ch - 64);
      }
      tagCol = col - 1;
      let num = 0;
      for (; k < ve; k++) {
        const d = data[k]! - 48;
        if (d < 0 || d > 9) {
          num = NaN;
          break;
        }
        num = num * 10 + d;
      }
      tagNum = num;
    } else if (name === 115 && !seenS) {
      seenS = true;
      let st = 0;
      for (let k = vs; k < ve; k++) {
        const d = data[k]! - 48;
        if (d < 0 || d > 9) {
          st = -1;
          break;
        }
        st = st * 10 + d;
      }
      tagStyle = st;
    } else if (name === 116 && !seenT) {
      seenT = true;
      tagType = typeCode(data, vs, ve - vs);
    }
  }
  return n;
}

function typeCode(data: Uint8Array, at: number, len: number): number {
  const c0 = data[at];
  if (len === 1) {
    switch (c0) {
      case 115: return T_SHARED;
      case 98: return T_BOOL;
      case 101: return T_ERROR;
      case 100: return T_DATE;
      case 110: return T_NONE;
    }
  } else if (len === 3 && c0 === 115 && data[at + 1] === 116 && data[at + 2] === 114) {
    return T_STR;
  } else if (len === 9 && c0 === 105 && matches(data, INLINE_STR, at)) {
    return T_INLINE;
  }
  return T_OTHER;
}

function isNameEnd(c: number | undefined): boolean {
  return c === 32 || c === 62 || c === 47 || c === 10 || c === 13 || c === 9;
}

// Exact powers of ten: a mantissa below 1e15 divided by one of these is correctly rounded, hence equal to Number(text).
const POW10 = [1, 1e1, 1e2, 1e3, 1e4, 1e5, 1e6, 1e7, 1e8, 1e9, 1e10, 1e11, 1e12, 1e13, 1e14, 1e15, 1e16, 1e17, 1e18, 1e19, 1e20, 1e21, 1e22];

/**
 * Parses `-?digits[.digits]` at `at` followed by exactly `</v>`, returning NaN when the text is anything else
 * (exponent, sign, whitespace, over 15 significant digits, empty), so the caller falls back to Number(). A real NaN never
 * reaches here because `NaN` text is not digits.
 */
function fastNumber(data: Uint8Array, at: number, bodyTo: number): number {
  let i = at;
  let c = data[i]!;
  const neg = c === 45;
  if (neg) c = data[++i]!;
  let mant = 0;
  let sig = 0;
  let digits = 0;
  while (c >= 48 && c <= 57) {
    if (mant !== 0 || c !== 48) {
      if (++sig > 15) return NaN;
      mant = mant * 10 + (c - 48);
    }
    digits++;
    c = data[++i]!;
  }
  let frac = 0;
  if (c === 46) {
    c = data[++i]!;
    while (c >= 48 && c <= 57) {
      if (mant !== 0 || c !== 48) {
        if (++sig > 15) return NaN;
        mant = mant * 10 + (c - 48);
      }
      frac++;
      digits++;
      c = data[++i]!;
    }
  }
  if (digits === 0 || frac > 22 || c !== LT || i + 4 > bodyTo || data[i + 1] !== SLASH || data[i + 2] !== 118 || data[i + 3] !== GT) return NaN;
  const v = frac === 0 ? mant : mant / POW10[frac]!;
  return neg ? -v : v;
}

function cellValue(view: ByteView, type: number, style: number, bodyFrom: number, bodyTo: number, ctx: SheetContext, scan: Scan): CellValue {
  const data = view.data;
  // bodyFrom is the '>' of the cell start tag. `<v>` right after it is the common shape; anything else takes the slow path.
  if ((type === T_NONE || type === T_SHARED) && data[bodyFrom + 1] === LT && data[bodyFrom + 2] === 118 && data[bodyFrom + 3] === GT) {
    if (type === T_NONE) {
      const n = fastNumber(data, bodyFrom + 4, bodyTo);
      if (n === n) {
        if (style >= 0 && (scan.xf ??= ctx.xfKind)[style]) return numberCell(n, style, ctx, scan);
        return n;
      }
    } else {
      let i = bodyFrom + 4;
      let c = data[i]!;
      let idx = 0;
      while (c >= 48 && c <= 57 && i - bodyFrom < 13) {
        idx = idx * 10 + (c - 48);
        c = data[++i]!;
      }
      if (i > bodyFrom + 4 && c === LT && i + 4 <= bodyTo && data[i + 1] === SLASH && data[i + 2] === 118 && data[i + 3] === GT) return (scan.sst ??= ctx.sharedStrings)[idx];
    }
  }
  if (type === T_INLINE) {
    if (matches(data, IS_T, bodyFrom + 1)) {
      // The first '<' after the text must be `</t></is>` itself; searching for that string directly would scan the whole sheet when it is absent.
      // The same pass notes whether the text needs decoding at all: `&`, CR or `_x`.
      const from = bodyFrom + 8;
      let lt = from;
      let special = 0; // 1: only named entities, 2: also CR, `_x` or a numeric reference
      for (;;) {
        while (lt < bodyTo && STOP[data[lt]!] === 0) lt++;
        const c = data[lt];
        if (c === LT || lt >= bodyTo) break;
        if (c === 38) special |= data[lt + 1] === 35 ? 2 : 1; // a numeric reference can spell `_`
        else if (c === 13 || data[lt + 1] === 120) special |= 2;
        lt++;
      }
      if (lt < bodyTo && lt + 9 <= bodyTo && matches(data, T_IS_END, lt)) {
        const text = view.sliceText(from, lt);
        return capped(special === 0 ? text : special === 1 ? decodeEntities(text) : decodeText(text), scan.maxChars);
      }
    }
    const is = find(view, IS, bodyFrom, scan.is);
    if (is >= 0 && is < bodyTo) {
      const isEnd = find(view, IS_END, is, scan.isEnd);
      const text = view.sliceText(is, isEnd < 0 || isEnd > bodyTo ? bodyTo : isEnd);
      return capped(readText(text, 0, text.length, true), scan.maxChars); // the check runs on this one slice, so it costs nothing when absent
    }
  }
  // A formula cell yields its cached <v>; returning the formula itself is a later phase.
  const vs = find(view, V_OPEN, bodyFrom, scan.v);
  if (vs < 0 || vs >= bodyTo || !isNameEnd(data[vs + 2])) return undefined;
  let vgt = vs;
  while (vgt < bodyTo && data[vgt] !== GT) vgt++;
  if (vgt >= bodyTo || data[vgt - 1] === SLASH) return undefined;
  const vEnd = find(view, V_END, vgt, scan.vEnd);
  let raw = view.sliceText(vgt + 1, vEnd < 0 || vEnd > bodyTo ? bodyTo : vEnd);
  if (raw === '') return undefined;
  if (data[vgt + 1] === LT && type !== T_STR && type !== T_INLINE && type !== T_ERROR && raw.includes('<![CDATA[')) raw = decodeText(raw).trim();

  switch (type) {
    case T_NONE: {
      return numberCell(Number(raw), style, ctx, scan);
    }
    case T_SHARED:
      return (scan.sst ??= ctx.sharedStrings)[+raw];
    case T_BOOL:
      return raw === '1' || raw === 'true';
    case T_DATE:
      return isoToDate(raw);
    case T_ERROR:
      return { error: decodeText(raw) };
    case T_INLINE: // <v> without <is>
    case T_STR:
    default:
      return capped(decodeText(raw), scan.maxChars);
  }
}

/** Hand-written scanner over the inflated sheet part's bytes: rows are yielded one at a time, so only the current row is materialised. */
export interface GapBudget {
  /** Called once per run of empty rows the scanner is about to synthesize; throws when the workbook's budget is spent. */
  add(count: number): void;
}

/** Rows of the sheet part in order; row N is the Nth array, a skipped row number yielding an empty array. */
export function scanRows(bytes: Uint8Array, ctx: SheetContext, gaps?: GapBudget): IterableIterator<CellValue[]> {
  return new RowIterator(scanRuns(bytes, ctx, gaps));
}

/** Yields each present row, and for a run of skipped row numbers one number, the length of the run. */
function* scanRuns(bytes: Uint8Array, ctx: SheetContext, gaps?: GapBudget): Generator<CellValue[] | number, void> {
  // The scanner is written for the unprefixed spelling and is the hot loop; a prefixed part is renamed to it once, here, so it never pays for the generality.
  const prefix = rootPrefixOf(bytes);
  if (prefix !== '') bytes = stripPrefix(bytes, prefix);
  const view = viewOf(bytes);
  const data = view.data;
  const lim = ctx.limits ?? DEFAULT_LIMITS;
  const maxRows = lim.maxRows;
  const maxCols = lim.maxColumns;
  const scan: Scan = { maxChars: lim.maxCellChars, sst: undefined, xf: undefined, v: memo(), vEnd: memo(), is: memo(), isEnd: memo(), cellEnd: memo(), cell: memo(), row: memo(), rowEnd: memo() };
  const sd = view.indexOfSeq(SHEET_DATA, 0);
  if (sd < 0) return;
  let sdGt = sd;
  while (sdGt < data.length && data[sdGt] !== GT) sdGt++;
  if (data[sdGt - 1] === SLASH) return;
  const sdEndAt = view.lastIndexOfSeq(SHEET_DATA_END); // from the tail: a forward search reads the whole part
  const limit = sdEndAt < sdGt ? data.length : sdEndAt;

  let emitted = 0;
  let rowNum = 0;
  let p = view.indexOfSeq(ROW, sdGt);
  while (p >= 0 && p < limit) {
    if (!isNameEnd(data[p + 4])) {
      p = find(view, ROW, p + 4, scan.row);
      continue;
    }
    const gt = scanTag(data, p + 4);
    rowNum = tagHasR ? tagNum : rowNum + 1;
    // The gap below is at most maxRows long, so a row number is never taken on trust to size a loop or an array.
    if (rowNum > maxRows) throw new XlsxLimitError('maxRows', `row number ${tagHasR ? rowNum : 'sequence'} passes ${maxRows}`);
    if (gaps !== undefined && emitted < rowNum - 1) gaps.add(rowNum - 1 - emitted);
    if (emitted < rowNum - 1) {
      yield rowNum - 1 - emitted;
      emitted = rowNum - 1;
    }
    if (emitted >= maxRows) throw new XlsxLimitError('maxRows', `sheet has more than ${maxRows} rows`);
    emitted++;

    const out: CellValue[] = [];
    if (data[gt - 1] === SLASH) {
      yield out;
      p = find(view, ROW, gt, scan.row);
      continue;
    }
    let col = -1;
    let at = gt + 1;
    for (;;) {
      let c = -1;
      for (; at < limit; at++) {
        const ch = data[at]!;
        if (ch !== 32 && ch !== 10 && ch !== 13 && ch !== 9) break;
      }
      if (data[at] === LT && data[at + 1] === 99 && isNameEnd(data[at + 2]) && at < limit) {
        c = at;
      } else if (matches(data, ROW_END, at)) {
        at += 6;
        break;
      } else {
        // Something other than a cell or </row> (extLst, an unclosed last row): fall back to searching this one row.
        const rowEnd = find(view, ROW_END, at, scan.rowEnd);
        const end = rowEnd < 0 || rowEnd > limit ? limit : rowEnd;
        c = find(view, CELL, at, scan.cell);
        while (c >= 0 && c < end && !isNameEnd(data[c + 2])) c = find(view, CELL, c + 2, scan.cell);
        if (c < 0 || c >= end) {
          at = end;
          break;
        }
      }
      const cgt = scanTag(data, c + 2);
      col = tagHasR ? tagCol : col + 1;
      if (col >= maxCols) throw new XlsxLimitError('maxColumns', `column ${col + 1} passes ${maxCols}`);
      let v: CellValue;
      if (data[cgt - 1] === SLASH) {
        at = cgt + 1;
      } else {
        const cEnd = find(view, CELL_END, cgt, scan.cellEnd);
        v = cellValue(view, tagType, tagStyle, cgt, cEnd < 0 ? limit : cEnd, ctx, scan);
        at = cEnd < 0 ? limit : cEnd + 4;
      }
      if (v !== undefined) {
        while (out.length < col) out.push(undefined);
        out[col] = v;
      }
    }
    yield out;
    p = find(view, ROW, at, scan.row);
  }
}
