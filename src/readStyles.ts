import { dialectOfText } from './dialect.js';
import { attr, decodeEntities, startTags } from './xmlRead.js';
import { DEFAULT_LIMITS, XlsxLimitError, type XlsxLimits } from './limits.js';

/** What a cell's number format makes of its numeric value. */
export const XF = { NONE: 0, DATE: 1, TIME: 2 } as const;

const BUILTIN_KIND = new Uint8Array(59);
for (const id of [14, 15, 16, 17, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 50, 51, 52, 53, 54, 55, 56, 57, 58]) BUILTIN_KIND[id] = XF.DATE;
// h:mm AM/PM, h:mm:ss AM/PM, h:mm, h:mm:ss, mm:ss, [h]:mm:ss, mmss.0
for (const id of [18, 19, 20, 21, 45, 46, 47]) BUILTIN_KIND[id] = XF.TIME;

const QUOTED_ESCAPED_FILL = /"[^"]*"|\\.|[_*]./g;
// [h] [mm] [ss] are elapsed-time fields; every other bracket is a color, condition or locale.
const ELAPSED = /\[(h+)\]|\[(m+)\]|\[(s+)\]/g;

/** Removes every `[...]` group. An unterminated `[` and the rest stay, as a lazy regex would leave them, without rescanning from every opener. */
function stripBrackets(s: string): string {
  let open = s.indexOf('[');
  if (open < 0) return s;
  let out = '';
  let pos = 0;
  while (open >= 0) {
    const close = s.indexOf(']', open + 1);
    if (close < 0) break;
    out += s.slice(pos, open);
    pos = close + 1;
    open = s.indexOf('[', pos);
  }
  return out + s.slice(pos);
}
const AM_PM = /am\/pm|a\/p/g;

/** Classifies a number format code by the fields it displays: any year or day makes a date, a month beside time fields stays a date, time fields alone make a time or duration. */
export function classifyFormat(code: string): number {
  if (code === 'General') return XF.NONE;
  let s = code.replace(QUOTED_ESCAPED_FILL, '').toLowerCase();
  // \u0001 marks minutes that are known to be minutes, so the month/minute rule below only judges the plain `m`.
  s = s.replace(ELAPSED, (_, h, m) => (h ? 'h' : m ? '\u0001' : 's'));
  s = stripBrackets(s);
  let time = false;
  if (s.includes('am/pm') || s.includes('a/p')) {
    time = true;
    s = s.replace(AM_PM, '');
  }
  let month = false;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c === 121 || c === 100) return XF.DATE; // y d
    if (c === 104 || c === 115 || c === 1) time = true; // h s, and the elapsed-minutes marker
    else if (c === 109) {
      let j = i + 1;
      while (s.charCodeAt(j) === 109) j++;
      if (j - i >= 3 || !isMinute(s, i, j)) month = true;
      else time = true;
      i = j - 1;
    }
  }
  return month ? XF.DATE : time ? XF.TIME : XF.NONE;
}

/** ISO rule: `m` right after `h` or right before `s` (allowing `:`, `.` and spaces between) is minutes. */
function isMinute(s: string, from: number, to: number): boolean {
  let i = from - 1;
  while (i >= 0 && (s[i] === ':' || s[i] === ' ' || s[i] === '.')) i--;
  if (i >= 0 && s[i] === 'h') return true;
  i = to;
  while (i < s.length && (s[i] === ':' || s[i] === ' ' || s[i] === '.')) i++;
  return s[i] === 's';
}

/** `kind[s]` for a cell's `s` attribute: one table lookup per cell, all format analysis happens here once. */
export function readXfKinds(xml: string, limits: Readonly<XlsxLimits> = DEFAULT_LIMITS): Uint8Array {
  const d = dialectOfText(xml);
  const custom = new Map<number, number>();
  // CT_Stylesheet orders numFmts before cellXfs; later numFmt tags (inside dxfs) must not redefine ids the xfs use.
  const xfsAt = xml.indexOf(d.cellXfs);
  for (const tag of startTags(xml, d.numFmt, 0, xfsAt < 0 ? xml.length : xfsAt)) {
    if (custom.size >= limits.maxCellXfs) throw new XlsxLimitError('maxCellXfs', `styles define more than ${limits.maxCellXfs} number formats`);
    const id = attr(tag, 'numFmtId');
    const code = attr(tag, 'formatCode');
    if (id !== undefined && code !== undefined) custom.set(Number(id), classifyFormat(decodeEntities(code)));
  }
  const from = xml.indexOf(d.cellXfs);
  if (from < 0) return new Uint8Array(0);
  const to = xml.indexOf(d.cellXfsEnd, from);
  const kinds: number[] = [];
  for (const tag of startTags(xml, d.xf, from, to < 0 ? xml.length : to)) {
    const id = Number(attr(tag, 'numFmtId') ?? 0);
    if (kinds.length >= limits.maxCellXfs) throw new XlsxLimitError('maxCellXfs', `styles define more than ${limits.maxCellXfs} cell formats`);
    kinds.push(custom.get(id) ?? BUILTIN_KIND[id] ?? XF.NONE);
  }
  return Uint8Array.from(kinds);
}
