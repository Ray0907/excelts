import { decodeAddress, encodeAddress, l2n, n2l } from './colCache.js';
import { ValueType } from './enums.js';

/** Stored in a row's value array where a merged cell is not the master. */
export class MergeSlave {
  constructor(
    readonly masterRow: number,
    readonly masterCol: number,
  ) {}
}

export type Raw = unknown;

/** Set once any shared formula is stored, so writers only pre-scan for them when it can matter. */
export const flags = { shared: false };

/** The ValueType of a raw stored value, decided on demand and never stored (exceljs `Value.getType`). */
export function typeOf(v: Raw): number {
  if (v === null || v === undefined) return ValueType.Null;
  switch (typeof v) {
    case 'string': return ValueType.String;
    case 'number': return ValueType.Number;
    case 'boolean': return ValueType.Boolean;
  }
  if (v instanceof Date) return ValueType.Date;
  if (v instanceof MergeSlave) return ValueType.Merge;
  if (v instanceof String) return ValueType.String;
  const o = v as Record<string, unknown>;
  if (o.text && o.hyperlink) return ValueType.Hyperlink;
  if (o.formula || o.sharedFormula) return ValueType.Formula;
  if (o.richText) return ValueType.RichText;
  if (o.sharedString) return ValueType.SharedString;
  if (o.error) return ValueType.Error;
  return ValueType.JSON;
}

/** `cell.type` hides the JSON detail: it is a String to callers. */
export function publicType(v: Raw): number {
  const t = typeOf(v);
  return t === ValueType.JSON ? ValueType.String : t;
}

/** The stored form of an assigned value: formulas and hyperlinks are copied so later edits to the caller's object do not leak in. */
export function storeValue(v: unknown): Raw {
  if (v === undefined) return null;
  if (v === null || typeof v !== 'object' || v instanceof Date) return v;
  if (v instanceof String) return v.toString();
  const t = typeOf(v);
  const o = v as Record<string, unknown>;
  if (t === ValueType.Formula) {
    if (o.sharedFormula || o.shareType === 'shared') flags.shared = true;
    const c: Record<string, unknown> = {};
    for (const k of ['formula', 'result', 'ref', 'shareType', 'sharedFormula']) if (o[k] !== undefined && o[k] !== null && o[k] !== '') c[k] = o[k];
    return c;
  }
  if (t === ValueType.Hyperlink) {
    const c: Record<string, unknown> = { text: o.text, hyperlink: o.hyperlink };
    if (o.tooltip) c.tooltip = o.tooltip;
    return c;
  }
  return v;
}

export function valueText(v: Raw): string {
  switch (typeOf(v)) {
    case ValueType.Null: return '';
    case ValueType.Merge: return '';
    case ValueType.RichText: return (v as { richText: { text: string }[] }).richText.map((t) => t.text).join('');
    case ValueType.Hyperlink: return String((v as { text: string }).text);
    case ValueType.Formula: {
      const r = (v as { result?: unknown }).result;
      return r ? String(r) : '';
    }
    case ValueType.Error: return String((v as { error: string }).error);
    case ValueType.JSON: return JSON.stringify(v);
    default: return String(v);
  }
}

const candidate = /(([a-z_\-0-9]*)!)?([a-z0-9_$]{2,})([(])?/gi;
const cr = /^([$])?([a-z]+)([$])?([1-9][0-9]*)$/i;

/** Shifts the relative references of a shared formula from its master cell to another cell. */
export function slideFormula(formula: string, from: string, to: string): string {
  const o = decodeAddress(from);
  const t = decodeAddress(to);
  return formula.replace(candidate, (whole, sheet: string | undefined, _s, addr: string, paren: string | undefined) => {
    if (paren) return whole;
    const m = cr.exec(addr);
    if (!m) return whole;
    const colStr = m[2]!.toUpperCase();
    if (colStr.length > 3 || (colStr.length === 3 && colStr > 'XFD')) return whole;
    let col = l2n(colStr);
    let row = parseInt(m[4]!, 10);
    if (!m[1]) col += t.col - o.col;
    if (!m[3]) row += t.row - o.row;
    return (sheet || '') + (m[1] || '') + n2l(col) + (m[3] || '') + row;
  });
}

export function addressOf(row: number, col: number): string {
  return encodeAddress(row, col);
}
