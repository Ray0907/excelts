import { encodeAddress, decodeEx } from './colCache.js';
import type { Row } from './row.js';
import { MergeSlave, storeValue, type Raw } from './values.js';
import type { Worksheet } from './worksheet.js';

function copyRow(src: Row, dst: Row): void {
  dst._v = src._v.slice();
  dst._s = src._s ? src._s.slice() : null;
  dst._xf = src._xf;
  dst._n = src._n ? new Map(src._n) : null;
  dst.height = src.height;
}

export function spliceRows(ws: Worksheet, start: number, count: number, inserts: Raw[]): void {
  const nKeep = start + count;
  const nInserts = inserts.length;
  const nExpand = nInserts - count;
  const nEnd = ws._rows.length;
  if (nExpand < 0) {
    if (start === nEnd) ws._rows[nEnd - 1] = undefined;
    for (let i = nKeep; i <= nEnd; i++) {
      const src = ws._rows[i - 1];
      if (src) {
        copyRow(src, ws.getRow(i + nExpand));
        ws._rows[i - 1] = undefined;
      } else ws._rows[i + nExpand - 1] = undefined;
    }
  } else if (nExpand > 0) {
    for (let i = nEnd; i >= nKeep; i--) {
      const src = ws._rows[i - 1];
      if (src) {
        const dst = ws.getRow(i + nExpand);
        copyRow(src, dst);
        // merged cells move with their master
        dst._v = dst._v.map((v) => (v instanceof MergeSlave ? new MergeSlave(v.masterRow + nInserts, v.masterCol) : v));
      } else ws._rows[i + nExpand - 1] = undefined;
    }
    const moved = new Map<string, import('./range.js').Range>();
    for (const [k, r] of ws._merges) {
      if (r.top >= nKeep) {
        r.top += nInserts;
        r.bottom += nInserts;
        moved.set(encodeAddress(r.top, r.left), r);
      } else moved.set(k, r);
    }
    ws._merges.clear();
    for (const [k, r] of moved) ws._merges.set(k, r);
  }
  for (let i = 0; i < nInserts; i++) {
    const dst = ws.getRow(start + i);
    dst._xf = 0;
    dst._n = null;
    dst.values = inserts[i] as never;
  }
  ws.workbook?._names?.spliceRows(ws.name, start, count, nInserts);
}

export function insertRow(ws: Worksheet, pos: number, value: Raw, style: string) {
  spliceRows(ws, pos, 0, [value]);
  ws._setStyleOption(pos, style);
  return ws.getRow(pos);
}

export function insertRows(ws: Worksheet, pos: number, values: Raw[], style: string) {
  spliceRows(ws, pos, 0, values);
  if (style !== 'n') {
    for (let i = 0; i < values.length; i++) {
      if (style[0] === 'o' && ws.findRow(values.length + pos + i) !== undefined) ws._copyStyleRow(values.length + pos + i, pos + i, style[1] === '+');
      else if (style[0] === 'i' && ws.findRow(pos - 1) !== undefined) ws._copyStyleRow(pos - 1, pos + i, style[1] === '+');
    }
  }
  return ws.getRows(pos, values.length);
}

export function duplicateRow(ws: Worksheet, rowNum: number, count: number, insert: boolean): void {
  const src = ws._rows[rowNum - 1]!;
  const inserts = new Array<Raw>(count).fill(src.values);
  spliceRows(ws, rowNum + 1, insert ? 0 : count, inserts);
  for (let i = 0; i < count; i++) {
    const dst = ws._rows[rowNum + i]!;
    dst._xf = src._xf;
    dst.height = src.height;
    src.eachCell({ includeEmpty: true }, (cell, col) => {
      dst.getCell(col)._xf = cell._xf;
    });
  }
}

export function rowSplice(row: Row, start: number, count: number, inserts: Raw[]): void {
  const j = start - 1;
  const v = row._v;
  while (v.length < j) v.push(undefined);
  const stored = inserts.map((x) => storeValue(x));
  v.splice(j, count, ...stored);
  if (row._s) {
    while (row._s.length < j) row._s.push(undefined);
    row._s.splice(j, count, ...stored.map(() => 0));
  } else if (inserts.length) {
    row._s = [];
    for (let k = 0; k < inserts.length; k++) row._s[j + k] = 0;
  }
  while (v.length && v[v.length - 1] === undefined) v.pop();
  if (row._n) {
    const moved = new Map<number, import('./notes.js').Note>();
    for (const [c, note] of row._n) {
      if (c < start) moved.set(c, note);
      else if (c >= start + count) moved.set(c - count + inserts.length, note);
    }
    row._n = moved.size ? moved : null;
  }
  row._cells = null;
}

export function fillFormula(ws: Worksheet, range: string, formula: string, results: unknown, shareType = 'shared'): void {
  const d = decodeEx(range);
  const { top, left, bottom, right } = d as Required<Pick<typeof d, 'top' | 'left' | 'bottom' | 'right'>>;
  const width = right - left + 1;
  const master = encodeAddress(top, left);
  const shared = shareType === 'shared';
  let get: (r: number, c: number) => unknown = () => undefined;
  if (typeof results === 'function') get = results as typeof get;
  else if (Array.isArray(results)) {
    get = Array.isArray(results[0]) ? (r, c) => (results as unknown[][])[r - top]![c - left] : (r, c) => results[(r - top) * width + (c - left)];
  }
  let first = true;
  for (let r = top; r <= bottom; r++) {
    for (let c = left; c <= right; c++) {
      const cell = ws.getCell(r, c);
      if (first) {
        cell.value = { shareType, formula, ref: range, result: get(r, c) };
        first = false;
      } else cell.value = shared ? { sharedFormula: master, result: get(r, c) } : get(r, c);
    }
  }
}

export function spliceColumns(ws: Worksheet, start: number, count: number, inserts: Raw[][]): void {
  const rows = ws._rows;
  const nRows = rows.length;
  if (inserts.length > 0) {
    for (let i = 0; i < nRows; i++) {
      const args: Raw[] = inserts.map((ins) => ins[i] || null);
      rowSplice(ws.getRow(i + 1), start, count, args);
    }
  } else {
    for (const r of rows) if (r) rowSplice(r, start, count, []);
  }
  const nExpand = inserts.length - count;
  const nKeep = start + count;
  const nEnd = ws._columns?.length ?? 0;
  if (nExpand < 0) {
    for (let i = start + inserts.length; i <= nEnd; i++) ws.getColumn(i).defn = ws.getColumn(i - nExpand).defn;
  } else if (nExpand > 0) {
    for (let i = nEnd; i >= nKeep; i--) ws.getColumn(i + nExpand).defn = ws.getColumn(i).defn;
  }
  for (let i = start; i < start + inserts.length; i++) ws.getColumn(i).defn = null;
  ws.workbook?._names?.spliceColumns(ws.name, start, count, inserts.length);
}
