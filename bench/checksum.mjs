// Order-sensitive digest of what a reader yields. Every arm pays the same visit cost.
export class Checksum {
  constructor() { this.rows = 0; this.cells = 0; this.numSum = 0; this.strLen = 0; this.strHash = 0; this.trues = 0; this.dateSum = 0; this.dates = 0; this.skipHeader = true; }
  // values: array; `from` is the index of column A (exceljs row.values is 1-based).
  row(values, from = 0) {
    if (this.skipHeader) { this.skipHeader = false; return; }
    this.rows++;
    for (let i = from; i < values.length; i++) {
      const v = values[i];
      if (v == null) continue;
      this.cells++;
      switch (typeof v) {
        case 'number': this.numSum += v; break;
        case 'string': {
          this.strLen += v.length;
          let h = 0;
          for (let k = 0; k < v.length; k++) h = (Math.imul(h, 31) + v.charCodeAt(k)) | 0;
          this.strHash = (this.strHash + h) | 0;
          break;
        }
        case 'boolean': if (v) this.trues++; break;
        default:
          if (v instanceof Date) { this.dates++; this.dateSum += Math.round(v.getTime()); }
          else throw new TypeError(`unexpected cell value ${JSON.stringify(v)}`);
      }
    }
  }
  result() { const { rows, cells, numSum, strLen, strHash, trues, dateSum, dates } = this; return { rows, cells, numSum, strLen, strHash, trues, dateSum, dates }; }
}

export function sameResult(a, b) {
  const bad = [];
  for (const k of Object.keys(b)) {
    const ok = k === 'numSum' ? Math.abs(a[k] - b[k]) <= 1e-9 * Math.abs(b[k]) : a[k] === b[k];
    if (!ok) bad.push(`${k} got ${a[k]} want ${b[k]}`);
  }
  return bad;
}
