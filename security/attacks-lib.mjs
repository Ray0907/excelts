import { deflateRawSync } from 'node:zlib';
import { NS, XML, skeleton } from './lib/skeleton.mjs';

export const rep = (s, n) => Buffer.from(s.repeat(n));
export const cache = new Map();

/** Prefix, `n` copies of one byte, suffix, deflated once and reused: a 1 GB run of one byte costs about 1 MB. */
export function bigPart(key, prefix, fill, n, suffix, level = 9) {
  if (!cache.has(key)) {
    const p = Buffer.from(prefix);
    const s = Buffer.from(suffix);
    const buf = Buffer.alloc(p.length + n + s.length, fill);
    p.copy(buf, 0);
    s.copy(buf, p.length + n);
    cache.set(key, { comp: deflateRawSync(buf, { level }), usize: buf.length });
  }
  return cache.get(key);
}

export const WS_PRE = `${XML}<worksheet ${NS}><sheetData>`;
export const WS_SUF = '</sheetData></worksheet>';
export const sst = (items, attrs = '') => `${XML}<sst ${NS}${attrs}>${items}</sst>`;
export const inline = (t) => `<c t="inlineStr"><is><t>${t}</t></is></c>`;
export const BASE_BODY = '<row r="1"><c r="A1" t="inlineStr"><is><t>hi</t></is></c><c r="B1"><v>2</v></c></row>';
export const base = () => skeleton({ sheets: [{ name: 'S', body: BASE_BODY }] });
export const mb = (n) => Math.round(n * 1048576);

export function patchEocd(buf, fn) {
  const b = Buffer.from(buf);
  const at = b.lastIndexOf(Buffer.from([0x50, 0x4b, 5, 6]));
  fn(b, at);
  return b;
}
export function patchFirstCd(buf, fn) {
  const b = Buffer.from(buf);
  const at = b.indexOf(Buffer.from([0x50, 0x4b, 1, 2]));
  fn(b, at);
  return b;
}

export const laughs = (() => {
  let dtd = '<!ENTITY lol0 "lollollollollollollollollollol">';
  for (let i = 1; i <= 9; i++) dtd += `<!ENTITY lol${i} "${`&lol${i - 1};`.repeat(10)}">`;
  return `<!DOCTYPE lolz [${dtd}]>`;
})();
export const noExpansion = (r) => (r.maxStr > 100000 ? `entity expanded to a ${r.maxStr}-char string` : undefined);
