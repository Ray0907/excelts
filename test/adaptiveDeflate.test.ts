import { describe, expect, it } from 'vitest';
import { constants } from 'node:zlib';
import { adaptiveDeflateOptions, defaultCompressor, nativeCompressor } from '../src/zip.js';
import { readXlsx, writeXlsx } from '../src/index.js';

const MiB = 1 << 20;
const lcg = (seed: number) => () => (seed = (Math.imul(seed, 1103515245) + 12345) >>> 0) / 4294967296;

function randomText(bytes: number): Uint8Array {
  const r = lcg(7);
  const u = new Uint8Array(bytes);
  for (let i = 0; i < bytes; i++) u[i] = 33 + Math.floor(r() * 60);
  return u;
}

describe('adaptiveDeflateOptions', () => {
  it('keeps the requested level for small parts and for compressible data', () => {
    expect(adaptiveDeflateOptions(randomText(7 * MiB), 4)).toEqual({ level: 4 });
    const xml = new TextEncoder().encode('<c r="A1" t="s"><v>12345</v></c>'.repeat(400000));
    expect(xml.length).toBeGreaterThan(8 * MiB);
    expect(adaptiveDeflateOptions(xml, 4)).toEqual({ level: 4 });
  });
  it('uses entropy-only coding for near-random text, where matching finds nothing', () => {
    expect(adaptiveDeflateOptions(randomText(9 * MiB), 4)).toEqual({ level: 1, strategy: constants.Z_HUFFMAN_ONLY });
  });
  it('uses the cheapest matching level for text that is matchable but compresses poorly', () => {
    const r = lcg(3);
    const words = Array.from({ length: 5000 }, () => Array.from({ length: 3 + Math.floor(r() * 8) }, () => String.fromCharCode(97 + Math.floor(r() * 26))).join(''));
    let s = '';
    while (s.length < 9 * MiB) s += words[Math.floor(r() * words.length)] + ' ';
    const o = adaptiveDeflateOptions(new TextEncoder().encode(s), 4);
    expect(o.level).toBeLessThanOrEqual(4);
  });
});

describe.runIf(nativeCompressor !== undefined)('writeXlsx with a large, poorly compressible string part', () => {
  it('round-trips every string', () => {
    const r = lcg(11);
    const rows: string[][] = [];
    for (let i = 0; i < 320; i++) rows.push([Array.from({ length: 30000 }, () => String.fromCharCode(97 + Math.floor(r() * 26)) + String.fromCharCode(48 + Math.floor(r() * 10))).join('').slice(0, 30000)]);
    const bytes = writeXlsx({ name: 'S', columns: [{ header: 'h' }], rows }, defaultCompressor);
    const wb = readXlsx(bytes);
    const got = [...wb.sheets[0]!.rows()].slice(1).map((x) => x[0]);
    expect(got).toEqual(rows.map((x) => x[0]));
  });
});
