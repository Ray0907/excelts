import { deflateSync, inflateSync, Inflate } from 'fflate';
import { OutputOverflow } from './limits.js';

export interface Compressor {
  deflateRaw(data: Uint8Array, level: number): Uint8Array;
  crc32(data: Uint8Array): number;
  /** Throws an error `isOutputOverflow` recognises when the output would pass `maxOutput` bytes; the output is never allocated past it. */
  inflateRaw(data: Uint8Array, maxOutput?: number): Uint8Array;
}

/** Speed first: level 7 is 2x slower for an 8% smaller file (100k rows, 5.7 vs 6.3 MB), so level 4 stays unless a change shrinks output by 50% or more. */
export const NATIVE_LEVEL = 4;
/** fflate's JS deflate costs far more per level, so the fallback stays at the cheapest level. */
export const FALLBACK_LEVEL = 1;

const CRC_TABLE = new Uint32Array(256);
for (let n = 0; n < 256; n++) {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  CRC_TABLE[n] = c;
}

export const fallbackCompressor: Compressor = {
  deflateRaw: (data, level) => deflateSync(data, { level: level as 1 }),
  inflateRaw: (data, maxOutput = Infinity) => {
    // Deflate expands at most 1032:1, so input this small cannot overshoot the cap and the one-shot call is safe and fastest.
    if (data.length * 1032 <= maxOutput) return inflateSync(data);
    // Otherwise feed small slices so the output fflate holds at once stays near the cap.
    const chunks: Uint8Array[] = [];
    let total = 0;
    const inf = new Inflate((chunk) => {
      total += chunk.length;
      if (total > maxOutput) throw new OutputOverflow();
      chunks.push(chunk);
    });
    const SLICE = 16 * 1024;
    for (let i = 0; i < data.length; i += SLICE) inf.push(data.subarray(i, i + SLICE), i + SLICE >= data.length);
    return concat(chunks);
  },
  crc32(data) {
    let c = 0xffffffff;
    for (let i = 0; i < data.length; i++) c = CRC_TABLE[(c ^ data[i]!) & 0xff]! ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  },
};

/** Below this the sampling (two fresh zlib states, ~0.4 ms) would show as about 1% on 4 MB sheets, where level 4 is still only a few ms of work. */
const ADAPT_MIN_BYTES = 8 << 20;
const SAMPLE_BYTES = 32 << 10;
/** Each sample costs a fresh zlib state (~0.1 ms), so few and larger beat many and small. */
const SAMPLE_AT = [0.25, 0.75];
/** Level-1 output above this share of the input means matching finds little (random-looking text such as hashes or base64); typical prose and XML stay well below. */
const POOR_RATIO = 0.45;
/** Entropy-only coding is used when it comes within this factor of level-1 size: it then costs nothing in file size and skips the match search. */
const HUFFMAN_TOLERANCE = 1.03;

export interface DeflateOptions {
  level: number;
  strategy?: number;
}

/**
 * Chooses the deflate settings for one part. Level 4 finds long matches that level 1 misses, but on text where nothing matches it
 * spends 3-4x the time for the same bytes. For large parts, two spread 32 KB samples are deflated at level 1; only a poor ratio (>= 45%) changes
 * the settings, so compressible XML keeps `level` and its output bytes exactly.
 */
export function adaptiveDeflateOptions(data: Uint8Array, level: number): DeflateOptions {
  if (level <= 1 || data.length < ADAPT_MIN_BYTES) return { level };
  const zlib = nodeZlib();
  if (!zlib) return { level };
  let raw = 0;
  let fast = 0;
  const samples: Uint8Array[] = [];
  for (const frac of SAMPLE_AT) {
    const at = Math.floor((data.length - SAMPLE_BYTES) * frac);
    const part = data.subarray(at, at + SAMPLE_BYTES);
    samples.push(part);
    raw += part.length;
    fast += zlib.deflateRawSync(part, { level: 1 }).length;
  }
  if (fast < POOR_RATIO * raw) return { level }; // the common case stops here, after 64 KB of level-1 work
  let entropy = 0;
  for (const part of samples) entropy += zlib.deflateRawSync(part, { level: 1, strategy: zlib.constants.Z_HUFFMAN_ONLY }).length;
  return entropy <= HUFFMAN_TOLERANCE * fast ? { level: 1, strategy: zlib.constants.Z_HUFFMAN_ONLY } : { level: 1 };
}

type Zlib = typeof import('node:zlib');
function nodeZlib(): Zlib | undefined {
  const z = (globalThis as { process?: { getBuiltinModule?: (id: string) => unknown } }).process?.getBuiltinModule?.('node:zlib') as Partial<Zlib> | undefined;
  return typeof z?.deflateRawSync === 'function' && typeof z.crc32 === 'function' && typeof z.inflateRawSync === 'function' ? (z as Zlib) : undefined;
}

function loadNative(): Compressor | undefined {
  const zlib = nodeZlib();
  if (!zlib) return undefined;
  const { deflateRawSync, inflateRawSync, crc32 } = zlib;
  return {
    deflateRaw: (data, level) => deflateRawSync(data, adaptiveDeflateOptions(data, level)),
    inflateRaw: (data, maxOutput) => (maxOutput === undefined ? inflateRawSync(data) : inflateRawSync(data, { maxOutputLength: Math.min(Math.max(1, maxOutput), 0xffffffff) })),
    crc32: (data) => crc32(data),
  };
}

export const nativeCompressor: Compressor | undefined = loadNative();
export const defaultCompressor: Compressor = nativeCompressor ?? fallbackCompressor;
export const levelFor = (c: Compressor): number => (c === fallbackCompressor ? FALLBACK_LEVEL : NATIVE_LEVEL);

export interface ZipEntry {
  name: string;
  data: Uint8Array;
}

const LIMIT_32 = 0xffffffff;
const DOS_TIME = 0;
const DOS_DATE = (0 << 9) | (1 << 5) | 1; // 1980-01-01, fixed so output is reproducible
const UTF8_FLAG = 0x0800;
const encoder = new TextEncoder();

/** Writes a zip with every entry deflated; no zip64, so total size and entry count stay under 4 GiB and 65535. */
export function writeZip(entries: readonly ZipEntry[], compressor: Compressor, level: number): Uint8Array {
  const locals: Uint8Array[] = [];
  const centrals: Uint8Array[] = [];
  let offset = 0;
  let centralSize = 0;

  for (const { name, data } of entries) {
    const nameBytes = encoder.encode(name);
    const packed = compressor.deflateRaw(data, level);
    const crc = compressor.crc32(data);

    const local = new Uint8Array(30 + nameBytes.length + packed.length);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, 20, true);
    lv.setUint16(6, UTF8_FLAG, true);
    lv.setUint16(8, 8, true);
    lv.setUint16(10, DOS_TIME, true);
    lv.setUint16(12, DOS_DATE, true);
    lv.setUint32(14, crc, true);
    lv.setUint32(18, packed.length, true);
    lv.setUint32(22, data.length, true);
    lv.setUint16(26, nameBytes.length, true);
    local.set(nameBytes, 30);
    local.set(packed, 30 + nameBytes.length);

    const central = new Uint8Array(46 + nameBytes.length);
    const cv = new DataView(central.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint16(4, 20, true);
    cv.setUint16(6, 20, true);
    cv.setUint16(8, UTF8_FLAG, true);
    cv.setUint16(10, 8, true);
    cv.setUint16(12, DOS_TIME, true);
    cv.setUint16(14, DOS_DATE, true);
    cv.setUint32(16, crc, true);
    cv.setUint32(20, packed.length, true);
    cv.setUint32(24, data.length, true);
    cv.setUint16(28, nameBytes.length, true);
    cv.setUint32(42, offset, true);
    central.set(nameBytes, 46);

    locals.push(local);
    centrals.push(central);
    offset += local.length;
    centralSize += central.length;
  }

  if (entries.length > 0xffff || offset + centralSize + 22 > LIMIT_32) throw new RangeError('zip exceeds the 4 GiB limit (zip64 is not supported)');

  const end = new Uint8Array(22);
  const ev = new DataView(end.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, entries.length, true);
  ev.setUint16(10, entries.length, true);
  ev.setUint32(12, centralSize, true);
  ev.setUint32(16, offset, true);

  return concat([...locals, ...centrals, end]);
}

export function concat(parts: readonly Uint8Array[]): Uint8Array {
  let total = 0;
  for (const p of parts) total += p.length;
  const out = new Uint8Array(total);
  let off = 0;
  for (const p of parts) {
    out.set(p, off);
    off += p.length;
  }
  return out;
}
