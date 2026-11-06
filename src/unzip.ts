import { defaultCompressor, type Compressor } from './zip.js';
import { DEFAULT_LIMITS, isOutputOverflow, XlsxLimitError, type XlsxLimits } from './limits.js';

export interface ZipReader {
  readonly names: readonly string[];
  /** Inflates the entry on each call; undefined when the name is absent. */
  get(name: string): Uint8Array | undefined;
}

interface Entry {
  method: number;
  compressedSize: number;
  uncompressedSize: number;
  localOffset: number;
}

const SIG_EOCD = 0x06054b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_LOCAL = 0x04034b50;
const SIG_ZIP64_LOCATOR = 0x07064b50;
const SIG_ZIP64_EOCD = 0x06064b50;
const MAX_COMMENT = 0xffff;
const EOCD_SIZE = 22;
const ZIP64_LOCATOR_SIZE = 20;
const ZIP64_EOCD_SIZE = 56;
const FIELD_ZIP64 = 0x0001;
const decoder = new TextDecoder();

function corrupt(why: string): Error {
  return new Error(`invalid zip: ${why}`);
}

/** A 64-bit little-endian field as a safe integer; anything larger cannot index a buffer and is corrupt. */
function u64(view: DataView, at: number, what: string): number {
  const v = view.getBigUint64(at, true);
  if (v > BigInt(Number.MAX_SAFE_INTEGER)) throw corrupt(`${what} out of range`);
  return Number(v);
}

/**
 * Entry sizes come from the central directory only: exceljs streaming and openpyxl write data descriptors,
 * leaving the local header sizes zero. The local header is read just for its variable-length tail.
 * Sizes are used to refuse early and to cap, never to size an allocation: the inflater stops at the caps whatever the directory claims.
 */
export function unzip(bytes: Uint8Array, compressor: Compressor = defaultCompressor, limits: Readonly<XlsxLimits> = DEFAULT_LIMITS): ZipReader {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (bytes.length < EOCD_SIZE) throw corrupt('too short');
  let eocd = -1;
  for (let i = bytes.length - EOCD_SIZE, stop = Math.max(0, i - MAX_COMMENT); i >= stop; i--) {
    if (view.getUint32(i, true) === SIG_EOCD) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw corrupt('end of central directory not found (truncated?)');

  let count = view.getUint16(eocd + 10, true);
  let cdSize = view.getUint32(eocd + 12, true);
  let cdOffset = view.getUint32(eocd + 16, true);
  const saturated = count === 0xffff || cdSize === 0xffffffff || cdOffset === 0xffffffff;
  if (eocd >= ZIP64_LOCATOR_SIZE && view.getUint32(eocd - ZIP64_LOCATOR_SIZE, true) === SIG_ZIP64_LOCATOR) {
    if (saturated) {
      const at = u64(view, eocd - ZIP64_LOCATOR_SIZE + 8, 'zip64 locator offset');
      if (at + ZIP64_EOCD_SIZE > eocd || view.getUint32(at, true) !== SIG_ZIP64_EOCD) throw corrupt('zip64 end of central directory record not found');
      count = u64(view, at + 32, 'zip64 entry count');
      cdSize = u64(view, at + 40, 'zip64 central directory size');
      cdOffset = u64(view, at + 48, 'zip64 central directory offset');
    }
  } else if (saturated) {
    throw new RangeError('zip64 fields present but the zip64 end of central directory locator is missing');
  }
  if (count > limits.maxEntries) throw new XlsxLimitError('maxEntries', `zip has ${count} entries`);
  if (cdOffset + cdSize > eocd) throw corrupt('central directory out of bounds');

  const entries = new Map<string, Entry>();
  let p = cdOffset;
  for (let n = 0; n < count; n++) {
    if (p + 46 > eocd || view.getUint32(p, true) !== SIG_CENTRAL) throw corrupt('bad central directory entry');
    const method = view.getUint16(p + 10, true);
    let compressedSize = view.getUint32(p + 20, true);
    let uncompressedSize = view.getUint32(p + 24, true);
    const nameLen = view.getUint16(p + 28, true);
    const extraLen = view.getUint16(p + 30, true);
    const commentLen = view.getUint16(p + 32, true);
    let localOffset = view.getUint32(p + 42, true);
    if (p + 46 + nameLen + extraLen > eocd) throw corrupt('central directory entry out of bounds');
    if (compressedSize === 0xffffffff || uncompressedSize === 0xffffffff || localOffset === 0xffffffff) {
      // The zip64 extra field carries, in this order, only the values whose 32-bit slot is saturated.
      let q = p + 46 + nameLen;
      const extraEnd = q + extraLen;
      let found = false;
      while (q + 4 <= extraEnd) {
        const id = view.getUint16(q, true);
        const size = view.getUint16(q + 2, true);
        const body = q + 4;
        if (body + size > extraEnd) break;
        if (id === FIELD_ZIP64) {
          let o = body;
          const take = (what: string): number => {
            if (o + 8 > body + size) throw corrupt(`zip64 extra field too short for ${what}`);
            const v = u64(view, o, what);
            o += 8;
            return v;
          };
          if (uncompressedSize === 0xffffffff) uncompressedSize = take('uncompressed size');
          if (compressedSize === 0xffffffff) compressedSize = take('compressed size');
          if (localOffset === 0xffffffff) localOffset = take('local header offset');
          found = true;
          break;
        }
        q = body + size;
      }
      if (!found) throw corrupt('zip64 entry without a zip64 extra field');
    }
    if (p + 46 + nameLen > eocd) throw corrupt('central directory entry out of bounds');
    entries.set(decoder.decode(bytes.subarray(p + 46, p + 46 + nameLen)), { method, compressedSize, uncompressedSize, localOffset });
    p += 46 + nameLen + extraLen + commentLen;
  }

  // Largest inflated size seen per entry; the sum of these is what counts toward maxTotalBytes, so reading a sheet again costs nothing extra.
  const charged = new Map<string, number>();
  let totalCharged = 0;

  return {
    names: [...entries.keys()],
    get(name) {
      const e = entries.get(name);
      if (!e) return undefined;
      const l = e.localOffset;
      if (l + 30 > bytes.length || view.getUint32(l, true) !== SIG_LOCAL) throw corrupt(`bad local header for ${name}`);
      const start = l + 30 + view.getUint16(l + 26, true) + view.getUint16(l + 28, true);
      if (start + e.compressedSize > cdOffset) throw corrupt(`data of ${name} out of bounds`);
      const data = bytes.subarray(start, start + e.compressedSize);
      // CRC-32 is not verified in v0; its cost is to be measured before deciding.
      if (e.method !== 0 && e.method !== 8) throw new Error(`unsupported zip compression method ${e.method} for ${name}`);
      if (e.uncompressedSize > limits.maxEntryBytes) throw new XlsxLimitError('maxEntryBytes', `${name} declares ${e.uncompressedSize} bytes`);

      const before = charged.get(name) ?? 0;
      const room = limits.maxTotalBytes - (totalCharged - before);
      // Ratio cap: up to ratioMinBytes any ratio is fine, beyond it output may be at most maxRatio times the compressed size.
      const ratioCap = Math.max(limits.ratioMinBytes, data.length * limits.maxRatio);
      const cap = Math.min(limits.maxEntryBytes, room, ratioCap);
      let out: Uint8Array;
      if (e.method === 0) {
        out = data;
        if (out.length > cap) throw overflow(name, cap, limits, room);
      } else {
        try {
          out = compressor.inflateRaw(data, cap);
        } catch (err) {
          if (isOutputOverflow(err)) throw overflow(name, cap, limits, room);
          throw err;
        }
      }
      if (out.length > before) {
        totalCharged += out.length - before;
        charged.set(name, out.length);
      }
      return out;
    },
  };
}

function overflow(name: string, cap: number, limits: Readonly<XlsxLimits>, room: number): XlsxLimitError {
  if (cap === limits.maxEntryBytes) return new XlsxLimitError('maxEntryBytes', `${name} inflates past ${cap} bytes`);
  if (cap === room) return new XlsxLimitError('maxTotalBytes', `inflating ${name} would pass ${limits.maxTotalBytes} bytes in total`);
  return new XlsxLimitError('maxRatio', `${name} inflates past ${cap} bytes, more than ${limits.maxRatio}:1`);
}
