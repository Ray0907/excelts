import { describe, expect, it } from 'vitest';
import { strToU8 } from 'fflate';
import { readXlsx, writeXlsx } from '../src/index.js';
import { unzip } from '../src/unzip.js';
import { fallbackCompressor } from '../src/zip.js';

interface Item {
  name: string;
  data: Uint8Array;
}

const u16 = (v: DataView, o: number, n: number) => v.setUint16(o, n, true);
const u32 = (v: DataView, o: number, n: number) => v.setUint32(o, n, true);
const u64 = (v: DataView, o: number, n: number) => v.setBigUint64(o, BigInt(n), true);

/** A zip whose sizes, offsets and counts all live in zip64 fields: 32-bit slots hold the 0xFFFFFFFF markers. */
function zip64(items: Item[], { saturateCount = true, junkExtraFirst = true } = {}): Uint8Array {
  const chunks: Uint8Array[] = [];
  let offset = 0;
  const push = (b: Uint8Array) => {
    chunks.push(b);
    offset += b.length;
  };
  const central: Uint8Array[] = [];
  for (const { name, data } of items) {
    const nameBytes = strToU8(name);
    const raw = deflateRawOf(data);
    const local = new Uint8Array(30 + nameBytes.length + raw.length);
    const lv = new DataView(local.buffer);
    u32(lv, 0, 0x04034b50);
    u16(lv, 4, 45);
    u16(lv, 8, 8);
    u32(lv, 18, 0xffffffff);
    u32(lv, 22, 0xffffffff);
    u16(lv, 26, nameBytes.length);
    local.set(nameBytes, 30);
    local.set(raw, 30 + nameBytes.length);
    const localOffset = offset;
    push(local);

    // extra: an unrelated field first, then 0x0001 with usize, csize, offset in that order
    const junk = junkExtraFirst ? [0x99, 0x99, 2, 0, 7, 7] : [];
    const z = new Uint8Array(4 + 24);
    const zv = new DataView(z.buffer);
    u16(zv, 0, 1);
    u16(zv, 2, 24);
    u64(zv, 4, data.length);
    u64(zv, 12, raw.length);
    u64(zv, 20, localOffset);
    const extra = new Uint8Array(junk.length + z.length);
    extra.set(junk);
    extra.set(z, junk.length);
    const c = new Uint8Array(46 + nameBytes.length + extra.length);
    const cv = new DataView(c.buffer);
    u32(cv, 0, 0x02014b50);
    u16(cv, 4, 45);
    u16(cv, 6, 45);
    u16(cv, 10, 8);
    u32(cv, 20, 0xffffffff);
    u32(cv, 24, 0xffffffff);
    u16(cv, 28, nameBytes.length);
    u16(cv, 30, extra.length);
    u32(cv, 42, 0xffffffff);
    c.set(nameBytes, 46);
    c.set(extra, 46 + nameBytes.length);
    central.push(c);
  }
  const cdOffset = offset;
  for (const c of central) push(c);
  const cdSize = offset - cdOffset;
  const eocd64Offset = offset;
  const e64 = new Uint8Array(56);
  const ev = new DataView(e64.buffer);
  u32(ev, 0, 0x06064b50);
  u64(ev, 4, 44);
  u16(ev, 12, 45);
  u16(ev, 14, 45);
  u64(ev, 24, items.length);
  u64(ev, 32, items.length);
  u64(ev, 40, cdSize);
  u64(ev, 48, cdOffset);
  push(e64);
  const loc = new Uint8Array(20);
  const lv = new DataView(loc.buffer);
  u32(lv, 0, 0x07064b50);
  u64(lv, 8, eocd64Offset);
  u32(lv, 16, 1);
  push(loc);
  const end = new Uint8Array(22);
  const dv = new DataView(end.buffer);
  u32(dv, 0, 0x06054b50);
  u16(dv, 8, saturateCount ? 0xffff : items.length);
  u16(dv, 10, saturateCount ? 0xffff : items.length);
  u32(dv, 12, 0xffffffff);
  u32(dv, 16, 0xffffffff);
  push(end);
  const out = new Uint8Array(offset);
  let at = 0;
  for (const ch of chunks) {
    out.set(ch, at);
    at += ch.length;
  }
  return out;
}

function deflateRawOf(data: Uint8Array): Uint8Array {
  return fallbackCompressor.deflateRaw(data, 1);
}

/** The parts of an ordinary xlsx, re-wrapped as a zip64 container. */
function asZip64(xlsx: Uint8Array, opts?: Parameters<typeof zip64>[1]): Uint8Array {
  const z = unzip(xlsx);
  return zip64(
    z.names.map((name) => ({ name, data: z.get(name)! })),
    opts,
  );
}

describe('zip64 reading', () => {
  const source = writeXlsx({ name: 'Z', columns: [{ header: 'a' }, { header: 'b' }], rows: [['x', 1], ['y', 2.5], [null, 3]] });
  const want = [...readXlsx(source).sheets[0]!.rows()];

  it('reads sizes, offsets and entry count from the zip64 records', () => {
    const big = asZip64(source);
    expect(unzip(big).names).toEqual(unzip(source).names);
    expect([...readXlsx(big).sheets[0]!.rows()]).toEqual(want);
  });

  it('reads it with the fallback inflater and without the unrelated extra field', () => {
    expect([...readXlsx(asZip64(source, { junkExtraFirst: false }), {}, fallbackCompressor).sheets[0]!.rows()]).toEqual(want);
  });

  it('reads a zip64 central directory whose classic count is not saturated', () => {
    expect([...readXlsx(asZip64(source, { saturateCount: false })).sheets[0]!.rows()]).toEqual(want);
  });

  it('rejects a locator that points outside the file', () => {
    const big = asZip64(source).slice();
    const dv = new DataView(big.buffer);
    dv.setBigUint64(big.length - 22 - 20 + 8, 1n << 40n, true);
    expect(() => readXlsx(big)).toThrow(/zip64/);
  });

  it('rejects a zip64 record with a corrupt signature', () => {
    const big = asZip64(source).slice();
    const dv = new DataView(big.buffer);
    const at = Number(dv.getBigUint64(big.length - 22 - 20 + 8, true));
    dv.setUint32(at, 0, true);
    expect(() => readXlsx(big)).toThrow(/zip64/);
  });

  it('rejects zip64 sizes beyond 2^53', () => {
    const big = asZip64(source).slice();
    const dv = new DataView(big.buffer);
    const at = Number(dv.getBigUint64(big.length - 22 - 20 + 8, true));
    dv.setBigUint64(at + 48, 1n << 60n, true);
    expect(() => readXlsx(big)).toThrow(/zip64|central directory/);
  });

  it('still reports saturated classic fields without a locator as zip64', () => {
    const b = source.slice();
    new DataView(b.buffer).setUint32(b.length - 22 + 16, 0xffffffff, true);
    expect(() => readXlsx(b)).toThrow(RangeError);
  });
});
