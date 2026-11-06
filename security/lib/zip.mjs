import { crc32, deflateRawSync, inflateRawSync } from 'node:zlib';

const u32 = (v, n) => v.setUint32(n, 0, true);

/**
 * Hand-rolled zip writer that can lie: every size, offset, name and header field can be overridden per entry,
 * and entries can share one local record, none of which a well-behaved zip library allows.
 * entry: { name, data? | comp?+usize?, method?, crc?, csize?, usize?, flags?, localExtra?, cdExtra?, comment?, pointTo?, localNameLen?, localExtraLen?, versionNeeded? }
 * opts:  { count?, cdSize?, cdOffset?, comment?, zip64Locator?, prefix?, trailer?, skipEocd? }
 */
export function zipBuild(entries, opts = {}) {
  const enc = (s) => (typeof s === 'string' ? Buffer.from(s, 'utf8') : Buffer.from(s));
  const parts = [];
  let pos = 0;
  const push = (b) => {
    parts.push(b);
    pos += b.length;
  };
  if (opts.prefix) push(Buffer.from(opts.prefix));
  const prepared = entries.map((e) => {
    const method = e.method ?? 8;
    let comp = e.comp;
    let usize = e.usize;
    if (!comp) {
      const data = Buffer.from(e.data ?? []);
      usize ??= data.length;
      comp = method === 8 ? deflateRawSync(data, { level: e.level ?? 6 }) : data;
    }
    usize ??= comp.length;
    const crc = e.crc ?? (e.data ? crc32(e.data) : 0);
    return { e, method, comp, usize, crc, name: enc(e.name), offset: -1 };
  });
  for (const p of prepared) {
    if (p.e.pointTo !== undefined) continue;
    p.offset = pos;
    const nameLen = p.e.localNameLen ?? p.name.length;
    const extra = Buffer.from(p.e.localExtra ?? []);
    const h = Buffer.alloc(30);
    const v = new DataView(h.buffer, h.byteOffset, 30);
    v.setUint32(0, 0x04034b50, true);
    v.setUint16(4, p.e.versionNeeded ?? 20, true);
    v.setUint16(6, p.e.flags ?? 0x800, true);
    v.setUint16(8, p.method, true);
    v.setUint32(14, p.crc, true);
    v.setUint32(18, p.e.csize ?? p.comp.length, true);
    v.setUint32(22, p.usize, true);
    v.setUint16(26, nameLen, true);
    v.setUint16(28, p.e.localExtraLen ?? extra.length, true);
    push(h);
    push(p.name);
    push(extra);
    push(Buffer.from(p.comp));
  }
  for (const p of prepared) {
    if (p.e.pointTo === undefined) continue;
    const t = prepared[p.e.pointTo];
    p.offset = t.offset;
    if (p.e.comp === undefined || p.e.comp.length === 0) Object.assign(p, { comp: t.comp, method: t.method, usize: t.usize, crc: t.crc });
  }
  const cdStart = pos;
  for (const p of prepared) {
    const comment = Buffer.from(p.e.comment ?? []);
    const extra = Buffer.from(p.e.cdExtra ?? []);
    const h = Buffer.alloc(46);
    const v = new DataView(h.buffer, h.byteOffset, 46);
    v.setUint32(0, 0x02014b50, true);
    v.setUint16(4, 20, true);
    v.setUint16(6, p.e.versionNeeded ?? 20, true);
    v.setUint16(8, p.e.flags ?? 0x800, true);
    v.setUint16(10, p.method, true);
    v.setUint32(16, p.crc, true);
    v.setUint32(20, p.e.csize ?? p.comp.length, true);
    v.setUint32(24, p.usize, true);
    v.setUint16(28, p.name.length, true);
    v.setUint16(30, extra.length, true);
    v.setUint16(32, comment.length, true);
    v.setUint32(42, p.e.localOffset ?? p.offset, true);
    push(h);
    push(p.name);
    push(extra);
    push(comment);
  }
  const cdSize = pos - cdStart;
  if (opts.zip64Locator) {
    const l = Buffer.alloc(20);
    l.writeUInt32LE(0x07064b50, 0);
    push(l);
  }
  if (!opts.skipEocd) {
    const comment = Buffer.from(opts.comment ?? []);
    const h = Buffer.alloc(22);
    const v = new DataView(h.buffer, h.byteOffset, 22);
    v.setUint32(0, 0x06054b50, true);
    v.setUint16(8, Math.min(entries.length, 0xffff), true);
    v.setUint16(10, opts.count ?? Math.min(entries.length, 0xffff), true);
    v.setUint32(12, opts.cdSize ?? cdSize, true);
    v.setUint32(16, opts.cdOffset ?? cdStart, true);
    v.setUint16(20, comment.length, true);
    push(h);
    push(comment);
  }
  if (opts.trailer) push(Buffer.from(opts.trailer));
  return Buffer.concat(parts);
}

/** Entries of a well-formed zip as { name, method, comp, usize, crc }, or undefined when the central directory is unusable. */
export function zipParse(bytes) {
  try {
    const b = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let eocd = -1;
    for (let i = b.length - 22; i >= Math.max(0, b.length - 22 - 0xffff); i--) {
      if (b.readUInt32LE(i) === 0x06054b50) {
        eocd = i;
        break;
      }
    }
    if (eocd < 0) return undefined;
    const count = b.readUInt16LE(eocd + 10);
    let p = b.readUInt32LE(eocd + 16);
    const out = [];
    for (let n = 0; n < count; n++) {
      if (b.readUInt32LE(p) !== 0x02014b50) return undefined;
      const method = b.readUInt16LE(p + 10);
      const crc = b.readUInt32LE(p + 16);
      const csize = b.readUInt32LE(p + 20);
      const usize = b.readUInt32LE(p + 24);
      const nl = b.readUInt16LE(p + 28);
      const xl = b.readUInt16LE(p + 30);
      const cl = b.readUInt16LE(p + 32);
      const lo = b.readUInt32LE(p + 42);
      const name = b.toString('utf8', p + 46, p + 46 + nl);
      const start = lo + 30 + b.readUInt16LE(lo + 26) + b.readUInt16LE(lo + 28);
      out.push({ name, method, crc, usize, comp: b.subarray(start, start + csize) });
      p += 46 + nl + xl + cl;
    }
    return out;
  } catch {
    return undefined;
  }
}

/** Inflated bytes of a parsed entry, or undefined when it does not inflate or exceeds `max`. */
export function entryData(entry, max = 40e6) {
  try {
    if (entry.method === 0) return entry.comp;
    if (entry.method === 8) return inflateRawSync(entry.comp, { maxOutputLength: max });
  } catch {
    /* corrupt seed part */
  }
  return undefined;
}

/** Byte offsets of every fixed-size field of the end record, the central directory entries and their local headers. */
export function zipFields(bytes) {
  const b = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = -1;
  for (let i = b.length - 22; i >= Math.max(0, b.length - 22 - 0xffff); i--) {
    if (b.readUInt32LE(i) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) return undefined;
  const fields = [];
  const add = (label, at, size) => fields.push({ label, at, size });
  add('eocd.count', eocd + 10, 2);
  add('eocd.cdSize', eocd + 12, 4);
  add('eocd.cdOffset', eocd + 16, 4);
  add('eocd.commentLen', eocd + 20, 2);
  let p = b.readUInt32LE(eocd + 16);
  const count = b.readUInt16LE(eocd + 10);
  for (let n = 0; n < count && p + 46 <= b.length && b.readUInt32LE(p) === 0x02014b50; n++) {
    add('cd.flags', p + 8, 2);
    add('cd.method', p + 10, 2);
    add('cd.csize', p + 20, 4);
    add('cd.usize', p + 24, 4);
    add('cd.nameLen', p + 28, 2);
    add('cd.extraLen', p + 30, 2);
    add('cd.commentLen', p + 32, 2);
    add('cd.localOffset', p + 42, 4);
    const lo = b.readUInt32LE(p + 42);
    if (lo + 30 <= b.length && b.readUInt32LE(lo) === 0x04034b50) {
      add('local.method', lo + 8, 2);
      add('local.csize', lo + 18, 4);
      add('local.usize', lo + 22, 4);
      add('local.nameLen', lo + 26, 2);
      add('local.extraLen', lo + 28, 2);
    }
    p += 46 + b.readUInt16LE(p + 28) + b.readUInt16LE(p + 30) + b.readUInt16LE(p + 32);
  }
  return fields;
}
