import { deflateRawSync } from 'node:zlib';
import { zipBuild } from './lib/zip.mjs';
import { NS, XML, sheetXml, skeleton } from './lib/skeleton.mjs';
import { BASE_BODY, WS_PRE, WS_SUF, base, bigPart, inline, laughs, mb, noExpansion, patchEocd, patchFirstCd, rep, sst } from './attacks-lib.mjs';

export function zipAttacks({ bombMb = 1024 } = {}) {
  const A = [];
  const add = (id, group, note, build, extra = {}) => A.push({ id, group, note, build, ...extra });
  const bomb = mb(bombMb);

  const target = (t, entry) => skeleton({ sheets: [{ name: 'S', body: BASE_BODY }], relsXml: `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${t}"/></Relationships>`, extra: [{ name: entry, data: sheetXml(BASE_BODY) }] });
  add('name-traversal', 'zip-structure', 'rel Target ../../../etc/passwd and an entry with that name', () => target('../../../etc/passwd', '../../../etc/passwd'));
  add('name-absolute', 'zip-structure', 'absolute Target and entry /etc/passwd, C:\\windows\\win.ini', () => target('/etc/passwd', '/etc/passwd'));
  add('name-nul', 'zip-structure', 'entry name containing NUL', () => target('worksheets/s\u0000.xml', 'xl/worksheets/s\u0000.xml'));
  add('name-backslash-case', 'zip-structure', 'Target with backslashes and different case', () => target('Worksheets\\\\SHEET1.XML', 'xl/worksheets/sheet1.xml'));
  add('name-65535-chars', 'zip-structure', 'entry name of 65535 characters', () => skeleton({ sheets: [{ name: 'S', body: BASE_BODY }], extra: [{ name: 'x'.repeat(65535), data: 'a' }] }));
  add('name-collision-65534-entries', 'zip-structure', '65534 entries whose names differ only in case', () => {
    const extra = Array.from({ length: 65_000 }, (_, i) => ({ name: `XL/Worksheets/${i.toString(36)}.XML`, comp: Buffer.alloc(0), usize: 0, pointTo: 0 }));
    return skeleton({ sheets: [{ name: 'S', body: BASE_BODY }], extra });
  });
  add('cd-outside-file', 'zip-structure', 'central directory offset beyond the file', () => patchEocd(base(), (b, at) => b.writeUInt32LE(b.length + 1000, at + 16)));
  add('cd-size-overflow', 'zip-structure', 'central directory size 0xFFFFFFFE', () => patchEocd(base(), (b, at) => b.writeUInt32LE(0xfffffffe, at + 12)));
  add('cd-count-65535', 'zip-structure', 'entry count 65535 over a normal directory', () => patchEocd(base(), (b, at) => b.writeUInt16LE(0xffff, at + 10)));
  add('cd-count-65534-garbage', 'zip-structure', 'entry count 65534 with 4 real entries then garbage', () => patchEocd(base(), (b, at) => b.writeUInt16LE(0xfffe, at + 10)));
  add('cd-count-65534-valid', 'zip-structure', '65534 valid central entries sharing one local record', () => {
    const extra = Array.from({ length: 65_529 }, (_, i) => ({ name: `p/${i}`, comp: Buffer.alloc(0), usize: 0, pointTo: 0 }));
    return skeleton({ sheets: [{ name: 'S', body: BASE_BODY }], extra });
  });
  add('cd-random-garbage', 'zip-structure', 'valid EOCD over 2 MB of random bytes', () => {
    let s = 12345;
    const g = Buffer.alloc(2 * 1048576);
    for (let i = 0; i < g.length; i++) g[i] = (s = (Math.imul(s, 1103515245) + 12345) >>> 0) >>> 24;
    return zipBuild([], { prefix: g, count: 100, cdSize: g.length, cdOffset: 0 });
  });
  add('eocd-fake-in-comment', 'zip-structure', '65535-byte comment full of EOCD signatures', () => zipBuild([{ name: 'x', data: 'x' }], { comment: Buffer.from([0x50, 0x4b, 5, 6]).toString('latin1').repeat(16000) }));
  add('zip64-locator', 'zip-structure', 'zip64 end-of-central-directory locator present', () => zipBuild([{ name: 'xl/workbook.xml', data: '<workbook/>' }], { zip64Locator: true }));
  add('zip64-entry-sizes', 'zip-structure', 'entry sizes 0xFFFFFFFF', () => patchFirstCd(base(), (b, at) => {
    b.writeUInt32LE(0xffffffff, at + 20);
    b.writeUInt32LE(0xffffffff, at + 24);
  }));
  add('method-unknown', 'zip-structure', 'compression methods 9, 12, 14, 99 on the sheet', () => {
    const b = base();
    return patchFirstCd(b, (buf) => {
      let at = 0;
      for (let n = 0; n < 4; n++) {
        at = buf.indexOf(Buffer.from([0x50, 0x4b, 1, 2]), at + 1);
        buf.writeUInt16LE([9, 12, 14, 99][n], at + 10);
      }
    });
  });
  add('local-header-lies', 'zip-structure', 'local header name/extra lengths 65535', () => {
    const b = Buffer.from(base());
    b.writeUInt16LE(0xffff, 26);
    b.writeUInt16LE(0xffff, 28);
    return b;
  });
  add('stored-huge-csize', 'zip-structure', 'method 0 with csize larger than the file', () => zipBuild([{ name: 'xl/workbook.xml', data: '<workbook/>', method: 0, csize: 0x7fffffff }]));
  add('truncated-deflate', 'zip-structure', 'valid directory over a truncated deflate stream', () => {
    const part = deflateRawSync(Buffer.from(sheetXml(BASE_BODY))).subarray(0, 10);
    return skeleton({ sheets: [{ name: 'S', part: { comp: part, usize: 999 } }] });
  });
  add('gzip-bomb-in-stored', 'zip-structure', 'stored (method 0) entry equal to a 1 MB raw deflate of 1 GB, read as XML', () => {
    const part = bigPart('sheet', WS_PRE, 32, bomb, WS_SUF);
    return skeleton({ sheets: [{ name: 'S', part: { comp: part.comp, usize: part.comp.length, entry: { method: 0 } } }] });
  });

  return A;
}
