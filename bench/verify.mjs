import { readFileSync } from 'node:fs';
import { unzipSync, strFromU8 } from 'fflate';
import { row } from './data.mjs';

const unesc = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
const serial = (d) => d.getTime() / 86400000 + 25569;

// Independent of both writers: counts <row> elements and decodes a few sampled rows.
export function verify(path, n) {
  const files = unzipSync(readFileSync(path));
  const sheetName = Object.keys(files).find((k) => /^xl\/worksheets\/sheet1\.xml$/.test(k));
  if (!sheetName) return { ok: false, why: 'no sheet1.xml' };
  const xml = strFromU8(files[sheetName]);
  const rowCount = (xml.match(/<row[ >]/g) || []).length;
  if (rowCount !== n + 1) return { ok: false, why: `rows ${rowCount} != ${n + 1}`, rowCount };
  const ssFile = files['xl/sharedStrings.xml'];
  const shared = ssFile ? [...strFromU8(ssFile).matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) => unesc([...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((t) => t[1]).join(''))) : [];
  const cellsOf = (r) => {
    const m = xml.match(new RegExp(`<row r="${r}"[^>]*>([\\s\\S]*?)</row>`));
    if (!m) throw new Error(`row ${r} missing`);
    return [...m[1].matchAll(/<c r="([A-Z]+)\d+"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)].map((c) => {
      const t = /t="(\w+)"/.exec(c[2])?.[1];
      const body = c[3] || '';
      if (t === 's') return shared[Number(/<v>([\s\S]*?)<\/v>/.exec(body)[1])];
      if (t === 'inlineStr') return unesc(/<t[^>]*>([\s\S]*?)<\/t>/.exec(body)[1]);
      if (t === 'str') return unesc(/<v>([\s\S]*?)<\/v>/.exec(body)[1]);
      if (t === 'b') return /<v>1<\/v>/.test(body);
      return Number(/<v>([\s\S]*?)<\/v>/.exec(body)[1]);
    });
  };
  const sample = [0, 1, Math.floor(n / 2), n - 1].filter((i) => i < n);
  for (const i of sample) {
    const got = cellsOf(i + 2);
    const want = row(i).map((v) => (v instanceof Date ? serial(v) : v));
    for (let c = 0; c < want.length; c++) {
      const ok = typeof want[c] === 'number' ? Math.abs(got[c] - want[c]) < 1e-6 : got[c] === want[c];
      if (!ok) return { ok: false, why: `row ${i} col ${c}: got ${JSON.stringify(got[c])} want ${JSON.stringify(want[c])}`, rowCount };
    }
  }
  const styled = /numFmt[^>]*formatCode="yyyy-mm-dd hh:mm:ss"/.test(strFromU8(files['xl/styles.xml']));
  if (!styled) return { ok: false, why: 'date numFmt missing in styles.xml', rowCount };
  return { ok: true, rowCount };
}
