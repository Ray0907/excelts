import { entryData, zipBuild, zipParse } from './lib/zip.mjs';
import { int, pick } from './fuzz-rng.mjs';

const BIG = 10 * 1048576;
const randomTag = (text, rng) => {
  const m = [...text.matchAll(/<[A-Za-z][^<>]*>/g)];
  return m.length ? pick(rng, m.slice(0, 20000)) : undefined;
};
const CELL_REFS = ['ZZZZZZ9999999999', 'A0', 'A-1', '', 'a1', 'XFE1', 'ZZZZZZZZ1', 'A4294967296', 'A99999999999999999999999', '1', 'AAAAAAAAAAAAAAAAAAAA1'];
const ROW_RS = ['4294967296', '-1', '0', '1048577', '18446744073709551616', '1e5', '99999999999999999999', '2147483648', '4294967295'];
const STYLE_IDX = ['99999999', '-1', '4294967296', '1e3', 'abc', '2147483648'];
const SST_IDX = ['-1', '4294967296', '1e999', '99999999', '0x10', 'NaN', '', '-0', '1.5', '9999999999999999999999'];
const ENTITIES = ['&#xFFFFFFFF;', '&#x110000;', '&#0;', '&#xD800;', '&#99999999999;', '&bogus;', '&amp', '&#x;', '&#;', '&#xDFFF;', '&#1114111;'];

const XML_OPS = {
  delClose(t, rng, r) {
    const m = [...t.matchAll(/<\/[^>]+>/g)];
    if (!m.length) return t;
    const x = pick(rng, m);
    r.push(`delClose@${x.index}`);
    return t.slice(0, x.index) + t.slice(x.index + x[0].length);
  },
  nest(t, rng, r) {
    const tag = pick(rng, ['is', 'c', 'row', 'si', 'a', 'sst', 'r', 't']);
    const n = pick(rng, [10000, 100000, 300000]);
    const at = t.indexOf('>', int(rng, t.length)) + 1;
    r.push(`nest<${tag}>x${n}@${at}`);
    return t.slice(0, at) + `<${tag}>`.repeat(n) + t.slice(at);
  },
  nestIs(t, rng, r) {
    const at = t.indexOf('<sheetData>') >= 0 ? t.indexOf('<sheetData>') + 11 : int(rng, t.length);
    r.push('nestIs100k');
    return t.slice(0, at) + '<row r="1"><c r="A1" t="inlineStr">' + '<is>'.repeat(100000) + '<t>x</t>' + '</is>'.repeat(100000) + '</c></row>' + t.slice(at);
  },
  hugeAttr(t, rng, r) {
    const m = [...t.matchAll(/(\w+)="([^"]*)"/g)];
    if (!m.length) return t;
    const x = pick(rng, m.slice(0, 20000));
    const size = pick(rng, [1 << 20, BIG]);
    r.push(`hugeAttr ${x[1]}@${x.index} ${size}`);
    return t.slice(0, x.index) + `${x[1]}="${'A'.repeat(size)}"` + t.slice(x.index + x[0].length);
  },
  manyAttrs(t, rng, r) {
    const x = randomTag(t, rng);
    if (!x) return t;
    const n = pick(rng, [50000, 200000]);
    r.push(`manyAttrs x${n}@${x.index}`);
    const at = x.index + x[0].length - (x[0].endsWith('/>') ? 2 : 1);
    let s = '';
    for (let i = 0; i < n; i++) s += ` a${i}="1"`;
    return t.slice(0, at) + s + t.slice(at);
  },
  bigText(t, rng, r) {
    const m = [...t.matchAll(/<(t|v)(\s[^>]*)?>/g)];
    const at = m.length ? pick(rng, m.slice(0, 20000)).index : int(rng, t.length);
    const ins = m.length ? at + t.slice(at).indexOf('>') + 1 : at;
    r.push(`bigText@${ins}`);
    return t.slice(0, ins) + pick(rng, ['A', '&amp;', '9', ' ', '\u00e9']).repeat(rng() < 0.5 ? BIG : BIG / 4) + t.slice(ins);
  },
  entity(t, rng, r) {
    const e = pick(rng, ENTITIES);
    const m = [...t.matchAll(/<(t|v)(\s[^>]*)?>/g)];
    const at = m.length ? pick(rng, m.slice(0, 20000)).index : int(rng, t.length);
    const ins = m.length ? at + t.slice(at).indexOf('>') + 1 : at;
    r.push(`entity ${e}@${ins}`);
    return t.slice(0, ins) + e + t.slice(ins);
  },
  doctype(t, rng, r) {
    const decl = pick(rng, [
      '<!DOCTYPE x [<!ENTITY a "aaaaaaaaaa"><!ENTITY b "&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;"><!ENTITY c "&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;">]>',
      '<!DOCTYPE x SYSTEM "file:///etc/passwd">',
      '<!DOCTYPE x [<!ENTITY a SYSTEM "file:///etc/passwd">]>',
      '<!DOCTYPE x [<!ELEMENT x ANY>]>',
    ]);
    const at = t.indexOf('?>') >= 0 && rng() < 0.7 ? t.indexOf('?>') + 2 : 0;
    r.push(`doctype@${at}`);
    return t.slice(0, at) + decl + t.slice(at).replace(/<(t|v)(\s[^>]*)?>/, (s) => s + '&c;&a;');
  },
  cdata(t, rng, r) {
    const at = t.indexOf('>', int(rng, t.length)) + 1;
    r.push(`cdata@${at}`);
    return t.slice(0, at) + pick(rng, ['<![CDATA[', '<![CDATA[]]', '<![CDATA[x]]>]]>', '<!--', '<?pi']) + t.slice(at);
  },
  cellRef(t, rng, r) {
    const m = [...t.matchAll(/<c\s[^>]*?r="[^"]*"/g)];
    const v = pick(rng, CELL_REFS);
    r.push(`cellRef=${v}`);
    if (!m.length) return t.replace('<sheetData>', `<sheetData><row r="1"><c r="${v}"><v>1</v></c></row>`);
    const x = pick(rng, m.slice(0, 20000));
    return t.slice(0, x.index) + x[0].replace(/r="[^"]*"/, `r="${v}"`) + t.slice(x.index + x[0].length);
  },
  rowR(t, rng, r) {
    const v = pick(rng, ROW_RS);
    r.push(`row r=${v}`);
    const m = [...t.matchAll(/<row\b[^>]*?>/g)];
    if (!m.length) return t.replace('<sheetData>', `<sheetData><row r="${v}"/>`);
    const x = pick(rng, m.slice(0, 20000));
    const tag = /\sr="[^"]*"/.test(x[0]) ? x[0].replace(/\sr="[^"]*"/, ` r="${v}"`) : x[0].replace('<row', `<row r="${v}"`);
    return t.slice(0, x.index) + tag + t.slice(x.index + x[0].length);
  },
  styleIdx(t, rng, r) {
    const v = pick(rng, STYLE_IDX);
    r.push(`s=${v}`);
    const m = [...t.matchAll(/<c\b[^>]*?>/g)];
    if (!m.length) return t;
    const x = pick(rng, m.slice(0, 20000));
    const tag = /\ss="[^"]*"/.test(x[0]) ? x[0].replace(/\ss="[^"]*"/, ` s="${v}"`) : x[0].replace('<c', `<c s="${v}"`);
    return t.slice(0, x.index) + tag + t.slice(x.index + x[0].length);
  },
  sstIdx(t, rng, r) {
    const v = pick(rng, SST_IDX);
    r.push(`sst idx=${v}`);
    const m = [...t.matchAll(/<c\b[^>]*?t="s"[^>]*>\s*<v>[^<]*<\/v>/g)];
    if (!m.length) return t.replace('<sheetData>', `<sheetData><row r="1"><c r="A1" t="s"><v>${v}</v></c></row>`);
    const x = pick(rng, m.slice(0, 20000));
    return t.slice(0, x.index) + x[0].replace(/<v>[^<]*<\/v>/, `<v>${v}</v>`) + t.slice(x.index + x[0].length);
  },
  countLie(t, rng, r) {
    r.push('countLie');
    return t.replace(/(count|uniqueCount|ref|dimension)="[^"]*"/, (_, k) => `${k}="4294967295"`);
  },
  textChunk(t, rng, r) {
    const len = 1 + int(rng, Math.min(5000, t.length));
    const at = int(rng, t.length - len + 1);
    const times = pick(rng, [2, 10, 1000]);
    r.push(`xmlDup@${at}+${len}x${times}`);
    return t.slice(0, at) + t.slice(at, at + len).repeat(times) + t.slice(at + len);
  },
};

export function xmlMutation(b, rng, r) {
  const entries = zipParse(b);
  if (!entries) return undefined;
  const xmlIdx = entries.map((e, i) => (/\.xml(\.rels)?$|\.rels$/i.test(e.name) ? i : -1)).filter((i) => i >= 0);
  if (!xmlIdx.length) return undefined;
  const preferred = xmlIdx.filter((i) => /workbook|sheet|sharedStrings|styles/i.test(entries[i].name));
  const i = pick(rng, rng() < 0.8 && preferred.length ? preferred : xmlIdx);
  const data = entryData(entries[i]);
  if (!data) return undefined;
  const op = pick(rng, Object.keys(XML_OPS));
  r.push(`xml:${entries[i].name}:${op}`);
  const text = XML_OPS[op](Buffer.from(data).toString('utf8'), rng, r);
  return zipBuild(entries.map((e, k) => (k === i ? { name: e.name, data: text, level: 1 } : { name: e.name, comp: e.comp, method: e.method, usize: e.usize, crc: e.crc })));
}
