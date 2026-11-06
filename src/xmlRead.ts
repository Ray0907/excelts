import { PLAIN, type Dialect } from './dialect.js';

export function decodeEntities(s: string): string {
  let i = s.indexOf('&');
  if (i < 0) return s;
  let out = '';
  let last = 0;
  while (i >= 0) {
    let rep = '';
    let end = -1;
    const c1 = s.charCodeAt(i + 1);
    if (c1 === 97) {
      const c2 = s.charCodeAt(i + 2);
      if (c2 === 109) {
        if (s.charCodeAt(i + 3) === 112 && s.charCodeAt(i + 4) === 59) { rep = '&'; end = i + 4; }
      } else if (c2 === 112) {
        if (s.charCodeAt(i + 3) === 111 && s.charCodeAt(i + 4) === 115 && s.charCodeAt(i + 5) === 59) { rep = "'"; end = i + 5; }
      }
    } else if (c1 === 108) {
      if (s.charCodeAt(i + 2) === 116 && s.charCodeAt(i + 3) === 59) { rep = '<'; end = i + 3; }
    } else if (c1 === 103) {
      if (s.charCodeAt(i + 2) === 116 && s.charCodeAt(i + 3) === 59) { rep = '>'; end = i + 3; }
    } else if (c1 === 113) {
      if (s.charCodeAt(i + 2) === 117 && s.charCodeAt(i + 3) === 111 && s.charCodeAt(i + 4) === 116 && s.charCodeAt(i + 5) === 59) { rep = '"'; end = i + 5; }
    } else if (c1 === 35) {
      let j = i + 2;
      const hex = s.charCodeAt(j) === 120;
      if (hex) j++;
      const first = j;
      let code = 0;
      for (;; j++) {
        const d = s.charCodeAt(j);
        let v: number;
        if (d >= 48 && d <= 57) v = d - 48;
        else if (hex && d >= 97 && d <= 102) v = d - 87;
        else if (hex && d >= 65 && d <= 70) v = d - 55;
        else break;
        if (code <= 0x10ffff) code = code * (hex ? 16 : 10) + v;
      }
      if (j > first && s.charCodeAt(j) === 59 && code <= 0x10ffff) { rep = String.fromCodePoint(code); end = j; }
    }
    if (end < 0) {
      i = s.indexOf('&', i + 1);
    } else {
      out += s.slice(last, i) + rep;
      last = end + 1;
      i = s.indexOf('&', last);
    }
  }
  return last === 0 ? s : out + s.slice(last);
}

const NEEDS_DECODE = /[&\r<]|_x/;
const LONE_CR = /\r\n?/g;
const XSTRING_ESCAPE = /_x([0-9a-fA-F]{4})_/g;
const CDATA_OPEN = '<![CDATA[';

function plainText(seg: string): string {
  // XML 1.0 normalizes literal CR LF and CR to LF before references expand, so `&#13;` survives as CR.
  return decodeEntities(seg.indexOf('\r') < 0 ? seg : seg.replace(LONE_CR, '\n'));
}

/** Decodes element content: CDATA unwrapped, end-of-line normalized, entities expanded, then the OOXML `_xHHHH_` escapes (ECMA-376 18.4.6)
 * resolved last, so `_x005F_x000D_` comes out as the literal `_x000D_`. One regex scan lets the common plain string through unchanged. */
export function decodeText(s: string): string {
  if (!NEEDS_DECODE.test(s)) return s;
  let out = '';
  let pos = 0;
  for (;;) {
    const open = s.indexOf(CDATA_OPEN, pos);
    if (open < 0) break;
    const close = s.indexOf(']]>', open + 9);
    if (close < 0) break;
    out += plainText(s.slice(pos, open)) + s.slice(open + 9, close).replace(LONE_CR, '\n');
    pos = close + 3;
  }
  out += plainText(pos === 0 ? s : s.slice(pos));
  return out.indexOf('_x') < 0 ? out : out.replace(XSTRING_ESCAPE, (_, hex: string) => String.fromCharCode(parseInt(hex, 16)));
}

/** Raw (undecoded) value of ` name="..."` inside a start tag's text. The leading space keeps `id` from matching `r:id`. */
export function attr(tag: string, name: string): string | undefined {
  const at = tag.indexOf(` ${name}=`);
  if (at < 0) return undefined;
  const q = at + name.length + 2;
  const quote = tag[q];
  const end = tag.indexOf(quote!, q + 1);
  return end < 0 ? undefined : tag.slice(q + 1, end);
}

/** Start tags `<name ...>` of the given element, in document order; the text is `<name ... ` up to (excluding) `>`. */
export function* startTags(xml: string, name: string, from = 0, to = xml.length): Generator<string> {
  const open = `<${name}`;
  let p = xml.indexOf(open, from);
  while (p >= 0 && p < to) {
    const c = xml.charCodeAt(p + open.length);
    const gt = xml.indexOf('>', p);
    if (gt < 0) return;
    if (c === 32 || c === 62 || c === 47 || c === 10 || c === 13 || c === 9) yield xml.slice(p, gt);
    p = xml.indexOf(open, gt);
  }
}

/** `seg` without its `<rPh>...</rPh>` runs; an `<rPh>` with no closing tag, and everything after it, stays. One forward pass (a lazy regex rescans to the end from every unclosed opener). */
function stripPhonetic(seg: string, d: Dialect): string {
  let out = '';
  let pos = 0;
  let open = seg.indexOf(d.rPh);
  while (open >= 0) {
    const close = seg.indexOf(d.rPhEnd, open);
    if (close < 0) break;
    out += seg.slice(pos, open);
    pos = close + d.rPhEnd.length;
    open = seg.indexOf(d.rPh, pos);
  }
  return pos === 0 ? seg : out + seg.slice(pos);
}

/** Concatenated text of every `<t>` in xml[from, to), skipping `<rPh>` phonetic runs. Whitespace is kept as written (xml:space only matters to parsers that would trim).
 * `docHasPhonetic` is decided once per part: an indexOf for `<rPh` per string scans to the end of the whole part when none exists, which is quadratic. */
export function readText(xml: string, from: number, to: number, docHasPhonetic = false, d: Dialect = PLAIN): string {
  if (docHasPhonetic) {
    const seg = xml.slice(from, to);
    if (seg.includes(d.rPh)) {
      const stripped = stripPhonetic(seg, d);
      return readText(stripped, 0, stripped.length, false, d);
    }
  }
  const open = d.t;
  const tLen = open.length;
  const close = d.tEnd;
  let out = '';
  let p = from;
  for (;;) {
    const t = xml.indexOf(open, p);
    if (t < 0 || t >= to) break;
    const c = xml.charCodeAt(t + tLen);
    const gt = xml.indexOf('>', t);
    if (gt < 0) break;
    if (c !== 62 && c !== 32 && c !== 47) {
      p = gt;
      continue;
    }
    if (xml.charCodeAt(gt - 1) === 47) {
      p = gt;
      continue;
    }
    const end = xml.indexOf(close, gt);
    if (end < 0 || end > to) break;
    out += decodeText(xml.slice(gt + 1, end));
    p = end + close.length;
  }
  return out;
}
