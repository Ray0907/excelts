import { describe, expect, it } from 'vitest';
import { decodeEntities } from '../src/xmlRead.js';

// Pre-R8 reference implementation (regexp + callback), kept as the oracle.
const ENTITY = /&(#x[0-9a-fA-F]+|#[0-9]+|amp|lt|gt|quot|apos);/g;
function oldEntity(whole: string, body: string): string {
  switch (body) {
    case 'amp': return '&';
    case 'lt': return '<';
    case 'gt': return '>';
    case 'quot': return '"';
    case 'apos': return "'";
  }
  const code = body.charCodeAt(1) === 120 ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10);
  return code <= 0x10ffff ? String.fromCodePoint(code) : whole;
}
function oldDecode(s: string): string {
  return s.indexOf('&') < 0 ? s : s.replace(ENTITY, oldEntity);
}

describe('decodeEntities', () => {
  it.each([
    ['&amp;amp;', '&amp;'],
    ['&#x1F389;', '\u{1F389}'],
    ['&#0;', '\0'],
    ['&#xD800;', '\ud800'],
    ['&#1114112;', '&#1114112;'],
    ['&#x;', '&#x;'],
    ['&#;', '&#;'],
    ['&AMP;', '&AMP;'],
    ['&amp', '&amp'],
    ['x&', 'x&'],
    ['&&amp;', '&&'],
    ['&#00000000000065;', 'A'],
    ['&#X41;', '&#X41;'],
    ['&foo; &quot;&apos;&lt;&gt;', '&foo; "\'<>'],
    ['note &amp; &lt;7&gt;', 'note & <7>'],
    ['&#' + '9'.repeat(400) + ';', '&#' + '9'.repeat(400) + ';'],
    ['&#x' + 'F'.repeat(400) + ';', '&#x' + 'F'.repeat(400) + ';'],
  ])('fixed case %j', (input, expected) => {
    expect(oldDecode(input)).toBe(expected);
    expect(decodeEntities(input)).toBe(expected);
  });

  it('matches the pre-R8 implementation on 100k random strings', () => {
    let seed = 12345;
    const rnd = () => ((seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0) / 0x100000000);
    const atoms = ['&', '&', '&', ';', ';', '#', 'x', 'X', '0', '1', '9', 'a', 'f', 'A', 'F', 'amp', 'lt', 'gt', 'quot', 'apos', 'AMP', 'q', 'z', ' ', '中', '日本', '\u{1F389}', '\ud83c', '10FFFF', '110000', 'D800'];
    for (let n = 0; n < 100_000; n++) {
      const len = Math.floor(rnd() * 14);
      let s = '';
      for (let k = 0; k < len; k++) s += atoms[Math.floor(rnd() * atoms.length)];
      expect(decodeEntities(s), JSON.stringify(s)).toBe(oldDecode(s));
    }
  });
});
