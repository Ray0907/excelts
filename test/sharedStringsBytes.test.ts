import { describe, expect, it } from 'vitest';
import { SharedStrings } from '../src/sharedStrings.js';
import { escapeText } from '../src/xml.js';

const SAMPLES = [
  'plain', '', ' lead', 'trail ', 'a\nb', 'a\rb', 'a&b<c>d"e', 'tab\there', '\u0001ctl\u001f', '_x0041_ literal', '_x005F_', '_xABCD\u0001', 'x_x', '_x00', '日本語 Привет 😀', 'lone\ud800surrogate', '￿￾',
  'x'.repeat(40000), 'ab&'.repeat(30000), 'é'.repeat(70000),
];

describe('SharedStrings.toBytes', () => {
  it('equals the UTF-8 encoding of toXml() for any mix of strings, rich entries and sizes', () => {
    const ss = new SharedStrings();
    for (const s of SAMPLES.filter((s) => s.length <= 32767)) ss.add(s);
    ss.addRich('<r><t>rich &amp; text</t></r>');
    for (const s of SAMPLES.filter((s) => s.length <= 32767)) ss.add(s + '!');
    expect(Buffer.from(ss.toBytes()).equals(Buffer.from(new TextEncoder().encode(ss.toXml())))).toBe(true);
  });
  it('is the empty table for no strings', () => {
    const ss = new SharedStrings();
    expect(new TextDecoder().decode(ss.toBytes())).toBe(ss.toXml());
  });
  it('handles many chunks of small and large strings', () => {
    const ss = new SharedStrings();
    for (let i = 0; i < 5000; i++) ss.add(`s${i}${'z'.repeat(i % 700)}`);
    for (let i = 0; i < 100; i++) ss.add(`${i}`.repeat(9000));
    expect(Buffer.from(ss.toBytes()).equals(Buffer.from(new TextEncoder().encode(ss.toXml())))).toBe(true);
  });
});

describe('escapeText fast path', () => {
  it('returns its input unchanged when nothing needs escaping and escapes exactly as before otherwise', () => {
    const ILLEGAL_OR_SPECIAL = /[&<>"\r\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]|_(?=x[0-9a-fA-F]{4}(?:_|[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]))/;
    for (const s of SAMPLES) {
      const e = escapeText(s);
      if (!ILLEGAL_OR_SPECIAL.test(s)) expect(e).toBe(s);
      else expect(e).not.toBe(s);
    }
    expect(escapeText('_x0041_')).toBe('_x005F_x0041_');
    expect(escapeText('a\u0001b')).toBe('a_x0001_b');
    expect(escapeText('a&b')).toBe('a&amp;b');
  });
});
