import { describe, expect, it } from 'vitest';
import fc from 'fast-check';
import { attributes, readRelationships, targetPaths } from '../src/workbookParts.js';

/** The general resolution, as it was before the plain-path shortcut. */
function referenceTargetPaths(folder: string, target: string): string[] {
  const slashed = target.replace(/\\/g, '/');
  const spellings = [slashed];
  try {
    const decoded = decodeURIComponent(slashed);
    if (decoded !== slashed) spellings.unshift(decoded);
  } catch {
    // literal spelling stands
  }
  return spellings.map((t) => {
    const parts: string[] = [];
    for (const seg of (t.startsWith('/') ? t.slice(1) : `${folder}${t}`).split('/')) {
      if (seg === '..') parts.pop();
      else if (seg !== '.' && seg !== '') parts.push(seg);
    }
    return parts.join('/');
  });
}

const piece = fc.constantFrom('a', 'B', '0', '_', '-', '.', '..', '/', '//', '%41', '%zz', '\\', ' ', 'é', 'xl', 'sheet1', '.xml', '&amp;');
const text = fc.array(piece, { maxLength: 8 }).map((a) => a.join(''));
const folders = fc.oneof(fc.constantFrom('', 'xl/', 'a/b/', 'xl/worksheets/'), text);

describe('targetPaths', () => {
  it('agrees with the general resolution for any folder and target', () => {
    fc.assert(
      fc.property(folders, text, (folder, target) => {
        expect(targetPaths(folder, target)).toEqual(referenceTargetPaths(folder, target));
      }),
      { numRuns: 20000 },
    );
  });
  it('agrees on the plain shapes real packages use', () => {
    for (const [f, t] of [['xl/', 'worksheets/sheet1.xml'], ['xl/', 'sharedStrings.xml'], ['', 'xl/workbook.xml'], ['xl/', 'theme/theme1.xml'], ['xl/', '../docProps/app.xml'], ['xl/', '/xl/styles.xml'], ['xl/', 'sheet 1.xml'], ['xl/', 'a%20b.xml']] as const) {
      expect(targetPaths(f, t)).toEqual(referenceTargetPaths(f, t));
    }
  });
});

const TYPES = ['http://x/officeDocument', 'http://x/worksheet', 'http://x/sharedStrings', 'http://x/styles', 'http://x/theme', 'http://x/calcChain', 'http://x/worksheetX', 'http://x/mystyles', 'http://x/styles2'];
const rel = fc.record({ id: fc.constantFrom('rId1', 'rId2', 'x'), type: fc.constantFrom(...TYPES), target: fc.constantFrom('a.xml', 'b/c.xml', 'x&amp;y.xml'), ext: fc.boolean(), sp: fc.constantFrom(' ', '\n', '\t') });
const relsXml = fc.array(rel, { maxLength: 8 }).map((rs) => `<Relationships xmlns="p">${rs.map((r) => `<Relationship${r.sp}Id="${r.id}"${r.sp}Type="${r.type}"${r.sp}Target="${r.target}"${r.ext ? `${r.sp}TargetMode="External"` : ''}/>`).join('')}</Relationships>`);

describe('readRelationships pre-filter', () => {
  const wanted = (suffixes: string[]) => (r: { type: string }) => suffixes.some((s) => r.type.endsWith(s));
  it('never drops a relationship whose type ends with a wanted suffix, and keeps order', () => {
    fc.assert(
      fc.property(relsXml, (xml) => {
        const all = [...readRelationships(xml)];
        const sets: [string[], string[]][] = [[['officeDocument'], ['/officeDocument']], [['worksheet', 'sharedStrings', 'styles'], ['/worksheet', '/sharedStrings', '/styles']]];
        for (const [needles, suffixes] of sets) {
          const filtered = [...readRelationships(xml, needles)];
          expect(filtered.filter(wanted(suffixes))).toEqual(all.filter(wanted(suffixes)));
        }
      }),
      { numRuns: 3000 },
    );
  });
  it('is lazy: taking the first relationship does not parse the rest', () => {
    const it = readRelationships('<Relationships><Relationship Id="a" Type="t/officeDocument" Target="x"/><Relationship Id="b" Type="t" Target="y"/></Relationships>', ['officeDocument']);
    expect(it.next().value).toMatchObject({ id: 'a' });
  });
});

/** The attribute scanner as it was before the character-class table. */
function referenceAttributes(tag: string): Map<string, string> {
  const WS = /\s/;
  const isSpace = (c: number): boolean => c === 32 || (c >= 9 && c <= 13) || (c > 127 && WS.test(String.fromCharCode(c)));
  const isNameChar = (c: number): boolean => !(c === 61 || c === 47 || c === 34 || c === 39 || c === 60 || c === 62 || isSpace(c));
  const out = new Map<string, string>();
  const n = tag.length;
  let i = 0;
  while (i < n) {
    if (!isNameChar(tag.charCodeAt(i))) {
      i++;
      continue;
    }
    const nameStart = i;
    while (i < n && isNameChar(tag.charCodeAt(i))) i++;
    const nameEnd = i;
    let j = i;
    while (j < n && isSpace(tag.charCodeAt(j))) j++;
    if (tag.charCodeAt(j) !== 61) continue;
    j++;
    while (j < n && isSpace(tag.charCodeAt(j))) j++;
    const quote = tag.charCodeAt(j);
    if (quote !== 34 && quote !== 39) continue;
    const close = tag.indexOf(quote === 34 ? '"' : "'", j + 1);
    if (close < 0) continue;
    const name = tag.slice(nameStart, nameEnd);
    if (!out.has(name)) out.set(name, tag.slice(j + 1, close));
    i = close + 1;
  }
  return out;
}

describe('attributes', () => {
  const bit = fc.constantFrom('<a', ' ', '\t', '\n', '\r', '\u000b', '\u00a0', '\u2028', '\ufeff', '\u3000', '=', '"', "'", '/', '>', 'x', 'r:id', 'xmlns:r', 'Id', 'é', '😀', '\u0000', '&amp;', '1');
  const tags = fc.array(bit, { maxLength: 24 }).map((a) => a.join(''));
  it('matches the scanner it replaced on arbitrary tag text', () => {
    fc.assert(
      fc.property(tags, (t) => {
        expect([...attributes(t)]).toEqual([...referenceAttributes(t)]);
      }),
      { numRuns: 30000 },
    );
  });
  it('reads ordinary tags', () => {
    expect([...attributes(`<sheet name="A b" sheetId='2' r:id = "rId1"`)]).toEqual([['name', 'A b'], ['sheetId', '2'], ['r:id', 'rId1']]);
  });
});
