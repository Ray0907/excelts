import { describe, expect, it } from 'vitest';
import { scanRows, type SheetContext } from '../src/readSheet.js';
import { readSharedStrings } from '../src/readSharedStrings.js';

const ctx: SheetContext = { sharedStrings: [], xfKind: [], date1904: false };
const inline = (...bodies: string[]) => {
  const cells = bodies.map((b) => `<c t="inlineStr">${b}</c>`).join('');
  return [...scanRows(new TextEncoder().encode(`<sheetData><row r="1">${cells}</row></sheetData>`), ctx)][0]!;
};
const sst = (...entries: string[]) => readSharedStrings(`<sst count="1">${entries.join('')}</sst>`);

describe('inline strings', () => {
  it('reads each shape', () => {
    expect(
      inline(
        '<is><t>user_1</t></is>',
        '<is><t>a &amp; b &lt; c &#13; &#x1F389;</t></is>',
        '<is><t xml:space="preserve">  pad  </t></is>',
        '<is><r><t>a</t></r><r><t>b</t></r></is>',
        '<is><t/></is>',
        '<is></is>',
        '<is><t></t></is>',
        '<is><t>a</t><t>b</t></is>',
        '<is><t>x &gt; y > z</t></is>',
        '<is><t>漢字 🎉</t></is>',
        '<is><t>keep</t><rPh sb="0" eb="1"><t>drop</t></rPh></is>',
      ),
    ).toEqual(['user_1', 'a & b < c \r 🎉', '  pad  ', 'ab', '', '', '', 'ab', 'x > y > z', '漢字 🎉', 'keep']);
  });

  it('phonetic runs in one cell leave plain cells alone', () => {
    const row = inline('<is><t>plain</t></is>', '<is><t>base</t><rPh sb="0" eb="1"><t>ph</t></rPh></is>', '<is><t>after</t></is>');
    expect(row).toEqual(['plain', 'base', 'after']);
  });
});

describe('shared strings', () => {
  it('mixes plain, rich, empty and phonetic entries', () => {
    expect(
      sst(
        '<si><t>plain</t></si>',
        '<si><r><t>a</t></r><r><t>b</t></r></si>',
        '<si/>',
        '<si><t></t></si>',
        '<si><t>base</t><rPh sb="0" eb="1"><t>ph</t></rPh></si>',
        '<si><t>next</t></si>',
        '<si><t xml:space="preserve"> sp </t></si>',
        '<si><t>&amp; &#x1F389; &lt;</t></si>',
        '<si><t>漢字 🎉</t></si>',
        '<si><t>x > y</t></si>',
        '<si><t>a</t><t>b</t></si>',
      ),
    ).toEqual(['plain', 'ab', '', '', 'base', 'next', ' sp ', '& 🎉 <', '漢字 🎉', 'x > y', 'ab']);
  });

  it('plain entries without any phonetic run', () => {
    expect(sst('<si><t>one</t></si>', '<si><t>two</t></si>')).toEqual(['one', 'two']);
  });
});
