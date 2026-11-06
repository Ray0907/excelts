import { describe, expect, it } from 'vitest';
import { strToU8, zipSync } from 'fflate';
import { readXlsx, type CellValue } from '../src/index.js';

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
function book(worksheet: string): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(
      `${XML}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/s1.xml"/></Relationships>`,
    ),
    'xl/worksheets/s1.xml': strToU8(worksheet),
  });
}
const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const grid = (worksheet: string): CellValue[][] => [...readXlsx(book(worksheet)).sheets[0]!.rows()];
const withData = (rows: string, after = '</sheetData></worksheet>') => grid(`${XML}<worksheet ${NS}><sheetData>${rows}${after}`);
const c = (ref: string, v: number) => `<c r="${ref}"><v>${v}</v></c>`;

describe('row and cell boundaries', () => {
  it('tolerates whitespace between cells, rows and before </row>', () => {
    const rows = `\n  <row r="1">\r\n\t${c('A1', 1)}\r\n\t${c('B1', 2)}\n  </row>\r\n  <row r="2">\n${c('A2', 3)}\n</row>\n`;
    expect(withData(rows)).toEqual([[1, 2], [3]]);
  });

  it('ignores an extLst child after the cells', () => {
    const rows = `<row r="1">${c('A1', 1)}<extLst><ext uri="x"><x14:foo xmlns:x14="y"/></ext></extLst></row><row r="2">${c('A2', 2)}</row>`;
    expect(withData(rows)).toEqual([[1], [2]]);
  });

  it('keeps cells that follow a non-cell child', () => {
    const rows = `<row r="1"><extLst/>${c('A1', 1)}${c('B1', 2)}</row><row r="2">${c('A2', 3)}</row>`;
    expect(withData(rows)).toEqual([[1, 2], [3]]);
  });

  it('handles self-closing and empty rows', () => {
    const rows = `<row r="1"/><row r="2"></row><row r="3"/><row r="5" spans="1:3"></row><row r="6" spans="1:3">\n</row><row r="7">${c('A7', 7)}</row>`;
    expect(withData(rows)).toEqual([[], [], [], [], [], [], [7]]);
  });

  it('handles two consecutive empty rows between data rows', () => {
    const rows = `<row r="1">${c('A1', 1)}</row><row r="2"></row><row r="3"></row><row r="4">${c('A4', 4)}</row>`;
    expect(withData(rows)).toEqual([[1], [], [], [4]]);
  });

  it('separates cells by CRLF and tabs, including self-closing cells', () => {
    const rows = `<row r="1">${c('A1', 1)}\r\n<c r="B1" s="0"/>\t\t${c('C1', 3)}\r\n</row>`;
    expect(withData(rows)).toEqual([[1, undefined, 3]]);
  });

  it('reads a last row that has no </row> before </sheetData>', () => {
    expect(withData(`<row r="1">${c('A1', 1)}</row><row r="2">${c('A2', 2)}${c('B2', 3)}`)).toEqual([[1], [2, 3]]);
  });

  it('does not read <row inside content after </sheetData>', () => {
    const after = `</sheetData><rowBreaks count="1"><brk id="3"/></rowBreaks><row r="9">${c('A9', 9)}</row></worksheet>`;
    expect(withData(`<row r="1">${c('A1', 1)}</row>`, after)).toEqual([[1]]);
  });

  it('reads a sheet without </sheetData>', () => {
    expect(grid(`${XML}<worksheet ${NS}><sheetData><row r="1">${c('A1', 1)}</row><row r="2">${c('A2', 2)}</row>`)).toEqual([[1], [2]]);
  });

  it('does not take <col-like or escaped text for a cell', () => {
    const rows = `<row r="1"><c r="A1" t="inlineStr"><is><t>&lt;c r="B1"&gt;</t></is></c><c r="B1" t="inlineStr"><is><r><rPr><charset val="1"/><color rgb="FF"/></rPr><t>x</t></r></is></c></row>`;
    expect(withData(rows)).toEqual([['<c r="B1">', 'x']]);
  });
});

describe('inline strings scale linearly', () => {
  const n = 30000;
  it('30k inline cells without phonetic runs read in well under a second', () => {
    const rows = Array.from({ length: n }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}" t="inlineStr"><is><r><t>s${i}</t></r></is></c></row>`).join('');
    const t0 = performance.now();
    const g = withData(rows);
    expect(g.length).toBe(n);
    expect(g[n - 1]).toEqual([`s${n - 1}`]);
    expect(performance.now() - t0).toBeLessThan(1000);
  });

  it('skips phonetic text only in the one cell that carries it', () => {
    const plain = (i: number) => `<c r="A${i}" t="inlineStr"><is><r><t>p${i}</t></r></is></c>`;
    const rich = '<c r="B1" t="inlineStr"><is><r><t>a</t></r><rPh sb="0" eb="1"><t>PHON</t></rPh><r><t>b</t></r></is></c>';
    const rows = `<row r="1">${plain(1)}${rich}</row>` + Array.from({ length: 2000 }, (_, i) => `<row r="${i + 2}">${plain(i + 2)}</row>`).join('');
    const g = withData(rows);
    expect(g[0]).toEqual(['p1', 'ab']);
    expect(g[2000]).toEqual(['p2001']);
  });
});
