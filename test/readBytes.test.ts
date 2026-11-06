import { describe, expect, it } from 'vitest';
import { strToU8, zipSync } from 'fflate';
import { readXlsx, type CellValue } from '../src/index.js';
import { fallbackCompressor } from '../src/zip.js';

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';

function build(sheetPart: Uint8Array, sst?: string): Uint8Array {
  const files: Record<string, Uint8Array> = {
    'xl/workbook.xml': strToU8(
      `${XML}<workbook ${NS} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    ),
    'xl/worksheets/sheet1.xml': sheetPart,
  };
  if (sst) files['xl/sharedStrings.xml'] = strToU8(sst);
  return zipSync(files);
}

const sheetXml = (body: string, bom = false): Uint8Array => strToU8(`${bom ? '﻿' : ''}${XML}<worksheet ${NS}><sheetData>${body}</sheetData></worksheet>`);
const grid = (bytes: Uint8Array, compressor?: Parameters<typeof readXlsx>[1]): CellValue[][] => [...readXlsx(bytes, compressor).sheets[0]!.rows()];

const SST = `${XML}<sst ${NS}><si><t>共有 \u{1F600}</t></si><si><r><t>日本</t></r><r><t>語</t></r></si></sst>`;

const MIXED =
  '<row r="1"><c r="A1" t="inlineStr"><is><t>日本語のテキスト</t></is></c><c r="B1" t="inlineStr"><is><t>party \u{1F389} &amp; &#x1F389;</t></is></c><c r="C1" t="str"><f>A1</f><v>中文 \u{1F680}</v></c><c r="D1" t="s"><v>0</v></c><c r="E1" t="s"><v>1</v></c></row>' +
  '<row r="2"><c r="A2" t="inlineStr"><is><r><t>豆</t></r><r><t xml:space="preserve"> \u{1F600}</t></r><rPh sb="0" eb="1"><t>ふり</t></rPh></r></is></c><c r="B2"><v>1.25</v></c><c r="C2" t="b"><v>1</v></c><c r="F2"><f>1+1</f><v>2</v></c></row>';

const MIXED_WANT = [
  ['日本語のテキスト', 'party \u{1F389} & \u{1F389}', '中文 \u{1F680}', '共有 \u{1F600}', '日本語'],
  ['豆 \u{1F600}', 1.25, true, undefined, undefined, 2],
];

describe('sheet scanning over bytes', () => {
  it('reads CJK, emoji and numeric references in inline strings, t="str" and shared strings', () => {
    expect(grid(build(sheetXml(MIXED), SST))).toEqual(MIXED_WANT);
  });

  it('skips a UTF-8 BOM at the start of the part', () => {
    expect(grid(build(sheetXml(MIXED, true), SST))).toEqual(MIXED_WANT);
  });

  it('handles CRLF between rows and cells', () => {
    const body = '<row r="1">\r\n<c r="A1"><v>1</v></c>\r\n<c r="B1" t="inlineStr"><is><t>漢字</t></is></c>\r\n</row>\r\n<row r="2">\r\n<c r="A2"><v>2</v></c>\r\n</row>\r\n';
    expect(grid(build(sheetXml(body)))).toEqual([[1, '漢字'], [2]]);
  });

  it('accepts extra row attributes in any order', () => {
    const body =
      '<row spans="1:2" r="1" ht="15" customHeight="1"><c r="A1"><v>1</v></c></row>' +
      '<row ht="15" x14ac:dyDescent="0.25" r="2" s="1" customFormat="1" spans="1:2"><c r="B2" t="inlineStr"><is><t>ok 🎉</t></is></c></row>' +
      '<row   r="3"  ><c r="A3"><v>3</v></c></row>';
    expect(grid(build(sheetXml(body)))).toEqual([[1], [undefined, 'ok 🎉'], [3]]);
  });

  it('yields no rows for a self-closing sheetData', () => {
    const part = strToU8(`${XML}<worksheet ${NS}><sheetData/></worksheet>`);
    expect(grid(build(part))).toEqual([]);
  });

  it('crosses a 3 MB sheet', () => {
    const rows: string[] = [];
    const want: CellValue[][] = [];
    let bytes = 0;
    for (let i = 1; bytes < 3_000_000; i++) {
      const s = i % 3 === 0 ? `行${i} \u{1F389} &amp; x` : `row${i}`;
      const x = `<row r="${i}"><c r="A${i}" t="inlineStr"><is><t>${s}</t></is></c><c r="B${i}"><v>${i}.5</v></c><c r="C${i}" t="s"><v>${i % 2}</v></c></row>`;
      rows.push(x);
      bytes += x.length;
      want.push([i % 3 === 0 ? `行${i} \u{1F389} & x` : s, i + 0.5, i % 2 ? '日本語' : '共有 \u{1F600}']);
    }
    const out = grid(build(sheetXml(rows.join('\n')), SST));
    expect(out.length).toBe(want.length);
    expect(out).toEqual(want);
  });

  it('gives the same grid with and without global Buffer, and with the fallback inflater', () => {
    const bytes = build(sheetXml(MIXED), SST);
    const node = grid(bytes);
    const saved = globalThis.Buffer;
    try {
      (globalThis as { Buffer?: unknown }).Buffer = undefined;
      expect(grid(bytes, fallbackCompressor)).toEqual(node);
      expect(grid(bytes)).toEqual(node);
    } finally {
      globalThis.Buffer = saved;
    }
    expect(grid(bytes, fallbackCompressor)).toEqual(node);
    expect(node).toEqual(MIXED_WANT);
  });
});
