import { describe, expect, it } from 'vitest';
import ExcelJS from 'exceljs';
import { strToU8, zipSync } from 'fflate';
import { readXlsx, writeXlsx, type CellValue } from '../src/index.js';
import { fallbackCompressor } from '../src/zip.js';
// @ts-expect-error untyped fixture module
import { FIXTURES, expectedGrid } from '../verify/fixtures.mjs';
// @ts-expect-error untyped bench module
import { COLUMNS, row } from './helpers/data.mjs';

const grid = (bytes: Uint8Array, sheet = 0): CellValue[][] => [...readXlsx(bytes).sheets[sheet]!.rows()];

const norm = (g: readonly (readonly CellValue[])[]): unknown[][] =>
  g.map((r) => {
    const out = r.map((v) => (v instanceof Date ? v.getTime() : typeof v === 'number' && !Number.isFinite(v) ? null : (v ?? null)));
    while (out.length && out[out.length - 1] === null) out.pop();
    return out;
  });

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
function handBuilt(sheetXml: string, opts: { styles?: string; sst?: string; workbookPr?: string; target?: string } = {}): Uint8Array {
  const files: Record<string, Uint8Array> = {
    'xl/workbook.xml': strToU8(
      `${XML}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${opts.workbookPr ?? ''}<sheets><sheet name="A &amp; B" sheetId="7" r:id="rId9"/></sheets></workbook>`,
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${opts.target ?? '/xl/worksheets/s1.xml'}"/></Relationships>`,
    ),
    'xl/worksheets/s1.xml': strToU8(`${XML}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${sheetXml}</sheetData></worksheet>`),
  };
  if (opts.styles) files['xl/styles.xml'] = strToU8(opts.styles);
  if (opts.sst) files['xl/sharedStrings.xml'] = strToU8(opts.sst);
  return zipSync(files);
}

describe('readXlsx round trip of writeXlsx', () => {
  for (const f of FIXTURES.filter((x: { big?: boolean }) => !x.big)) {
    it(`fixture ${f.name}`, () => {
      expect(norm(grid(writeXlsx(f)))).toEqual(norm(expectedGrid(f)));
    });
  }

  it('bench-shaped 1000 rows equal the input, with the fallback inflater too', () => {
    const rows = Array.from({ length: 1000 }, (_, i) => row(i));
    const bytes = writeXlsx({ name: 'bench', columns: COLUMNS, rows });
    const want = [COLUMNS.map((c: { header: string }) => c.header), ...rows];
    expect(grid(bytes)).toEqual(want);
    expect([...readXlsx(bytes, {}, fallbackCompressor).sheets[0]!.rows()]).toEqual(want);
  });

  it('exposes the sheet name', () => {
    expect(readXlsx(writeXlsx({ name: 'My & Sheet', columns: [{ header: 'h' }], rows: [] })).sheets.map((s) => s.name)).toEqual(['My & Sheet']);
  });

  it('rows() can be iterated twice', () => {
    const ws = readXlsx(writeXlsx({ name: 'S', columns: [{ header: 'h' }], rows: [['x']] })).sheets[0]!;
    expect([...ws.rows()]).toEqual([...ws.rows()]);
  });
});

describe('readXlsx on files written by the exceljs writer', () => {
  it('reads every cell type, two sheets, a gap, and a 1900 date', async () => {
    const wb = new ExcelJS.Workbook();
    const a = wb.addWorksheet('Alpha');
    const when = new Date(Date.UTC(2021, 5, 15, 12, 30, 0));
    a.getCell('A1').value = 'text';
    a.getCell('B1').value = 42.5;
    a.getCell('C1').value = true;
    a.getCell('D1').value = when;
    a.getCell('D1').numFmt = 'yyyy-mm-dd hh:mm:ss';
    a.getCell('E1').value = { richText: [{ text: 'Hello ' }, { font: { bold: true }, text: 'wor' }, { text: 'ld' }] };
    a.getCell('F1').value = { formula: 'B1*2', result: 85 };
    a.getCell('G1').value = { error: '#DIV/0!' } as ExcelJS.CellErrorValue;
    a.getCell('H1').value = { formula: 'A1&"!"', result: 'text!' };
    a.getCell('A2').value = 'left';
    a.getCell('C2').value = 'right of gap';
    const b = wb.addWorksheet('Beta - 2');
    b.getCell('A1').value = 'only';
    const bytes = new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);
    const r = readXlsx(bytes);
    expect(r.sheets.map((s) => s.name)).toEqual(['Alpha', 'Beta - 2']);
    expect(grid(bytes, 0)).toEqual([
      ['text', 42.5, true, when, 'Hello world', 85, { error: '#DIV/0!' }, 'text!'],
      ['left', undefined, 'right of gap'],
    ]);
    expect(grid(bytes, 1)).toEqual([['only']]);
  });

  it('honours date1904', async () => {
    const wb = new ExcelJS.Workbook();
    wb.properties.date1904 = true;
    const ws = wb.addWorksheet('S');
    const when = new Date(Date.UTC(2021, 5, 15, 12, 30, 0));
    ws.getCell('A1').value = when;
    ws.getCell('A1').numFmt = 'yyyy-mm-dd hh:mm:ss';
    const bytes = new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);
    expect(grid(bytes)).toEqual([[when]]);
  });
});

describe('readXlsx on hand-built files', () => {
  it('accepts rows and cells without r, and every cell type', () => {
    const sst = `${XML}<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>plain</t></si><si><r><t xml:space="preserve">a </t></r><r><rPr><b/></rPr><t>b</t></r><rPh sb="0" eb="1"><t>PHON</t></rPh></si><si><t>x &amp; y &lt;z&gt; &#13;&#x1F389;</t></si></sst>`;
    const bytes = handBuilt(
      '<row><c t="inlineStr"><is><t>a &amp; b</t></is></c><c t="str"><f>A1</f><v>x &#13; y</v></c><c t="b"><v>1</v></c><c t="b"><v>0</v></c><c t="e"><v>#DIV/0!</v></c><c t="d"><v>2021-06-15T12:30:00Z</v></c><c><v>1.5</v></c><c/><c t="s"><v>0</v></c></row>' +
        '<row><c t="inlineStr"><is><r><t xml:space="preserve">rich </t></r><r><t>&#x1F389;</t></r></is></c><c t="s"><v>1</v></c><c t="s"><v>2</v></c><c><f>1+1</f></c></row>' +
        '<row/><row><c t="d"><v>2021-06-15</v></c></row>',
      { sst },
    );
    expect(grid(bytes)).toEqual([
      ['a & b', 'x \r y', true, false, { error: '#DIV/0!' }, new Date(Date.UTC(2021, 5, 15, 12, 30)), 1.5, undefined, 'plain'],
      ['rich \u{1F389}', 'a b', 'x & y <z> \r\u{1F389}'],
      [],
      [new Date(Date.UTC(2021, 5, 15))],
    ]);
    expect(readXlsx(bytes).sheets[0]!.name).toBe('A & B');
  });

  it('places cells by column letters and fills missing row numbers with empty rows', () => {
    const bytes = handBuilt('<row r="3"><c r="C3"><v>1</v></c><c r="B3" s="0"/><c r="AA3"><v>2</v></c><c r="AB3" s="0"/></row><row r="5"><c r="A5"><v>3</v></c></row>');
    const g = grid(bytes);
    expect(g.length).toBe(5);
    expect(g[0]).toEqual([]);
    expect(g[1]).toEqual([]);
    expect(g[2]!.length).toBe(27);
    expect(g[2]![2]).toBe(1);
    expect(g[2]![1]).toBeUndefined();
    expect(g[2]![26]).toBe(2);
    expect(g[3]).toEqual([]);
    expect(g[4]).toEqual([3]);
  });

  it('classifies date formats via styles', () => {
    const numFmts = [
      [164, 'yyyy-mm-dd'],
      [165, '0.00&quot;d&quot;'],
      [166, '[h]:mm:ss'],
      [167, '\\d0'],
      [168, '[Red]#,##0'],
      [169, 'General'],
      [170, '&quot;year&quot; 0'],
    ];
    const xfs = [0, 14, 2, 164, 165, 166, 167, 168, 169, 170, 22, 49];
    const styles =
      `${XML}<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="${numFmts.length}">` +
      numFmts.map(([id, code]) => `<numFmt numFmtId="${id}" formatCode="${code}"/>`).join('') +
      '</numFmts><cellStyleXfs count="1"><xf numFmtId="14"/></cellStyleXfs><cellXfs count="' + xfs.length + '">' +
      xfs.map((id) => `<xf numFmtId="${id}" fontId="0"><alignment/></xf>`).join('') +
      '</cellXfs></styleSheet>';
    const cells = xfs.map((_, s) => `<c s="${s}"><v>44000</v></c>`).join('');
    const isDate = grid(handBuilt(`<row>${cells}</row>`, { styles }))[0]!.map((v) => v instanceof Date);
    expect(isDate).toEqual([false, true, false, true, false, true, false, false, false, false, true, false]);
  });

  it('converts serials for the 1900 system including the Lotus leap bug, and the 1904 system', () => {
    const styles = `${XML}<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>`;
    const sheet = '<row><c s="1"><v>25569</v></c><c s="1"><v>61</v></c><c s="1"><v>59</v></c><c s="1"><v>1</v></c><c s="1"><v>44000.5</v></c></row>';
    expect(grid(handBuilt(sheet, { styles }))[0]).toEqual([
      new Date(Date.UTC(1970, 0, 1)),
      new Date(Date.UTC(1900, 2, 1)),
      new Date(Date.UTC(1900, 1, 28)),
      new Date(Date.UTC(1900, 0, 1)),
      new Date(Date.UTC(2020, 5, 18, 12)),
    ]);
    expect(grid(handBuilt('<row><c s="1"><v>24107</v></c><c s="1"><v>0</v></c></row>', { styles, workbookPr: '<workbookPr date1904="1"/>' }))[0]).toEqual([
      new Date(Date.UTC(1970, 0, 1)),
      new Date(Date.UTC(1904, 0, 1)),
    ]);
  });

  it('reads r, s and t in any attribute order, quoting and with extra attributes', () => {
    const sst = `${XML}<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>zero</t></si><si><t>one</t></si></sst>`;
    const xfs = Array.from({ length: 13 }, (_, i) => `<xf numFmtId="${i === 12 ? 14 : 0}"/>`).join('');
    const styles = `${XML}<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="13">${xfs}</cellXfs></styleSheet>`;
    const cells =
      '<c t=\'s\' r=\'A1\'><v>1</v></c>' +
      '<c s="12" r="B1"><v>25569</v></c>' +
      '<c cm="1" vm="2" r="C1" s="12" ph="a>b"><v>25569</v></c>' +
      '<c r="D1" t="s" s="3" ph="x&gt;y"><v>0</v></c>' +
      '<c  r="E1"\n t="inlineStr"><is><t>ws</t></is></c>' +
      '<c r="F1" s="11"><v>7</v></c>';
    expect(grid(handBuilt(`<row r="1">${cells}</row>`, { sst, styles }))[0]).toEqual([
      'one',
      new Date(Date.UTC(1970, 0, 1)),
      new Date(Date.UTC(1970, 0, 1)),
      'zero',
      'ws',
      7,
    ]);
  });

  it('keeps multi-letter columns apart across gaps', () => {
    const bytes = handBuilt('<row r="2"><c r="Z2"><v>1</v></c><c r="AA2"><v>2</v></c><c r="AZ2"><v>3</v></c><c r="XFD2"><v>4</v></c></row>');
    const row = grid(bytes)[1]!;
    expect(row.length).toBe(16384);
    expect([row[25], row[26], row[51], row[16383]]).toEqual([1, 2, 3, 4]);
    expect(row.filter((v) => v !== undefined).length).toBe(4);
  });

  it('tells t="s", "str", "inlineStr", "b", "e", "d" and "n" apart', () => {
    const sst = `${XML}<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>sst0</t></si></sst>`;
    const bytes = handBuilt(
      '<row><c t="s"><v>0</v></c><c t="str"><v>0</v></c><c t="inlineStr"><v>0</v></c><c t="b"><v>1</v></c><c t="e"><v>0</v></c><c t="d"><v>2021-06-15</v></c><c t="n"><v>5</v></c><c t="zz"><v>0</v></c><c t=""><v>9</v></c></row>',
      { sst },
    );
    expect(grid(bytes)[0]).toEqual(['sst0', '0', '0', true, { error: '0' }, new Date(Date.UTC(2021, 5, 15)), 5, '0', '9']);
  });

  it('accepts single-quoted row r and row attributes in any order', () => {
    const bytes = handBuilt("<row spans=\"1:1\" r='2'><c><v>1</v></c></row><row ht=\"15\" r=\"4\" customHeight=\"1\"><c><v>2</v></c></row>");
    expect(grid(bytes)).toEqual([[], [1], [], [2]]);
  });

  it('resolves relative targets', () => {
    expect(grid(handBuilt('<row><c><v>1</v></c></row>', { target: 'worksheets/s1.xml' }))).toEqual([[1]]);
  });
});

describe('readXlsx rejects bad zips', () => {
  const good = () => writeXlsx({ name: 'S', columns: [{ header: 'h' }], rows: [['x']] });

  it('truncated', () => {
    const b = good();
    expect(() => readXlsx(b.subarray(0, b.length - 30))).toThrow();
    expect(() => readXlsx(new Uint8Array(0))).toThrow();
    expect(() => readXlsx(strToU8('not a zip at all, just text'))).toThrow();
  });

  it('zip64 marker throws RangeError', () => {
    const b = good().slice();
    new DataView(b.buffer).setUint32(b.length - 22 + 16, 0xffffffff, true);
    expect(() => readXlsx(b)).toThrow(RangeError);
  });
});

describe('readXlsx shared strings scale linearly', () => {
  it('30k unique strings without phonetic runs read in well under a second', () => {
    const n = 30000;
    const sst = `<sst>${Array.from({ length: n }, (_, i) => `<si><t>s${i}</t></si>`).join('')}</sst>`;
    const sheet = `<worksheet><sheetData><row><c t="s"><v>${n - 1}</v></c></row></sheetData></worksheet>`;
    const t0 = performance.now();
    const rows = grid(handBuilt(sheet, { sst }));
    expect(rows[0]).toEqual([`s${n - 1}`]);
    expect(performance.now() - t0).toBeLessThan(1000);
  });

  it('phonetic runs are skipped only in the strings that carry them', () => {
    const sst = '<sst><si><t>plain</t></si><si><r><t>a</t></r><rPh sb="0" eb="1"><t>PHON</t></rPh><r><t>b</t></r></si><si><t>after</t></si></sst>';
    const sheet = '<worksheet><sheetData><row><c t="s"><v>0</v></c><c t="s"><v>1</v></c><c t="s"><v>2</v></c></row></sheetData></worksheet>';
    expect(grid(handBuilt(sheet, { sst }))[0]).toEqual(['plain', 'ab', 'after']);
  });
});
