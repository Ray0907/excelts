import { strFromU8, unzipSync } from 'fflate';
import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { readXlsx, writeXlsx, type CellValue } from '../src/index.js';
import { workbookOf } from './handBuilt.js';

const first = (bytes: Uint8Array): CellValue => [...readXlsx(bytes).sheets[0]!.rows()][0]![0];
const inlineT = (body: string) => first(workbookOf(`<row r="1"><c r="A1" t="inlineStr"><is>${body}</is></c></row>`));
const shared = (body: string) => first(workbookOf('<row r="1"><c r="A1" t="s"><v>0</v></c></row>', { sst: `<sst count="1" uniqueCount="1"><si>${body}</si></sst>` }));
const str = (v: string) => first(workbookOf(`<row r="1"><c r="A1" t="str"><f>x</f><v>${v}</v></c></row>`));

const ALL: [string, (body: string) => CellValue][] = [
  ['inline <t>', (b) => inlineT(`<t>${b}</t>`)],
  ['inline rich run', (b) => inlineT(`<r><t>${b}</t></r>`)],
  ['shared <t>', (b) => shared(`<t>${b}</t>`)],
  ['shared rich run', (b) => shared(`<r><t>${b}</t></r>`)],
  ['t="str" <v>', str],
];

describe.each(ALL)('%s', (_name, read) => {
  describe('end-of-line handling', () => {
    it.each([
      ['a\r\nb', 'a\nb'],
      ['a\rb', 'a\nb'],
      ['a\n\rb', 'a\n\nb'],
      ['a\r\n\r\nb', 'a\n\nb'],
      ['a\r', 'a\n'],
      ['\r\n漢字\r\n', '\n漢字\n'],
      ['a&#13;b', 'a\rb'],
      ['a&#xD;&#10;b', 'a\r\nb'],
      ['a\r&#13;\r\nb', 'a\n\r\nb'],
    ])('%j -> %j', (raw, want) => {
      expect(read(raw)).toBe(want);
    });
  });

  describe('_xHHHH_ escapes', () => {
    it.each([
      ['_x000D_', '\r'],
      ['_x000d_', '\r'],
      ['a_x0041_b', 'aAb'],
      ['_x005F_x000D_', '_x000D_'],
      ['_x005f_x000D_', '_x000D_'],
      ['_x005F_', '_'],
      ['_x005F_x005F_', '_x005F_'],
      ['__x0041_', '_A'],
      ['_x0041__x0042_', 'AB'],
      ['Data_x001B_[32m1', 'Data\u001b[32m1'],
      ['_x00G0_', '_x00G0_'],
      ['_x12_', '_x12_'],
      ['_x0041', '_x0041'],
      ['_x00411_', '_x00411_'],
      ['x0041_', 'x0041_'],
      ['_X0041_', '_X0041_'],
      ['_x 041_', '_x 041_'],
      ['snake_case_x_y', 'snake_case_x_y'],
      ['&#95;x0041_', 'A'],
      ['_x0026_', '&'],
      ['_x0020_', ' '],
    ])('%j -> %j', (raw, want) => {
      expect(read(raw)).toBe(want);
    });
  });

  describe('CDATA', () => {
    it.each([
      ['<![CDATA[Hello]]>', 'Hello'],
      ['<![CDATA[]]>', ''],
      ['<![CDATA[a & <b> &amp; "c"]]>', 'a & <b> &amp; "c"'],
      ['pre <![CDATA[mid]]> post', 'pre mid post'],
      ['a&amp;b<![CDATA[<c>]]>d&lt;e', 'a&b<c>d<e'],
      ['<![CDATA[x]]><![CDATA[y]]>', 'xy'],
      ['<![CDATA[line1\r\nline2]]>', 'line1\nline2'],
      ['<![CDATA[&#13;]]>', '&#13;'],
      ['<![CDATA[_x0041_]]>', 'A'],
    ])('%j -> %j', (raw, want) => {
      expect(read(raw)).toBe(want);
    });
  });
});

describe('CDATA in <v>', () => {
  it('unwraps shared index, number, boolean', () => {
    const wb = workbookOf(
      '<row r="1"><c r="A1" t="s"><v><![CDATA[0]]></v></c><c r="B1"><v><![CDATA[12.5]]></v></c><c r="C1" t="b"><v><![CDATA[1]]></v></c><c r="D1" t="e"><v><![CDATA[#N/A]]></v></c></row>',
      { sst: '<sst><si><t>zero</t></si></sst>' },
    );
    expect([...readXlsx(wb).sheets[0]!.rows()][0]).toEqual(['zero', 12.5, true, { error: '#N/A' }]);
  });
});

describe('plain text keeps the cheap path', () => {
  it('leaves ordinary strings, underscores and "x" alone', () => {
    expect(shared('<t>user_xyz_1 x_ _x</t>')).toBe('user_xyz_1 x_ _x');
  });
});

describe('writer round trip', () => {
  function readableOf(s: string): Uint8Array {
    return writeXlsx({ name: 's', columns: [{ header: 'h' }], rows: [[s]] });
  }
  const back = (s: string): CellValue => [...readXlsx(readableOf(s)).sheets[0]!.rows()][1]![0];

  const control = Array.from({ length: 0x20 }, (_, i) => String.fromCharCode(i));
  it.each(control.map((c) => [c.charCodeAt(0).toString(16).padStart(2, '0'), c]))('control character 0x%s', (_hex, c) => {
    expect(back(`a${c}b`)).toBe(`a${c}b`);
  });

  it('all control characters in one string', () => {
    const s = control.join('');
    expect(back(s)).toBe(s);
  });

  it.each([
    '_x0041_',
    'a_x000D_b',
    '_x005F_x000D_',
    '_x005F_',
    '_x0041__x0042_',
    '__x0041_',
    '_x00G0_',
    '_x12_',
    '_x',
    'a\r\nb',
    'a\rb',
    '￾￿',
    '😀',
    '&<>"\' _x0026_',
  ])('%j', (s) => {
    expect(back(s)).toBe(s);
  });

  it('writes the underscore of a literal escape as _x005F_', () => {
    const xml = strFromU8(unzipSync(readableOf('_x0041_ and a_x000D_b'))['xl/sharedStrings.xml']!);
    expect(xml).toContain('<t>_x005F_x0041_ and a_x005F_x000D_b</t>');
  });

  it('writes a control character as an escape, not dropped', () => {
    const xml = strFromU8(unzipSync(readableOf('a\u0000b\u001Fc'))['xl/sharedStrings.xml']!);
    expect(xml).toContain('<t>a_x0000_b_x001F_c</t>');
  });

  it('keeps tab and newline literal, CR as a character reference', () => {
    const xml = strFromU8(unzipSync(readableOf('a\tb\nc\rd'))['xl/sharedStrings.xml']!);
    expect(xml).toContain('a\tb\nc&#13;d');
  });

  it('exceljs reads the escaped control character back', async () => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(readableOf('a\u0001b_x0041_') as unknown as Buffer);
    expect(wb.worksheets[0]!.getCell('A2').value).toBe('a\u0001b_x0041_');
  });
});
