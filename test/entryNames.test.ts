import { describe, expect, it } from 'vitest';
import { strToU8, zipSync } from 'fflate';
import { unzip } from '../src/unzip.js';
import { readXlsx } from '../src/index.js';

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

describe('zip entry names', () => {
  it('decode as UTF-8 whether plain ASCII, accented, CJK, or a long path', () => {
    const long = `xl/${'d/'.repeat(150)}file.xml`;
    const names = ['a.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/Blatt-ß.xml', 'xl/工作表/表1.xml', 'emoji-😀.xml', long, 'with space/x y.xml', ''];
    const z = unzip(zipSync(Object.fromEntries(names.filter(Boolean).map((n) => [n, strToU8('x')]))));
    expect([...z.names].sort()).toEqual(names.filter(Boolean).sort());
    for (const n of names.filter(Boolean)) expect(Array.from(z.get(n)!)).toEqual([120]);
  });

  it('reads a workbook whose sheet part has a non-ASCII name', () => {
    const files = {
      'xl/workbook.xml': strToU8(`${XML}<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`),
      'xl/_rels/workbook.xml.rels': strToU8(`${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/Blatt-%C3%9F.xml"/><Relationship Id="rId2" Type="${REL}/theme" Target="theme/theme1.xml"/><Relationship Id="rId3" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/></Relationships>`),
      'xl/worksheets/Blatt-ß.xml': strToU8(`${XML}<worksheet xmlns="${MAIN}"><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>`),
    };
    expect([...readXlsx(zipSync(files)).sheets[0]!.rows()]).toEqual([[7]]);
  });
});
