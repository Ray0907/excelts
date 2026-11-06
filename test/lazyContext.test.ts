import { describe, expect, it } from 'vitest';
import { strToU8, zipSync } from 'fflate';
import { readXlsx, XlsxLimitError, type XlsxLimits } from '../src/index.js';

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function pkg(sheetData: string, sst?: string, styles?: string): Uint8Array {
  const files: Record<string, Uint8Array> = {
    'xl/workbook.xml': strToU8(`${XML}<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/s1.xml"/></Relationships>`),
    'xl/worksheets/s1.xml': strToU8(`${XML}<worksheet xmlns="${MAIN}"><sheetData>${sheetData}</sheetData></worksheet>`),
  };
  if (sst !== undefined) files['xl/sharedStrings.xml'] = strToU8(`${XML}<sst xmlns="${MAIN}">${sst}</sst>`);
  if (styles !== undefined) files['xl/styles.xml'] = strToU8(`${XML}<styleSheet xmlns="${MAIN}">${styles}</styleSheet>`);
  return zipSync(files);
}
const read = (bytes: Uint8Array, limits: Partial<XlsxLimits>) => [...readXlsx(bytes, { limits }).sheets[0]!.rows()];

const FIVE_STRINGS = '<si><t>a</t></si>'.repeat(5);
const FIVE_XFS = `<cellXfs count="5">${'<xf numFmtId="14"/>'.repeat(5)}</cellXfs>`;

describe('shared strings and styles are read when a cell needs them', () => {
  it('a sheet with no shared-string cell never touches the table', () => {
    const bytes = pkg('<row r="1"><c r="A1"><v>1</v></c><c r="B1" t="inlineStr"><is><t>x</t></is></c></row>', FIVE_STRINGS);
    expect(read(bytes, { maxSharedStrings: 2 })).toEqual([[1, 'x']]);
  });
  it('the first shared-string cell does, and the table limits still apply', () => {
    const bytes = pkg('<row r="1"><c r="A1" t="s"><v>0</v></c></row>', FIVE_STRINGS);
    expect(() => read(bytes, { maxSharedStrings: 2 })).toThrow(XlsxLimitError);
    expect(read(bytes, {})).toEqual([['a']]);
  });
  it('a sheet whose cells carry no style never touches the styles', () => {
    expect(read(pkg('<row r="1"><c r="A1"><v>1</v></c></row>', undefined, FIVE_XFS), { maxCellXfs: 2 })).toEqual([[1]]);
  });
  it('a styled cell does, including style 0, which may itself be a date format', () => {
    const bytes = pkg('<row r="1"><c r="A1" s="0"><v>45000</v></c></row>', undefined, FIVE_XFS);
    expect(() => read(bytes, { maxCellXfs: 2 })).toThrow(XlsxLimitError);
    expect(read(bytes, {})[0]![0]).toBeInstanceOf(Date);
  });
  it('a styled text cell needs no styles; a missing table reads as no formats', () => {
    expect(read(pkg('<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>x</t></is></c></row>', undefined, FIVE_XFS), { maxCellXfs: 2 })).toEqual([['x']]);
    expect(read(pkg('<row r="1"><c r="A1" s="3"><v>5</v></c></row>'), {})).toEqual([[5]]);
  });
});
