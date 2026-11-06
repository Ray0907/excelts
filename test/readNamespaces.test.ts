import { describe, expect, it } from 'vitest';
import { existsSync, readFileSync } from 'node:fs';
import { strToU8, zipSync } from 'fflate';
import { readXlsx, type CellValue } from '../src/index.js';

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const S_MAIN = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
const S_REL = 'http://purl.oclc.org/ooxml/officeDocument/relationships';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const BOM = '﻿';

const pkg = (files: Record<string, string>): Uint8Array => zipSync(Object.fromEntries(Object.entries(files).map(([k, v]) => [k, strToU8(v)])));
const read = (bytes: Uint8Array) => readXlsx(bytes).sheets.map((s) => ({ name: s.name, rows: [...s.rows()] }));
const D = (y: number, m: number, d: number) => new Date(Date.UTC(y, m - 1, d));

const relsXml = (type: string, target: string, id = 'rId1', pkgNs = PKG) =>
  `${XML}<Relationships xmlns="${pkgNs}"><Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`;
const WS_TYPE = `${REL}/worksheet`;

describe('fully prefixed parts (ClosedXML style)', () => {
  const files = {
    'xl/workbook.xml': `${XML}<x:workbook xmlns:x="${MAIN}" xmlns:r="${REL}"><x:workbookPr date1904="0"/><x:sheets><x:sheet name="A &amp; B" sheetId="1" r:id="rId2"/></x:sheets></x:workbook>`,
    'xl/_rels/workbook.xml.rels': `${XML}<rel:Relationships xmlns:rel="${PKG}"><rel:Relationship Id="rId2" Type="${WS_TYPE}" Target="worksheets/sheet1.xml"/><rel:Relationship Id="rId3" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/><rel:Relationship Id="rId4" Type="${REL}/styles" Target="styles.xml"/></rel:Relationships>`,
    'xl/sharedStrings.xml': `${XML}<x:sst xmlns:x="${MAIN}" count="2" uniqueCount="2"><x:si><x:t>hello</x:t></x:si><x:si><x:r><x:t>a</x:t></x:r><x:r><x:t>b</x:t></x:r><x:rPh sb="0" eb="1"><x:t>ignored</x:t></x:rPh></x:si></x:sst>`,
    'xl/styles.xml': `${XML}<x:styleSheet xmlns:x="${MAIN}"><x:numFmts count="1"><x:numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></x:numFmts><x:cellStyleXfs count="1"><x:xf numFmtId="14"/></x:cellStyleXfs><x:cellXfs count="3"><x:xf numFmtId="0"/><x:xf numFmtId="164"/><x:xf numFmtId="14"/></x:cellXfs></x:styleSheet>`,
    'xl/worksheets/sheet1.xml':
      `${XML}<x:worksheet xmlns:x="${MAIN}"><x:cols><x:col min="1" max="1" width="9"/></x:cols><x:sheetData>` +
      `<x:row r="1"><x:c r="A1" t="s"><x:v>0</x:v></x:c><x:c r="B1" t="s"><x:v>1</x:v></x:c><x:c r="C1" t="inlineStr"><x:is><x:t>inl</x:t></x:is></x:c>` +
      `<x:c r="D1" t="inlineStr"><x:is><x:r><x:t>r1</x:t></x:r><x:r><x:t>r2</x:t></x:r></x:is></x:c></x:row>` +
      `<x:row r="2"><x:c r="A2" s="1"><x:v>43831</x:v></x:c><x:c r="B2" s="2"><x:v>43832</x:v></x:c><x:c r="C2"><x:f>1+1</x:f><x:v>2</x:v></x:c><x:c r="D2" t="b"><x:v>1</x:v></x:c><x:c r="E2" t="str"><x:v>s &amp; t</x:v></x:c><x:c r="F2" s="1"/></x:row>` +
      `</x:sheetData><x:conditionalFormatting sqref="A1"><x:cfRule type="expression"/></x:conditionalFormatting></x:worksheet>`,
  };

  it('reads names, shared strings, inline strings, date styles, formulas', () => {
    expect(read(pkg(files))).toEqual([
      {
        name: 'A & B',
        rows: [
          ['hello', 'ab', 'inl', 'r1r2'],
          [D(2020, 1, 1), D(2020, 1, 2), 2, true, 's & t'],
        ],
      },
    ]);
  });
});

describe('relationship id attribute is found by namespace, not by the text r:id', () => {
  const sheetXml = `${XML}<worksheet xmlns="${MAIN}"><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>`;
  const build = (sheetTag: string, wbAttrs = '') =>
    pkg({
      'xl/workbook.xml': `${XML}<workbook xmlns="${MAIN}"${wbAttrs}><sheets>${sheetTag}</sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': relsXml(WS_TYPE, 'worksheets/sheet1.xml', 'rId1'),
      'xl/worksheets/sheet1.xml': sheetXml,
    });
  const want = [{ name: 'Sheet1', rows: [[7]] }];

  it('relationships:id declared on the sheet element', () => {
    expect(read(build(`<sheet name="Sheet1" sheetId="1" xmlns:relationships="${REL}" relationships:id="rId1"/>`))).toEqual(want);
  });
  it('ns:id declared on the sheet element', () => {
    expect(read(build(`<sheet name="Sheet1" sheetId="1" xmlns:ns="${REL}" ns:id="rId1"/>`))).toEqual(want);
  });
  it('prefix declared on the root element', () => {
    expect(read(build(`<sheet name="Sheet1" sheetId="1" q:id="rId1"/>`, ` xmlns:q="${REL}"`))).toEqual(want);
  });
  it('the declaration may follow the attribute that uses it', () => {
    expect(read(build(`<sheet name="Sheet1" ns:id="rId1" sheetId="1" xmlns:ns="${REL}"/>`))).toEqual(want);
  });
  it('strict relationships namespace', () => {
    expect(read(build(`<sheet name="Sheet1" sheetId="1" xmlns:o="${S_REL}" o:id="rId1"/>`))).toEqual(want);
  });
  it('an id attribute bound to another namespace is ignored', () => {
    expect(read(build(`<sheet name="Sheet1" sheetId="1" xmlns:foo="urn:other" foo:id="rIdBad" xmlns:r="${REL}" r:id="rId1"/>`))).toEqual(want);
  });
  it('single-quoted attributes and newlines between attributes', () => {
    expect(read(build(`<sheet\n name='Sheet1'\n sheetId='1'\n xmlns:ns='${REL}'\n ns:id='rId1'/>`))).toEqual(want);
  });
});

describe('UTF-8 BOM at the start of a part', () => {
  it('is tolerated in workbook, rels, sharedStrings, styles and sheet', () => {
    const bytes = pkg({
      '_rels/.rels': BOM + relsXml(`${REL}/officeDocument`, 'xl/workbook.xml'),
      'xl/workbook.xml': `${BOM}${XML}<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': BOM + relsXml(WS_TYPE, 'worksheets/sheet1.xml'),
      'xl/sharedStrings.xml': `${BOM}${XML}<sst xmlns="${MAIN}"><si><t>bom</t></si></sst>`,
      'xl/styles.xml': `${BOM}${XML}<styleSheet xmlns="${MAIN}"><cellXfs><xf numFmtId="14"/></cellXfs></styleSheet>`,
      'xl/worksheets/sheet1.xml': `${BOM}${XML}<worksheet xmlns="${MAIN}"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" s="0"><v>43831</v></c></row></sheetData></worksheet>`,
    });
    expect(read(bytes)).toEqual([{ name: 'S', rows: [['bom', D(2020, 1, 1)]] }]);
  });

  it('is tolerated before a prefixed sheet root', () => {
    const bytes = pkg({
      'xl/workbook.xml': `${XML}<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': relsXml(WS_TYPE, 'worksheets/sheet1.xml'),
      'xl/worksheets/sheet1.xml': `${BOM}<x:worksheet xmlns:x="${MAIN}"><x:sheetData><x:row r="1"><x:c r="A1"><x:v>3</x:v></x:c></x:row></x:sheetData></x:worksheet>`,
    });
    expect(read(bytes)).toEqual([{ name: 'S', rows: [[3]] }]);
  });
});

describe('workbook part and relationship targets', () => {
  const sheet = (v: number) => `${XML}<worksheet xmlns="${MAIN}"><sheetData><row r="1"><c r="A1"><v>${v}</v></c></row></sheetData></worksheet>`;

  it('finds the workbook through _rels/.rels when it is not xl/workbook.xml, with its own sharedStrings and styles', () => {
    const bytes = pkg({
      '_rels/.rels': relsXml(`${REL}/officeDocument`, '/xl/book.xml', 'rId9'),
      'xl/book.xml': `${XML}<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="Book" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/book.xml.rels':
        `${XML}<Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${WS_TYPE}" Target="worksheets/s.xml"/>` +
        `<Relationship Id="rId2" Type="${REL}/sharedStrings" Target="strings/all.xml"/><Relationship Id="rId3" Type="${REL}/styles" Target="look.xml"/></Relationships>`,
      'xl/strings/all.xml': `${XML}<sst xmlns="${MAIN}"><si><t>moved</t></si></sst>`,
      'xl/look.xml': `${XML}<styleSheet xmlns="${MAIN}"><cellXfs><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>`,
      'xl/worksheets/s.xml': `${XML}<worksheet xmlns="${MAIN}"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" s="1"><v>43831</v></c></row></sheetData></worksheet>`,
    });
    expect(read(bytes)).toEqual([{ name: 'Book', rows: [['moved', D(2020, 1, 1)]] }]);
  });

  it('a workbook in a different folder resolves its targets from that folder', () => {
    const bytes = pkg({
      '_rels/.rels': relsXml(`${REL}/officeDocument`, 'docs/main/wb.xml'),
      'docs/main/wb.xml': `${XML}<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="D" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'docs/main/_rels/wb.xml.rels': relsXml(WS_TYPE, '../sheets/one.xml'),
      'docs/sheets/one.xml': sheet(11),
    });
    expect(read(bytes)).toEqual([{ name: 'D', rows: [[11]] }]);
  });

  it('falls back to xl/workbook.xml when _rels/.rels is absent or names no officeDocument', () => {
    const files = {
      'xl/workbook.xml': `${XML}<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="F" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': relsXml(WS_TYPE, 'worksheets/sheet1.xml'),
      'xl/worksheets/sheet1.xml': sheet(1),
    };
    expect(read(pkg(files))).toEqual([{ name: 'F', rows: [[1]] }]);
    expect(read(pkg({ ...files, '_rels/.rels': relsXml(`${REL}/extended-properties`, 'docProps/app.xml') }))).toEqual([{ name: 'F', rows: [[1]] }]);
  });

  it('absolute, ../, backslash and percent-encoded targets', () => {
    const rel = (id: string, target: string) => `<Relationship Id="${id}" Type="${WS_TYPE}" Target="${target}"/>`;
    const sh = (n: string, id: string) => `<sheet name="${n}" sheetId="1" r:id="${id}"/>`;
    const bytes = pkg({
      'xl/workbook.xml': `${XML}<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets>${sh('abs', 'a')}${sh('up', 'b')}${sh('back', 'c')}${sh('pct', 'd')}${sh('dot', 'e')}</sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': `${XML}<Relationships xmlns="${PKG}">${rel('a', '/xl/worksheets/abs.xml')}${rel('b', '../xl/worksheets/up.xml')}${rel('c', 'worksheets\\back.xml')}${rel('d', 'worksheets/with%20space.xml')}${rel('e', './worksheets/dot.xml')}</Relationships>`,
      'xl/worksheets/abs.xml': sheet(1),
      'xl/worksheets/up.xml': sheet(2),
      'xl/worksheets/back.xml': sheet(3),
      'xl/worksheets/with space.xml': sheet(4),
      'xl/worksheets/dot.xml': sheet(5),
    });
    expect(read(bytes).map((s) => [s.name, s.rows[0]![0]])).toEqual([['abs', 1], ['up', 2], ['back', 3], ['pct', 4], ['dot', 5]]);
  });
});

describe('strict OOXML namespaces', () => {
  it('reads a strict-namespace package', () => {
    const bytes = pkg({
      '_rels/.rels': relsXml('http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument', 'xl/workbook.xml', 'rId1', PKG),
      'xl/workbook.xml': `${XML}<workbook xmlns="${S_MAIN}" xmlns:r="${S_REL}"><workbookPr date1904="false"/><sheets><sheet name="Strict" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': relsXml(`${S_REL}/worksheet`, 'worksheets/sheet1.xml'),
      'xl/sharedStrings.xml': `${XML}<sst xmlns="${S_MAIN}"><si><t>strict</t></si></sst>`,
      'xl/styles.xml': `${XML}<styleSheet xmlns="${S_MAIN}"><cellXfs><xf numFmtId="14"/></cellXfs></styleSheet>`,
      'xl/worksheets/sheet1.xml': `${XML}<worksheet xmlns="${S_MAIN}"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" s="0"><v>43831</v></c><c r="C1" t="d"><v>2020-01-02T00:00:00Z</v></c></row></sheetData></worksheet>`,
    });
    expect(read(bytes)).toEqual([{ name: 'Strict', rows: [['strict', D(2020, 1, 1), D(2020, 1, 2)]] }]);
  });

  it('strict namespace with prefixed elements', () => {
    const bytes = pkg({
      'xl/workbook.xml': `${XML}<s:workbook xmlns:s="${S_MAIN}" xmlns:r="${S_REL}"><s:sheets><s:sheet name="SP" sheetId="1" r:id="rId1"/></s:sheets></s:workbook>`,
      'xl/_rels/workbook.xml.rels': relsXml(`${S_REL}/worksheet`, 'worksheets/sheet1.xml'),
      'xl/worksheets/sheet1.xml': `${XML}<s:worksheet xmlns:s="${S_MAIN}"><s:sheetData><s:row r="1"><s:c r="A1"><s:v>9</s:v></s:c></s:row></s:sheetData></s:worksheet>`,
    });
    expect(read(bytes)).toEqual([{ name: 'SP', rows: [[9]] }]);
  });
});

const CORPUS = 'corpus/raw';
// Expected cells were looked up in oracle/out/openpyxl.ndjson (the independent openpyxl reader), not produced by this package.
const firstCells = (rows: CellValue[][], n: number): [number, number, CellValue][] => {
  const out: [number, number, CellValue][] = [];
  for (let r = 0; r < rows.length && out.length < n; r++) {
    const row = rows[r]!;
    for (let c = 0; c < row.length && out.length < n; c++) if (row[c] != null) out.push([r + 1, c + 1, row[c]]);
  }
  return out;
};

describe('real files that other readers open (skipped when corpus/raw is absent)', () => {
  const closedXml = `${CORPUS}/closedxml/ClosedXML.Tests_Resource_Examples_AutoFilter_CustomAutoFilter.xlsx`;
  it.skipIf(!existsSync(closedXml))('ClosedXML fully prefixed workbook', () => {
    const sheets = readXlsx(readFileSync(closedXml)).sheets;
    expect(sheets.map((s) => s.name)).toEqual(['Single Column Numbers', 'Single Column Strings', 'Single Column Mixed', 'Multi Column']);
    expect(firstCells([...sheets[0]!.rows()], 3)).toEqual([[1, 1, 'Numbers'], [2, 1, 5], [3, 1, 4]]);
    expect(firstCells([...sheets[1]!.rows()], 3)).toEqual([[1, 1, 'Strings'], [2, 1, 'E'], [3, 1, 'D']]);
  });

  const nsPrefix = `${CORPUS}/calamine/tests_nonstandard-xml-ns-prefix.xlsx`;
  it.skipIf(!existsSync(nsPrefix))('calamine ns:id workbook', () => {
    const sheets = readXlsx(readFileSync(nsPrefix)).sheets;
    expect(sheets.map((s) => s.name)).toEqual(['Sheet1']);
    expect(firstCells([...sheets[0]!.rows()], 3)).toEqual([[1, 1, 'a'], [1, 2, 'b'], [2, 1, 1]]);
  });

  const simpleRange = `${CORPUS}/calamine/tests_simple_range.xlsm`;
  it.skipIf(!existsSync(simpleRange))('calamine relationships:id workbook', () => {
    const sheets = readXlsx(readFileSync(simpleRange)).sheets;
    expect(sheets.map((s) => s.name)).toEqual(['Sheet1']);
    expect(firstCells([...sheets[0]!.rows()], 3)).toEqual([[1, 1, 'Cell_A1'], [1, 2, 'Cell_B1'], [2, 1, 'Cell_A2']]);
  });
});
