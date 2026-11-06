import { zipBuild } from './lib/zip.mjs';
import { NS, XML, sheetXml, skeleton } from './lib/skeleton.mjs';
import { BASE_BODY, WS_PRE, WS_SUF, base, bigPart, inline, laughs, mb, noExpansion, patchEocd, patchFirstCd, rep, sst } from './attacks-lib.mjs';

export function xmlAttacks({ bombMb = 1024 } = {}) {
  const A = [];
  const add = (id, group, note, build, extra = {}) => A.push({ id, group, note, build, ...extra });
  const bomb = mb(bombMb);

  add('rows-10m-empty', 'rows', '10M <row/> elements', () => skeleton({ sheets: [{ name: 'S', body: rep('<row/>', 10_000_000).toString() }] }));
  add('rows-10m-numbered', 'rows', '10M <row r="N"/> elements (more than the 1048576 Excel limit)', () => {
    const parts = [];
    for (let i = 1; i <= 10_000_000; i += 1) parts.push(`<row r="${i}"/>`);
    return skeleton({ sheets: [{ name: 'S', body: parts.join('') }] });
  });
  add('row-jump-1048576', 'rows', 'single row r=1048576 (API pads 1048575 empty rows)', () => skeleton({ sheets: [{ name: 'S', body: '<row r="1048576"><c r="A1"><v>1</v></c></row>' }] }));
  add('row-jump-1048576-x50-sheets', 'rows', '50 sheets, each one row at r=1048576', () => skeleton({ sheets: Array.from({ length: 50 }, (_, i) => ({ name: `S${i}`, body: '<row r="1048576"/>' })) }));
  add('row-jump-2^32', 'rows', 'row r=4294967296', () => skeleton({ sheets: [{ name: 'S', body: '<row r="4294967296"/>' }] }));
  add('row-jump-2^31', 'rows', 'row r=2147483648', () => skeleton({ sheets: [{ name: 'S', body: '<row r="2147483648"/>' }] }));
  add('row-jump-1e19', 'rows', 'row r=9999999999999999999 (beyond 2^53)', () => skeleton({ sheets: [{ name: 'S', body: '<row r="9999999999999999999"/>' }] }));
  add('row-negative-and-zero', 'rows', 'row r=-5, r=0, r=abc', () => skeleton({ sheets: [{ name: 'S', body: '<row r="-5"><c><v>1</v></c></row><row r="0"><c><v>2</v></c></row><row r="abc"><c><v>3</v></c></row>' }] }));

  add('cell-XFD1048576', 'cells', 'single cell at the last Excel position', () => skeleton({ sheets: [{ name: 'S', body: '<row r="1048576"><c r="XFD1048576"><v>1</v></c></row>' }] }));
  add('cell-col-ZZZZZ', 'cells', 'cell r=ZZZZZ1 (column 12M)', () => skeleton({ sheets: [{ name: 'S', body: '<row r="1"><c r="ZZZZZ1"><v>1</v></c></row>' }] }));
  add('cell-col-ZZZZZZZZ', 'cells', 'cell r=ZZZZZZZZ1 (column 2e11)', () => skeleton({ sheets: [{ name: 'S', body: '<row r="1"><c r="ZZZZZZZZ1"><v>1</v></c></row>' }] }));
  add('cell-ref-ZZZZZZ9999999999', 'cells', 'cell r=ZZZZZZ9999999999', () => skeleton({ sheets: [{ name: 'S', body: '<row r="1"><c r="ZZZZZZ9999999999"><v>1</v></c></row>' }] }));
  add('cell-row-1m-cells', 'cells', 'one row of 2M cells without r', () => skeleton({ sheets: [{ name: 'S', body: `<row r="1">${rep('<c><v>1</v></c>', 2_000_000)}</row>` }] }));
  add('cell-10m-in-rows', 'cells', '10M numeric cells over 100k rows', () => {
    const row = `<row>${rep('<c><v>1</v></c>', 100)}</row>`;
    return skeleton({ sheets: [{ name: 'S', body: row.repeat(100_000) }] });
  });
  add('cell-style-oob', 'cells', 's="4294967296", s="-1", s="99999999" with a styles part', () =>
    skeleton({ sheets: [{ name: 'S', body: '<row><c s="4294967296"><v>1</v></c><c s="-1"><v>1</v></c><c s="99999999"><v>1</v></c></row>' }], styles: `${XML}<styleSheet ${NS}><cellXfs count="1"><xf numFmtId="14"/></cellXfs></styleSheet>` }));
  add('styles-5m-xf', 'cells', '5M <xf/> entries in cellXfs', () => skeleton({ sheets: [{ name: 'S', body: BASE_BODY }], styles: `${XML}<styleSheet ${NS}><cellXfs count="5000000">${rep('<xf numFmtId="14"/>', 5_000_000)}</cellXfs></styleSheet>` }));

  add('xml-nest-is-100k', 'xml', '100000 nested <is> in one cell', () => skeleton({ sheets: [{ name: 'S', body: `<row r="1"><c t="inlineStr">${'<is>'.repeat(100000)}<t>x</t>${'</is>'.repeat(100000)}</c></row>` }] }));
  add('xml-nest-1m-tags', 'xml', '1M nested unknown tags in sheetData', () => skeleton({ sheets: [{ name: 'S', body: `${rep('<a>', 1_000_000)}${rep('</a>', 1_000_000)}` }] }));
  add('xml-10mb-text', 'xml', '10 MB single text node in an inline string and in <v>', () => skeleton({ sheets: [{ name: 'S', body: `<row r="1">${inline('A'.repeat(mb(10)))}<c t="str"><v>${'B'.repeat(mb(10))}</v></c></row>` }] }));
  add('xml-100mb-text', 'xml', '100 MB inline string', () => skeleton({ sheets: [{ name: 'S', part: bigPart('is100', `${WS_PRE}<row r="1"><c t="inlineStr"><is><t>`, 65, mb(100), `</t></is></c></row>${WS_SUF}`) }] }));
  add('xml-1m-attrs-on-row', 'xml', '1M attributes on a <row> and a <c>', () => {
    const attrs = Array.from({ length: 1_000_000 }, (_, i) => `a${i}="1"`).join(' ');
    return skeleton({ sheets: [{ name: 'S', body: `<row ${attrs}><c ${attrs}><v>1</v></c></row>` }] });
  });
  add('xml-10mb-attr-value', 'xml', '10 MB attribute value on <row> and <c>', () => skeleton({ sheets: [{ name: 'S', body: `<row r="1" x="${'A'.repeat(mb(10))}"><c r="A1" x="${'B'.repeat(mb(10))}"><v>1</v></c></row>` }] }));
  add('xml-unterminated-cdata', 'xml', 'unterminated CDATA in <v>', () => skeleton({ sheets: [{ name: 'S', body: '<row r="1"><c r="A1"><v><![CDATA[never closed</v></c></row>' }] }));
  add('xml-prefixed-50mb', 'xml', '50 MB sheet using an x: prefix (the reader rewrites it)', () => {
    const rows = `<x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row>`.repeat(Math.floor(mb(50) / 50));
    return skeleton({ sheets: [{ name: 'S', xml: `${XML}<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData>${rows}</x:sheetData></x:worksheet>` }] });
  });

  add('quad-empty-cells', 'quadratic', '300k <c r="A1" s="1"></c> (not self-closing, no <v> anywhere)', () => skeleton({ sheets: [{ name: 'S', body: `<row r="1">${rep('<c s="1"></c>', 300_000)}</row>` }] }));
  add('quad-interleaved-row', 'quadratic', '300k <c/> each followed by an unknown element in one row', () => skeleton({ sheets: [{ name: 'S', body: `<row r="1">${rep('<c/><x/>', 300_000)}</row>` }] }));
  add('quad-inline-no-is', 'quadratic', '300k inlineStr cells with neither <is> nor <v>', () => skeleton({ sheets: [{ name: 'S', body: `<row r="1">${rep('<c t="inlineStr"></c>', 300_000)}</row>` }] }));
  add('quad-v-unclosed', 'quadratic', '300k cells whose <v> never closes', () => skeleton({ sheets: [{ name: 'S', body: `<row r="1">${rep('<c t="str"><v>x</c>', 300_000)}</row>` }] }));
  add('quad-rows-unclosed', 'quadratic', '300k <row> with a cell and no </row>', () => skeleton({ sheets: [{ name: 'S', body: rep('<row><c><v>1</v></c>', 300_000).toString() }] }));

  add('workbook-100k-sheets', 'workbook', '100k <sheet> entries all pointing at one tiny part', () => {
    const sheets = Array.from({ length: 100_000 }, (_, i) => `<sheet name="s${i}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('');
    const rels = Array.from({ length: 100_000 }, (_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>`).join('');
    return skeleton({
      sheets: [{ name: 'S', body: BASE_BODY }],
      workbookXml: `${XML}<workbook ${NS} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets}</sheets></workbook>`,
      relsXml: `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`,
    });
  });
  add('workbook-1m-attrs-on-sheet', 'workbook', '1M attributes on the <sheet> tag', () => {
    const attrs = Array.from({ length: 1_000_000 }, (_, i) => `a${i}="1"`).join(' ');
    return skeleton({ sheets: [{ name: 'S', body: BASE_BODY }], workbookXml: `${XML}<workbook ${NS} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet ${attrs} name="S" sheetId="1" r:id="rId1"/></sheets></workbook>` });
  });
  add('workbook-100mb-comment', 'workbook', '100 MB XML comment inside workbook.xml', () => {
    const wb = bigPart('wbc', `${XML}<workbook ${NS} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><!--`, 65, mb(100), '--><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>');
    return zipBuild([
      { name: 'xl/workbook.xml', ...wb },
      { name: 'xl/_rels/workbook.xml.rels', data: `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>` },
      { name: 'xl/worksheets/sheet1.xml', data: sheetXml(BASE_BODY) },
    ]);
  });
  add('workbook-attr-name-run', 'regex', '200 KB run of name characters inside the <sheet> tag (attribute regex backtracks at every start)', () =>
    skeleton({ sheets: [{ name: 'S', body: BASE_BODY }], workbookXml: `${XML}<workbook ${NS} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1" ${'A'.repeat(200_000)}/></sheets></workbook>` }));
  add('styles-numfmt-brackets', 'regex', 'numFmt formatCode of 200k "[" with no "]"', () =>
    skeleton({ sheets: [{ name: 'S', body: BASE_BODY }], styles: `${XML}<styleSheet ${NS}><numFmts count="1"><numFmt numFmtId="164" formatCode="${'['.repeat(200_000)}"/></numFmts><cellXfs count="1"><xf numFmtId="164"/></cellXfs></styleSheet>` }));
  add('styles-numfmt-quotes', 'regex', 'numFmt formatCode of 200k unmatched quote characters', () =>
    skeleton({ sheets: [{ name: 'S', body: BASE_BODY }], styles: `${XML}<styleSheet ${NS}><numFmts count="1"><numFmt numFmtId="164" formatCode="${'&quot;'.repeat(200_000)}"/></numFmts><cellXfs count="1"><xf numFmtId="164"/></cellXfs></styleSheet>` }));
  add('sst-rph-unclosed', 'regex', 'one shared string holding 100k <rPh> with no </rPh>', () =>
    skeleton({ sheets: [{ name: 'S', body: '<row r="1"><c t="s"><v>0</v></c></row>' }], sst: sst(`<si><t>x</t>${'<rPh>'.repeat(100_000)}</si>`) }));
  add('workbook-duplicate-parts', 'workbook', 'two xl/workbook.xml entries (first-wins vs last-wins differential)', () =>
    skeleton({ sheets: [{ name: 'S', body: BASE_BODY }], extra: [{ name: 'xl/workbook.xml', data: `${XML}<workbook ${NS} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="EVIL" sheetId="1" r:id="rId1"/></sheets></workbook>` }] }));
  add('workbook-duplicate-sheets', 'workbook', 'two sheet entries with the same name and id', () =>
    skeleton({ sheets: [{ name: 'S', body: BASE_BODY }, { name: 'S', body: BASE_BODY }] }));
  add('workbook-no-sheets', 'workbook', 'workbook with an empty <sheets/>', () => skeleton({ sheets: [{ name: 'S', body: '' }], workbookXml: `${XML}<workbook ${NS}><sheets/></workbook>` }));

  return A;
}
