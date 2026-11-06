import { existsSync, readFileSync } from 'node:fs';
import path from 'node:path';
import { ROOT } from './harness.mjs';
import { zipBuild } from './lib/zip.mjs';
import { NS, XML, skeleton } from './lib/skeleton.mjs';
import { BASE_BODY, WS_PRE, WS_SUF, base, bigPart, inline, laughs, mb, noExpansion, patchFirstCd, rep, sst } from './attacks-lib.mjs';

export function bombAttacks({ bombMb = 1024 } = {}) {
  const A = [];
  const add = (id, group, note, build, extra = {}) => A.push({ id, group, note, build, ...extra });
  const bomb = mb(bombMb);

  add('bomb-sheet', 'zip-bomb', `${bombMb} MB sheet part of spaces`, () => skeleton({ sheets: [{ name: 'S', part: bigPart('sheet', WS_PRE, 32, bomb, WS_SUF) }] }));
  add('bomb-sharedstrings', 'zip-bomb', `${bombMb} MB sharedStrings part`, () => skeleton({ sheets: [{ name: 'S', body: BASE_BODY }], sst: bigPart('sst', `${XML}<sst ${NS}>`, 32, bomb, '</sst>') }));
  add('bomb-styles', 'zip-bomb', `${bombMb} MB styles part`, () => skeleton({ sheets: [{ name: 'S', body: '<row r="1"><c r="A1"><v>1</v></c></row>' }], styles: bigPart('styles', `${XML}<styleSheet ${NS}>`, 32, bomb, '</styleSheet>') }));
  add('bomb-workbook', 'zip-bomb', '256 MB workbook.xml of spaces', () => {
    const big = bigPart('wb', `${XML}<workbook ${NS}><sheets>`, 32, mb(256), '</sheets></workbook>');
    return zipBuild([
      { name: 'xl/workbook.xml', comp: big.comp, usize: big.usize },
    ]);
  });
  add('bomb-unreferenced', 'zip-bomb', `${bombMb} MB part nothing references (must stay uninflated)`, () => skeleton({ sheets: [{ name: 'S', body: BASE_BODY }], extra: [{ name: 'xl/media/bomb.bin', ...bigPart('sheet', WS_PRE, 32, bomb, WS_SUF) }] }));
  add('bomb-8-sheets', 'zip-bomb', '8 sheets x 256 MB each, 2 GB in total across entries', () => {
    const part = bigPart('sheet256', WS_PRE, 32, mb(256), WS_SUF);
    return skeleton({ sheets: Array.from({ length: 8 }, (_, i) => ({ name: `S${i}`, part })) });
  });
  add('bomb-overlap-600', 'zip-bomb', '600 sheets whose central entries all point at one 128 MB local record (75 GB total, each part only 128 MB)', () => {
    const part = bigPart('sheet128', WS_PRE, 32, mb(128), WS_SUF);
    const first = { comp: part.comp, usize: part.usize };
    return skeleton({ sheets: Array.from({ length: 600 }, (_, i) => ({ name: `S${i}`, part: i === 0 ? first : { comp: Buffer.alloc(0), usize: part.usize, entry: { pointTo: 4 } } })) });
  });
  add('bomb-usize-lie', 'zip-bomb', 'central directory claims 4 GB uncompressed for a tiny real stream', () => {
    const b = base();
    return patchFirstCd(b, (buf, at) => buf.writeUInt32LE(0xfffffffe, at + 24));
  });
  add('bomb-csize-lie', 'zip-bomb', 'compressed size claims 4 GB', () => patchFirstCd(base(), (buf, at) => buf.writeUInt32LE(0xfffffffe, at + 20)));

  const poc = path.join(ROOT, 'corpus/raw/poi/test-data_spreadsheet_poc-shared-strings.xlsx');
  add('sst-corpus-poc', 'sst-bomb', 'real PoC from the corpus', () => (existsSync(poc) ? readFileSync(poc) : null));
  add('sst-3m-entries', 'sst-bomb', '3M tiny <si>, ~40 MB inflated', () => skeleton({ sheets: [{ name: 'S', body: '<row r="1"><c r="A1" t="s"><v>2999999</v></c></row>' }], sst: sst(rep('<si><t>x</t></si>', 3_000_000)) }));
  add('sst-count-lie', 'sst-bomb', 'count/uniqueCount claim 4294967295', () => skeleton({ sheets: [{ name: 'S', body: '<row r="1"><c r="A1" t="s"><v>0</v></c></row>' }], sst: sst('<si><t>x</t></si>', ' count="4294967295" uniqueCount="4294967295"') }));
  add('sst-100mb-string-x1m-refs', 'sst-bomb', 'one 100 MB shared string referenced by 1M cells (each reference must stay a pointer)', () =>
    skeleton({ sheets: [{ name: 'S', body: `<row r="1">${rep('<c t="s"><v>0</v></c>', 1_000_000).toString()}</row>` }], sst: sst(`<si><t>${'a'.repeat(mb(100))}</t></si>`) }));
  add('sst-index-oob', 'sst-bomb', 'shared string indices -1, 2^32, 1e999 and beyond the table', () => skeleton({ sheets: [{ name: 'S', body: '<row r="1"><c t="s"><v>-1</v></c><c t="s"><v>4294967296</v></c><c t="s"><v>1e999</v></c><c t="s"><v>99</v></c></row>' }], sst: sst('<si><t>x</t></si>') }));

  add('laughs-sheet-inline', 'entity', 'DOCTYPE with nested entities, referenced from an inline string', () => skeleton({ sheets: [{ name: 'S', xml: `${XML}${laughs}<worksheet ${NS}><sheetData><row r="1">${inline('&lol9;')}</row></sheetData></worksheet>` }] }), { check: noExpansion });
  add('laughs-sharedstrings', 'entity', 'DOCTYPE with nested entities in sharedStrings', () => skeleton({ sheets: [{ name: 'S', body: '<row r="1"><c t="s"><v>0</v></c></row>' }], sst: `${XML}${laughs}<sst ${NS}><si><t>&lol9;</t></si></sst>` }), { check: noExpansion });
  add('laughs-workbook', 'entity', 'DOCTYPE in workbook.xml, entity in a sheet name', () =>
    skeleton({ sheets: [{ name: 'S', body: BASE_BODY }], workbookXml: `${XML}${laughs}<workbook ${NS} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="&lol9;" sheetId="1" r:id="rId1"/></sheets></workbook>` }), { check: noExpansion });
  add('laughs-styles', 'entity', 'DOCTYPE in styles, entity in a numFmt code', () =>
    skeleton({ sheets: [{ name: 'S', body: BASE_BODY }], styles: `${XML}${laughs}<styleSheet ${NS}><numFmts count="1"><numFmt numFmtId="164" formatCode="&lol9;"/></numFmts><cellXfs count="1"><xf numFmtId="164"/></cellXfs></styleSheet>` }), { check: noExpansion });
  add('entity-xxe-external', 'entity', 'SYSTEM entity naming a local file must stay unresolved', () =>
    skeleton({ sheets: [{ name: 'S', xml: `${XML}<!DOCTYPE x [<!ENTITY xxe SYSTEM "file:///etc/passwd">]><worksheet ${NS}><sheetData><row r="1">${inline('&xxe;')}</row></sheetData></worksheet>` }] }), { check: (r) => (r.maxStr > 200 ? 'external entity content appeared' : undefined) });
  add('entity-numeric-bad', 'entity', 'numeric references &#xFFFFFFFF; &#x110000; &#xD800; &#0; in text', () =>
    skeleton({ sheets: [{ name: 'S', body: `<row r="1">${inline('&#xFFFFFFFF;')}${inline('&#x110000;')}${inline('&#xD800;')}${inline('&#0;')}${inline('&#99999999999999999999;')}</row>` }] }));
  add('entity-10m-amp', 'entity', '20 MB of &amp; in one inline string', () => skeleton({ sheets: [{ name: 'S', body: `<row r="1">${inline('&amp;'.repeat(4_000_000))}</row>` }] }));

  return A;
}
