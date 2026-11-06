import { zipBuild } from './zip.mjs';

export const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
export const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const NS_R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL_WS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';

export const sheetXml = (body) => `${XML}<worksheet ${NS}><sheetData>${body}</sheetData></worksheet>`;

/**
 * A minimal xlsx whose parts are passed as strings, Buffers or { comp, usize } (already deflated).
 * sheets: [{ name, body | xml | part }]; extra: additional zip entries.
 */
export function skeleton({ sheets, sst, styles, workbookXml, relsXml, extra = [], zipOpts } = {}) {
  sheets ??= [{ name: 'S', body: '' }];
  const entries = [
    { name: '[Content_Types].xml', data: `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>` },
    { name: '_rels/.rels', data: `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>` },
    {
      name: 'xl/workbook.xml',
      data: workbookXml ?? `${XML}<workbook ${NS} ${NS_R}><sheets>${sheets.map((s, i) => `<sheet name="${s.name}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>`,
    },
    {
      name: 'xl/_rels/workbook.xml.rels',
      data:
        relsXml ??
        `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="${REL_WS}" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}${sst !== undefined ? `<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>` : ''}${styles !== undefined ? `<Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` : ''}</Relationships>`,
    },
  ];
  const asEntry = (name, part) => (part && part.comp ? { name, comp: part.comp, usize: part.usize, ...part.entry } : { name, data: part });
  sheets.forEach((s, i) => entries.push(asEntry(`xl/worksheets/sheet${i + 1}.xml`, s.part ?? s.xml ?? sheetXml(s.body ?? ''))));
  if (sst !== undefined) entries.push(asEntry('xl/sharedStrings.xml', sst));
  if (styles !== undefined) entries.push(asEntry('xl/styles.xml', styles));
  entries.push(...extra);
  return zipBuild(entries, zipOpts);
}
