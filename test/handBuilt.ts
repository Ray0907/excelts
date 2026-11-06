import { strToU8, zipSync } from 'fflate';

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

export function workbookOf(sheetData: string, opts: { styles?: string; sst?: string; date1904?: boolean } = {}): Uint8Array {
  const files: Record<string, Uint8Array> = {
    'xl/workbook.xml': strToU8(
      `${XML}<workbook xmlns="${MAIN}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${opts.date1904 ? '<workbookPr date1904="1"/>' : ''}<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/s1.xml"/></Relationships>`,
    ),
    'xl/worksheets/s1.xml': strToU8(`${XML}<worksheet xmlns="${MAIN}"><sheetData>${sheetData}</sheetData></worksheet>`),
  };
  if (opts.styles) files['xl/styles.xml'] = strToU8(opts.styles);
  if (opts.sst) files['xl/sharedStrings.xml'] = strToU8(opts.sst);
  return zipSync(files);
}

/** A styles part with the given custom formats (ids from 164) and one cellXfs entry per format reference: a built-in id or a custom code. */
export function stylesOf(formats: readonly (number | string)[]): string {
  const custom = formats.filter((f): f is string => typeof f === 'string');
  const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
  const numFmts = custom.length ? `<numFmts count="${custom.length}">${custom.map((c, i) => `<numFmt numFmtId="${164 + i}" formatCode="${esc(c)}"/>`).join('')}</numFmts>` : '';
  const xfs = formats.map((f) => `<xf numFmtId="${typeof f === 'number' ? f : 164 + custom.indexOf(f)}"/>`).join('');
  return `${XML}<styleSheet xmlns="${MAIN}">${numFmts}<cellXfs count="${formats.length}">${xfs}</cellXfs></styleSheet>`;
}
