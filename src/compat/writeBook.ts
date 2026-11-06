import { strToU8 } from 'fflate';
import { SharedStrings } from '../sharedStrings.js';
import { escapeXml } from '../xml.js';
import { concat, defaultCompressor, levelFor, writeZip, type Compressor, type ZipEntry } from '../zip.js';
import { NotSupportedError } from './errors.js';
import { NS_MAIN, NS_REL, SheetWriter, type WriteCtx } from './writeSheet.js';
import type { Workbook } from './workbook.js';
import type { Worksheet } from './worksheet.js';

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml';

export interface WriteOptions {
  zip?: { compression?: 'STORE' | 'DEFLATE'; compressionOptions?: { level?: number } };
  /** Drop features of a loaded file that cannot be written back instead of throwing. */
  ignoreUnsupported?: boolean;
}

export function levelOf(options: WriteOptions | undefined, compressor: Compressor): number {
  const z = options?.zip;
  if (z?.compression === 'STORE') return 0;
  const l = z?.compressionOptions?.level;
  return typeof l === 'number' ? Math.max(0, Math.min(9, l)) : levelFor(compressor);
}

export function assertWritable(wb: Workbook, options: WriteOptions | undefined): void {
  if (options?.ignoreUnsupported) return;
  const found = new Set<string>(wb._unsupported);
  for (const ws of wb._worksheets) if (ws) for (const f of ws._unsupported) found.add(f);
  if (found.size) {
    throw new NotSupportedError(
      `saving a loaded file that uses ${[...found].join(', ')}`,
      'these parts would be lost; pass { ignoreUnsupported: true } to write without them',
    );
  }
}

const dateXml = (tag: string, d: Date | undefined): string => (d ? `<${tag} xsi:type="dcterms:W3CDTF">${d.toISOString().replace(/\.\d{3}Z$/, 'Z')}</${tag}>` : '');
const textXml = (tag: string, v: unknown): string => (v === undefined || v === null || v === '' ? '' : `<${tag}>${escapeXml(String(v))}</${tag}>`);

function coreXml(wb: Workbook): string {
  return (
    DECL +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" ' +
    'xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    textXml('dc:creator', wb.creator) + textXml('cp:lastModifiedBy', wb.lastModifiedBy) +
    dateXml('dcterms:created', wb.created) + dateXml('dcterms:modified', wb.modified) +
    textXml('dc:title', wb.title) + textXml('dc:subject', wb.subject) + textXml('cp:keywords', wb.keywords) +
    textXml('cp:category', wb.category) + textXml('dc:description', wb.description) + textXml('cp:contentStatus', wb.contentStatus) +
    textXml('dc:language', wb.language) + textXml('cp:revision', wb.revision) +
    (wb.lastPrinted ? dateXml('cp:lastPrinted', wb.lastPrinted) : '') +
    '</cp:coreProperties>'
  );
}

function appXml(wb: Workbook, sheets: Worksheet[]): string {
  return (
    DECL +
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
    '<Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>' +
    `<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>${sheets.length}</vt:i4></vt:variant></vt:vector></HeadingPairs>` +
    `<TitlesOfParts><vt:vector size="${sheets.length}" baseType="lpstr">${sheets.map((s) => `<vt:lpstr>${escapeXml(s.name)}</vt:lpstr>`).join('')}</vt:vector></TitlesOfParts>` +
    textXml('Company', wb.company) + textXml('Manager', wb.manager) +
    '<LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>'
  );
}

function quoteSheet(name: string): string {
  return `'${name.replace(/'/g, "''")}'`;
}

function definedNamesXml(wb: Workbook, sheets: Worksheet[]): string {
  let s = '';
  const names = wb._names;
  if (names) {
    for (const d of names.model) s += `<definedName name="${escapeXml(d.name)}">${escapeXml(d.ranges.join(','))}</definedName>`;
    for (const r of names.raw) {
      const at = r.scope ? sheets.indexOf(r.scope) : -1;
      if (r.scope && at < 0) continue;
      let a = '';
      for (const [k, v] of Object.entries(r.attrs)) if (/^[A-Za-z]+$/.test(k)) a += ` ${k}="${escapeXml(v)}"`;
      s += `<definedName name="${escapeXml(r.name)}"${r.scope ? ` localSheetId="${at}"` : ''}${a}>${escapeXml(r.text)}</definedName>`;
    }
  }
  sheets.forEach((ws, i) => {
    const q = quoteSheet(ws.name);
    const ps = ws.pageSetup;
    const abs = (r: string): string => r.split(':').map((p) => p.replace(/^\$?([A-Za-z]*)\$?(\d*)$/, (_m, c: string, n: string) => `${c ? '$' + c : ''}${n ? '$' + n : ''}`)).join(':');
    if (ps.printArea) {
      const areas = String(ps.printArea).split('&&').map((a) => `${q}!${abs(a)}`);
      s += `<definedName name="_xlnm.Print_Area" localSheetId="${i}">${escapeXml(areas.join(','))}</definedName>`;
    }
    const titles: string[] = [];
    if (ps.printTitlesColumn) titles.push(`${q}!${abs(String(ps.printTitlesColumn))}`);
    if (ps.printTitlesRow) titles.push(`${q}!${abs(String(ps.printTitlesRow))}`);
    if (titles.length) s += `<definedName name="_xlnm.Print_Titles" localSheetId="${i}">${escapeXml(titles.join(','))}</definedName>`;
  });
  return s ? `<definedNames>${s}</definedNames>` : '';
}

function bookViewsXml(wb: Workbook): string {
  const views = wb.views.length ? wb.views : [{ x: 0, y: 0, width: 12000, height: 24000, firstSheet: 0, activeTab: 0 }];
  const map: [string, string][] = [['x', 'xWindow'], ['y', 'yWindow'], ['width', 'windowWidth'], ['height', 'windowHeight'], ['firstSheet', 'firstSheet'], ['activeTab', 'activeTab'], ['visibility', 'visibility'], ['tabRatio', 'tabRatio']];
  const bools: [string, string][] = [['showHorizontalScroll', 'showHorizontalScroll'], ['showVerticalScroll', 'showVerticalScroll'], ['showSheetTabs', 'showSheetTabs'], ['minimized', 'minimized'], ['autoFilterDateGrouping', 'autoFilterDateGrouping']];
  return (
    '<bookViews>' +
    views
      .map((v) => {
        const r = v as Record<string, unknown>;
        let a = '';
        for (const [k, x] of map) if (r[k] !== undefined) a += ` ${x}="${escapeXml(String(r[k]))}"`;
        for (const [k, x] of bools) if (r[k] !== undefined) a += ` ${x}="${r[k] ? 1 : 0}"`;
        return `<workbookView${a}/>`;
      })
      .join('') +
    '</bookViews>'
  );
}

/** The package for a workbook, written straight from its rows. */
export function writeWorkbook(wb: Workbook, options?: WriteOptions, compressor: Compressor = defaultCompressor): Uint8Array {
  assertWritable(wb, options);
  const sheets = wb.worksheets;
  const ctx: WriteCtx = { strings: new SharedStrings(), book: wb._styles, date1904: !!wb.properties.date1904 };
  const parts: ZipEntry[] = [];
  const rels: string[] = [];
  const overrides: string[] = [];
  sheets.forEach((ws, i) => {
    const n = i + 1;
    const w = new SheetWriter(ws, ctx);
    w.writeAll();
    parts.push({ name: `xl/worksheets/sheet${n}.xml`, data: concat(w.chunks) });
    const r = w.relsXml();
    if (r) parts.push({ name: `xl/worksheets/_rels/sheet${n}.xml.rels`, data: strToU8(r) });
    rels.push(`<Relationship Id="rId${n}" Type="${NS_REL}/worksheet" Target="worksheets/sheet${n}.xml"/>`);
    overrides.push(`<Override PartName="/xl/worksheets/sheet${n}.xml" ContentType="${CT}.worksheet+xml"/>`);
  });
  collectExtras(wb, ctx, parts, overrides);
  return assemble(wb, sheets, ctx, parts, rels, overrides, options, compressor);
}

/** Adds the parts that sheets wrote beside themselves (pictures, notes) once every sheet is finished. */
export function collectExtras(wb: Workbook, ctx: WriteCtx, parts: ZipEntry[], overrides: string[]): void {
  if (!ctx.extras) return;
  parts.push(...ctx.extras.finish(wb));
  overrides.push(...ctx.extras.contentTypes());
}

/** Shared by the document and streaming writers once the sheet parts exist. */
export function assemble(
  wb: Workbook, sheets: Worksheet[], ctx: WriteCtx, sheetParts: ZipEntry[], sheetRels: string[], overrides: string[],
  options: WriteOptions | undefined, compressor: Compressor,
): Uint8Array {
  const n = sheets.length;
  const stylesId = n + 1, sstId = n + 2;
  const wbXml =
    `${DECL}<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">` +
    `<workbookPr${wb.properties.date1904 ? ' date1904="1"' : ''}/>` + bookViewsXml(wb) +
    '<sheets>' +
    sheets.map((ws, i) => `<sheet name="${escapeXml(ws.name)}" sheetId="${ws.id}"${ws.state !== 'visible' ? ` state="${ws.state}"` : ''} r:id="rId${i + 1}"/>`).join('') +
    '</sheets>' + definedNamesXml(wb, sheets) +
    `<calcPr calcId="171027"${wb.calcProperties.fullCalcOnLoad === false ? '' : ' fullCalcOnLoad="1"'}/></workbook>`;
  const files: ZipEntry[] = [
    {
      name: '[Content_Types].xml',
      data: strToU8(
        `${DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
          `<Override PartName="/xl/workbook.xml" ContentType="${CT}.sheet.main+xml"/>` + overrides.join('') +
          `<Override PartName="/xl/styles.xml" ContentType="${CT}.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="${CT}.sharedStrings+xml"/>` +
          '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
          '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>',
      ),
    },
    {
      name: '_rels/.rels',
      data: strToU8(
        `${DECL}<Relationships xmlns="${NS_PKG_REL}"><Relationship Id="rId1" Type="${NS_REL}/officeDocument" Target="xl/workbook.xml"/>` +
          '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
          `<Relationship Id="rId3" Type="${NS_REL}/extended-properties" Target="docProps/app.xml"/></Relationships>`,
      ),
    },
    { name: 'xl/workbook.xml', data: strToU8(wbXml) },
    {
      name: 'xl/_rels/workbook.xml.rels',
      data: strToU8(
        `${DECL}<Relationships xmlns="${NS_PKG_REL}">${sheetRels.join('')}<Relationship Id="rId${stylesId}" Type="${NS_REL}/styles" Target="styles.xml"/>` +
          `<Relationship Id="rId${sstId}" Type="${NS_REL}/sharedStrings" Target="sharedStrings.xml"/></Relationships>`,
      ),
    },
    ...sheetParts,
    { name: 'xl/styles.xml', data: strToU8(ctx.book.toXml()) },
    { name: 'xl/sharedStrings.xml', data: ctx.strings.toBytes() },
    { name: 'docProps/core.xml', data: strToU8(coreXml(wb)) },
    { name: 'docProps/app.xml', data: strToU8(appXml(wb, sheets)) },
  ];
  return writeZip(files, compressor, levelOf(options, compressor));
}
