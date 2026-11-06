import { resolveLimits, XlsxLimitError, type ReadOptions } from '../limits.js';
import { rootPrefixOf, stripPrefix } from '../dialect.js';
import { unzip } from '../unzip.js';
import { readWorkbookParts } from '../workbookParts.js';
import { defaultCompressor } from '../zip.js';
import { loadSheetData } from './loadSheet.js';
import { loadSheetExtras, readRels } from './loadExtras.js';
import { loadSheetHead, loadSheetTail } from './loadParts.js';
import { loadSharedStrings } from './loadStrings.js';
import { loadStyleBook } from './styleLoad.js';
import { StyleBook } from './styleBook.js';
import { child, kids, parseXml, textOf } from './xmlTree.js';
import type { Workbook } from './workbook.js';
import { Worksheet } from './worksheet.js';

export interface LoadOptions { base64?: boolean; /** exceljs-style ceilings; exceeding one rejects with `Max row count (n) exceeded`. */ maxRows?: number; maxCols?: number }

const decoder = new TextDecoder();
const UNSUPPORTED_PARTS: [RegExp, string][] = [
  [/^xl\/charts\//, 'charts'],
  [/^xl\/tables\//, 'tables'], [/^xl\/pivot(Tables|Cache)/, 'pivot tables'], [/^xl\/(threadedComments|persons)\//, 'threaded comments'],
  [/^xl\/externalLinks\//, 'external links'], [/^xl\/vbaProject/, 'macros'], [/^xl\/chartsheets\//, 'chart sheets'], [/^xl\/slicers/, 'slicers'],
];

/** Text before the root element (other than whitespace, comments and declarations) is malformed XML; reject it as a SAX-style parser would. */
function assertRootStart(xml: string): void {
  let i = 0;
  for (;;) {
    while (i < xml.length && xml.charCodeAt(i) <= 32) i++;
    if (xml.startsWith('<?', i)) i = xml.indexOf('?>', i) + 2 || xml.length;
    else if (xml.startsWith('<!--', i)) i = xml.indexOf('-->', i) + 3 || xml.length;
    else break;
  }
  if (i >= xml.length || xml.charCodeAt(i) === 60 || xml.charCodeAt(i) === 0xfeff) return;
  const lt = xml.indexOf('<', i);
  const upto = lt < 0 ? xml.length : lt;
  const line = xml.slice(0, upto).split('\n').length;
  const col = upto - (xml.lastIndexOf('\n', upto - 1) + 1) + 1;
  throw new Error(`${line}:${col}: text data outside of root node.`);
}

function dirname(p: string): string {
  const i = p.lastIndexOf('/');
  return i < 0 ? '' : p.slice(0, i + 1);
}

/** Replaces the content of `wb` with a workbook read from `bytes`, within the same limits as `readXlsx`. */
export function loadWorkbook(wb: Workbook, bytes: Uint8Array, options: ReadOptions & LoadOptions = {}): void {
  const limits = resolveLimits(options.limits);
  const zip = unzip(bytes, defaultCompressor, limits);
  let parts: ReturnType<typeof readWorkbookParts>;
  try {
    parts = readWorkbookParts(zip, limits);
  } catch (e) {
    // a zip without a workbook part (another document type) loads as an empty workbook, as exceljs does
    if (e instanceof Error && /missing xl\/workbook\.xml/.test(e.message) && !(e instanceof XlsxLimitError)) {
      wb._worksheets = [];
      return;
    }
    throw e;
  }
  const text = (name: string): string | undefined => {
    const d = zip.get(name);
    return d ? decoder.decode(d) : undefined;
  };

  wb._worksheets = [];
  wb._unsupported = [];
  wb.media = [];
  wb._names = undefined;
  wb.properties = { date1904: parts.date1904 };
  const stylesXml = text(parts.stylesPath);
  wb._styles = stylesXml ? loadStyleBook(stylesXml, limits) : new StyleBook();
  const sstXml = text(parts.sharedStringsPath);
  const sst = sstXml ? loadSharedStrings(sstXml, limits) : [];

  for (const [re, what] of UNSUPPORTED_PARTS) if (zip.names.some((n) => re.test(n)) && !wb._unsupported.includes(what)) wb._unsupported.push(what);

  const wbXml = text('xl/workbook.xml') ?? '';
  const root = parseXml(wbXml);
  const sheetNodes = kids(child(root, 'sheets'), 'sheet');
  const views = kids(child(root, 'bookViews'), 'workbookView').map((v) => {
    const a = v.attrs;
    const o: Record<string, unknown> = {};
    const num = (k: string, to: string): void => { if (a[k] !== undefined) o[to] = parseInt(a[k], 10); };
    num('xWindow', 'x'); num('yWindow', 'y'); num('windowWidth', 'width'); num('windowHeight', 'height'); num('firstSheet', 'firstSheet'); num('activeTab', 'activeTab');
    if (a.tabRatio !== undefined) o.tabRatio = parseInt(a.tabRatio, 10);
    o.visibility = a.visibility ?? 'visible';
    for (const k of ['showHorizontalScroll', 'showVerticalScroll', 'showSheetTabs', 'minimized', 'autoFilterDateGrouping']) if (a[k] !== undefined) o[k] = a[k] === '1' || a[k] === 'true';
    return o;
  });
  wb.views = views as never;
  const calc = child(root, 'calcPr');
  wb.calcProperties = calc ? { fullCalcOnLoad: calc.attrs.fullCalcOnLoad === '1' || calc.attrs.fullCalcOnLoad === 'true' } : {};

  const styles = wb._styles;
  const ctx = { sst, book: styles, date1904: parts.date1904, limits, maxRows: options.maxRows, maxCols: options.maxCols };
  const order: Worksheet[] = [];
  const env = { zip, limits, text, media: new Map<string, number>() };
  for (const ref of parts.sheets) {
    const node = sheetNodes.find((n) => n.attrs.name === ref.name);
    const id = node?.attrs.sheetId ? parseInt(node.attrs.sheetId, 10) : wb.nextId;
    const state = node?.attrs.state === 'hidden' || node?.attrs.state === 'veryHidden' ? node.attrs.state : 'visible';
    const ws = new Worksheet(wb, { id: wb._worksheets[id] ? wb.nextId : id, name: ref.name, orderNo: order.length + 1, state });
    wb._worksheets[ws.id] = ws;
    order.push(ws);
    let data = zip.get(ref.path);
    if (!data) throw new Error(`invalid xlsx: missing worksheet part ${ref.path}`);
    const prefix = rootPrefixOf(data);
    if (prefix !== '') data = stripPrefix(data, prefix);
    const xml = decoder.decode(data);
    assertRootStart(xml);
    const sd = xml.indexOf('<sheetData');
    let headEnd = sd < 0 ? xml.length : sd;
    let tailStart = xml.length;
    let dataFrom = headEnd;
    let dataTo = headEnd;
    if (sd >= 0) {
      const gt = xml.indexOf('>', sd);
      if (xml.charCodeAt(gt - 1) === 47) tailStart = gt + 1;
      else {
        const end = xml.lastIndexOf('</sheetData>');
        dataFrom = gt + 1;
        dataTo = end < 0 ? xml.length : end;
        tailStart = end < 0 ? xml.length : end + 12;
      }
    }
    headEnd = Math.min(headEnd, xml.length);
    loadSheetHead(ws, xml.slice(0, headEnd), limits);
    loadSheetData(ws, xml, dataFrom, dataTo, ctx);
    const rels = readRels(text(`${dirname(ref.path)}_rels/${ref.path.slice(ref.path.lastIndexOf('/') + 1)}.rels`));
    loadSheetTail(ws, xml.slice(tailStart), { get: (id) => (rels.get(id)?.external ? rels.get(id)!.target : undefined) }, limits, (tail) => loadSheetExtras(ws, tail, ref.path, rels, env));
  }

  // defined names: print settings belong to pageSetup, anything else cannot be written back
  for (const dn of kids(child(root, 'definedNames'), 'definedName')) {
    const name = dn.attrs.name ?? '';
    const ws = dn.attrs.localSheetId !== undefined ? order[+dn.attrs.localSheetId] : undefined;
    const value = textOf(dn);
    const strip = (s: string): string => s.replace(/^(?:'(?:[^']|'')+'|[^!]+)!/, '').replace(/\$/g, '');
    if (name === '_xlnm.Print_Area' && ws) ws.pageSetup.printArea = value.split(',').map(strip).join('&&');
    else if (name === '_xlnm.Print_Titles' && ws) {
      for (const part of value.split(',').map(strip)) {
        if (/^\d+:\d+$/.test(part)) ws.pageSetup.printTitlesRow = part;
        else if (/^[A-Za-z]+:[A-Za-z]+$/.test(part)) ws.pageSetup.printTitlesColumn = part;
      }
    } else if (name !== '_xlnm._FilterDatabase') {
      const attrs: Record<string, string> = {};
      for (const [k, v] of Object.entries(dn.attrs)) if (k !== 'name' && k !== 'localSheetId') attrs[k] = v;
      if (dn.attrs.localSheetId === undefined || ws) wb.definedNames.load(name, value, ws, attrs); // a scope naming no sheet is an invalid file: drop the name
    }
  }

  // document properties
  const core = text('docProps/core.xml');
  if (core) {
    const c = parseXml(core);
    const get = (n: string): string | undefined => {
      const x = child(c, n);
      return x ? textOf(x) : undefined;
    };
    wb.creator = get('creator') ?? '';
    wb.lastModifiedBy = get('lastModifiedBy') ?? '';
    const d = (n: string): Date | undefined => {
      const v = get(n);
      return v ? new Date(v) : undefined;
    };
    wb.created = d('created') ?? wb.created;
    wb.modified = d('modified') ?? wb.modified;
    wb.lastPrinted = d('lastPrinted');
    wb.title = get('title'); wb.subject = get('subject'); wb.keywords = get('keywords'); wb.category = get('category');
    wb.description = get('description'); wb.contentStatus = get('contentStatus'); wb.language = get('language');
    const rev = get('revision');
    if (rev !== undefined) wb.revision = parseInt(rev, 10);
  }
  const app = text('docProps/app.xml');
  if (app) {
    const a = parseXml(app);
    const co = child(a, 'Company');
    const ma = child(a, 'Manager');
    if (co) wb.company = textOf(co);
    if (ma) wb.manager = textOf(ma);
  }
}
