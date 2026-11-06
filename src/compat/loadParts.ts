import { XlsxLimitError, type XlsxLimits } from '../limits.js';
import { Range } from './range.js';
import { MergeSlave } from './values.js';
import { child, kids, parseXml, textOf, type XNode } from './xmlTree.js';
import type { Worksheet } from './worksheet.js';

const int = (v: string | undefined): number | undefined => (v === undefined ? undefined : parseInt(v, 10));
const flt = (v: string | undefined): number | undefined => (v === undefined ? undefined : parseFloat(v));
const bool = (v: string | undefined): boolean | undefined => (v === undefined ? undefined : v === '1' || v === 'true');

function color(n: XNode | undefined): { argb?: string; theme?: number; indexed?: number; tint?: number } | undefined {
  if (!n) return undefined;
  const c: { argb?: string; theme?: number; indexed?: number; tint?: number } = {};
  if (n.attrs.rgb !== undefined) c.argb = n.attrs.rgb;
  if (n.attrs.theme !== undefined) c.theme = +n.attrs.theme;
  if (n.attrs.indexed !== undefined) c.indexed = +n.attrs.indexed;
  if (n.attrs.tint !== undefined) c.tint = +n.attrs.tint;
  return c;
}

const UNSUPPORTED: Record<string, string> = {
  conditionalFormatting: 'conditional formatting', legacyDrawingHF: 'header/footer images',
  tableParts: 'tables', sheetProtection: 'sheet protection', protectedRanges: 'protected ranges', colBreaks: 'column page breaks',
  controls: 'form controls', oleObjects: 'embedded objects', customSheetViews: 'custom sheet views',
};

/** Everything before `<sheetData>`: sheet properties, views, format, columns. */
export function loadSheetHead(ws: Worksheet, head: string, limits: Readonly<XlsxLimits>): void {
  const root = parseXml(head);
  const pr = child(root, 'sheetPr');
  const fmt = child(root, 'sheetFormatPr');
  const props: Record<string, unknown> = {
    defaultRowHeight: parseFloat(fmt?.attrs.defaultRowHeight ?? '15'),
    dyDescent: parseFloat(fmt?.attrs.dyDescent ?? '55'),
    outlineLevelRow: parseInt(fmt?.attrs.outlineLevelRow ?? '0', 10),
    outlineLevelCol: parseInt(fmt?.attrs.outlineLevelCol ?? '0', 10),
  };
  if (fmt?.attrs.defaultColWidth !== undefined) props.defaultColWidth = parseFloat(fmt.attrs.defaultColWidth);
  const tab = color(child(pr ?? root, 'tabColor'));
  if (pr && tab) props.tabColor = tab;
  const op = pr && child(pr, 'outlinePr');
  if (op) {
    const o: Record<string, boolean> = {};
    if (op.attrs.summaryBelow !== undefined) o.summaryBelow = bool(op.attrs.summaryBelow)!;
    if (op.attrs.summaryRight !== undefined) o.summaryRight = bool(op.attrs.summaryRight)!;
    props.outlineProperties = o;
  }
  ws.properties = props as never;
  ws.pageSetup = { ...(pr && child(pr, 'pageSetUpPr')?.attrs.fitToPage ? { fitToPage: bool(child(pr, 'pageSetUpPr')!.attrs.fitToPage) } : {}) };
  ws.headerFooter = {};

  const views = kids(child(root, 'sheetViews'), 'sheetView').map((v) => {
    const a = v.attrs;
    const pane = child(v, 'pane');
    const sels = new Map<string, XNode>();
    for (const s of kids(v, 'selection')) sels.set(s.attrs.pane ?? 'topLeft', s);
    const m: Record<string, unknown> = {
      workbookViewId: parseInt(a.workbookViewId ?? 'NaN', 10), rightToLeft: a.rightToLeft === '1', tabSelected: a.tabSelected === '1',
      showRuler: a.showRuler !== '0', showRowColHeaders: a.showRowColHeaders !== '0', showGridLines: a.showGridLines !== '0',
      zoomScale: parseInt(a.zoomScale ?? '100', 10), zoomScaleNormal: parseInt(a.zoomScaleNormal ?? '100', 10),
    };
    delete m.tabSelected;
    if (pane) {
      const st = pane.attrs.state;
      const out: Record<string, unknown> = {
        workbookViewId: m.workbookViewId, rightToLeft: m.rightToLeft, state: st === 'frozen' || st === 'frozenSplit' ? 'frozen' : 'split',
        xSplit: parseInt(pane.attrs.xSplit ?? '0', 10), ySplit: parseInt(pane.attrs.ySplit ?? '0', 10), topLeftCell: pane.attrs.topLeftCell,
        showRuler: m.showRuler, showRowColHeaders: m.showRowColHeaders, showGridLines: m.showGridLines, zoomScale: m.zoomScale, zoomScaleNormal: m.zoomScaleNormal,
      };
      const active = pane.attrs.activePane || 'topLeft';
      if (out.state === 'split') out.activePane = active;
      const sel = sels.get(active);
      if (sel?.attrs.activeCell) out.activeCell = sel.attrs.activeCell;
      if (a.view) out.style = a.view;
      return out;
    }
    const out: Record<string, unknown> = { ...m, state: 'normal' };
    const sel = sels.get('topLeft');
    if (sel?.attrs.activeCell) out.activeCell = sel.attrs.activeCell;
    if (a.view) out.style = a.view;
    return out;
  });
  ws.views = views as never;

  for (const c of kids(child(root, 'cols'), 'col')) {
    const min = int(c.attrs.min) ?? 0;
    const max = int(c.attrs.max) ?? min;
    if (min < 1 || max < min) continue;
    if (max > limits.maxColumns) throw new XlsxLimitError('maxColumns', `column definition up to ${max} passes ${limits.maxColumns}`);
    for (let n = min; n <= max; n++) {
      const col = ws.getColumn(n);
      if (c.attrs.width !== undefined) col.width = parseFloat(c.attrs.width);
      if (c.attrs.style) col._xf = Math.min(+c.attrs.style, ws.workbook._styles.xfs.length - 1);
      if (bool(c.attrs.hidden)) col._hidden = true;
      if (c.attrs.outlineLevel) col._outline = parseInt(c.attrs.outlineLevel, 10) || 0;
    }
  }
}

export interface SheetRels { get(id: string): string | undefined }

/** Everything after `</sheetData>`. */
export function loadSheetTail(ws: Worksheet, tail: string, rels: SheetRels, limits: Readonly<XlsxLimits>, extras?: (root: XNode) => void): void {
  const root = parseXml(`<worksheet>${tail}`);
  extras?.(root);
  for (const n of root.children) {
    const u = UNSUPPORTED[n.name];
    if (u && !ws._unsupported.includes(u)) ws._unsupported.push(u);
  }
  const af = child(root, 'autoFilter');
  if (af?.attrs.ref) ws.autoFilter = af.attrs.ref;

  let covered = 0;
  for (const m of kids(child(root, 'mergeCells'), 'mergeCell')) {
    if (!m.attrs.ref) continue;
    const r = new Range(m.attrs.ref);
    covered += r.count;
    if (covered > limits.maxMergedCells) throw new XlsxLimitError('maxMergedCells', `merged ranges cover more than ${limits.maxMergedCells} cells`);
    const master = ws.getCell(r.top, r.left);
    for (let i = r.top; i <= r.bottom; i++) {
      for (let j = r.left; j <= r.right; j++) {
        if (i === r.top && j === r.left) continue;
        const row = ws.getRow(i);
        row._v[j - 1] = new MergeSlave(master.row, master.col);
        row._cells = row._cells ?? null;
        row._setXfAt(j - 1, row._s?.[j - 1] ?? master._xf);
      }
    }
    ws._merges.set(master.address, r);
  }

  for (const h of kids(child(root, 'hyperlinks'), 'hyperlink')) {
    const ref = h.attrs.ref;
    if (!ref) continue;
    const id = h.attrs.id;
    let target = id !== undefined ? rels.get(id) : undefined;
    if (target === undefined && h.attrs.location) target = `#${h.attrs.location}`;
    if (target === undefined) continue;
    const rng = new Range(ref);
    if (rng.count > 10_000) continue;
    rng.forEachAddress((addr) => {
      const c = ws.findCell(addr);
      const v = c?.value;
      const f = v !== null && typeof v === 'object' && 'formula' in (v as object) ? (v as { formula: string; result?: unknown }) : undefined;
      const text = f ? f.result : v;
      if (c && typeof text === 'string') {
        const hv: Record<string, unknown> = { text, hyperlink: target };
        if (f) hv.formula = f.formula;
        if (h.attrs.tooltip) hv.tooltip = h.attrs.tooltip;
        ws.getRow(c.row)._v[c.col - 1] = hv;
      }
    });
  }

  for (const d of kids(child(root, 'dataValidations'), 'dataValidation')) {
    const sq = d.attrs.sqref;
    if (!sq) continue;
    const dv: Record<string, unknown> = {};
    for (const k of ['type', 'operator', 'errorStyle', 'errorTitle', 'error', 'promptTitle', 'prompt']) if (d.attrs[k] !== undefined) dv[k] = d.attrs[k];
    for (const k of ['allowBlank', 'showInputMessage', 'showErrorMessage']) if (d.attrs[k] !== undefined) dv[k] = bool(d.attrs[k]);
    const f: unknown[] = [];
    for (const k of ['formula1', 'formula2']) {
      const n = child(d, k);
      if (n) {
        const t = textOf(n);
        // typed as exceljs reads them: whole numbers, decimals and dates; anything else (and a formula such as $A$1) stays text
        const type = d.attrs.type;
        const num = type === 'whole' || type === 'textLength' ? parseInt(t, 10) : type === 'decimal' || type === 'date' ? parseFloat(t) : NaN;
        f.push(Number.isNaN(num) ? t : type === 'date' ? new Date(Math.round((num - 25569) * 86400000)) : num);
      }
    }
    if (f.length) dv.formulae = f;
    for (const tok of sq.split(/\s+/).filter(Boolean)) {
      const r = new Range(tok);
      if (r.count === 1 || r.count > 10_000) ws.dataValidations.model[r.count === 1 ? r.tl : tok] = dv;
      else r.forEachAddress((a) => (ws.dataValidations.model[a] = dv));
    }
  }

  const ps = ws.pageSetup;
  const margins = child(root, 'pageMargins');
  if (margins) ps.margins = { left: flt(margins.attrs.left), right: flt(margins.attrs.right), top: flt(margins.attrs.top), bottom: flt(margins.attrs.bottom), header: flt(margins.attrs.header), footer: flt(margins.attrs.footer) };
  const po = child(root, 'printOptions');
  if (po) {
    if (po.attrs.headings !== undefined) ps.showRowColHeaders = bool(po.attrs.headings);
    if (po.attrs.gridLines !== undefined) ps.showGridLines = bool(po.attrs.gridLines);
    if (po.attrs.horizontalCentered !== undefined) ps.horizontalCentered = bool(po.attrs.horizontalCentered);
    if (po.attrs.verticalCentered !== undefined) ps.verticalCentered = bool(po.attrs.verticalCentered);
  }
  const pg = child(root, 'pageSetup');
  if (pg) {
    const a = pg.attrs;
    const set = (k: string, v: unknown): void => {
      if (v !== undefined) ps[k] = v;
    };
    set('paperSize', int(a.paperSize)); set('orientation', a.orientation); set('horizontalDpi', int(a.horizontalDpi)); set('verticalDpi', int(a.verticalDpi));
    set('pageOrder', a.pageOrder); set('blackAndWhite', bool(a.blackAndWhite)); set('draft', bool(a.draft)); set('cellComments', a.cellComments);
    set('errors', a.errors); set('scale', int(a.scale)); set('fitToWidth', int(a.fitToWidth)); set('fitToHeight', int(a.fitToHeight));
    set('firstPageNumber', int(a.firstPageNumber)); set('useFirstPageNumber', bool(a.useFirstPageNumber)); set('usePrinterDefaults', bool(a.usePrinterDefaults));
    set('copies', int(a.copies));
  }
  const hf = child(root, 'headerFooter');
  if (hf) {
    const h = ws.headerFooter;
    if (hf.attrs.differentFirst !== undefined) h.differentFirst = bool(hf.attrs.differentFirst);
    if (hf.attrs.differentOddEven !== undefined) h.differentOddEven = bool(hf.attrs.differentOddEven);
    for (const k of ['oddHeader', 'oddFooter', 'evenHeader', 'evenFooter', 'firstHeader', 'firstFooter']) {
      const n = child(hf, k);
      if (n) h[k] = textOf(n);
    }
  }
  for (const b of kids(child(root, 'rowBreaks'), 'brk')) {
    const e: { id: number; max: number; man: number; min?: number } = { id: int(b.attrs.id) ?? 0, max: int(b.attrs.max) ?? 16383, man: int(b.attrs.man) ?? 1 };
    if (b.attrs.min) e.min = int(b.attrs.min)!;
    ws.rowBreaks.push(e);
  }
}
