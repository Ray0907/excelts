import { escapeXml } from '../xml.js';
import { encodeAddress } from './colCache.js';
import type { StyleBook } from './styleBook.js';
import { dateToSerial1900 } from '../serial.js';
import { colorXml } from './styleEmit.js';
import type { Worksheet } from './worksheet.js';

const num = (v: unknown): string => String(v);
const bool = (v: unknown): string => (v ? '1' : '0');

function viewXml(v: Record<string, unknown>, selected: boolean): string {
  let a = '';
  if (v.rightToLeft === true) a += ' rightToLeft="1"';
  if (selected || v.tabSelected) a += ' tabSelected="1"';
  if (v.showRuler === false) a += ' showRuler="0"';
  if (v.showRowColHeaders === false) a += ' showRowColHeaders="0"';
  if (v.showGridLines === false) a += ' showGridLines="0"';
  if (v.zoomScale) a += ` zoomScale="${num(v.zoomScale)}"`;
  if (v.zoomScaleNormal) a += ` zoomScaleNormal="${num(v.zoomScaleNormal)}"`;
  if (v.style) a += ` view="${escapeXml(String(v.style))}"`;
  a += ` workbookViewId="${num(v.workbookViewId || 0)}"`;
  const state = v.state as string | undefined;
  const x = (v.xSplit as number) || 0;
  const y = (v.ySplit as number) || 0;
  const cell = (k: string): string => ` ${k}="${escapeXml(String(v[k]))}"`;
  let inner = '';
  if (state === 'frozen') {
    const pane = x && y ? 'bottomRight' : x ? 'topRight' : 'bottomLeft';
    const tl = (v.topLeftCell as string | undefined) ?? encodeAddress(y + 1, x + 1);
    inner += `<pane${x ? ` xSplit="${x}"` : ''}${y ? ` ySplit="${y}"` : ''} topLeftCell="${escapeXml(tl)}" activePane="${pane}" state="frozen"/>`;
    inner += `<selection pane="${pane}"${v.activeCell ? `${cell('activeCell')} sqref="${escapeXml(String(v.activeCell))}"` : ''}/>`;
  } else if (state === 'split') {
    const ap = v.activePane && v.activePane !== 'topLeft' ? String(v.activePane) : undefined;
    inner += `<pane${x ? ` xSplit="${x}"` : ''}${y ? ` ySplit="${y}"` : ''}${v.topLeftCell ? cell('topLeftCell') : ''}${ap ? ` activePane="${ap}"` : ''}/>`;
    inner += `<selection${ap ? ` pane="${ap}"` : ''}${v.activeCell ? `${cell('activeCell')} sqref="${escapeXml(String(v.activeCell))}"` : ''}/>`;
  } else if (v.activeCell) {
    inner += `<selection${cell('activeCell')} sqref="${escapeXml(String(v.activeCell))}"/>`;
  }
  return inner ? `<sheetView${a}>${inner}</sheetView>` : `<sheetView${a}/>`;
}

function colsXml(ws: Worksheet): string {
  const cols = ws._columns;
  if (!cols) return '';
  let out = '';
  let run: { min: number; max: number; at: (typeof cols)[number] } | null = null;
  const emit = (): void => {
    if (!run) return;
    const c = run.at;
    out +=
      `<col min="${run.min}" max="${run.max}" width="${c.width ?? 9}"` +
      (c._xf ? ` style="${c._xf}"` : '') + (c._hidden ? ' hidden="1"' : '') + (c.isCustomWidth ? ' customWidth="1"' : '') +
      (c._outline ? ` outlineLevel="${c._outline}"` : '') + (c.collapsed ? ' collapsed="1"' : '') + '/>';
    run = null;
  };
  cols.forEach((c, i) => {
    if (c.isDefault) {
      emit();
    } else if (run && c.equivalentTo(run.at)) run.max = i + 1;
    else {
      emit();
      run = { min: i + 1, max: i + 1, at: c };
    }
  });
  emit();
  return out ? `<cols>${out}</cols>` : '';
}

const formulaVal = (v: unknown): string => (v instanceof Date ? String(dateToSerial1900(v.getTime())) : escapeXml(String(v)));

function dvXml(ws: Worksheet): string {
  const entries = Object.entries(ws.dataValidations.model).filter(([, v]) => v);
  if (!entries.length) return '';
  return (
    `<dataValidations count="${entries.length}">` +
    entries
      .map(([sqref, dv]) => {
        const d = dv!;
        let a = '';
        if (d.type) a += ` type="${escapeXml(d.type)}"`;
        if (d.errorStyle) a += ` errorStyle="${escapeXml(d.errorStyle)}"`;
        if (d.operator) a += ` operator="${escapeXml(d.operator)}"`;
        if (d.allowBlank) a += ' allowBlank="1"';
        if (d.showInputMessage) a += ' showInputMessage="1"';
        if (d.showErrorMessage) a += ' showErrorMessage="1"';
        if (d.errorTitle) a += ` errorTitle="${escapeXml(d.errorTitle)}"`;
        if (d.error) a += ` error="${escapeXml(d.error)}"`;
        if (d.promptTitle) a += ` promptTitle="${escapeXml(d.promptTitle)}"`;
        if (d.prompt) a += ` prompt="${escapeXml(d.prompt)}"`;
        a += ` sqref="${escapeXml(sqref)}"`;
        const f = d.formulae ?? [];
        return `<dataValidation${a}>${f[0] !== undefined ? `<formula1>${formulaVal(f[0])}</formula1>` : ''}${f[1] !== undefined ? `<formula2>${formulaVal(f[1])}</formula2>` : ''}</dataValidation>`;
      })
      .join('') +
    '</dataValidations>'
  );
}

function afCell(v: unknown): string {
  if (typeof v === 'string') return v;
  const o = v as { row: number; column: number };
  return encodeAddress(o.row, o.column);
}

export const sheetTail = {
  headParts(ws: Worksheet, _book: StyleBook, maxRow: number) {
    const p = ws.properties;
    let pr = '';
    if (p.tabColor) pr += colorXml('tabColor', p.tabColor);
    const op = p.outlineProperties;
    if (op) pr += `<outlinePr${op.summaryBelow !== undefined ? ` summaryBelow="${bool(op.summaryBelow)}"` : ''}${op.summaryRight !== undefined ? ` summaryRight="${bool(op.summaryRight)}"` : ''}/>`;
    if (ws.pageSetup.fitToPage) pr += '<pageSetUpPr fitToPage="1"/>';
    const sheetPr = pr ? `<sheetPr>${pr}</sheetPr>` : '';
    const views = ws.views.length ? ws.views : [{}];
    const sheetViews = `<sheetViews>${views.map((v, i) => viewXml(v, i === 0 && ws.workbook._firstVisible() === ws)).join('')}</sheetViews>`;
    let rowOutline = p.outlineLevelRow || 0;
    for (let i = 0; i < ws._rows.length && i < maxRow; i++) {
      const o = ws._rows[i]?._outline;
      if (o && o > rowOutline) rowOutline = o;
    }
    let colOutline = p.outlineLevelCol || 0;
    for (const c of ws._columns ?? []) if (c._outline > colOutline) colOutline = c._outline;
    const fmt =
      `<sheetFormatPr defaultRowHeight="${p.defaultRowHeight ?? 15}"` +
      (p.defaultColWidth !== undefined ? ` defaultColWidth="${p.defaultColWidth}"` : '') +
      (rowOutline ? ` outlineLevelRow="${rowOutline}"` : '') + (colOutline ? ` outlineLevelCol="${colOutline}"` : '') +
      (p.dyDescent !== undefined ? ` x14ac:dyDescent="${p.dyDescent}"` : '') + '/>';
    return { sheetPr, views: sheetViews, fmt, cols: colsXml(ws) };
  },

  tail(ws: Worksheet, _book: StyleBook): string {
    let s = '';
    const af = ws.autoFilter;
    if (af) {
      const ref = typeof af === 'string' ? af : `${afCell(af.from)}:${afCell(af.to)}`;
      s += `<autoFilter ref="${escapeXml(ref)}"/>`;
    }
    if (ws._merges.size) {
      s += `<mergeCells count="${ws._merges.size}">${[...ws._merges.values()].map((r) => `<mergeCell ref="${r.range}"/>`).join('')}</mergeCells>`;
    }
    return s + dvXml(ws);
  },

  printParts(ws: Worksheet): string {
    const ps = ws.pageSetup;
    let s = '';
    const po =
      (ps.showRowColHeaders ? ' headings="1"' : '') + (ps.showGridLines ? ' gridLines="1"' : '') +
      (ps.horizontalCentered ? ' horizontalCentered="1"' : '') + (ps.verticalCentered ? ' verticalCentered="1"' : '');
    if (po) s += `<printOptions${po}/>`;
    const m = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3, ...ps.margins };
    s += `<pageMargins left="${m.left}" right="${m.right}" top="${m.top}" bottom="${m.bottom}" header="${m.header}" footer="${m.footer}"/>`;
    let a = '';
    const add = (k: string, v: unknown): void => {
      if (v !== undefined && v !== null) a += ` ${k}="${escapeXml(String(v))}"`;
    };
    add('paperSize', ps.paperSize);
    add('scale', ps.scale);
    add('firstPageNumber', ps.firstPageNumber);
    add('fitToWidth', ps.fitToWidth);
    add('fitToHeight', ps.fitToHeight);
    add('pageOrder', ps.pageOrder);
    add('orientation', ps.orientation);
    if (ps.usePrinterDefaults !== undefined) add('usePrinterDefaults', bool(ps.usePrinterDefaults));
    if (ps.blackAndWhite !== undefined) add('blackAndWhite', bool(ps.blackAndWhite));
    if (ps.draft !== undefined) add('draft', bool(ps.draft));
    add('cellComments', ps.cellComments === 'None' ? 'none' : ps.cellComments);
    if (ps.useFirstPageNumber !== undefined) add('useFirstPageNumber', bool(ps.useFirstPageNumber));
    add('errors', ps.errors);
    add('horizontalDpi', ps.horizontalDpi);
    add('verticalDpi', ps.verticalDpi);
    add('copies', ps.copies);
    s += `<pageSetup${a}/>`;
    const hf = ws.headerFooter;
    let h = '';
    for (const k of ['oddHeader', 'oddFooter', 'evenHeader', 'evenFooter', 'firstHeader', 'firstFooter']) {
      if (hf[k]) h += `<${k}>${escapeXml(String(hf[k]))}</${k}>`;
    }
    if (h || hf.differentFirst || hf.differentOddEven) {
      s += `<headerFooter${hf.differentOddEven ? ' differentOddEven="1"' : ''}${hf.differentFirst ? ' differentFirst="1"' : ''}>${h}</headerFooter>`;
    }
    if (ws.rowBreaks.length) {
      s += `<rowBreaks count="${ws.rowBreaks.length}" manualBreakCount="${ws.rowBreaks.length}">${ws.rowBreaks.map((b) => `<brk id="${b.id}"${b.min ? ` min="${b.min}"` : ''} max="${b.max}" man="1"/>`).join('')}</rowBreaks>`;
    }
    return s;
  },
};


