import { escapeXml } from '../xml.js';
import type { Alignment, Border, BorderEdge, Color, Fill, Font, Protection } from './styleTypes.js';

export const BUILTIN_FORMATS: Record<number, string> = {
  1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'mm-dd-yy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm',
  21: 'h:mm:ss', 22: 'm/d/yy h:mm', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@',
};
export const BUILTIN_IDS = new Map<string, number>(Object.entries(BUILTIN_FORMATS).map(([k, v]) => [v, +k]));

export function colorXml(tag: string, c: Color | undefined): string {
  if (!c) return '';
  let a = '';
  if (c.argb !== undefined) a += ` rgb="${escapeXml(String(c.argb))}"`;
  if (c.theme !== undefined) a += ` theme="${c.theme}"`;
  if (c.indexed !== undefined) a += ` indexed="${c.indexed}"`;
  if (c.auto) a += ' auto="1"';
  if (c.tint !== undefined) a += ` tint="${c.tint}"`;
  return `<${tag}${a}/>`;
}

export function fontXml(f: Font): string {
  let s = '<font>';
  if (f.bold) s += '<b/>';
  if (f.italic) s += '<i/>';
  if (f.strike) s += '<strike/>';
  if (f.condense) s += '<condense/>';
  if (f.extend) s += '<extend/>';
  if (f.outline) s += '<outline/>';
  if (f.shadow) s += '<shadow/>';
  if (f.underline) s += f.underline === true || f.underline === 'single' ? '<u/>' : `<u val="${escapeXml(String(f.underline))}"/>`;
  if (f.vertAlign) s += `<vertAlign val="${escapeXml(f.vertAlign)}"/>`;
  if (f.size !== undefined) s += `<sz val="${f.size}"/>`;
  s += colorXml('color', f.color);
  if (f.name) s += `<name val="${escapeXml(f.name)}"/>`;
  if (f.family !== undefined) s += `<family val="${f.family}"/>`;
  if (f.charset !== undefined) s += `<charset val="${f.charset}"/>`;
  if (f.scheme && f.scheme !== 'none') s += `<scheme val="${escapeXml(f.scheme)}"/>`;
  return s + '</font>';
}

export function fillXml(f: Fill): string {
  if (f.type === 'pattern') {
    return `<fill><patternFill patternType="${escapeXml(f.pattern)}">${colorXml('fgColor', f.fgColor)}${colorXml('bgColor', f.bgColor)}</patternFill></fill>`;
  }
  let a = '';
  if (f.gradient === 'angle') a = ` degree="${f.degree ?? 0}"`;
  else {
    a = ' type="path"';
    const c = f.center ?? {};
    const left = c.left, right = c.right ?? c.left, top = c.top, bottom = c.bottom ?? c.top;
    if (left) a += ` left="${left}"`;
    if (right) a += ` right="${right}"`;
    if (top) a += ` top="${top}"`;
    if (bottom) a += ` bottom="${bottom}"`;
  }
  return `<fill><gradientFill${a}>${f.stops.map((s) => `<stop position="${s.position}">${colorXml('color', s.color)}</stop>`).join('')}</gradientFill></fill>`;
}

function edgeXml(tag: string, e: BorderEdge | undefined): string {
  if (!e || !e.style) return `<${tag}/>`;
  const c = colorXml('color', e.color);
  return c ? `<${tag} style="${escapeXml(e.style)}">${c}</${tag}>` : `<${tag} style="${escapeXml(e.style)}"/>`;
}

export function borderXml(b: Border): string {
  const d = b.diagonal;
  const dattr = d && d.style ? (d.up ? ' diagonalUp="1"' : '') + (d.down ? ' diagonalDown="1"' : '') : '';
  return `<border${dattr}>${edgeXml('left', b.left)}${edgeXml('right', b.right)}${edgeXml('top', b.top)}${edgeXml('bottom', b.bottom)}${edgeXml('diagonal', d)}</border>`;
}

export function alignmentXml(a: Alignment): string {
  let s = '<alignment';
  if (a.horizontal) s += ` horizontal="${a.horizontal}"`;
  if (a.vertical) s += ` vertical="${a.vertical === 'middle' ? 'center' : a.vertical}"`;
  if (a.wrapText) s += ' wrapText="1"';
  if (a.shrinkToFit) s += ' shrinkToFit="1"';
  if (a.indent) s += ` indent="${a.indent}"`;
  if (a.textRotation !== undefined) {
    const t = a.textRotation;
    const v = t === 'vertical' ? 255 : t < 0 ? 90 - t : t;
    if (v) s += ` textRotation="${v}"`;
  }
  const ro = a.readingOrder;
  if (ro !== undefined) s += ` readingOrder="${ro === 'rtl' || ro === 2 ? 2 : 1}"`;
  return s + '/>';
}

export function protectionXml(p: Protection): string {
  let s = '<protection';
  if (p.locked !== undefined) s += ` locked="${p.locked ? 1 : 0}"`;
  if (p.hidden !== undefined) s += ` hidden="${p.hidden ? 1 : 0}"`;
  return s + '/>';
}

const H = new Set(['left', 'center', 'right', 'fill', 'centerContinuous', 'distributed', 'justify']);
const V = new Set(['top', 'middle', 'bottom', 'distributed', 'justify', 'center']);

/** The alignment as it survives a save; undefined when nothing valid is left (exceljs drops invalid values the same way). */
export function normalizeAlignment(a: Alignment | undefined): Alignment | undefined {
  if (!a || typeof a !== 'object') return undefined;
  const o: Alignment = {};
  if (a.horizontal && H.has(a.horizontal)) o.horizontal = a.horizontal;
  if (a.vertical && V.has(a.vertical)) o.vertical = a.vertical === 'center' ? 'middle' : a.vertical;
  if (a.wrapText) o.wrapText = true;
  if (a.shrinkToFit) o.shrinkToFit = true;
  const ind = Math.max(0, Math.trunc(Number(a.indent)) || 0);
  if (ind) o.indent = ind;
  const t = a.textRotation;
  if (t === 'vertical') o.textRotation = t;
  else if (typeof t === 'number' && t >= -90 && t <= 90 && t !== 0) o.textRotation = t;
  if (a.readingOrder === 'rtl' || a.readingOrder === 2) o.readingOrder = 'rtl';
  else if (a.readingOrder === 'ltr' || a.readingOrder === 1) o.readingOrder = 'ltr';
  return Object.keys(o).length ? o : undefined;
}
