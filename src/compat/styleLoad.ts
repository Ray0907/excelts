import { XlsxLimitError, type XlsxLimits } from '../limits.js';
import { StyleBook } from './styleBook.js';
import { borderXml, fillXml, fontXml, normalizeAlignment } from './styleEmit.js';
import type { Alignment, Border, BorderEdge, Color, Fill, Font, Protection } from './styleTypes.js';
import { child, kids, parseXml, type XNode } from './xmlTree.js';

const bool = (v: string | undefined, dflt = true): boolean => (v === undefined ? dflt : v === '1' || v === 'true');

function color(n: XNode | undefined): Color | undefined {
  if (!n) return undefined;
  const a = n.attrs;
  const c: Color = {};
  if (a.rgb !== undefined) c.argb = a.rgb;
  if (a.theme !== undefined) c.theme = +a.theme;
  if (a.indexed !== undefined) c.indexed = +a.indexed;
  if (a.auto !== undefined && bool(a.auto)) c.auto = true;
  if (a.tint !== undefined) c.tint = +a.tint;
  return Object.keys(c).length ? c : undefined;
}

function font(n: XNode): Font {
  const f: Font = {};
  for (const c of n.children) {
    const v = c.attrs.val;
    switch (c.name) {
      case 'b': f.bold = bool(v); break;
      case 'i': f.italic = bool(v); break;
      case 'strike': f.strike = bool(v); break;
      case 'outline': f.outline = bool(v); break;
      case 'shadow': f.shadow = bool(v); break;
      case 'condense': f.condense = bool(v); break;
      case 'extend': f.extend = bool(v); break;
      case 'u': f.underline = v === undefined || v === 'single' ? true : v; break;
      case 'vertAlign': if (v) f.vertAlign = v; break;
      case 'sz': f.size = parseFloat(v ?? ''); break;
      case 'name': if (v !== undefined) f.name = v; break;
      case 'family': f.family = parseInt(v ?? '', 10); break;
      case 'charset': f.charset = parseInt(v ?? '', 10); break;
      case 'scheme': if (v) f.scheme = v; break;
      case 'color': { const k = color(c); if (k) f.color = k; break; }
    }
  }
  return f;
}

function fill(n: XNode): Fill | undefined {
  const p = child(n, 'patternFill');
  if (p) {
    const o: Fill = { type: 'pattern', pattern: p.attrs.patternType ?? 'none' };
    const fg = color(child(p, 'fgColor'));
    const bg = color(child(p, 'bgColor'));
    if (fg) o.fgColor = fg;
    if (bg) o.bgColor = bg;
    return o;
  }
  const g = child(n, 'gradientFill');
  if (!g) return undefined;
  const stops = kids(g, 'stop').map((s) => ({ position: parseFloat(s.attrs.position ?? '0'), color: color(child(s, 'color')) ?? {} }));
  const a = g.attrs;
  if (a.degree) return { type: 'gradient', gradient: 'angle', degree: parseInt(a.degree, 10), stops };
  if (a.type === 'path') {
    const center: NonNullable<Extract<Fill, { type: 'gradient' }>['center']> = { left: a.left ? parseFloat(a.left) : 0, top: a.top ? parseFloat(a.top) : 0 };
    if (a.right !== a.left) center.right = a.right ? parseFloat(a.right) : 0;
    if (a.bottom !== a.top) center.bottom = a.bottom ? parseFloat(a.bottom) : 0;
    return { type: 'gradient', gradient: 'path', center, stops };
  }
  return { type: 'gradient', gradient: 'angle', degree: 0, stops };
}

function edge(n: XNode | undefined): BorderEdge | undefined {
  if (!n || !n.attrs.style) return undefined;
  const e: BorderEdge = { style: n.attrs.style };
  const c = color(child(n, 'color'));
  if (c) e.color = c;
  return e;
}

function border(n: XNode): Border {
  const b: Border = {};
  const l = edge(child(n, 'left') ?? child(n, 'start'));
  const r = edge(child(n, 'right') ?? child(n, 'end'));
  const t = edge(child(n, 'top'));
  const bo = edge(child(n, 'bottom'));
  if (t) b.top = t;
  if (l) b.left = l;
  if (bo) b.bottom = bo;
  if (r) b.right = r;
  const d = edge(child(n, 'diagonal'));
  if (d) {
    const dd: NonNullable<Border['diagonal']> = { ...d };
    if (bool(n.attrs.diagonalUp, false)) dd.up = true;
    if (bool(n.attrs.diagonalDown, false)) dd.down = true;
    b.diagonal = dd;
  }
  return b;
}

function alignment(n: XNode): Alignment | undefined {
  const a = n.attrs;
  const o: Alignment = {};
  if (a.horizontal) o.horizontal = a.horizontal;
  if (a.vertical) o.vertical = a.vertical === 'center' ? 'middle' : a.vertical;
  if (a.wrapText && bool(a.wrapText)) o.wrapText = true;
  if (a.shrinkToFit && bool(a.shrinkToFit)) o.shrinkToFit = true;
  if (a.indent) o.indent = parseInt(a.indent, 10);
  if (a.textRotation) {
    const t = parseInt(a.textRotation, 10);
    if (t === 255) o.textRotation = 'vertical';
    else if (t >= 0 && t <= 90) o.textRotation = t;
    else if (t > 90 && t <= 180) o.textRotation = 90 - t;
  }
  if (a.readingOrder) o.readingOrder = a.readingOrder === '2' ? 'rtl' : 'ltr';
  return normalizeAlignment(o);
}

function protection(n: XNode): Protection | undefined {
  const p: Protection = {};
  if (n.attrs.locked !== undefined) p.locked = bool(n.attrs.locked);
  if (n.attrs.hidden !== undefined) p.hidden = bool(n.attrs.hidden);
  return Object.keys(p).length ? p : undefined;
}

/** Rebuilds a StyleBook from a styles part so cellXfs indexes stay the style ids. */
export function loadStyleBook(xml: string, limits: Readonly<XlsxLimits>): StyleBook {
  const root = parseXml(xml, 4_000_000);
  const book = new StyleBook();
  book.resetForLoad();
  const need = (what: keyof XlsxLimits, n: number): void => {
    if (n > limits[what]) throw new XlsxLimitError(what, `styles part has ${n} entries`);
  };
  const fmts = kids(child(root, 'numFmts'), 'numFmt');
  need('maxCellXfs', fmts.length);
  for (const f of fmts) if (f.attrs.numFmtId && f.attrs.formatCode !== undefined) book.defineFormat(+f.attrs.numFmtId, f.attrs.formatCode);
  const fonts = kids(child(root, 'fonts'), 'font');
  need('maxCellXfs', fonts.length);
  fonts.forEach((n, i) => {
    const f = font(n);
    book.fonts.addLoaded(i === 0 ? undefined : f, fontXml(f));
  });
  const fills = kids(child(root, 'fills'), 'fill');
  need('maxCellXfs', fills.length);
  fills.forEach((n, i) => {
    const f = fill(n);
    book.fills.addLoaded(i === 0 ? undefined : f, f ? fillXml(f) : '<fill><patternFill patternType="none"/></fill>');
  });
  const borders = kids(child(root, 'borders'), 'border');
  need('maxCellXfs', borders.length);
  borders.forEach((n, i) => {
    const b = border(n);
    book.borders.addLoaded(i === 0 ? undefined : b, borderXml(b));
  });
  const xfs = kids(child(root, 'cellXfs'), 'xf');
  need('maxCellXfs', xfs.length);
  const a2 = (n: XNode | undefined): number => {
    const a = n && alignment(n);
    return a ? book.aligns.add(a) : 0;
  };
  for (const x of xfs) {
    const p = child(x, 'protection');
    const pr = p && protection(p);
    book.pushLoadedXf({
      n: +(x.attrs.numFmtId ?? 0), f: Math.min(+(x.attrs.fontId ?? 0), Math.max(0, fonts.length - 1)),
      fi: Math.min(+(x.attrs.fillId ?? 0), Math.max(0, fills.length - 1)), b: Math.min(+(x.attrs.borderId ?? 0), Math.max(0, borders.length - 1)),
      a: a2(child(x, 'alignment')), p: pr ? book.prots.add(pr) : 0,
    });
  }
  if (book.xfs.length === 0) book.pushLoadedXf({ n: 0, f: 0, fi: 0, b: 0, a: 0, p: 0 });
  // Parts carried over verbatim.
  const grab = (name: string): string | undefined => {
    const at = xml.search(new RegExp(`<${name}[\\s>/]`));
    if (at < 0) return undefined;
    const gt = xml.indexOf('>', at);
    if (xml.charCodeAt(gt - 1) === 47) return xml.slice(at, gt + 1);
    const close = xml.indexOf(`</${name}>`, gt);
    return close < 0 ? undefined : xml.slice(at, close + name.length + 3);
  };
  for (const k of ['cellStyleXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors'] as const) {
    const g = grab(k);
    if (g && !/xmlns:/.test(g) && !/<[a-z0-9]+:/i.test(g)) book.raw[k] = g;
  }
  return book;
}
