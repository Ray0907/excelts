import { readSharedStrings } from '../readSharedStrings.js';
import type { XlsxLimits } from '../limits.js';
import { child, kids, parseXml, textOf, type XNode } from './xmlTree.js';
import type { Font } from './styleTypes.js';

export function runFont(rpr: XNode): Font {
  const f: Font = {};
  for (const c of rpr.children) {
    const v = c.attrs.val;
    switch (c.name) {
      case 'rFont': if (v !== undefined) f.name = v; break;
      case 'charset': f.charset = parseInt(v ?? '', 10); break;
      case 'family': f.family = parseInt(v ?? '', 10); break;
      case 'b': f.bold = v !== '0' && v !== 'false'; break;
      case 'i': f.italic = v !== '0' && v !== 'false'; break;
      case 'strike': f.strike = v !== '0' && v !== 'false'; break;
      case 'outline': f.outline = v !== '0' && v !== 'false'; break;
      case 'shadow': f.shadow = v !== '0' && v !== 'false'; break;
      case 'u': f.underline = v === undefined || v === 'single' ? true : v; break;
      case 'vertAlign': if (v) f.vertAlign = v; break;
      case 'sz': f.size = parseFloat(v ?? ''); break;
      case 'scheme': if (v) f.scheme = v; break;
      case 'color': {
        const k: NonNullable<Font['color']> = {};
        if (c.attrs.rgb !== undefined) k.argb = c.attrs.rgb;
        if (c.attrs.theme !== undefined) k.theme = +c.attrs.theme;
        if (c.attrs.indexed !== undefined) k.indexed = +c.attrs.indexed;
        if (c.attrs.tint !== undefined) k.tint = +c.attrs.tint;
        f.color = k;
        break;
      }
    }
  }
  return f;
}

/** Shared strings as cells hold them: a string, or `{ richText }` where the item has formatted runs. */
export function loadSharedStrings(xml: string, limits: Readonly<XlsxLimits>): unknown[] {
  const plain: unknown[] = readSharedStrings(xml, limits);
  if (!/<(?:[A-Za-z0-9]+:)?r[\s>]/.test(xml)) return plain;
  let i = 0;
  let p = xml.indexOf('<si');
  while (p >= 0 && i < plain.length) {
    const c = xml.charCodeAt(p + 3);
    if (c !== 62 && c !== 32 && c !== 47) {
      p = xml.indexOf('<si', p + 3);
      continue;
    }
    const gt = xml.indexOf('>', p);
    const selfClose = xml.charCodeAt(gt - 1) === 47;
    const end = selfClose ? gt + 1 : xml.indexOf('</si>', gt) + 5;
    if (!selfClose && end > 4) {
      const block = xml.slice(p, end);
      if (/<(?:[A-Za-z0-9]+:)?r[\s>]/.test(block)) {
        const si = parseXml(block);
        const runs = kids(si, 'r').map((r) => {
          const run: { text: string; font?: Font } = { text: textOf(child(r, 't') ?? { name: 't', attrs: {}, children: [], text: '' }) };
          const rpr = child(r, 'rPr');
          if (rpr) run.font = runFont(rpr);
          return run;
        });
        if (runs.length) plain[i] = { richText: runs };
      }
    }
    i++;
    p = xml.indexOf('<si', end);
  }
  return plain;
}
