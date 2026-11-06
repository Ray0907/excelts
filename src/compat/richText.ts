import { escapeText } from '../xml.js';
import { colorXml } from './styleEmit.js';
import type { Font } from './styleTypes.js';

export interface RichRun { text: string; font?: Font }

function rPr(f: Font): string {
  let s = '';
  if (f.name) s += `<rFont val="${escapeText(f.name)}"/>`;
  if (f.charset !== undefined) s += `<charset val="${f.charset}"/>`;
  if (f.family !== undefined) s += `<family val="${f.family}"/>`;
  if (f.bold) s += '<b/>';
  if (f.italic) s += '<i/>';
  if (f.strike) s += '<strike/>';
  if (f.outline) s += '<outline/>';
  if (f.shadow) s += '<shadow/>';
  if (f.condense) s += '<condense/>';
  if (f.extend) s += '<extend/>';
  s += colorXml('color', f.color);
  if (f.size !== undefined) s += `<sz val="${f.size}"/>`;
  if (f.underline) s += f.underline === true || f.underline === 'single' ? '<u/>' : `<u val="${escapeText(String(f.underline))}"/>`;
  if (f.vertAlign) s += `<vertAlign val="${escapeText(f.vertAlign)}"/>`;
  if (f.scheme && f.scheme !== 'none') s += `<scheme val="${escapeText(f.scheme)}"/>`;
  return s ? `<rPr>${s}</rPr>` : '';
}

/** The contents of an `<si>` for rich-text runs. */
export function richTextXml(runs: readonly RichRun[]): string {
  return runs
    .map((r) => {
      const text = String(r.text ?? '');
      const space = /^\s|\s$|\n|\r/.test(text) ? ' xml:space="preserve"' : '';
      return `<r>${r.font ? rPr(r.font) : ''}<t${space}>${escapeText(text)}</t></r>`;
    })
    .join('');
}
