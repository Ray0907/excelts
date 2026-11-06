import { XlsxLimitError, type XlsxLimits } from '../limits.js';
import { escapeXml } from '../xml.js';
import { decodeAddress, encodeAddress } from './colCache.js';
import { runFont } from './loadStrings.js';
import { richTextXml } from './richText.js';
import type { Font } from './styleTypes.js';
import { child, kids, parseXml, textOf } from './xmlTree.js';

export interface NoteObject {
  texts: { text: string; font?: Font }[];
  margins?: { insetmode?: string; inset?: number[] | string };
  protection?: { locked?: string; lockText?: string };
  editAs?: string;
}
/** What `cell.note` holds: the text, or the full description. */
export type Note = string | NoteObject;
export interface NoteAt { row: number; col: number; note: Note }

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

export function checkNote(v: unknown): Note {
  if (typeof v === 'string') return v;
  if (v && typeof v === 'object' && Array.isArray((v as NoteObject).texts)) return v as NoteObject;
  throw new TypeError('cell.note must be a string or { texts: [{ text, font? }], margins?, protection?, editAs? }');
}

function full(n: Note): Required<NoteObject> {
  const o: NoteObject = typeof n === 'string' ? { texts: [{ text: n }] } : n;
  return {
    texts: o.texts,
    margins: { insetmode: 'auto', inset: [0.13, 0.13, 0.25, 0.25], ...o.margins },
    protection: { locked: 'True', lockText: 'True', ...o.protection },
    editAs: o.editAs ?? 'absolute',
  };
}

export function commentsXml(list: readonly NoteAt[]): string {
  return (
    `${DECL}<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Author</author></authors><commentList>` +
    list.map((n) => `<comment ref="${encodeAddress(n.row, n.col)}" authorId="0"><text>${richTextXml(full(n.note).texts)}</text></comment>`).join('') +
    '</commentList></comments>'
  );
}

const mm = (v: unknown): string => `${Number((parseFloat(String(v)) * 10).toFixed(4))}mm`;

/** The legacy drawing that makes Excel show the note marker; `block` keeps shape ids distinct between sheets. */
export function vmlXml(list: readonly NoteAt[], block: number): string {
  let s =
    `${DECL}<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">` +
    `<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${block}"/></o:shapelayout>` +
    '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>';
  list.forEach((n, i) => {
    const f = full(n.note);
    const inset = Array.isArray(f.margins.inset) ? f.margins.inset.map(mm).join(',') : f.margins.inset;
    const l = n.col;
    const t = Math.max(n.row - 2, 0);
    const move = f.editAs === 'absolute' ? '<x:MoveWithCells/><x:SizeWithCells/>' : f.editAs === 'oneCells' ? '<x:SizeWithCells/>' : '';
    s +=
      `<v:shape id="_x0000_s${block * 1024 + 1 + i}" type="#_x0000_t202" style="position:absolute; margin-left:105.3pt;margin-top:10.5pt;width:97.8pt;height:59.1pt;z-index:1;visibility:hidden" fillcolor="infoBackground [80]" strokecolor="none [81]"` +
      `${f.margins.insetmode ? ` o:insetmode="${escapeXml(f.margins.insetmode)}"` : ''}><v:fill color2="infoBackground [80]"/><v:shadow color="none [81]" obscured="t"/><v:path o:connecttype="none"/>` +
      `<v:textbox style="mso-direction-alt:auto"${inset ? ` inset="${escapeXml(inset)}"` : ''}><div style="text-align:left"/></v:textbox>` +
      `<x:ClientData ObjectType="Note">${move}<x:Anchor>${l}, 6, ${t}, 14, ${l + 2}, 2, ${t + 4}, 16</x:Anchor><x:Locked>${escapeXml(f.protection.locked ?? 'True')}</x:Locked>` +
      `<x:AutoFill>False</x:AutoFill><x:LockText>${escapeXml(f.protection.lockText ?? 'True')}</x:LockText><x:Row>${n.row - 1}</x:Row><x:Column>${n.col - 1}</x:Column></x:ClientData></v:shape>`;
  });
  return `${s}</xml>`;
}

interface Vml { insetmode?: string; inset?: number[]; locked?: string; lockText?: string; editAs: string }
export interface VmlNotes { shapes: Map<string, Vml>; other: boolean }

/** Placement details of Note shapes, by `row,col` (1-based); `other` is set when the part also holds shapes that are not notes. */
export function parseVml(vml: string, limits: Readonly<XlsxLimits>): VmlNotes {
  const shapes = new Map<string, Vml>();
  let other = false;
  const chunks = vml.split(/<(?:[\w.-]+:)?shape\b/).slice(1, limits.maxDrawingObjects + 2);
  for (const c of chunks) {
    const row = /<(?:[\w.-]+:)?Row>\s*(\d+)/.exec(c);
    const col = /<(?:[\w.-]+:)?Column>\s*(\d+)/.exec(c);
    if (!/ObjectType\s*=\s*["']Note["']/.test(c) || !row || !col) {
      other = true;
      continue;
    }
    const get = (re: RegExp): string | undefined => re.exec(c)?.[1];
    const inset = get(/<(?:[\w.-]+:)?textbox\b[^>]*?\binset\s*=\s*"([^"]*)"/);
    const flags = ['MoveWithCells', 'SizeWithCells'].filter((t) => new RegExp(`<(?:[\\w.-]+:)?${t}\\b`).test(c)).length;
    shapes.set(`${+row[1]! + 1},${+col[1]! + 1}`, {
      insetmode: get(/\binsetmode\s*=\s*"([^"]*)"/),
      inset: inset ? inset.split(',').map((m) => Number((parseFloat(m) * 0.1).toFixed(2))) : undefined,
      locked: get(/<(?:[\w.-]+:)?Locked>([^<]*)</),
      lockText: get(/<(?:[\w.-]+:)?LockText>([^<]*)</),
      editAs: ['twoCells', 'oneCells', 'absolute'][flags]!,
    });
  }
  return { shapes, other };
}

/** Notes of a comments part, joined with their placement details from the legacy drawing (matched by cell, not by order). */
export function parseNotes(comments: string, vml: VmlNotes | undefined, limits: Readonly<XlsxLimits>): NoteAt[] {
  const list = kids(child(parseXml(comments), 'commentList'), 'comment');
  if (list.length > limits.maxDrawingObjects) throw new XlsxLimitError('maxDrawingObjects', `comments part holds more than ${limits.maxDrawingObjects} notes`);
  const out: NoteAt[] = [];
  for (const c of list) {
    if (!c.attrs.ref) continue;
    const at = decodeAddress(c.attrs.ref);
    const text = child(c, 'text');
    const texts: { text: string; font?: Font }[] = [];
    if (text) {
      const runs = kids(text, 'r');
      if (runs.length) {
        for (const r of runs) {
          const t = child(r, 't');
          const rpr = child(r, 'rPr');
          texts.push({ ...(rpr ? { font: runFont(rpr) } : {}), text: t ? textOf(t) : '' });
        }
      } else {
        const t = child(text, 't');
        if (t) texts.push({ text: textOf(t) });
      }
    }
    const v = vml?.shapes.get(`${at.row},${at.col}`);
    if (texts.length === 1 && Object.keys(texts[0]!).length === 1) out.push({ row: at.row, col: at.col, note: texts[0]!.text });
    else {
      const note: NoteObject = { texts };
      if (v) {
        note.margins = { ...(v.insetmode !== undefined ? { insetmode: v.insetmode } : {}), ...(v.inset ? { inset: v.inset } : {}) };
        note.protection = { ...(v.locked !== undefined ? { locked: v.locked } : {}), ...(v.lockText !== undefined ? { lockText: v.lockText } : {}) };
        note.editAs = v.editAs;
      }
      out.push({ row: at.row, col: at.col, note });
    }
  }
  return out;
}
