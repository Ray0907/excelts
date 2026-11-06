import type { XlsxLimits } from '../limits.js';
import type { ZipReader } from '../unzip.js';
import { parseDrawing } from './drawing.js';
import { asBuffer, Image, type Medium } from './images.js';
import { parseNotes, parseVml } from './notes.js';
import { child, kids, parseXml, type XNode } from './xmlTree.js';
import type { Worksheet } from './worksheet.js';

export interface PartRel { type: string; target: string; external: boolean }

export function dirOf(path: string): string {
  return path.slice(0, path.lastIndexOf('/') + 1);
}

/** The package path a relationship `target` names, relative to the part at `from`. */
export function resolvePart(from: string, target: string): string {
  const parts = (target.startsWith('/') ? target.slice(1) : dirOf(from) + target).split('/');
  const out: string[] = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p !== '.' && p !== '') out.push(p);
  }
  return out.join('/');
}

export function readRels(xml: string | undefined): Map<string, PartRel> {
  const m = new Map<string, PartRel>();
  if (xml) for (const r of kids(parseXml(xml), 'Relationship')) if (r.attrs.Id && r.attrs.Target) m.set(r.attrs.Id, { type: r.attrs.Type ?? '', target: r.attrs.Target, external: r.attrs.TargetMode === 'External' });
  return m;
}

export interface ExtrasEnv {
  zip: ZipReader;
  limits: Readonly<XlsxLimits>;
  text(name: string): string | undefined;
  /** media part path -> index in workbook.media, shared by all sheets of the load */
  media: Map<string, number>;
}

function lazyMedium(path: string, zip: ZipReader): Medium | undefined {
  const base = path.slice(path.lastIndexOf('/') + 1);
  const dot = base.lastIndexOf('.');
  if (dot < 1) return undefined;
  let cache: Uint8Array | undefined;
  return {
    type: 'image',
    name: base.slice(0, dot),
    extension: base.slice(dot + 1),
    get buffer(): Uint8Array {
      return (cache ??= asBuffer(zip.get(path) ?? new Uint8Array()));
    },
  };
}

const relsPathOf = (part: string): string => `${dirOf(part)}_rels/${part.slice(part.lastIndexOf('/') + 1)}.rels`;
const ofType = (rels: Map<string, PartRel>, id: string | undefined, suffix: string): PartRel | undefined => {
  const r = id ? rels.get(id) : undefined;
  return r && !r.external && r.type.endsWith(suffix) ? r : undefined;
};

/** Pictures, the background picture and notes of one sheet, read from the elements after its data. */
export function loadSheetExtras(ws: Worksheet, tail: XNode, sheetPath: string, rels: Map<string, PartRel>, env: ExtrasEnv): void {
  const wb = ws.workbook;
  const mediaId = (path: string): number | undefined => {
    let id = env.media.get(path);
    if (id === undefined && env.zip.names.includes(path)) {
      const m = lazyMedium(path, env.zip);
      if (!m) return undefined;
      id = wb.media.push(m) - 1;
      env.media.set(path, id);
    }
    return id;
  };
  const lost = (what: string): void => {
    if (!ws._unsupported.includes(what)) ws._unsupported.push(what);
  };

  const drel = ofType(rels, child(tail, 'drawing')?.attrs.id, '/drawing');
  if (drel) {
    const part = resolvePart(sheetPath, drel.target);
    const xml = env.text(part);
    if (xml !== undefined) {
      const dr = readRels(env.text(relsPathOf(part)));
      const get = (id: string): { target: string; external: boolean } | undefined => {
        const r = dr.get(id);
        return r && { target: r.external ? r.target : resolvePart(part, r.target), external: r.external };
      };
      const res = parseDrawing(ws, xml, { get }, mediaId, env.limits);
      ws._media.push(...res.images);
      res.lost.forEach(lost);
    }
  }

  const prel = ofType(rels, child(tail, 'picture')?.attrs.id, '/image');
  if (prel) {
    const id = mediaId(resolvePart(sheetPath, prel.target));
    if (id !== undefined) ws._media.push(new Image(ws, 'background', id));
  }

  const crel = [...rels.values()].find((r) => !r.external && r.type.endsWith('/comments'));
  const vrel = ofType(rels, child(tail, 'legacyDrawing')?.attrs.id, '/vmlDrawing');
  const vmlText = vrel ? env.text(resolvePart(sheetPath, vrel.target)) : undefined;
  const vml = vmlText !== undefined ? parseVml(vmlText, env.limits) : undefined;
  if (vml?.other) lost('legacy drawing shapes (form controls)');
  const comments = crel ? env.text(resolvePart(sheetPath, crel.target)) : undefined;
  if (comments !== undefined) {
    for (const n of parseNotes(comments, vml, env.limits)) {
      if (n.row > env.limits.maxRows || n.col > env.limits.maxColumns) continue;
      const cell = ws.getCell(n.row, n.col);
      cell.note = n.note;
    }
  }
}
