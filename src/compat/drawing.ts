import { XlsxLimitError, type XlsxLimits } from '../limits.js';
import { escapeXml } from '../xml.js';
import { Anchor, Image, type ImageRange } from './images.js';
import { NS_REL } from './rels.js';
import { child, parseXml, type XNode } from './xmlTree.js';
import type { Worksheet } from './worksheet.js';

const EMU_PER_PIXEL = 9525;
const NS_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const NS_XN = 'urn:xlsx-next:1';
const EDIT_AS = new Set(['twoCell', 'oneCell', 'absolute']);

const pos = (v: number): number => {
  if (!Number.isFinite(v) || v < 0) throw new RangeError(`invalid image position ${v}`);
  return Math.floor(v);
};
const cell = (tag: string, a: Anchor): string =>
  `<xdr:${tag}><xdr:col>${pos(a.nativeCol)}</xdr:col><xdr:colOff>${pos(a.nativeColOff)}</xdr:colOff><xdr:row>${pos(a.nativeRow)}</xdr:row><xdr:rowOff>${pos(a.nativeRowOff)}</xdr:rowOff></xdr:${tag}>`;

/** The drawing part for a sheet's pictures and its relationships. `target(imageId)` gives the media part relative to the drawing. */
export function drawingParts(images: readonly Image[], target: (imageId: number) => string): { xml: string; rels: string } {
  const rels: string[] = [];
  const byImage = new Map<number, string>();
  const rel = (type: string, to: string, ext = false): string => {
    rels.push(`<Relationship Id="rId${rels.length + 1}" Type="${NS_REL}/${type}" Target="${escapeXml(to)}"${ext ? ' TargetMode="External"' : ''}/>`);
    return `rId${rels.length}`;
  };
  let body = '';
  let usesPrivate = false;
  images.forEach((img, i) => {
    const r = img.range as ImageRange;
    let rid = byImage.get(img.imageId);
    if (!rid) byImage.set(img.imageId, (rid = rel('image', target(img.imageId))));
    const link = r.hyperlinks?.hyperlink ? `<a:hlinkClick r:id="${rel('hyperlink', r.hyperlinks.hyperlink, true)}"${r.hyperlinks.tooltip ? ` tooltip="${escapeXml(r.hyperlinks.tooltip)}"` : ''}/>` : '';
    const pic =
      `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${i + 1}" name="Picture ${i + 1}">${link}</xdr:cNvPr><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
      `<xdr:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
      '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>';
    const editAs = EDIT_AS.has(r.editAs ?? '') ? r.editAs : 'oneCell';
    if (r.br) body += `<xdr:twoCellAnchor editAs="${editAs}">${cell('from', r.tl)}${cell('to', r.br)}${pic}<xdr:clientData/></xdr:twoCellAnchor>`;
    else {
      const e = r.ext!;
      // CT_OneCellAnchor has no editAs attribute (exceljs writes one anyway, which the schema forbids), so a non-default value is kept in a private ignorable attribute
      const priv = editAs !== 'oneCell';
      usesPrivate ||= priv;
      body += `<xdr:oneCellAnchor${priv ? ` xn:editAs="${editAs}"` : ''}>${cell('from', r.tl)}<xdr:ext cx="${pos(e.width * EMU_PER_PIXEL)}" cy="${pos(e.height * EMU_PER_PIXEL)}"/>${pic}<xdr:clientData/></xdr:oneCellAnchor>`;
    }
  });
  return {
    xml:
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${NS_REL}"${usesPrivate ? ` xmlns:mc="${NS_MC}" xmlns:xn="${NS_XN}" mc:Ignorable="xn"` : ''}>${body}</xdr:wsDr>`,
    rels: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`,
  };
}

export interface DrawingRels { get(id: string): { target: string; external: boolean } | undefined }

const int = (n: XNode | undefined, name: string): number => parseInt((n && child(n, name)?.text) || '0', 10) || 0;
const anchorOf = (ws: Worksheet, n: XNode | undefined): Anchor =>
  new Anchor(ws, { nativeCol: int(n, 'col'), nativeColOff: int(n, 'colOff'), nativeRow: int(n, 'row'), nativeRowOff: int(n, 'rowOff') });

/**
 * Reads the pictures of a drawing part. Anything else it holds (charts, shapes, groups, absolute anchors, linked pictures) is returned in `lost`,
 * since it cannot be written back. `medium(rel target)` yields the workbook image id of a media part, or undefined when it is missing.
 */
export function parseDrawing(
  ws: Worksheet, xml: string, rels: DrawingRels, medium: (target: string) => number | undefined, limits: Readonly<XlsxLimits>,
): { images: Image[]; lost: string[] } {
  const root = parseXml(xml);
  const images: Image[] = [];
  const lost = new Set<string>();
  if (root.children.length > limits.maxDrawingObjects) throw new XlsxLimitError('maxDrawingObjects', `drawing holds more than ${limits.maxDrawingObjects} objects`);
  for (const a of root.children) {
    if (a.name !== 'oneCellAnchor' && a.name !== 'twoCellAnchor') {
      lost.add(a.name === 'absoluteAnchor' ? 'absolute-anchored drawing objects' : 'drawing objects');
      continue;
    }
    const pic = child(a, 'pic');
    if (a.children.some((c) => !['from', 'to', 'ext', 'pic', 'clientData'].includes(c.name)) || !pic) {
      lost.add('drawings (charts or shapes)');
      continue;
    }
    const blip = child(child(pic, 'blipFill') ?? pic, 'blip');
    const emb = blip?.attrs.embed ? rels.get(blip.attrs.embed) : undefined;
    const id = emb && !emb.external ? medium(emb.target) : undefined;
    if (id === undefined) {
      lost.add('linked or missing pictures');
      continue;
    }
    const cnv = child(child(pic, 'nvPicPr') ?? pic, 'cNvPr');
    const hl = cnv && child(cnv, 'hlinkClick');
    const hr = hl?.attrs.id ? rels.get(hl.attrs.id) : undefined;
    const ext = child(a, 'ext');
    if (a.name === 'oneCellAnchor' ? !ext : !child(a, 'to')) {
      lost.add('drawings (charts or shapes)');
      continue;
    }
    const range: ImageRange = {
      tl: anchorOf(ws, child(a, 'from')),
      br: a.name === 'twoCellAnchor' ? anchorOf(ws, child(a, 'to')) : undefined,
      ext: ext && a.name === 'oneCellAnchor' ? { width: (parseInt(ext.attrs.cx ?? '0', 10) || 0) / EMU_PER_PIXEL, height: (parseInt(ext.attrs.cy ?? '0', 10) || 0) / EMU_PER_PIXEL } : undefined,
      // attrs are keyed by local name: this reads the real twoCellAnchor editAs and the private xn:editAs of a oneCellAnchor
      editAs: a.attrs.editAs || 'oneCell',
    };
    if (hr) range.hyperlinks = { hyperlink: hr.target, ...(hl?.attrs.tooltip !== undefined ? { tooltip: hl.attrs.tooltip } : {}) };
    const img = new Image(ws, 'image', id, range as never);
    images.push(img);
  }
  return { images, lost: [...lost] };
}
