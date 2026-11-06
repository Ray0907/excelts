import { strToU8 } from 'fflate';
import type { ZipEntry } from '../zip.js';
import { drawingParts } from './drawing.js';
import { imageType, mediumBytes } from './images.js';
import { commentsXml, vmlXml } from './notes.js';
import type { SheetWriter } from './writeSheet.js';

const CT = 'application/vnd.openxmlformats-officedocument';

/** Parts that live beside the sheets (pictures with their drawings, notes with their legacy drawings), gathered while the sheets are written. */
export class Extras {
  private readonly parts: ZipEntry[] = [];
  private readonly types: string[] = [];
  private readonly defaults = new Set<string>();
  private readonly media = new Map<number, string>();
  private drawings = 0;
  private notes = 0;

  private defaultType(ext: string, type: string): void {
    if (this.defaults.has(ext)) return;
    this.defaults.add(ext);
    this.types.push(`<Default Extension="${ext}" ContentType="${type}"/>`);
  }

  private add(name: string, xml: string, type?: string): void {
    this.parts.push({ name, data: strToU8(xml) });
    if (type) this.types.push(`<Override PartName="/${name}" ContentType="${type}"/>`);
  }

  private mediaPath(w: SheetWriter, id: number): string {
    let path = this.media.get(id);
    if (!path) {
      const m = w.ws.workbook.media[id];
      if (!m) throw new Error(`image ${id} is not in the workbook; add it with workbook.addImage first`);
      const ext = m.extension.toLowerCase();
      this.defaultType(ext, imageType(ext));
      this.media.set(id, (path = `xl/media/image${id + 1}.${ext}`));
    }
    return path;
  }

  /** Writes what a sheet owns beyond its cells and returns the elements for the end of the worksheet (`drawing`, `legacyDrawing`, `picture`). */
  sheet(w: SheetWriter): string {
    const ws = w.ws;
    let out = '';
    const images = ws._media.filter((m) => m.type === 'image');
    if (images.length) {
      const n = ++this.drawings;
      const d = drawingParts(images, (id) => `../${this.mediaPath(w, id).slice(3)}`);
      this.add(`xl/drawings/drawing${n}.xml`, d.xml, `${CT}.drawing+xml`);
      this.add(`xl/drawings/_rels/drawing${n}.xml.rels`, d.rels);
      out += `<drawing r:id="${w.addRel('drawing', `../drawings/drawing${n}.xml`)}"/>`;
    }
    if (w.notes.length) {
      const n = ++this.notes;
      const list = [...w.notes].sort((a, b) => a.row - b.row || a.col - b.col);
      this.add(`xl/comments${n}.xml`, commentsXml(list), `${CT}.spreadsheetml.comments+xml`);
      this.add(`xl/drawings/vmlDrawing${n}.vml`, vmlXml(list, n));
      this.defaultType('vml', `${CT}.vmlDrawing`);
      w.addRel('comments', `../comments${n}.xml`);
      out += `<legacyDrawing r:id="${w.addRel('vmlDrawing', `../drawings/vmlDrawing${n}.vml`)}"/>`;
    }
    const bg = ws._media.find((m) => m.type === 'background');
    if (bg) out += `<picture r:id="${w.addRel('image', `../${this.mediaPath(w, bg.imageId).slice(3)}`)}"/>`;
    return out;
  }

  /** Every part gathered so far, including the pictures themselves. */
  finish(wb: { media: readonly import('./images.js').Medium[] }): ZipEntry[] {
    const media: ZipEntry[] = [];
    for (const [id, name] of this.media) media.push({ name, data: mediumBytes(wb.media[id]!) });
    return [...this.parts, ...media];
  }

  /** `Default` and `Override` entries for [Content_Types].xml. */
  contentTypes(): string[] {
    return this.types;
  }
}
