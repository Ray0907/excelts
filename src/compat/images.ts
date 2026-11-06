import { decodeAddress } from './colCache.js';
import { Range } from './range.js';
import type { Workbook } from './workbook.js';
import type { Worksheet } from './worksheet.js';

export function asBuffer(u: Uint8Array): Uint8Array {
  const B = (globalThis as { Buffer?: { from(b: ArrayBufferLike, o: number, l: number): Uint8Array } }).Buffer;
  return B ? B.from(u.buffer, u.byteOffset, u.length) : u;
}

/** A picture held by the workbook; `buffer` of a loaded one is inflated on first use. */
export interface Medium {
  type: 'image';
  name?: string;
  extension: string;
  buffer?: Uint8Array;
  base64?: string;
  filename?: string;
}

const TYPES: Record<string, string> = {
  png: 'image/png', jpeg: 'image/jpeg', jpg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', tif: 'image/tiff', tiff: 'image/tiff',
  svg: 'image/svg+xml', webp: 'image/webp', emf: 'image/x-emf', wmf: 'image/x-wmf', ico: 'image/x-icon',
};

export function imageType(extension: string): string {
  const e = String(extension ?? '').toLowerCase();
  if (!/^[a-z0-9]{2,5}$/.test(e)) throw new TypeError(`unsupported image extension ${JSON.stringify(extension)}`);
  return TYPES[e] ?? `image/${e}`;
}

/** Medium bytes for writing; `filename` media must have been read by `readMediaFiles` first. */
export function mediumBytes(m: Medium): Uint8Array {
  if (m.buffer) return m.buffer;
  if (m.base64) {
    const s = m.base64;
    return Uint8Array.from(atob(s.slice(s.indexOf(',') + 1)), (c) => c.charCodeAt(0));
  }
  const read = (m as { _read?: Uint8Array })._read;
  if (read) return read;
  throw new Error('Unsupported media');
}

/** Reads the files named by `filename` media, so the synchronous writer can embed them. Rejects when one cannot be read. */
export async function readMediaFiles(wb: Workbook): Promise<void> {
  for (const m of wb.media) {
    if (m.buffer || m.base64 || !m.filename) continue;
    const fs = await import('node:fs/promises');
    Object.defineProperty(m, '_read', { value: new Uint8Array(await fs.readFile(m.filename)), enumerable: false, configurable: true });
  }
}

export interface AnchorLike { col?: number; row?: number; nativeCol?: number; nativeColOff?: number; nativeRow?: number; nativeRowOff?: number }

/** One corner of a picture: a native cell plus an offset in EMU, readable as a fractional column/row. */
export class Anchor {
  nativeCol = 0;
  nativeColOff = 0;
  nativeRow = 0;
  nativeRowOff = 0;

  constructor(
    readonly worksheet: Worksheet | undefined,
    address?: string | AnchorLike | null,
    offset = 0,
  ) {
    if (typeof address === 'string') {
      const d = decodeAddress(address);
      this.nativeCol = d.col + offset;
      this.nativeRow = d.row + offset;
    } else if (address && address.nativeCol !== undefined) {
      this.nativeCol = address.nativeCol || 0;
      this.nativeColOff = address.nativeColOff || 0;
      this.nativeRow = address.nativeRow || 0;
      this.nativeRowOff = address.nativeRowOff || 0;
    } else if (address && address.col !== undefined) {
      this.col = address.col + offset;
      this.row = (address.row ?? 0) + offset;
    }
  }

  get colWidth(): number {
    const c = this.worksheet?._columns?.[this.nativeCol];
    return c && c.isCustomWidth ? Math.floor(c.width! * 10000) : 640000;
  }
  get rowHeight(): number {
    const h = this.worksheet?.findRow(this.nativeRow + 1)?.height;
    return h ? Math.floor(h * 10000) : 180000;
  }
  get col(): number { return this.nativeCol + Math.min(this.colWidth - 1, this.nativeColOff) / this.colWidth; }
  set col(v: number) {
    this.nativeCol = Math.floor(v);
    this.nativeColOff = Math.floor((v - this.nativeCol) * this.colWidth);
  }
  get row(): number { return this.nativeRow + Math.min(this.rowHeight - 1, this.nativeRowOff) / this.rowHeight; }
  set row(v: number) {
    this.nativeRow = Math.floor(v);
    this.nativeRowOff = Math.floor((v - this.nativeRow) * this.rowHeight);
  }
  get model(): { nativeCol: number; nativeColOff: number; nativeRow: number; nativeRowOff: number } {
    return { nativeCol: this.nativeCol, nativeColOff: this.nativeColOff, nativeRow: this.nativeRow, nativeRowOff: this.nativeRowOff };
  }
}

export interface ImageRangeInput {
  tl: AnchorLike | string;
  br?: AnchorLike | string;
  ext?: { width: number; height: number };
  editAs?: string;
  hyperlinks?: { hyperlink: string; tooltip?: string };
}
export interface ImageRange { tl: Anchor; br?: Anchor; ext?: { width: number; height: number }; editAs?: string; hyperlinks?: { hyperlink: string; tooltip?: string } }

/** What `worksheet.getImages()` returns, and what the sheet keeps for a background picture. */
export class Image {
  range?: ImageRange;

  constructor(
    readonly worksheet: Worksheet,
    readonly type: 'image' | 'background',
    public imageId: number,
    range?: string | ImageRangeInput | ImageRange,
    hyperlinks?: ImageRange['hyperlinks'],
  ) {
    if (type !== 'image') return;
    if (typeof range === 'string') {
      const d = new Range(range);
      this.range = {
        tl: new Anchor(worksheet, { col: d.left, row: d.top }, -1),
        br: new Anchor(worksheet, { col: d.right, row: d.bottom }, 0),
        editAs: 'oneCell',
      };
    } else if (range && range.tl && (range.br || range.ext)) {
      this.range = {
        tl: new Anchor(worksheet, range.tl as AnchorLike, 0),
        br: range.br ? new Anchor(worksheet, range.br as AnchorLike, 0) : undefined,
        ext: range.ext,
        editAs: range.editAs,
        hyperlinks: hyperlinks ?? range.hyperlinks,
      };
    } else {
      throw new TypeError('addImage needs a range string such as "B2:D6", or { tl, br } or { tl, ext }');
    }
  }
}
