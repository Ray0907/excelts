import { CsvIO } from './csv.js';
import { DefinedNames } from './definedNames.js';
import { asBuffer, readMediaFiles, type Medium } from './images.js';
import { NotSupportedError } from './errors.js';
import { StyleBook } from './styleBook.js';
import { Worksheet, type WorksheetOptions } from './worksheet.js';
import { writeWorkbook, type WriteOptions } from './writeBook.js';
import { loadWorkbook, type LoadOptions } from './load.js';
import type { ReadOptions } from '../limits.js';

export interface BookView {
  x?: number; y?: number; width?: number; height?: number; firstSheet?: number; activeTab?: number;
  visibility?: string; [k: string]: unknown;
}
type Bytes = Uint8Array | ArrayBuffer | string;

function toBytes(data: Bytes, options?: LoadOptions): Uint8Array {
  if (typeof data === 'string') {
    if (options?.base64 === false) throw new TypeError('load(string) expects base64 data');
    return Uint8Array.from(atob(data), (c) => c.charCodeAt(0));
  }
  if (data instanceof Uint8Array) return data;
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  if (ArrayBuffer.isView(data)) return new Uint8Array((data as ArrayBufferView).buffer, (data as ArrayBufferView).byteOffset, (data as ArrayBufferView).byteLength);
  throw new Error("Can't read the data of 'the loaded zip file'. Is it in a supported JavaScript type (String, Blob, ArrayBuffer, etc) ?");
}

async function readStream(stream: AsyncIterable<Uint8Array | string>): Promise<Uint8Array> {
  const parts: Uint8Array[] = [];
  for await (const c of stream) parts.push(typeof c === 'string' ? new TextEncoder().encode(c) : c);
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}

class XlsxIO {
  constructor(readonly _wb: Workbook) {}

  async writeBuffer(options?: WriteOptions): Promise<Uint8Array> {
    await readMediaFiles(this._wb);
    return asBuffer(writeWorkbook(this._wb, options));
  }
  async writeFile(filename: string, options?: WriteOptions): Promise<void> {
    await readMediaFiles(this._wb);
    const bytes = writeWorkbook(this._wb, options);
    const fs = await import('node:fs/promises');
    await fs.writeFile(filename, bytes);
  }
  async write(stream: { write(b: Uint8Array, cb?: (e?: Error | null) => void): unknown; end(cb?: () => void): unknown; once?: (e: string, f: (e?: Error) => void) => unknown }, options?: WriteOptions): Promise<void> {
    await readMediaFiles(this._wb);
    const bytes = asBuffer(writeWorkbook(this._wb, options));
    await new Promise<void>((resolve, reject) => {
      stream.once?.('error', reject);
      stream.write(bytes, (e) => {
        if (e) reject(e);
        else stream.end(() => resolve());
      });
    });
  }
  async load(data: Bytes, options?: LoadOptions & ReadOptions): Promise<Workbook> {
    loadWorkbook(this._wb, toBytes(data, options), options);
    return this._wb;
  }
  async read(stream: AsyncIterable<Uint8Array | string>, options?: LoadOptions & ReadOptions): Promise<Workbook> {
    return this.load(await readStream(stream), options);
  }
  async readFile(filename: string, options?: LoadOptions & ReadOptions): Promise<Workbook> {
    const fs = await import('node:fs/promises');
    let bytes: Uint8Array;
    try {
      bytes = await fs.readFile(filename);
    } catch (e) {
      if ((e as { code?: string }).code === 'ENOENT') throw new Error(`File not found: ${filename}`);
      throw e;
    }
    return this.load(bytes, options);
  }
}

export class Workbook {
  creator = '';
  lastModifiedBy = '';
  created: Date = new Date();
  modified: Date = new Date();
  lastPrinted?: Date;
  title?: string; subject?: string; keywords?: string; category?: string; description?: string;
  company?: string; manager?: string; language?: string; revision?: number; contentStatus?: string;
  properties: { date1904?: boolean; [k: string]: unknown } = {};
  calcProperties: { fullCalcOnLoad?: boolean; [k: string]: unknown } = {};
  views: BookView[] = [];
  /** @internal */ _worksheets: (Worksheet | undefined)[] = [];
  /** @internal */ _styles = new StyleBook();
  /** @internal what a loaded file carried that cannot be written back */
  _unsupported: string[] = [];
  readonly xlsx = new XlsxIO(this);

  get worksheets(): Worksheet[] {
    return this._worksheets.slice(1).filter((w): w is Worksheet => !!w).sort((a, b) => a.orderNo - b.orderNo);
  }
  /** @internal */
  _firstVisible(): Worksheet | undefined {
    return this.worksheets.find((w) => w.state === 'visible');
  }
  get nextId(): number {
    for (let i = 1; i < this._worksheets.length; i++) if (!this._worksheets[i]) return i;
    return this._worksheets.length || 1;
  }
  addWorksheet(name?: string, options?: WorksheetOptions | string): Worksheet {
    const id = this.nextId;
    const opts: WorksheetOptions = typeof options === 'string' ? { properties: { tabColor: { argb: options } } as never } : (options ?? {});
    let last = 0;
    for (const w of this._worksheets) if (w && w.orderNo > last) last = w.orderNo;
    const ws = new Worksheet(this, { ...opts, id, name: name ?? `sheet${id}`, orderNo: last + 1 });
    this._worksheets[id] = ws;
    return ws;
  }
  removeWorksheetEx(ws: Worksheet): void {
    delete this._worksheets[ws.id];
  }
  removeWorksheet(id: number | string): void {
    this.getWorksheet(id)?.destroy();
  }
  getWorksheet(id?: number | string): Worksheet | undefined {
    if (id === undefined) return this._worksheets.find(Boolean);
    if (typeof id === 'number') return this._worksheets[id];
    if (typeof id === 'string') return this._worksheets.find((w) => w && w.name === id);
    return undefined;
  }
  eachSheet(cb: (ws: Worksheet, id: number) => void): void {
    this.worksheets.forEach((s) => cb(s, s.id));
  }
  clearThemes(): void {}
  /** @internal */ _names?: DefinedNames;
  get definedNames(): DefinedNames { return (this._names ??= new DefinedNames()); }
  get pivotTables(): never { throw new NotSupportedError('workbook pivot tables'); }
  /** @internal */ _csv?: CsvIO;
  get csv(): CsvIO { return (this._csv ??= new CsvIO(this)); }
  /** Pictures of the workbook; `addImage` returns an index into this list. */
  media: Medium[] = [];
  addImage(image: Omit<Medium, 'type'>): number {
    this.media.push({ ...image, type: 'image' });
    return this.media.length - 1;
  }
  getImage(id: number): Medium | undefined { return this.media[id]; }
  get model(): never { throw new NotSupportedError('workbook.model (internal model)'); }
  set model(_v: unknown) { throw new NotSupportedError('workbook.model (internal model)'); }
}


