import { SharedStrings } from '../sharedStrings.js';
import { concat, defaultCompressor } from '../zip.js';
import { NotSupportedError } from './errors.js';
import { readMediaFiles } from './images.js';
import type { Row } from './row.js';
import { Workbook } from './workbook.js';
import { Worksheet, type WorksheetOptions } from './worksheet.js';
import { assemble, assertWritable, collectExtras, type WriteOptions } from './writeBook.js';
import { NS_REL, SheetWriter, type WriteCtx } from './writeSheet.js';

/** A worksheet whose committed rows are serialised and released at once, so memory stays flat. */
class StreamSheet extends Worksheet {
  /** @internal */ _w!: SheetWriter;
  /** @internal */ _committed = 0;
  /** @internal */ _started = false;
  /** @internal */ _finished = false;

  /** @internal */
  _init(ctx: WriteCtx): void {
    this._w = new SheetWriter(this, ctx, true);
  }
  override get _lastRowNumber(): number {
    let n = this._rows.length;
    while (n > this._committed && this._rows[n - 1] === undefined) n--;
    return Math.max(n, this._committed);
  }
  override _commitRow(row: Row): void {
    if (this._finished) throw new Error('Cannot commit rows to a committed worksheet');
    if (!this._started) {
      this._started = true;
      this._w.head();
    }
    const batch: Row[] = [];
    for (let n = this._committed + 1; n <= row.number; n++) {
      const r = this._rows[n - 1];
      if (r) batch.push(r);
      this._rows[n - 1] = undefined;
    }
    this._w.rows(batch);
    if (row.number > this._committed) this._committed = row.number;
  }
  commit(): Promise<void> {
    if (this._finished) return Promise.resolve();
    if (!this._started) {
      this._started = true;
      this._w.head();
    }
    const batch: Row[] = [];
    for (let n = this._committed + 1; n <= this._rows.length; n++) {
      const r = this._rows[n - 1];
      if (r) batch.push(r);
      this._rows[n - 1] = undefined;
    }
    this._w.rows(batch);
    this._committed = this._rows.length;
    this._w.finish();
    this._finished = true;
    return Promise.resolve();
  }
}

export interface WorkbookWriterOptions {
  filename?: string;
  stream?: { write(b: Uint8Array, cb?: (e?: Error | null) => void): unknown; end(cb?: () => void): unknown };
  useSharedStrings?: boolean;
  useStyles?: boolean;
  zip?: WriteOptions['zip'];
}

/** `stream.xlsx.WorkbookWriter`: rows are serialised when committed. Strings always go through the shared string table. */
class WorkbookWriter extends Workbook {
  readonly _ctx: WriteCtx;
  constructor(readonly _options: WorkbookWriterOptions = {}) {
    super();
    this._ctx = { strings: new SharedStrings(), book: this._styles, date1904: false };
  }
  override addWorksheet(name?: string, options?: WorksheetOptions | string): Worksheet {
    const id = this.nextId;
    const opts: WorksheetOptions = typeof options === 'string' ? { properties: { tabColor: { argb: options } } as never } : (options ?? {});
    let last = 0;
    for (const w of this._worksheets) if (w && w.orderNo > last) last = w.orderNo;
    const ws = new StreamSheet(this, { ...opts, id, name: name ?? `sheet${id}`, orderNo: last + 1 });
    ws._init(this._ctx);
    this._worksheets[id] = ws;
    return ws;
  }
  addStyles(): Promise<void> { return Promise.resolve(); }
  async commit(): Promise<void> {
    assertWritable(this, undefined);
    await readMediaFiles(this);
    this._ctx.date1904 = !!this.properties.date1904;
    const sheets = this.worksheets as StreamSheet[];
    const parts: { name: string; data: Uint8Array }[] = [];
    const rels: string[] = [];
    const overrides: string[] = [];
    for (let i = 0; i < sheets.length; i++) {
      const ws = sheets[i]!;
      await ws.commit();
      const n = i + 1;
      parts.push({ name: `xl/worksheets/sheet${n}.xml`, data: concat(ws._w.chunks) });
      const r = ws._w.relsXml();
      if (r) parts.push({ name: `xl/worksheets/_rels/sheet${n}.xml.rels`, data: new TextEncoder().encode(r) });
      rels.push(`<Relationship Id="rId${n}" Type="${NS_REL}/worksheet" Target="worksheets/sheet${n}.xml"/>`);
      overrides.push(`<Override PartName="/xl/worksheets/sheet${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
    }
    collectExtras(this, this._ctx, parts, overrides);
    const bytes = assemble(this, sheets, this._ctx, parts, rels, overrides, { zip: this._options.zip }, defaultCompressor);
    if (this._options.filename) {
      const fs = await import('node:fs/promises');
      await fs.writeFile(this._options.filename, bytes);
    } else if (this._options.stream) {
      const s = this._options.stream;
      await new Promise<void>((resolve, reject) => s.write(bytes, (e) => (e ? reject(e) : s.end(() => resolve()))));
    }
  }
}

class WorkbookReader {
  constructor() {
    throw new NotSupportedError('stream.xlsx.WorkbookReader (streaming reader)', 'use Workbook.xlsx.readFile / load');
  }
}

export const stream = { xlsx: { WorkbookWriter, WorkbookReader } };
