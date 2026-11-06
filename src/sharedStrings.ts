import { XML_DECLARATION, escapeText } from './xml.js';

export const MAX_CELL_CHARS = 32767;

const SST_OPEN = '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const NEEDS_SPACE = /^\s|\s$|\n|\r/;
const SINK_CHUNK = 1 << 20;
/** Strings at least this long are written as three pieces so the item is never concatenated (and flattened) into one more copy. */
const LONG_ITEM = 512;

/** UTF-8 sink over big Buffers: a string is written straight into the current chunk, with no intermediate whole-table string. */
class ByteSink {
  private buf = Buffer.allocUnsafe(SINK_CHUNK);
  private pos = 0;
  private readonly full: Buffer[] = [];

  put(s: string): void {
    const need = s.length * 3; // UTF-8 never needs more than 3 bytes per UTF-16 unit
    if (this.pos + need > this.buf.length) {
      this.full.push(this.buf.subarray(0, this.pos));
      this.buf = Buffer.allocUnsafe(Math.max(SINK_CHUNK, need));
      this.pos = 0;
    }
    this.pos += this.buf.write(s, this.pos, 'utf8');
  }

  finish(): Uint8Array {
    const last = this.buf.subarray(0, this.pos);
    if (this.full.length === 0) return last;
    this.full.push(last);
    return Buffer.concat(this.full);
  }
}

export class SharedStrings {
  private readonly index = new Map<string, number>();
  private readonly list: string[] = [];
  private refs = 0;
  /** Entries that are rich text: index -> the `<r>` runs XML. */
  private readonly rich = new Map<number, string>();

  add(s: string): number {
    if (s.length > MAX_CELL_CHARS) {
      throw new RangeError(`string of ${s.length} chars exceeds the ${MAX_CELL_CHARS}-char cell limit`);
    }
    this.refs++;
    let i = this.index.get(s);
    if (i === undefined) {
      i = this.list.length;
      this.index.set(s, i);
      this.list.push(s);
    }
    return i;
  }

  /** Adds a rich-text item given as the XML inside `<si>`; equal runs share an entry. */
  addRich(runsXml: string): number {
    this.refs++;
    const key = `\u0000r:${runsXml}`;
    let i = this.index.get(key);
    if (i === undefined) {
      i = this.list.length;
      this.index.set(key, i);
      this.list.push(key);
      this.rich.set(i, runsXml);
    }
    return i;
  }

  /** The table as UTF-8: exactly the bytes of `toXml()`, written without building the whole document as one string where Node's Buffer exists. */
  toBytes(): Uint8Array {
    if (typeof Buffer !== 'function') return new TextEncoder().encode(this.toXml());
    const sink = new ByteSink();
    sink.put(XML_DECLARATION + `${SST_OPEN} count="${this.refs}" uniqueCount="${this.list.length}">`);
    for (let i = 0; i < this.list.length; i++) {
      const s = this.list[i]!;
      const r = this.rich.get(i);
      if (r !== undefined) {
        sink.put(`<si>${r}</si>`);
        continue;
      }
      const open = NEEDS_SPACE.test(s) ? '<si><t xml:space="preserve">' : '<si><t>';
      if (s.length < LONG_ITEM) {
        sink.put(`${open}${escapeText(s)}</t></si>`);
      } else {
        sink.put(open);
        sink.put(escapeText(s));
        sink.put('</t></si>');
      }
    }
    sink.put('</sst>');
    return sink.finish();
  }

  toXml(): string {
    const items = this.list.map((s, i) => {
      const r = this.rich.get(i);
      if (r !== undefined) return `<si>${r}</si>`;
      const text = escapeText(s);
      const space = NEEDS_SPACE.test(s) ? ' xml:space="preserve"' : '';
      return `<si><t${space}>${text}</t></si>`;
    });
    return (
      XML_DECLARATION +
      `${SST_OPEN} count="${this.refs}" uniqueCount="${this.list.length}">` +
      items.join('') +
      '</sst>'
    );
  }
}
