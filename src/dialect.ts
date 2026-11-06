import { viewOf } from './bytes.js';

const encoder = new TextEncoder();
const decoder = new TextDecoder();

/**
 * Every element name a reader searches for, spelled the way one document spells the main spreadsheetml namespace:
 * `<si>` with a default namespace, `<x:si>` when the root declares `xmlns:x`. Built once per part.
 * Attribute names never carry the prefix, so only element names live here.
 */
export class Dialect {
  // Element names, for startTags().
  readonly sheet: string;
  readonly workbookPr: string;
  readonly numFmt: string;
  readonly xf: string;
  readonly sheets: string;
  // Start-tag openings and fragments for string searches.
  readonly cellXfs: string;
  readonly cellXfsEnd: string;
  readonly si: string;
  readonly siEnd: string;
  readonly siTShort: string;
  readonly tSiEnd: string;
  readonly t: string;
  readonly tEnd: string;
  readonly rPh: string;
  readonly rPhEnd: string;

  constructor(readonly prefix: string) {
    const p = prefix === '' ? '' : `${prefix}:`;
    this.sheet = `${p}sheet`;
    this.workbookPr = `${p}workbookPr`;
    this.numFmt = `${p}numFmt`;
    this.xf = `${p}xf`;
    this.sheets = `${p}sheets`;
    this.cellXfs = `<${p}cellXfs`;
    this.cellXfsEnd = `</${p}cellXfs>`;
    this.si = `<${p}si`;
    this.siEnd = `</${p}si>`;
    this.siTShort = `<${p}t>`;
    this.tSiEnd = `</${p}t></${p}si>`;
    this.t = `<${p}t`;
    this.tEnd = `</${p}t>`;
    this.rPh = `<${p}rPh`;
    this.rPhEnd = `</${p}rPh>`;
  }
}

/** The unprefixed dialect, by far the commonest; created once at module load so the hot path never builds anything. */
export const PLAIN = new Dialect('');
const prefixed = new Map<string, Dialect>();

export function dialectOf(prefix: string): Dialect {
  if (prefix === '') return PLAIN;
  let d = prefixed.get(prefix);
  if (!d) prefixed.set(prefix, (d = new Dialect(prefix)));
  return d;
}

/** Start tag text (`<x:workbook xmlns:x="..."`, without `>`) of the first element, which is always in the main namespace; a BOM, declaration, comment or doctype may precede it. */
export function rootTag(xml: string): string {
  let p = xml.indexOf('<');
  while (p >= 0) {
    const c = xml.charCodeAt(p + 1);
    if (c !== 63 && c !== 33) break; // `<?` and `<!`
    p = xml.indexOf('<', p + 2);
  }
  if (p < 0) return '';
  const gt = xml.indexOf('>', p);
  return xml.slice(p, gt < 0 ? xml.length : gt);
}

export function rootPrefix(xml: string): string {
  const tag = rootTag(xml);
  for (let q = 1; q < tag.length; q++) {
    const c = tag.charCodeAt(q);
    if (c === 58) return tag.slice(1, q);
    if (c <= 32 || c === 47) break;
  }
  return '';
}

export function dialectOfText(xml: string): Dialect {
  return dialectOf(rootPrefix(xml.slice(0, 2048)));
}

export function rootPrefixOf(bytes: Uint8Array): string {
  return rootPrefix(decoder.decode(bytes.subarray(0, 2048)));
}

/** Copy of `bytes` with the prefix removed from every element name (`<x:c` becomes `<c`, `</x:c` becomes `</c`).
 * A raw `<` cannot occur in attribute values or text, so only markup matches; a CDATA section quoting `<x:` would be rewritten too. */
export function stripPrefix(bytes: Uint8Array, prefix: string): Uint8Array {
  const open = encoder.encode(`<${prefix}:`);
  const close = encoder.encode(`</${prefix}:`);
  const drop = open.length - 1;
  const view = viewOf(bytes);
  const out = new Uint8Array(bytes.length);
  let n = 0;
  let last = 0;
  let a = view.indexOfSeq(open, 0);
  let b = view.indexOfSeq(close, 0);
  while (a >= 0 || b >= 0) {
    const isOpen = b < 0 || (a >= 0 && a < b);
    const keep = (isOpen ? a : b) + (isOpen ? 1 : 2);
    out.set(bytes.subarray(last, keep), n);
    n += keep - last;
    last = keep + drop;
    if (isOpen) a = view.indexOfSeq(open, last);
    else b = view.indexOfSeq(close, last);
  }
  out.set(bytes.subarray(last), n);
  return out.subarray(0, n + bytes.length - last);
}
