import { dialectOfText, rootPrefix, rootTag, type Dialect } from './dialect.js';
import { decodeEntities, startTags } from './xmlRead.js';
import type { ZipReader } from './unzip.js';
import { DEFAULT_LIMITS, XlsxLimitError, type XlsxLimits } from './limits.js';

export interface SheetRef {
  name: string;
  /** Zip entry name of the worksheet part. */
  path: string;
}
export interface WorkbookParts {
  sheets: SheetRef[];
  date1904: boolean;
  /** Zip entry names; the conventional locations when the workbook relationships name none. */
  sharedStringsPath: string;
  stylesPath: string;
}

const DEFAULT_WORKBOOK = 'xl/workbook.xml';
const REL_NAMESPACES = new Set(['http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'http://purl.oclc.org/ooxml/officeDocument/relationships']);

const decoder = new TextDecoder();

const WHITESPACE = /\s/;
const OFFICE_DOCUMENT = ['officeDocument'];
const PART_KINDS = ['worksheet', 'sharedStrings', 'styles'];

/** ASCII character classes: 0 name character, 1 whitespace, 2 `=` `/` quotes and angle brackets (neither whitespace nor part of a name). */
const CLASS = new Uint8Array(128);
for (let c = 9; c <= 13; c++) CLASS[c] = 1;
CLASS[32] = 1;
for (const ch of '=/"\'<>') CLASS[ch.charCodeAt(0)] = 2;

/** Whitespace beyond ASCII is whatever `\s` says (NBSP, U+FEFF, line separators ...). */
function isSpace(c: number): boolean {
  return c < 128 ? CLASS[c] === 1 : WHITESPACE.test(String.fromCharCode(c));
}

/** A name character: anything but whitespace, `=`, `/`, quotes and angle brackets. */
function isNameChar(c: number): boolean {
  return c < 128 ? CLASS[c] === 0 : !WHITESPACE.test(String.fromCharCode(c));
}

/** Every attribute of a start tag, the first occurrence of a name winning. Used on the few tags of the small parts, never per cell.
 * One forward pass: each character is looked at a bounded number of times, whatever the tag holds. */
export function attributes(tag: string): Map<string, string> {
  const out = new Map<string, string>();
  const n = tag.length;
  let i = 0;
  while (i < n) {
    if (!isNameChar(tag.charCodeAt(i))) {
      i++;
      continue;
    }
    const nameStart = i;
    while (i < n && isNameChar(tag.charCodeAt(i))) i++;
    const nameEnd = i;
    let j = i;
    while (j < n && isSpace(tag.charCodeAt(j))) j++;
    if (tag.charCodeAt(j) !== 61) continue;
    j++;
    while (j < n && isSpace(tag.charCodeAt(j))) j++;
    const quote = tag.charCodeAt(j);
    if (quote !== 34 && quote !== 39) continue;
    const close = tag.indexOf(quote === 34 ? '"' : "'", j + 1);
    if (close < 0) continue;
    const name = tag.slice(nameStart, nameEnd);
    if (!out.has(name)) out.set(name, tag.slice(j + 1, close));
    i = close + 1;
  }
  return out;
}

/** Entry lookup that forgives case and `\` separators, which some writers produce and most readers accept. */
class Parts {
  private folded: Map<string, string> | undefined;
  private readonly exact: Set<string>;
  constructor(private readonly zip: ZipReader) {
    this.exact = new Set(zip.names);
  }

  find(...candidates: string[]): string | undefined {
    for (const c of candidates) if (this.exact.has(c)) return c;
    this.folded ??= new Map(this.zip.names.map((n) => [n.replace(/\\/g, '/').toLowerCase(), n]));
    for (const c of candidates) {
      const hit = this.folded.get(c.toLowerCase());
      if (hit !== undefined) return hit;
    }
    return undefined;
  }

  text(name: string): string | undefined {
    const found = this.find(name);
    return found === undefined ? undefined : decoder.decode(this.zip.get(found)!);
  }
}

export interface Relationship {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

/**
 * The relationships of a `.rels` part, in document order. `needles` is a sound pre-filter, not a different meaning: a tag that contains none of them
 * cannot carry a `Type` that ends with one of the wanted suffixes, so its attributes are never parsed. Lazy, so a caller that stops early stops parsing.
 */
export function* readRelationships(xml: string, needles?: readonly string[]): Generator<Relationship> {
  for (const tag of startTags(xml, `${prefixed(rootPrefix(xml))}Relationship`)) {
    if (needles !== undefined && !needles.some((n) => tag.includes(n))) continue;
    const a = attributes(tag);
    const id = a.get('Id');
    const target = a.get('Target');
    const type = a.get('Type');
    if (id !== undefined && target !== undefined && type !== undefined) yield { id, type, target: decodeEntities(target), external: a.get('TargetMode') === 'External' };
  }
}

const prefixed = (prefix: string): string => (prefix === '' ? '' : `${prefix}:`);

const SIMPLE_TARGET = /^[\w-]+(?:\/[\w-]+)*\.\w+$/;
const SIMPLE_FOLDER = /^(?:[\w-]+\/)*$/;

/** Candidates for the zip entry a target names: relative to the source part's folder unless it starts with `/`. Both the percent-decoded and the literal spelling, because an entry may really contain `%`. */
export function targetPaths(folder: string, target: string): string[] {
  // Letters, digits, `_`, `-` and `/` only, with a plain extension: nothing to unescape, no `.` or `..` segment, so the general walk below would hand back the input.
  if (SIMPLE_TARGET.test(target) && SIMPLE_FOLDER.test(folder)) return [folder + target];
  const slashed = target.replace(/\\/g, '/');
  const spellings = [slashed];
  try {
    const decoded = decodeURIComponent(slashed);
    if (decoded !== slashed) spellings.unshift(decoded);
  } catch {
    // not valid percent-encoding: the literal spelling stands
  }
  return spellings.map((t) => {
    const parts: string[] = [];
    for (const seg of (t.startsWith('/') ? t.slice(1) : `${folder}${t}`).split('/')) {
      if (seg === '..') parts.pop();
      else if (seg !== '.' && seg !== '') parts.push(seg);
    }
    return parts.join('/');
  });
}

function folderOf(path: string): string {
  return path.slice(0, path.lastIndexOf('/') + 1);
}

/** The workbook is whatever the package relationships call the office document; `xl/workbook.xml` only when they do not say. */
function findWorkbook(parts: Parts): string {
  const rels = parts.text('_rels/.rels');
  if (rels !== undefined) {
    for (const r of readRelationships(rels, OFFICE_DOCUMENT)) {
      if (r.external || !r.type.endsWith('/officeDocument')) continue;
      const found = parts.find(...targetPaths('', r.target));
      if (found !== undefined) return found;
    }
  }
  return DEFAULT_WORKBOOK;
}

/** The relationship id of a `<sheet>`: the attribute whose prefix resolves to the relationships namespace, wherever the document declared it. */
function relationshipId(a: Map<string, string>, inherited: Map<string, string>): string | undefined {
  for (const [name, value] of a) {
    const colon = name.indexOf(':');
    if (colon <= 0 || name.slice(colon + 1) !== 'id') continue;
    const prefix = name.slice(0, colon);
    const uri = a.get(`xmlns:${prefix}`) ?? inherited.get(prefix);
    if (uri !== undefined && REL_NAMESPACES.has(uri)) return value;
  }
  return undefined;
}

/** `xmlns:p="uri"` declarations of the root and `<sheets>` tags, which are the ancestors of every `<sheet>`. */
function declarations(...tags: string[]): Map<string, string> {
  const out = new Map<string, string>();
  for (const tag of tags) for (const [name, value] of attributes(tag)) if (name.startsWith('xmlns:')) out.set(name.slice(6), value);
  return out;
}

export function readWorkbookParts(zip: ZipReader, limits: Readonly<XlsxLimits> = DEFAULT_LIMITS): WorkbookParts {
  const parts = new Parts(zip);
  const wbPath = findWorkbook(parts);
  const wb = parts.text(wbPath);
  if (wb === undefined) throw new Error(`invalid xlsx: missing ${wbPath}`);
  const folder = folderOf(wbPath);
  const relsPath = `${folder}_rels/${wbPath.slice(folder.length)}.rels`;
  const rels = parts.text(relsPath);
  if (rels === undefined) throw new Error(`invalid xlsx: missing ${relsPath}`);

  const worksheets = new Map<string, string>();
  let sharedStringsPath = 'xl/sharedStrings.xml';
  let stylesPath = 'xl/styles.xml';
  for (const r of readRelationships(rels, PART_KINDS)) {
    if (r.external) continue;
    // chartsheets and dialogsheets are listed in <sheets> too but hold no cells; themes, drawings and the rest never need a path resolved
    const kind = r.type.endsWith('/worksheet') ? 0 : r.type.endsWith('/sharedStrings') ? 1 : r.type.endsWith('/styles') ? 2 : -1;
    if (kind < 0) continue;
    const candidates = targetPaths(folder, r.target);
    const resolved = parts.find(...candidates) ?? candidates[0]!;
    if (kind === 0) worksheets.set(r.id, resolved);
    else if (kind === 1) sharedStringsPath = resolved;
    else stylesPath = resolved;
  }

  const d: Dialect = dialectOfText(wb);
  const sheetsTag = startTags(wb, d.sheets).next();
  const inherited = declarations(rootTag(wb), sheetsTag.done ? '' : sheetsTag.value);
  const sheets: SheetRef[] = [];
  for (const tag of startTags(wb, d.sheet)) {
    const a = attributes(tag);
    const name = a.get('name');
    const rid = relationshipId(a, inherited);
    const path = rid === undefined ? undefined : worksheets.get(rid);
    if (name !== undefined && path !== undefined) {
      if (sheets.length >= limits.maxSheets) throw new XlsxLimitError('maxSheets', `workbook lists more than ${limits.maxSheets} sheets`);
      sheets.push({ name: decodeEntities(name), path });
    }
  }

  const pr = startTags(wb, d.workbookPr).next();
  const flag = pr.done ? undefined : attributes(pr.value).get('date1904');
  return { sheets, date1904: flag === '1' || flag === 'true', sharedStringsPath, stylesPath };
}
