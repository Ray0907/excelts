import { dialectOfText } from './dialect.js';
import { decodeText, readText } from './xmlRead.js';
import { DEFAULT_LIMITS, XlsxLimitError, type XlsxLimits } from './limits.js';

/** Index-addressed table; an `<si>` is plain `<t>` or rich runs `<r><t>`, which concatenate. */
export function readSharedStrings(xml: string, limits: Readonly<XlsxLimits> = DEFAULT_LIMITS): string[] {
  const d = dialectOfText(xml);
  const si = d.si;
  const tShort = d.siTShort;
  const tShortLen = tShort.length;
  const tSiEnd = d.tSiEnd;
  const siEnd = d.siEnd;
  const siLen = si.length;
  const out: string[] = [];
  let chars = 0;
  const maxCount = limits.maxSharedStrings;
  const maxChars = limits.maxSharedStringBytes;
  const maxCell = limits.maxCellChars;
  // Counts and sizes are checked as each entry lands, so a table is refused before it is built, not after.
  const add = (text: string): void => {
    if (out.length >= maxCount) throw new XlsxLimitError('maxSharedStrings', `shared string table has more than ${maxCount} entries`);
    if (text.length > maxCell) throw new XlsxLimitError('maxCellChars', `shared string ${out.length} has ${text.length} characters`);
    chars += text.length;
    if (chars > maxChars) throw new XlsxLimitError('maxSharedStringBytes', `shared strings total more than ${maxChars} characters`);
    out.push(text);
  };
  const hasPhonetic = xml.includes(d.rPh);
  let p = xml.indexOf(si);
  while (p >= 0) {
    const c = xml.charCodeAt(p + siLen);
    if (c !== 62 && c !== 32 && c !== 47) {
      p = xml.indexOf(si, p + siLen);
      continue;
    }
    const gt = xml.indexOf('>', p);
    if (gt < 0) break;
    if (xml.charCodeAt(gt - 1) === 47) {
      add('');
      p = xml.indexOf(si, gt);
      continue;
    }
    if (gt === p + siLen && xml.startsWith(tShort, gt + 1)) {
      const from = gt + 1 + tShortLen;
      const lt = xml.indexOf('<', from);
      if (lt >= 0 && xml.startsWith(tSiEnd, lt)) {
        add(decodeText(xml.slice(from, lt)));
        p = xml.indexOf(si, lt + 4);
        continue;
      }
    }
    const end = xml.indexOf(siEnd, gt);
    if (end < 0) break;
    // The slice is a view, not a copy; it keeps readText from searching past this entry for a <t> that belongs to a later one.
    const body = xml.slice(gt, end);
    add(readText(body, 0, body.length, hasPhonetic, d));
    p = xml.indexOf(si, end);
  }
  return out;
}
