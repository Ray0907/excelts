import { describe, expect, it } from 'vitest';
import { strToU8, zipSync } from 'fflate';
import { readXlsx, XlsxLimitError, type CellValue, type XlsxLimits } from '../src/index.js';
import { fallbackCompressor } from '../src/zip.js';
import { classifyFormat } from '../src/readStyles.js';
import { readSharedStrings } from '../src/readSharedStrings.js';

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

interface Parts {
  sheets?: string[];
  sst?: string;
  styles?: string;
  workbook?: string;
  extra?: Record<string, Uint8Array>;
}

/** A minimal package; each sheet entry is the `<sheetData>` body. */
function pkg({ sheets = [''], sst, styles, workbook, extra = {} }: Parts = {}): Uint8Array {
  const files: Record<string, Uint8Array> = {
    'xl/workbook.xml': strToU8(workbook ?? `${XML}<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets>${sheets.map((_, i) => `<sheet name="S${i}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="worksheets/s${i}.xml"/>`).join('')}</Relationships>`),
    ...extra,
  };
  sheets.forEach((body, i) => (files[`xl/worksheets/s${i}.xml`] = strToU8(`${XML}<worksheet xmlns="${MAIN}"><sheetData>${body}</sheetData></worksheet>`)));
  if (sst !== undefined) files['xl/sharedStrings.xml'] = strToU8(`${XML}<sst xmlns="${MAIN}">${sst}</sst>`);
  if (styles !== undefined) files['xl/styles.xml'] = strToU8(`${XML}<styleSheet xmlns="${MAIN}">${styles}</styleSheet>`);
  return zipSync(files);
}

const readAll = (bytes: Uint8Array, limits?: Partial<XlsxLimits>, compressor?: typeof fallbackCompressor): CellValue[][][] =>
  readXlsx(bytes, { limits }, compressor).sheets.map((s) => [...s.rows()]);

function timed<T>(fn: () => T): { ms: number; value: T | undefined; error: unknown } {
  const t0 = performance.now();
  let value: T | undefined;
  let error: unknown;
  try {
    value = fn();
  } catch (e) {
    error = e;
  }
  return { ms: performance.now() - t0, value, error };
}

describe('XlsxLimitError', () => {
  it('is an Error with its own name and is exported', () => {
    const e = new XlsxLimitError('maxRows', 'too many rows');
    expect(e).toBeInstanceOf(Error);
    expect(e.name).toBe('XlsxLimitError');
    expect(e.limit).toBe('maxRows');
  });
});

describe('cell references', () => {
  it('rejects a column far beyond XFD instead of padding the row', () => {
    const bytes = pkg({ sheets: ['<row r="1"><c r="ZZZZZZZZ1"><v>1</v></c></row>'] });
    expect(() => readAll(bytes)).toThrow(XlsxLimitError);
  });

  it('rejects a megabyte run of letters in a reference', () => {
    const bytes = pkg({ sheets: [`<row r="1"><c r="${'Z'.repeat(1 << 20)}1"><v>1</v></c></row>`] });
    const r = timed(() => readAll(bytes));
    expect(r.error).toBeInstanceOf(XlsxLimitError);
    expect(r.ms).toBeLessThan(1000);
  });

  it('accepts XFD and rejects XFE (column 16384)', () => {
    expect(readAll(pkg({ sheets: ['<row r="1"><c r="XFD1"><v>1</v></c></row>'] }))[0]![0]).toHaveLength(16384);
    expect(() => readAll(pkg({ sheets: ['<row r="1"><c r="XFE1"><v>1</v></c></row>'] }))).toThrow(XlsxLimitError);
  });

  it('honours a smaller maxColumns, also for cells without a reference', () => {
    const ok = pkg({ sheets: ['<row r="1"><c r="C1"><v>1</v></c></row>'] });
    expect(readAll(ok, { maxColumns: 3 })[0]![0]).toEqual([undefined, undefined, 1]);
    expect(() => readAll(pkg({ sheets: ['<row r="1"><c r="D1"><v>1</v></c></row>'] }), { maxColumns: 3 })).toThrow(XlsxLimitError);
    expect(() => readAll(pkg({ sheets: ['<row><c><v>1</v></c><c><v>1</v></c><c><v>1</v></c><c><v>1</v></c></row>'] }), { maxColumns: 3 })).toThrow(XlsxLimitError);
  });
});

describe('row numbers', () => {
  it.each(['4294967296', '9007199254740993', '9999999999999999999', '1048577'])('rejects <row r="%s"/> quickly', (r) => {
    const out = timed(() => readAll(pkg({ sheets: [`<row r="${r}"/>`] })));
    expect(out.error).toBeInstanceOf(XlsxLimitError);
    expect(out.ms).toBeLessThan(1000);
  });

  it('accepts the last Excel row number', () => {
    const out = readAll(pkg({ sheets: ['<row r="1048576"><c r="A1048576"><v>7</v></c></row>'] }))[0]!;
    expect(out).toHaveLength(1048576);
    expect(out[1048575]).toEqual([7]);
  });

  it('honours a smaller maxRows for numbered and unnumbered rows', () => {
    expect(readAll(pkg({ sheets: ['<row r="10"/>'] }), { maxRows: 10 })[0]).toHaveLength(10);
    expect(() => readAll(pkg({ sheets: ['<row r="11"/>'] }), { maxRows: 10 })).toThrow(XlsxLimitError);
    expect(() => readAll(pkg({ sheets: ['<row/>'.repeat(11)] }), { maxRows: 10 })).toThrow(XlsxLimitError);
  });
});

describe('row gaps', () => {
  const sheet = '<row r="10"><c r="A10"><v>1</v></c></row>';
  it('are budgeted over the whole workbook', () => {
    const bytes = pkg({ sheets: [sheet, sheet, sheet] });
    expect(() => readAll(bytes, { maxGapRows: 20 })).toThrow(/maxGapRows/);
    expect(readAll(bytes, { maxGapRows: 27 })).toHaveLength(3);
  });

  it('count a sheet once however often it is read', () => {
    const wb = readXlsx(pkg({ sheets: [sheet] }), { limits: { maxGapRows: 9 } });
    for (let i = 0; i < 4; i++) expect([...wb.sheets[0]!.rows()]).toHaveLength(10);
  });

  it('do not count rows that are present', () => {
    expect(readAll(pkg({ sheets: ['<row r="1"/><row r="2"/><row r="3"/>'] }), { maxGapRows: 0 })[0]).toHaveLength(3);
  });
});

describe('scans stay inside their cell and row', () => {
  const N = 60000;
  const BUDGET = 1000;
  const run = (body: string) => timed(() => readAll(pkg({ sheets: [body] })));

  it('empty cells with a style and no <v> anywhere', () => {
    const r = run(`<row r="1">${'<c s="1"></c>'.repeat(N)}</row>`);
    expect(r.ms).toBeLessThan(BUDGET);
  });

  it('inlineStr cells without <is>', () => {
    const r = run(`<row r="1">${'<c t="inlineStr"></c>'.repeat(N)}</row>`);
    expect(r.ms).toBeLessThan(BUDGET);
  });

  it('a cell followed by an unknown element, repeated in one row', () => {
    const r = run(`<row r="1">${'<c/><x/>'.repeat(N)}</row>`);
    expect(r.ms).toBeLessThan(BUDGET);
  });

  it('<v> that never closes', () => {
    const r = run(`<row r="1">${'<c><v>x</c>'.repeat(N)}</row>`);
    expect(r.ms).toBeLessThan(BUDGET);
  });

  it('<v> that never closes, with a </v> only at the very end', () => {
    const r = run(`<row r="1">${'<c><v>x</c>'.repeat(N)}<c><v>1</v></c></row>`);
    expect(r.ms).toBeLessThan(BUDGET);
  });

  it('rows without </row>', () => {
    const r = run('<row><c><v>1</v></c>'.repeat(N));
    expect(r.ms).toBeLessThan(BUDGET);
  });

  it('cells without </c>', () => {
    const r = run(`<row r="1">${'<c><v>1</v>'.repeat(N)}</row>`);
    expect(r.ms).toBeLessThan(BUDGET);
  });

  it('shared-string entries that hold no <t> while a later one does', () => {
    const t0 = performance.now();
    const out = readSharedStrings(`<sst>${'<si><r/></si>'.repeat(N)}<si><t>x</t></si></sst>`);
    expect(out).toHaveLength(N + 1);
    expect(out[N]).toBe('x');
    expect(performance.now() - t0).toBeLessThan(BUDGET);
  });

  it('still reads the shapes it bounds', () => {
    const out = readAll(pkg({ sheets: ['<row r="1"><c r="A1" s="1"></c><c r="B1" t="inlineStr"><is><t>hi</t></is></c><c r="C1"><f>1+1</f><v>2</v></c><x/><c r="D1" t="str"><v>z</v></c></row>'] }))[0]![0]!;
    expect(out).toEqual([undefined, 'hi', 2, 'z']);
  });
});

describe('regexes that backtracked', () => {
  it('a long run of name characters in a <sheet> tag', () => {
    const workbook = `${XML}<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet ${'a'.repeat(60000)} name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`;
    const r = timed(() => readAll(pkg({ workbook, sheets: [''] })));
    expect(r.ms).toBeLessThan(500);
    expect(r.error).toBeUndefined();
  });

  it('keeps the first of two equal attributes, with either quote', () => {
    const workbook = `${XML}<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name='A' name="B" sheetId = "1" r:id='rId1'/></sheets></workbook>`;
    expect(readXlsx(pkg({ workbook, sheets: [''] })).sheets.map((s) => s.name)).toEqual(['A']);
  });

  it('an unterminated bracket run in a number format', () => {
    const r = timed(() => classifyFormat('['.repeat(80000)));
    expect(r.ms).toBeLessThan(300);
  });

  it('classifies bracketed colours, conditions and locales as before', () => {
    expect(classifyFormat('[Red]yyyy')).toBe(1);
    expect(classifyFormat('[$-409]h:mm')).toBe(2);
    expect(classifyFormat('[>100]0.00')).toBe(0);
    expect(classifyFormat('[h]:mm:ss')).toBe(2);
  });

  it('an unclosed <rPh> run in a shared string', () => {
    const t0 = performance.now();
    const out = readSharedStrings(`<sst><si><t>a</t>${'<rPh sb="0">'.repeat(80000)}</si></sst>`);
    expect(out).toHaveLength(1);
    expect(performance.now() - t0).toBeLessThan(300);
  });

  it('still drops closed phonetic runs', () => {
    expect(readSharedStrings('<sst><si><r><t>a</t></r><rPh sb="0" eb="1"><t>ignored</t></rPh><r><t>b</t></r></si></sst>')).toEqual(['ab']);
  });
});

describe('inflate limits', () => {
  const spaces = (n: number) => `<row r="1"><c r="A1" t="inlineStr"><is><t>${' '.repeat(n)}</t></is></c></row>`;

  it.each([
    ['native', undefined],
    ['fallback', fallbackCompressor],
  ] as const)('caps one entry (%s inflater)', (_, comp) => {
    const bytes = pkg({ sheets: [spaces(3_000_000)] });
    expect(() => readAll(bytes, { maxEntryBytes: 1_000_000 }, comp)).toThrow(XlsxLimitError);
    expect(() => readAll(bytes, { maxEntryBytes: 1_000_000 }, comp)).toThrow(/maxEntryBytes/);
    expect(readAll(bytes, { maxEntryBytes: 4_000_000, maxCellChars: 4_000_000, maxRatio: 1e9 }, comp)[0]).toHaveLength(1);
  });

  it('turns the engine overflow into the dedicated error, not a RangeError', () => {
    const bytes = pkg({ sheets: [spaces(3_000_000)] });
    let caught: unknown;
    try {
      readAll(bytes, { maxEntryBytes: 1000 });
    } catch (e) {
      caught = e;
    }
    expect(caught).toBeInstanceOf(XlsxLimitError);
    expect(caught).not.toBeInstanceOf(RangeError);
  });

  it('caps the total inflated across entries', () => {
    const bytes = pkg({ sheets: [spaces(400_000), spaces(400_000), spaces(400_000)] });
    const wb = readXlsx(bytes, { limits: { maxTotalBytes: 1_000_000, maxCellChars: 1e9 } });
    expect([...wb.sheets[0]!.rows()]).toHaveLength(1);
    expect([...wb.sheets[1]!.rows()]).toHaveLength(1);
    expect(() => [...wb.sheets[2]!.rows()]).toThrow(/maxTotalBytes/);
  });

  it('reading one sheet again does not count twice', () => {
    const bytes = pkg({ sheets: [spaces(400_000)] });
    const wb = readXlsx(bytes, { limits: { maxTotalBytes: 1_000_000, maxCellChars: 1e9 } });
    for (let i = 0; i < 5; i++) expect([...wb.sheets[0]!.rows()]).toHaveLength(1);
  });

  it('caps the ratio for entries over the minimum size', () => {
    const bytes = pkg({ sheets: [spaces(200_000)] });
    expect(() => readAll(bytes, { maxRatio: 20, ratioMinBytes: 10_000, maxCellChars: 1e9 })).toThrow(/maxRatio/);
    expect(readAll(bytes, { maxRatio: 20, ratioMinBytes: 1_000_000, maxCellChars: 1e9 })[0]).toHaveLength(1);
  });

  it('rejects a declared size beyond the cap without inflating, and still caps when the declaration lies low', () => {
    const bytes = pkg({ sheets: [spaces(300_000)] }).slice();
    const dv = new DataView(bytes.buffer);
    // find the sheet's central directory record and rewrite its uncompressed size
    const sig = strToU8('xl/worksheets/s0.xml');
    let at = -1;
    for (let i = bytes.length - 22; i >= 0 && at < 0; i--) {
      if (dv.getUint32(i, true) === 0x02014b50 && sig.every((b, k) => bytes[i + 46 + k] === b)) at = i;
    }
    expect(at).toBeGreaterThan(0);
    dv.setUint32(at + 24, 0xfffffffe, true);
    expect(() => readAll(bytes, { maxEntryBytes: 1_000_000 })).toThrow(/maxEntryBytes/);
    dv.setUint32(at + 24, 10, true);
    expect(() => readAll(bytes, { maxEntryBytes: 100_000 })).toThrow(XlsxLimitError);
  });

  it('caps the entry count and the sheet count', () => {
    const extra: Record<string, Uint8Array> = {};
    for (let i = 0; i < 20; i++) extra[`junk/${i}.bin`] = strToU8('x');
    expect(() => readXlsx(pkg({ extra }), { limits: { maxEntries: 10 } })).toThrow(/maxEntries/);
    expect(() => readXlsx(pkg({ sheets: ['', '', ''] }), { limits: { maxSheets: 2 } })).toThrow(/maxSheets/);
    expect(readXlsx(pkg({ sheets: ['', ''] }), { limits: { maxSheets: 2 } }).sheets).toHaveLength(2);
  });

  it('defaults are the documented ones', () => {
    const big = pkg({ sheets: [''] });
    expect(() => readXlsx(big)).not.toThrow();
  });
});

describe('allocation limits', () => {
  it('shared string count', () => {
    const sst = '<si><t>a</t></si>'.repeat(4);
    const bytes = pkg({ sheets: ['<row r="1"><c t="s"><v>0</v></c></row>'], sst });
    expect(() => readAll(bytes, { maxSharedStrings: 3 })).toThrow(/maxSharedStrings/);
    expect(readAll(bytes, { maxSharedStrings: 4 })[0]![0]).toEqual(['a']);
  });

  it('shared string text size', () => {
    const sst = '<si><t>aaaa</t></si>'.repeat(3);
    const bytes = pkg({ sheets: ['<row r="1"><c t="s"><v>0</v></c></row>'], sst });
    expect(() => readAll(bytes, { maxSharedStringBytes: 11 })).toThrow(/maxSharedStringBytes/);
    expect(readAll(bytes, { maxSharedStringBytes: 12 })[0]).toHaveLength(1);
  });

  it('cellXfs count', () => {
    const styles = `<cellXfs count="5">${'<xf numFmtId="0"/>'.repeat(5)}</cellXfs>`;
    const bytes = pkg({ sheets: ['<row r="1"><c s="1"><v>1</v></c></row>'], styles });
    expect(() => readAll(bytes, { maxCellXfs: 4 })).toThrow(/maxCellXfs/);
    expect(readAll(bytes, { maxCellXfs: 5 })[0]![0]).toEqual([1]);
  });

  it('cell text length, for inline, formula-string and shared strings', () => {
    const cells = '<row r="1"><c t="inlineStr"><is><t>123456</t></is></c></row>';
    expect(() => readAll(pkg({ sheets: [cells] }), { maxCellChars: 5 })).toThrow(/maxCellChars/);
    expect(readAll(pkg({ sheets: [cells] }), { maxCellChars: 6 })[0]![0]).toEqual(['123456']);
    expect(() => readAll(pkg({ sheets: ['<row r="1"><c t="str"><f>x</f><v>123456</v></c></row>'] }), { maxCellChars: 5 })).toThrow(/maxCellChars/);
    expect(() => readAll(pkg({ sheets: ['<row r="1"><c t="s"><v>0</v></c></row>'], sst: '<si><t>123456</t></si>' }), { maxCellChars: 5 })).toThrow(/maxCellChars/);
  });

  it('default cell text limit is 32767', () => {
    const body = (n: number) => `<row r="1"><c t="inlineStr"><is><t>${'a'.repeat(n)}</t></is></c></row>`;
    expect(readAll(pkg({ sheets: [body(32767)] }))[0]![0]![0]).toHaveLength(32767);
    expect(() => readAll(pkg({ sheets: [body(32768)] }))).toThrow(XlsxLimitError);
  });
});
