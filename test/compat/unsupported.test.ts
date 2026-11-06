import { describe, expect, it } from 'vitest';
import { strToU8, zipSync, unzipSync, strFromU8 } from 'fflate';
import { NotSupportedError, Workbook, XlsxLimitError } from '../../src/compat/index.js';

describe('Tier 3 refuses loudly', () => {
  const ws = new Workbook().addWorksheet('S');
  const cases: [string, () => unknown][] = [
    ['addTable', () => ws.addTable()],
    ['protect', () => ws.protect()],
    ['conditional formatting', () => ws.addConditionalFormatting()],
    ['pivot', () => ws.addPivotTable()],
    ['workbook.model', () => new Workbook().model],
    ['ws.model', () => ws.model],
    ['deep style mutation', () => {
      ws.getCell('A1').font = { color: { argb: 'FF000000' } };
      (ws.getCell('A1').font!.color as { argb: string }).argb = 'x';
    }],
  ];
  for (const [name, fn] of cases) {
    it(name, () => {
      expect(fn).toThrow(NotSupportedError);
    });
  }
  it('names the feature in the message', () => {
    try {
      ws.addTable();
    } catch (e) {
      expect((e as NotSupportedError).feature).toMatch(/addTable/);
      expect((e as Error).message).toMatch(/not supported: worksheet\.addTable/);
    }
  });
});

async function withDrawing(): Promise<Uint8Array> {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const files = unzipSync(await wb.xlsx.writeBuffer());
  files['xl/charts/chart1.xml'] = strToU8('<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>');
  const sheet = strFromU8(files['xl/worksheets/sheet1.xml']!).replace('</worksheet>', '<conditionalFormatting sqref="A1"/></worksheet>');
  files['xl/worksheets/sheet1.xml'] = strToU8(sheet);
  return zipSync(files);
}

describe('round trip of unsupported content', () => {
  it('loads, but saving throws unless ignoreUnsupported is set', async () => {
    const wb = await new Workbook().xlsx.load(await withDrawing());
    expect(wb.getWorksheet('S')!.getCell('A1').value).toBe('x');
    await expect(wb.xlsx.writeBuffer()).rejects.toThrow(/charts.*conditional formatting|conditional formatting.*charts/);
    await expect(wb.xlsx.writeBuffer()).rejects.toBeInstanceOf(NotSupportedError);
    const bytes = await wb.xlsx.writeBuffer({ ignoreUnsupported: true });
    expect(bytes.length).toBeGreaterThan(100);
  });
});

describe('security limits apply to compat load', () => {
  async function pkg(): Promise<Uint8Array> {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    for (let i = 0; i < 20; i++) ws.addRow([i, 'v' + i]);
    ws.mergeCells('A1:C3');
    return wb.xlsx.writeBuffer();
  }
  it('maxRows / maxSharedStrings / maxMergedCells / maxEntries surface as XlsxLimitError', async () => {
    const bytes = await pkg();
    await expect(new Workbook().xlsx.load(bytes, { limits: { maxRows: 10 } })).rejects.toBeInstanceOf(XlsxLimitError);
    await expect(new Workbook().xlsx.load(bytes, { limits: { maxSharedStrings: 5 } })).rejects.toThrow(/maxSharedStrings/);
    await expect(new Workbook().xlsx.load(bytes, { limits: { maxMergedCells: 4 } })).rejects.toThrow(/maxMergedCells/);
    await expect(new Workbook().xlsx.load(bytes, { limits: { maxEntries: 3 } })).rejects.toThrow(/maxEntries/);
  });
  it('maxCellXfs bounds the styles part', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    for (let i = 0; i < 10; i++) ws.getCell(1, i + 1).numFmt = `0.${'0'.repeat(i + 1)}`;
    const bytes = await wb.xlsx.writeBuffer();
    await expect(new Workbook().xlsx.load(bytes, { limits: { maxCellXfs: 5 } })).rejects.toThrow(/maxCellXfs/);
  });
});
