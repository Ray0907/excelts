import { describe, expect, it } from 'vitest';
import { unzipSync, strFromU8 } from 'fflate';
import ExcelJS from 'exceljs';
import { Workbook, ValueType } from '../../src/compat/index.js';
import { readXlsx } from '../../src/index.js';

async function roundTrip(wb: Workbook): Promise<Workbook> {
  const bytes = await wb.xlsx.writeBuffer();
  return new Workbook().xlsx.load(bytes);
}

describe('compat write then load', () => {
  it('values of every type, multi-sheet', async () => {
    const wb = new Workbook();
    const a = wb.addWorksheet('First');
    const b = wb.addWorksheet('Second');
    const d = new Date(Date.UTC(2020, 1, 3, 4, 5, 6));
    a.addRow([1, 'two', true, d, { formula: 'A1*2', result: 2 }, { error: '#N/A' }, { text: 'link', hyperlink: 'http://example.com/?a=1&b=2' }]);
    b.getCell('C3').value = { richText: [{ text: 'rich ', font: { bold: true } }, { text: 'text' }] };
    const out = await roundTrip(wb);
    const r = out.getWorksheet('First')!.getRow(1);
    expect(r.getCell(1).value).toBe(1);
    expect(r.getCell(2).value).toBe('two');
    expect(r.getCell(3).value).toBe(true);
    expect((r.getCell(4).value as Date).getTime()).toBe(d.getTime());
    expect(r.getCell(5).value).toEqual({ formula: 'A1*2', result: 2 });
    expect(r.getCell(6).value).toEqual({ error: '#N/A' });
    expect(r.getCell(7).value).toEqual({ text: 'link', hyperlink: 'http://example.com/?a=1&b=2' });
    expect(out.getWorksheet('Second')!.getCell('C3').value).toEqual({ richText: [{ text: 'rich ', font: { bold: true } }, { text: 'text' }] });
    expect(out.getWorksheet('Second')!.getCell('C3').type).toBe(ValueType.RichText);
  });
  it('layout: merges, widths, heights, frozen view, properties, print setup', async () => {
    const wb = new Workbook();
    wb.creator = 'Tester';
    wb.title = 'T';
    const ws = wb.addWorksheet('L', { properties: { tabColor: { argb: 'FF00FF00' } }, pageSetup: { orientation: 'landscape', fitToWidth: 1, fitToHeight: 0 } });
    ws.columns = [{ header: 'a', key: 'a', width: 20, hidden: false }, { header: 'b', key: 'b', width: 5 }];
    ws.getRow(2).height = 33;
    ws.mergeCells('A3:B4');
    ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 1 }];
    ws.autoFilter = 'A1:B1';
    const out = await roundTrip(wb);
    const o = out.getWorksheet('L')!;
    expect(out.creator).toBe('Tester');
    expect(out.title).toBe('T');
    expect(o.getColumn(1).width).toBe(20);
    expect(o.getRow(2).height).toBe(33);
    expect(o.getCell('B4').master.address).toBe('A3');
    expect(o.views[0]).toMatchObject({ state: 'frozen', xSplit: 1, ySplit: 1 });
    expect(o.autoFilter).toBe('A1:B1');
    expect(o.properties.tabColor).toEqual({ argb: 'FF00FF00' });
    expect(o.pageSetup).toMatchObject({ orientation: 'landscape', fitToWidth: 1, fitToHeight: 0, fitToPage: true });
  });
  it('styles survive: font, fill, border, alignment, numFmt, protection', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('St');
    const style = {
      numFmt: '#,##0.00',
      font: { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFF0000' }, underline: true },
      fill: { type: 'pattern' as const, pattern: 'solid', fgColor: { argb: 'FF00FF00' } },
      border: { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'double' } },
      alignment: { horizontal: 'center', vertical: 'middle', wrapText: true, textRotation: 30 },
      protection: { locked: false },
    };
    ws.getCell('B2').value = 1234.5;
    ws.getCell('B2').style = style;
    const out = await roundTrip(wb);
    expect(out.getWorksheet('St')!.getCell('B2').style).toEqual(style);
  });
  it('lean readXlsx and exceljs both read the compat output', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('X');
    ws.addRow(['h', 1.5, new Date(Date.UTC(2021, 0, 1))]);
    ws.getCell('A1').font = { bold: true };
    const bytes = await wb.xlsx.writeBuffer();
    expect([...readXlsx(bytes).sheets[0]!.rows()][0]).toEqual(['h', 1.5, new Date(Date.UTC(2021, 0, 1))]);
    const ex = new ExcelJS.Workbook();
    await ex.xlsx.load(bytes as never);
    expect(ex.getWorksheet('X')!.getCell('A1').font!.bold).toBe(true);
    expect(ex.getWorksheet('X')!.getCell('B1').value).toBe(1.5);
  });
  it('package has the expected parts and escapes text', async () => {
    const wb = new Workbook();
    wb.addWorksheet('P&Q').getCell('A1').value = '<a & "b">';
    const files = unzipSync(await wb.xlsx.writeBuffer());
    expect(Object.keys(files)).toEqual(expect.arrayContaining(['[Content_Types].xml', 'xl/workbook.xml', 'xl/styles.xml', 'xl/sharedStrings.xml', 'xl/worksheets/sheet1.xml', 'docProps/core.xml']));
    expect(strFromU8(files['xl/workbook.xml']!)).toContain('name="P&amp;Q"');
    expect(strFromU8(files['xl/sharedStrings.xml']!)).toContain('&lt;a &amp; &quot;b&quot;&gt;');
  });
  it('1904 date system round trips', async () => {
    const wb = new Workbook();
    wb.properties.date1904 = true;
    const d = new Date(Date.UTC(2000, 0, 1));
    wb.addWorksheet('D').getCell('A1').value = d;
    const out = await roundTrip(wb);
    expect(out.properties.date1904).toBe(true);
    expect((out.getWorksheet('D')!.getCell('A1').value as Date).getTime()).toBe(d.getTime());
  });
});
