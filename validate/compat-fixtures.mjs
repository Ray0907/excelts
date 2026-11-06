// Workbooks built through the exceljs-shaped facade, covering styles, merges, multi-sheet, views, print setup, formulas and links.
import { Workbook } from '../dist/compat/index.js';

export function compatFixtures() {
  const out = [];

  {
    const wb = new Workbook();
    wb.creator = 'Tester';
    wb.title = 'Styles & more';
    const ws = wb.addWorksheet('Styles', { properties: { tabColor: { argb: 'FF00FF00' }, outlineLevelRow: 1, outlineLevelCol: 1 } });
    ws.columns = [
      { header: 'Name', key: 'name', width: 24 },
      { header: 'Amount', key: 'amount', width: 12, style: { numFmt: '#,##0.00' } },
      { header: 'When', key: 'when', width: 20, outlineLevel: 1 },
    ];
    ws.addRow({ name: 'a < b & c', amount: 1234.5, when: new Date(Date.UTC(2021, 5, 7, 8, 9, 10)) });
    ws.addRow({ name: 'second', amount: -2, when: new Date(Date.UTC(1999, 11, 31)) });
    ws.getRow(1).font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFFFF' }, underline: 'double' };
    ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF336699' } };
    ws.getRow(1).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true, textRotation: 15 };
    ws.getRow(1).height = 30;
    ws.getRow(2).outlineLevel = 1;
    ws.getCell('B2').border = { top: { style: 'thin', color: { argb: 'FFFF0000' } }, bottom: { style: 'double' }, diagonal: { style: 'thin', up: true, down: true } };
    ws.getCell('C2').protection = { locked: false, hidden: true };
    ws.getCell('A4').fill = { type: 'gradient', gradient: 'angle', degree: 45, stops: [{ position: 0, color: { argb: 'FF0000FF' } }, { position: 1, color: { argb: 'FFFFFFFF' } }] };
    ws.getCell('A5').fill = { type: 'gradient', gradient: 'path', center: { left: 0.5, top: 0.5 }, stops: [{ position: 0, color: { argb: 'FFFF0000' } }, { position: 1, color: { argb: 'FF0000FF' } }] };
    ws.getCell('A6').value = 'merged';
    ws.mergeCells('A6:C7');
    ws.getCell('A6').alignment = { horizontal: 'center' };
    ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 1, activeCell: 'B2' }];
    ws.autoFilter = 'A1:C1';
    out.push({ name: 'compat-styles', wb });
  }

  {
    const wb = new Workbook();
    const a = wb.addWorksheet('Formulas', { pageSetup: { orientation: 'landscape', fitToWidth: 1, fitToHeight: 0, paperSize: 9, printArea: 'A1:D10', printTitlesRow: '1:1' }, headerFooter: { oddHeader: '&CHeader', oddFooter: '&P of &N' } });
    a.getCell('A1').value = 1;
    a.getCell('B1').value = { formula: 'A1+1', result: 2 };
    a.getCell('C1').value = { sharedFormula: 'B1', result: 3 };
    a.getCell('D1').value = { sharedFormula: 'B1', result: 4 };
    a.getCell('A2').value = { formula: 'IF(A1>0,"yes","no")', result: 'yes' };
    a.getCell('A3').value = { error: '#N/A' };
    a.getCell('A4').value = true;
    a.getCell('A5').value = { text: 'site', hyperlink: 'https://example.com/?q=1&r=2', tooltip: 'go' };
    a.getCell('A6').value = { text: 'inside', hyperlink: '#Other!A1' };
    a.getCell('A7').value = { richText: [{ text: 'bold ', font: { bold: true, size: 14, color: { argb: 'FFFF0000' } } }, { text: 'plain' }] };
    a.getRow(3).addPageBreak();
    a.dataValidations.add('E1', { type: 'list', allowBlank: true, formulae: ['"x,y,z"'], showErrorMessage: true, errorTitle: 'Bad', error: 'pick one' });
    const b = wb.addWorksheet('Other', { state: 'hidden' });
    b.addRow([1, 2, 3]);
    wb.addWorksheet('Third').addRow(['t']);
    wb.views = [{ x: 0, y: 0, width: 10000, height: 20000, firstSheet: 0, activeTab: 0, visibility: 'visible' }];
    out.push({ name: 'compat-formulas-multisheet', wb });
  }

  {
    const wb = new Workbook();
    wb.properties.date1904 = true;
    const ws = wb.addWorksheet('Dates');
    ws.addRow([new Date(Date.UTC(2000, 0, 1)), 'x']);
    ws.getColumn(1).width = 14;
    out.push({ name: 'compat-1904', wb });
  }
  {
    // pictures (two-cell, one-cell, linked), a background picture and notes, on two sheets so part numbering is exercised
    const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
    const wb = new Workbook();
    const pic = wb.addImage({ base64: `data:image/png;base64,${PNG}`, extension: 'png' });
    const jpg = wb.addImage({ buffer: Uint8Array.from(atob(PNG), (c) => c.charCodeAt(0)), extension: 'JPEG' });
    const a = wb.addWorksheet('Pictures');
    a.getCell('A1').value = 'with pictures';
    a.addImage(pic, 'C3:E6');
    a.addImage(jpg, { tl: { col: 0.25, row: 0.5 }, ext: { width: 120, height: 80 }, editAs: 'absolute', hyperlinks: { hyperlink: 'https://example.com/?a=1&b=2', tooltip: 'go' } });
    a.addBackgroundImage(pic);
    a.getCell('B2').note = 'plain note';
    const b = wb.addWorksheet('Notes');
    b.getCell('A1').value = 1;
    b.getCell('A1').note = { texts: [{ text: 'bold ', font: { bold: true, size: 11, name: 'Calibri' } }, { text: 'rest & <more>' }], margins: { insetmode: 'custom', inset: [0.2, 0.2, 0.3, 0.3] }, protection: { locked: 'False', lockText: 'False' }, editAs: 'oneCells' };
    b.getCell('D9').note = 'second';
    b.addImage(pic, { tl: { col: 3, row: 3 }, br: { col: 5, row: 6 }, editAs: 'twoCell' });
    out.push({ name: 'compat-pictures-notes', wb });
  }
  return out;
}
