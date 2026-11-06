import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { Workbook } from '../../src/compat/index.js';

async function roundTrip(wb: Workbook): Promise<Workbook> {
  return new Workbook().xlsx.load(await wb.xlsx.writeBuffer());
}

describe('defined names', () => {
  it('cells of a name join into rectangles; names survive a round trip', async () => {
    const wb = new Workbook();
    const a = wb.addWorksheet('blort');
    const b = wb.addWorksheet('My Sheet');
    for (const ref of ['C5', 'B6', 'C6', 'B5']) a.getCell(ref).name = 'squares';
    a.getCell('A1').name = 'five';
    a.getCell('B7').name = 'sheets';
    b.getCell('B7').name = 'sheets';
    a.getCell('G1').names = ['thing1', 'thing2'];
    a.getCell('G2').names = ['once', 'twice'];
    a.getCell('G2').removeName('once');
    const out = await roundTrip(wb);
    const names = out.definedNames;
    expect(names.getRanges('squares').ranges).toEqual(['blort!$B$5:$C$6']);
    expect(names.getRanges('five').ranges).toEqual(['blort!$A$1']);
    expect(names.getRanges('sheets').ranges.sort()).toEqual(["'My Sheet'!$B$7", 'blort!$B$7']);
    expect(out.getWorksheet('blort')!.getCell('G1').names.sort()).toEqual(['thing1', 'thing2']);
    expect(out.getWorksheet('blort')!.getCell('G2').names).toEqual(['twice']);
    expect(out.getWorksheet('blort')!.getCell('A2').name).toBeUndefined();
    expect(names.getRanges('once').ranges).toEqual([]);
  });

  it('names follow their cells through spliceRows and spliceColumns', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    for (const ref of ['A1', 'A2', 'A3', 'A4']) ws.getCell(ref).name = 'col';
    ws.getCell('B2').name = 'one';
    ws.spliceRows(2, 1);
    expect(wb.definedNames.getRanges('col').ranges).toEqual(['S!$A$1:$A$3']);
    expect(wb.definedNames.getRanges('one').ranges).toEqual([]);
    ws.insertRow(1, ['x']);
    expect(wb.definedNames.getRanges('col').ranges).toEqual(['S!$A$2:$A$4']);
    ws.spliceColumns(1, 0, ['new']);
    expect(wb.definedNames.getRanges('col').ranges).toEqual(['S!$B$2:$B$4']);
  });

  it('a large range is kept as a rectangle, not expanded cell by cell, and still shifts', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    wb.definedNames.add('S!$A$1:$XFD$1048576', 'everything');
    wb.definedNames.add('S!$B$10:$D$5000', 'block');
    expect(wb.definedNames.getRanges('everything').ranges).toEqual(['S!$A$1:$XFD$1048576']);
    expect(ws.getCell('C100').names.sort()).toEqual(['block', 'everything']);
    ws.spliceRows(5, 2);
    expect(wb.definedNames.getRanges('block').ranges).toEqual(['S!$B$8:$D$4998']);
    const out = await roundTrip(wb);
    expect(out.definedNames.getRanges('block').ranges).toEqual(['S!$B$8:$D$4998']);
  });

  it('a hostile file with a huge name range loads in bounded time and memory', async () => {
    const wb = new Workbook();
    wb.addWorksheet('S').getCell('A1').value = 1;
    const files = unzipSync(await wb.xlsx.writeBuffer());
    const xml = strFromU8(files['xl/workbook.xml']!).replace('<calcPr', `<definedNames>${'<definedName name="n$">S!$A$1:$XFD$1048576</definedName>'.replace('n$', 'big')}</definedNames><calcPr`);
    files['xl/workbook.xml'] = strToU8(xml);
    const out = await new Workbook().xlsx.load(zipSync(files));
    expect(out.definedNames.getRanges('big').ranges).toEqual(['S!$A$1:$XFD$1048576']);
  });

  it('names that are not plain cell ranges are written back as they were', async () => {
    const wb = new Workbook();
    const a = wb.addWorksheet('A');
    wb.addWorksheet('B');
    a.getCell('A1').value = 1;
    const files = unzipSync(await wb.xlsx.writeBuffer());
    const defs =
      '<definedNames><definedName name="rate">0.07</definedName><definedName name="whole_col">A!$A:$A</definedName>' +
      '<definedName name="local" localSheetId="1">B!$A$1</definedName><definedName name="hid" hidden="1">A!$A$1</definedName>' +
      '<definedName name="plain">A!$A$1:$B$2,B!$C$3</definedName><definedName name="f">SUM(A!$A$1:$A$3)</definedName></definedNames>';
    files['xl/workbook.xml'] = strToU8(strFromU8(files['xl/workbook.xml']!).replace('<calcPr', `${defs}<calcPr`));
    const loaded = await new Workbook().xlsx.load(zipSync(files));
    // nothing is lost, so saving needs no ignoreUnsupported
    const out = unzipSync(await loaded.xlsx.writeBuffer());
    const xml = strFromU8(out['xl/workbook.xml']!);
    for (const s of ['name="rate">0.07<', 'name="whole_col">A!$A:$A<', 'name="local" localSheetId="1">B!$A$1<', 'name="hid" hidden="1">A!$A$1<', 'name="f">SUM(A!$A$1:$A$3)<']) expect(xml).toContain(s);
    expect(loaded.definedNames.getRanges('plain').ranges.sort()).toEqual(['A!$A$1:$B$2', 'B!$C$3']);
    expect(loaded.getWorksheet('B')!.getCell('C3').names).toEqual(['plain']);
    expect(loaded.getWorksheet('A')!.getCell('A1').names.sort()).toEqual(['plain']);
  });

  it('a sheet-scoped name is dropped with its sheet, and keeps its sheet when sheets are reordered', async () => {
    const wb = new Workbook();
    wb.addWorksheet('A');
    wb.addWorksheet('B');
    const files = unzipSync(await wb.xlsx.writeBuffer());
    files['xl/workbook.xml'] = strToU8(strFromU8(files['xl/workbook.xml']!).replace('<calcPr', '<definedNames><definedName name="local" localSheetId="1">B!$A$1</definedName></definedNames><calcPr'));
    const loaded = await new Workbook().xlsx.load(zipSync(files));
    loaded.removeWorksheet('A');
    expect(strFromU8(unzipSync(await loaded.xlsx.writeBuffer())['xl/workbook.xml']!)).toContain('name="local" localSheetId="0"');
    loaded.removeWorksheet('B');
    expect(strFromU8(unzipSync(await loaded.xlsx.writeBuffer())['xl/workbook.xml']!)).not.toContain('local');
  });

  it('the name model can be read and set', () => {
    const wb = new Workbook();
    wb.definedNames.model = [{ name: 'x', ranges: ['S!$A$1:$B$1'] }, { name: 'skip', ranges: ['not a range'] }];
    expect(wb.definedNames.model).toEqual([{ name: 'x', ranges: ['S!$A$1:$B$1'] }]);
  });
});
