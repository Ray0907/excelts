import { describe, expect, it } from 'vitest';
import { mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Workbook, stream, NotSupportedError } from '../../src/compat/index.js';

describe('stream.xlsx.WorkbookWriter', () => {
  it('writes committed rows and reads back through Workbook', async () => {
    const filename = join(mkdtempSync(join(tmpdir(), 'xn-')), 's.xlsx');
    const wb = new stream.xlsx.WorkbookWriter({ filename, useSharedStrings: true, useStyles: true });
    const ws = wb.addWorksheet('Data');
    ws.columns = [{ header: 'n', key: 'n', style: { numFmt: '0.00' } }, { header: 's', key: 's' }];
    for (let i = 0; i < 1000; i++) ws.addRow({ n: i, s: 'v' + (i % 7) }).commit();
    ws.mergeCells('D1:E1');
    await ws.commit();
    await wb.commit();
    const back = await new Workbook().xlsx.readFile(filename);
    const r = back.getWorksheet('Data')!;
    expect(r.rowCount).toBe(1001);
    expect(r.getCell('A1').value).toBe('n');
    expect(r.getCell('B1001').value).toBe('v' + (999 % 7));
    expect(r.getCell('A5').numFmt).toBe('0.00');
    expect(r.getCell('E1').master.address).toBe('D1');
  });
  it('does not keep committed rows in memory', async () => {
    const wb = new stream.xlsx.WorkbookWriter({});
    const ws = wb.addWorksheet('S');
    for (let i = 0; i < 100; i++) ws.addRow([i]).commit();
    expect(ws.findRow(50)).toBeUndefined();
    expect(ws.rowCount).toBe(100);
  });
  it('the streaming reader is refused', () => {
    expect(() => new stream.xlsx.WorkbookReader()).toThrow(NotSupportedError);
  });
});
